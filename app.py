"""
Converteme — servidor de conversão (Google Cloud Run)

O que faz:
  POST /converter?para=docx   PDF  -> Word (.docx), com pdf2docx (MIT) + PyMuPDF (AGPL).
                              Relatório em colunas (ou se a detalhada travar ou falhar): conversão
                              "no lugar" (layout.py), cada linha na mesma posição do PDF.
                              O cabeçalho X-Modo diz qual foi: completo, posicao ou simples.
  POST /converter?para=pdf    Word, Excel, PowerPoint, OpenDocument, RTF -> PDF, com LibreOffice (MPL)
  GET  /saude                 "ok" (para o Cloud Run saber que está de pé)

Regras de privacidade (LGPD):
  - o arquivo chega no corpo da requisição, vai para uma pasta temporária só
    desta conversão e a pasta é apagada assim que a resposta sai (com erro ou sem);
  - nada do conteúdo nem o nome do arquivo vai para o registro (log): só o tipo,
    o tamanho arredondado e o tempo;
  - só aceita pedidos com a chave combinada (variável CHAVE), que fica nas
    configurações do Cloud Run e do Worker da Cloudflare, nunca no código.

Licença: AGPL-3.0 (veja LICENSE), por usar o PyMuPDF.
"""
import hmac
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

PORTA = int(os.environ.get("PORT", "8080"))
CHAVE = os.environ.get("CHAVE", "")
MAX_MB = float(os.environ.get("MAX_MB", "30"))   # o Cloud Run aceita até 32 MiB por pedido
MAX_PAGINAS = int(os.environ.get("MAX_PAGINAS", "300"))
# a resposta tem de sair em menos de ~100 s (limite de espera da Cloudflare):
# PDF -> Word: até 75 s na detalhada + poucos segundos na "no lugar"; Office -> PDF: até 85 s
PRAZO = int(os.environ.get("PRAZO_SEGUNDOS", "85"))
PRAZO_DOCX = int(os.environ.get("PRAZO_DOCX_SEGUNDOS", "75"))      # conversão detalhada de PDF para Word
MEMORIA_DOCX_MB = int(os.environ.get("MEMORIA_DOCX_MB", "900"))    # teto de memória de cada processo da detalhada (0 = sem teto)
# processos da conversão detalhada (no Cloud Run com 2 CPUs, 2 processos levam quase metade do tempo)
PROCESSOS_DOCX = max(1, min(int(os.environ.get("PROCESSOS_DOCX", "2")), len(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else 1))
SOFFICE = os.environ.get("SOFFICE", shutil.which("soffice") or "soffice")

TIPOS_SAIDA = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}
# uma conversão do LibreOffice por vez em cada cópia do servidor (ele não gosta de concorrência)
TRAVA_LO = threading.Lock()


class Problema(Exception):
    def __init__(self, status, codigo, mensagem):
        super().__init__(mensagem)
        self.status, self.codigo, self.mensagem = status, codigo, mensagem


def tipo_de(dados: bytes) -> str:
    """Descobre o tipo pelo conteúdo (não pelo nome)."""
    if dados[:5] == b"%PDF-" or b"%PDF-" in dados[:1024]:
        return "pdf"
    if dados[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "ole"          # .doc, .xls, .ppt antigos
    if dados[:2] == b"PK":
        inicio = dados[:4000]
        if b"word/" in inicio:
            return "docx"
        if b"xl/" in inicio:
            return "xlsx"
        if b"ppt/" in inicio:
            return "pptx"
        if b"mimetypeapplication/vnd.oasis.opendocument" in inicio:
            return "odf"
        return "zip-office"   # o LibreOffice tenta mesmo assim
    if dados[:5] == b"{\\rtf":
        return "rtf"
    return "desconhecido"


def _pdf2docx_filho(entrada: str, saida: str) -> None:
    """Roda num processo separado (veja pdf_para_docx): a conversão detalhada.
    Sem a busca de "tabelas sem linhas" (parse_stream_table): ela compara cada pedaço de texto
    com todos os outros e, num relatório com centenas de pedaços por página, levava ~55 s por
    página. As tabelas com bordas continuam sendo reconhecidas."""
    import logging
    logging.disable(logging.CRITICAL)      # o pdf2docx escreve o andamento no log; não precisamos
    from pdf2docx import Converter
    cv = Converter(entrada)
    try:
        varios = PROCESSOS_DOCX > 1 and len(cv.fitz_doc) >= 4
        cv.convert(saida, parse_stream_table=False, multi_processing=varios, cpu_count=PROCESSOS_DOCX)
    finally:
        cv.close()


def parece_relatorio(entrada: str) -> bool:
    """Relatório de sistema (colunas alinhadas, sem grade de tabela): vai direto para a conversão
    "no lugar", que fica igual ao PDF e sai em segundos. Olha só as 6 primeiras páginas."""
    import pymupdf
    import layout
    linhas = colunas = 0
    with pymupdf.open(entrada) as d:
        for pagina in list(d)[:6]:
            verticais = 0
            for desenho in pagina.get_drawings():
                for it in desenho["items"]:
                    if it[0] == "l" and abs(it[1].x - it[2].x) < 1 and abs(it[1].y - it[2].y) > 5:
                        verticais += 1
                    elif it[0] == "re" and it[1].width < 2 and it[1].height > 5:
                        verticais += 1
            if verticais >= 4:       # tem tabela com grade: a detalhada faz tabela de verdade no Word
                continue
            for ln in layout._linhas_da_pagina(pagina):
                linhas += 1
                it = ln["itens"]
                if sum(1 for a, b in zip(it, it[1:]) if b["x0"] - a["x1"] > 8) >= 2:   # 3 ou mais colunas
                    colunas += 1
    return colunas >= 15 and colunas / max(1, linhas) >= 0.3


def _limitar_memoria():
    if MEMORIA_DOCX_MB > 0:
        import resource
        teto = MEMORIA_DOCX_MB * 1024 * 1024
        resource.setrlimit(resource.RLIMIT_AS, (teto, teto))


def docx_simples(entrada: str, saida: str) -> None:
    """Plano B: texto corrido e imagens na ordem de leitura (sem tabelas nem colunas).
    Rápido e à prova de PDFs complicados; mantém tamanho da letra, negrito, itálico e cor."""
    import io
    import pymupdf
    from docx import Document
    from docx.shared import Pt, RGBColor, Emu
    doc = Document()
    with pymupdf.open(entrada) as pdf:
        for n, pagina in enumerate(pdf):
            if n == 0:
                sec = doc.sections[0]
                sec.page_width, sec.page_height = Emu(int(pagina.rect.width * 12700)), Emu(int(pagina.rect.height * 12700))
                for lado in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
                    setattr(sec, lado, Pt(50))
            else:
                doc.add_page_break()
            largura_util = pagina.rect.width - 100
            blocos = pagina.get_text("dict", sort=True).get("blocks", [])
            for b in blocos:
                if b.get("type") == 1:      # imagem
                    try:
                        w = min(largura_util, b["bbox"][2] - b["bbox"][0])
                        if w > 20 and b.get("image"):
                            doc.add_picture(io.BytesIO(b["image"]), width=Pt(w))
                    except Exception:  # noqa: BLE001 — imagem em formato que o Word não aceita: segue sem ela
                        pass
                    continue
                par = None
                for linha in b.get("lines", []):
                    trechos = [t for t in linha.get("spans", []) if t.get("text")]
                    if not "".join(t["text"] for t in trechos).strip():
                        continue
                    if par is None:
                        par = doc.add_paragraph()
                        par.paragraph_format.space_after = Pt(6)
                    elif par.runs and not par.runs[-1].text.endswith((" ", "-")):
                        par.add_run(" ")
                    for t in trechos:
                        r = par.add_run(t["text"])
                        r.font.size = Pt(max(5, min(72, round(t.get("size", 11) * 2) / 2)))
                        f = t.get("flags", 0)
                        r.bold, r.italic = bool(f & 16), bool(f & 2)
                        c = t.get("color", 0)
                        if c:
                            r.font.color.rgb = RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)
    doc.save(saida)


def _no_lugar(entrada: str, saida: str) -> str:
    """Conversão "no lugar" (layout.py); se até ela falhar, a de texto corrido. Devolve o modo."""
    import layout
    try:
        layout.converter(entrada, saida)
        return "posicao"
    except Exception as e:  # noqa: BLE001
        registrar(evento="posicao-falhou", classe=type(e).__name__)
    if os.path.exists(saida):
        os.remove(saida)
    try:
        docx_simples(entrada, saida)
        return "simples"
    except Exception as e:  # noqa: BLE001
        registrar(evento="simples-falhou", classe=type(e).__name__)
        raise Problema(422, "nao-converteu", "Não consegui converter este PDF em Word. Ele pode estar danificado. Tente abrir e salvar de novo como PDF, ou dividir em partes no Dividir PDF.")


def _rodar_com_prazo(cmd, prazo):
    """Roda e, se passar do prazo, mata o processo e os que ele abriu (os da conversão em paralelo)."""
    import signal
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True, preexec_fn=_limitar_memoria)
    try:
        return proc.wait(timeout=prazo)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(proc.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        proc.wait()
        raise


def pdf_para_docx(pasta: str, entrada: str):
    """Devolve (caminho, modo). modo = "completo" (pdf2docx), "posicao" (layout.py: cada linha
    no mesmo lugar) ou "simples" (texto corrido, último recurso)."""
    import pymupdf as fitz  # PyMuPDF
    with fitz.open(entrada) as d:
        if d.needs_pass:
            raise Problema(422, "senha", "Este PDF tem senha. Tire a senha primeiro e tente de novo.")
        if d.page_count > MAX_PAGINAS:
            raise Problema(413, "paginas", f"Este PDF tem {d.page_count} páginas. O limite é {MAX_PAGINAS}.")
        if d.page_count == 0:
            raise Problema(422, "vazio", "Este PDF não tem páginas.")
    saida = os.path.join(pasta, "saida.docx")
    try:
        relatorio = parece_relatorio(entrada)
    except Exception:  # noqa: BLE001 — na dúvida, segue o caminho normal
        relatorio = False
    if relatorio:
        registrar(evento="relatorio")
        return saida, _no_lugar(entrada, saida)
    # a conversão detalhada roda num processo separado, com prazo e teto de memória:
    # se travar, estourar a memória ou der erro, o servidor continua de pé e usa o plano B
    motivo = ""
    try:
        rc = _rodar_com_prazo([sys.executable, os.path.abspath(__file__), "--pdf2docx", entrada, saida], PRAZO_DOCX)
        if rc == 0 and os.path.exists(saida) and os.path.getsize(saida) > 0:
            return saida, "completo"
        motivo = "erro"
    except subprocess.TimeoutExpired:
        motivo = "prazo"
    registrar(evento="plano-b", motivo=motivo)
    if os.path.exists(saida):
        os.remove(saida)
    return saida, _no_lugar(entrada, saida)


def office_para_pdf(pasta: str, entrada: str) -> str:
    perfil = os.path.join(pasta, "perfil-lo")
    cmd = [SOFFICE, f"-env:UserInstallation=file://{perfil}", "--headless", "--norestore", "--nolockcheck",
           "--convert-to", "pdf", "--outdir", pasta, entrada]
    with TRAVA_LO:
        try:
            r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=PRAZO, env={**os.environ, "HOME": pasta})
        except subprocess.TimeoutExpired:
            raise Problema(504, "demorou", "A conversão demorou demais. Tente um arquivo menor.")
    saida = os.path.splitext(entrada)[0] + ".pdf"
    if r.returncode != 0 or not os.path.exists(saida) or os.path.getsize(saida) == 0:
        raise Problema(422, "nao-converteu", "Não consegui converter este arquivo. Ele pode estar danificado ou protegido por senha.")
    return saida


EXTENSAO = {"pdf": ".pdf", "ole": ".doc", "docx": ".docx", "xlsx": ".xlsx", "pptx": ".pptx", "odf": ".odt", "zip-office": ".docx", "rtf": ".rtf"}


def converter(dados: bytes, para: str, nome_ext: str):
    tipo = tipo_de(dados)
    pasta = tempfile.mkdtemp(prefix="cm-")
    try:
        # a extensão ajuda o LibreOffice a escolher o filtro (.doc x .xls x .ppt antigos)
        ext = EXTENSAO.get(tipo, ".bin")
        if tipo in ("ole", "odf", "zip-office") and nome_ext in (".doc", ".xls", ".ppt", ".odt", ".ods", ".odp", ".docx", ".xlsx", ".pptx"):
            ext = nome_ext
        entrada = os.path.join(pasta, "entrada" + ext)
        with open(entrada, "wb") as f:
            f.write(dados)
        if para == "docx":
            if tipo != "pdf":
                raise Problema(415, "tipo", "Para converter em Word, envie um PDF.")
            caminho, modo = pdf_para_docx(pasta, entrada)
        elif para == "pdf":
            if tipo in ("pdf", "desconhecido"):
                raise Problema(415, "tipo", "Envie um arquivo do Word, Excel, PowerPoint ou OpenDocument.")
            caminho, modo = office_para_pdf(pasta, entrada), "completo"
        else:
            raise Problema(400, "para", "Formato de saída desconhecido.")
        with open(caminho, "rb") as f:
            return f.read(), tipo, modo
    finally:
        shutil.rmtree(pasta, ignore_errors=True)


def registrar(**campos):
    """Registro sem conteúdo: só o que ajuda a acompanhar custo e erros."""
    print(json.dumps({"severity": "INFO", **campos}, ensure_ascii=False), file=sys.stdout, flush=True)


class Pedido(BaseHTTPRequestHandler):
    server_version = "converteme-conversor"
    sys_version = ""

    def log_message(self, *args):   # o padrão registraria o endereço e o caminho; não precisamos
        pass

    def responder(self, status, corpo: bytes, tipo="application/json; charset=utf-8", extra=None):
        self.send_response(status)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(corpo)

    def erro(self, p: Problema):
        self.responder(p.status, json.dumps({"erro": p.codigo, "mensagem": p.mensagem}, ensure_ascii=False).encode())

    def do_GET(self):
        if urlparse(self.path).path == "/saude":
            return self.responder(200, b"ok", "text/plain; charset=utf-8")
        self.responder(404, b'{"erro":"nao-existe"}')

    def do_POST(self):
        inicio = time.time()
        url = urlparse(self.path)
        if url.path != "/converter":
            return self.responder(404, b'{"erro":"nao-existe"}')
        if not CHAVE:
            return self.erro(Problema(503, "sem-chave", "Servidor sem chave configurada."))
        if not hmac.compare_digest(self.headers.get("X-Chave", ""), CHAVE):
            return self.erro(Problema(403, "chave", "Pedido não autorizado."))
        q = parse_qs(url.query)
        para = (q.get("para") or [""])[0]
        ext = (q.get("ext") or [""])[0].lower()
        nome_ext = ext if ext.startswith(".") and len(ext) <= 6 else ""
        try:
            tamanho = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            tamanho = 0
        if tamanho <= 0:
            return self.erro(Problema(411, "vazio", "Arquivo vazio."))
        if tamanho > MAX_MB * 1024 * 1024:
            return self.erro(Problema(413, "grande", f"Arquivo grande demais. O limite é {int(MAX_MB)} MB."))
        dados = self.rfile.read(tamanho)
        try:
            saida, tipo, modo = converter(dados, para, nome_ext)
        except Problema as p:
            registrar(evento="erro", para=para, codigo=p.codigo, mb=round(tamanho / 1048576, 1), s=round(time.time() - inicio, 1))
            return self.erro(p)
        except Exception as e:  # noqa: BLE001 — qualquer outra falha vira resposta clara, sem detalhes do arquivo
            registrar(evento="falha", para=para, classe=type(e).__name__, mb=round(tamanho / 1048576, 1), s=round(time.time() - inicio, 1))
            return self.erro(Problema(500, "falha", "Não consegui converter este arquivo."))
        finally:
            dados = None
        registrar(evento="convertido", de=tipo, para=para, modo=modo, mb=round(tamanho / 1048576, 1), s=round(time.time() - inicio, 1))
        self.responder(200, saida, TIPOS_SAIDA[para], {"X-Pedido": uuid.uuid4().hex[:8], "X-Modo": modo})


def main():
    servidor = ThreadingHTTPServer(("0.0.0.0", PORTA), Pedido)
    servidor.daemon_threads = True
    registrar(evento="iniciado", porta=PORTA)
    servidor.serve_forever()


if __name__ == "__main__":
    if len(sys.argv) == 4 and sys.argv[1] == "--pdf2docx":
        _pdf2docx_filho(sys.argv[2], sys.argv[3])
    else:
        main()
