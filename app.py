"""
Converteme — servidor de conversão (Google Cloud Run)

O que faz:
  POST /converter?para=docx   PDF  -> Word (.docx), com pdf2docx (MIT) + PyMuPDF (AGPL)
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
PRAZO = int(os.environ.get("PRAZO_SEGUNDOS", "150"))
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


def pdf_para_docx(pasta: str, entrada: str) -> str:
    import pymupdf as fitz  # PyMuPDF
    from pdf2docx import Converter
    with fitz.open(entrada) as d:
        if d.needs_pass:
            raise Problema(422, "senha", "Este PDF tem senha. Tire a senha primeiro e tente de novo.")
        if d.page_count > MAX_PAGINAS:
            raise Problema(413, "paginas", f"Este PDF tem {d.page_count} páginas. O limite é {MAX_PAGINAS}.")
        if d.page_count == 0:
            raise Problema(422, "vazio", "Este PDF não tem páginas.")
    saida = os.path.join(pasta, "saida.docx")
    cv = Converter(entrada)
    try:
        cv.convert(saida, multi_processing=False)
    finally:
        cv.close()
    return saida


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
            caminho = pdf_para_docx(pasta, entrada)
        elif para == "pdf":
            if tipo in ("pdf", "desconhecido"):
                raise Problema(415, "tipo", "Envie um arquivo do Word, Excel, PowerPoint ou OpenDocument.")
            caminho = office_para_pdf(pasta, entrada)
        else:
            raise Problema(400, "para", "Formato de saída desconhecido.")
        with open(caminho, "rb") as f:
            return f.read(), tipo
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
            saida, tipo = converter(dados, para, nome_ext)
        except Problema as p:
            registrar(evento="erro", para=para, codigo=p.codigo, mb=round(tamanho / 1048576, 1), s=round(time.time() - inicio, 1))
            return self.erro(p)
        except Exception as e:  # noqa: BLE001 — qualquer outra falha vira resposta clara, sem detalhes do arquivo
            registrar(evento="falha", para=para, classe=type(e).__name__, mb=round(tamanho / 1048576, 1), s=round(time.time() - inicio, 1))
            return self.erro(Problema(500, "falha", "Não consegui converter este arquivo."))
        finally:
            dados = None
        registrar(evento="convertido", de=tipo, para=para, mb=round(tamanho / 1048576, 1), s=round(time.time() - inicio, 1))
        self.responder(200, saida, TIPOS_SAIDA[para], {"X-Pedido": uuid.uuid4().hex[:8]})


def main():
    servidor = ThreadingHTTPServer(("0.0.0.0", PORTA), Pedido)
    servidor.daemon_threads = True
    registrar(evento="iniciado", porta=PORTA)
    servidor.serve_forever()


if __name__ == "__main__":
    main()
