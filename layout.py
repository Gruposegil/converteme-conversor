"""
Converteme — PDF para Word "no lugar" (conversão rápida).

Cada linha da página vira um parágrafo do Word. Cada pedaço de texto da linha fica
na mesma posição do PDF, com uma parada de tabulação (tab), e cada linha fica na
mesma altura, com espaço antes exato. As linhas horizontais do PDF viram bordas de
parágrafo. As imagens ficam na altura em que estavam.

Serve muito bem para relatórios de sistema (colunas, fichas, listagens), que são
justamente os PDFs em que a conversão detalhada (pdf2docx) fica lenta. Em texto
corrido, cada linha vira um parágrafo (o texto não "reflui" ao editar).

Usa só PyMuPDF e python-docx; leva poucos segundos mesmo com centenas de páginas.
"""
import io
import re

import pymupdf
from docx import Document
from docx.enum.text import WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

PT = 12700          # EMU por ponto
MARGEM_MIN = 18.0   # menor margem que o Word aceita bem (0,63 cm)

# fonte do PDF -> fonte do Word (as do Office; quem não tem usa a de mesmas medidas)
FONTES = [
    (r"arial|helvetica|liberationsans|arimo", "Arial"),
    (r"times|liberationserif|tinos", "Times New Roman"),
    (r"courier|liberationmono|cousine", "Courier New"),
    (r"calibri|carlito", "Calibri"),
    (r"cambria|caladea", "Cambria"),
    (r"verdana", "Verdana"),
    (r"tahoma", "Tahoma"),
    (r"georgia", "Georgia"),
    (r"segoe", "Segoe UI"),
]


def nome_fonte(pdf_nome: str) -> str:
    n = re.sub(r"^[A-Z]{6}\+", "", pdf_nome or "")       # tira o prefixo de subconjunto (ABCDEF+)
    base = re.split(r"[,\-]", n)[0]
    chave = re.sub(r"(MT|PS|PSMT)$", "", base).lower().replace(" ", "")
    for padrao, word in FONTES:
        if re.search(padrao, chave):
            return word
    return base or "Arial"


def _negrito(span) -> bool:
    return bool(span.get("flags", 0) & 16) or bool(re.search(r"bold|black|heavy|semibold", span.get("font", ""), re.I))


def _italico(span) -> bool:
    return bool(span.get("flags", 0) & 2) or bool(re.search(r"italic|oblique", span.get("font", ""), re.I))


def _linhas_da_pagina(pagina):
    """Junta os pedaços de texto que estão na mesma linha (mesma altura), da esquerda para a direita."""
    pedacos = []
    for b in pagina.get_text("dict").get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []):
            if abs(ln.get("dir", (1, 0))[1]) > 0.05:      # texto de lado ou de ponta-cabeça: fica de fora
                continue
            for s in ln.get("spans", []):
                if not s.get("text", "").strip():
                    continue
                x0, y0, x1, y1 = s["bbox"]
                pedacos.append({"x0": x0, "x1": x1, "topo": y0, "base": s.get("origin", (x0, y1))[1], "fundo": y1, "s": s})
    pedacos.sort(key=lambda p: (round(p["base"], 1), p["x0"]))
    linhas = []
    for p in pedacos:
        tam = p["s"].get("size", 10)
        alvo = None
        for ln in linhas[-3:]:
            if abs(ln["base"] - p["base"]) <= max(1.5, tam * 0.3):
                alvo = ln
        if alvo is None:
            alvo = {"base": p["base"], "topo": p["topo"], "fundo": p["fundo"], "itens": []}
            linhas.append(alvo)
        alvo["itens"].append(p)
        alvo["topo"], alvo["fundo"] = min(alvo["topo"], p["topo"]), max(alvo["fundo"], p["fundo"])
    for ln in linhas:
        ln["itens"].sort(key=lambda p: p["x0"])
    return linhas


def _regras(pagina):
    """Linhas horizontais finas (traços e retângulos achatados) da página: (y, x0, x1, espessura, cor)."""
    regras = []
    try:
        desenhos = pagina.get_drawings()
    except Exception:  # noqa: BLE001
        return regras
    for d in desenhos:
        r = d.get("rect")
        if r is None or r.width < 30:
            continue
        cor = d.get("color") or d.get("fill")
        if r.height <= 2.5:
            regras.append((r.y0 + r.height / 2, r.x0, r.x1, max(0.25, min(3, r.height or d.get("width") or 0.5)), cor))
    regras.sort()
    unicas = []
    for g in regras:      # a mesma linha desenhada duas vezes (traço + preenchimento) conta uma vez
        if unicas and abs(unicas[-1][0] - g[0]) < 0.8 and abs(unicas[-1][1] - g[1]) < 2:
            continue
        unicas.append(g)
    return unicas


def _imagens(pagina):
    out = []
    for b in pagina.get_text("dict").get("blocks", []):
        if b.get("type") == 1 and b.get("image"):
            x0, y0, x1, y1 = b["bbox"]
            if x1 - x0 > 8 and y1 - y0 > 8:
                out.append({"topo": y0, "fundo": y1, "x0": x0, "x1": x1, "img": b["image"], "ext": b.get("ext", "png")})
    return out


def _hex(cor):
    if cor is None:
        return "000000"
    if isinstance(cor, int):
        return f"{cor:06X}"
    r, g, b = (max(0, min(255, round(c * 255))) for c in list(cor)[:3])
    return f"{r:02X}{g:02X}{b:02X}"


def _borda_inferior(par, espessura, cor):
    pPr = par._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    lado = OxmlElement("w:bottom")
    lado.set(qn("w:val"), "single")
    lado.set(qn("w:sz"), str(max(2, min(24, round(espessura * 8)))))   # oitavos de ponto
    lado.set(qn("w:space"), "0")
    lado.set(qn("w:color"), _hex(cor))
    bdr.append(lado)
    pPr.append(bdr)


def _altura_borda(espessura):
    """Quanto a borda de parágrafo soma à altura no Word (a espessura, arredondada a oitavos de ponto)."""
    return max(2, min(24, round(espessura * 8))) / 8


def _formato(par, antes, altura):
    f = par.paragraph_format
    f.space_before = Pt(max(0, round(antes * 20) / 20))
    f.space_after = Pt(0)
    f.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    f.line_spacing = Pt(max(1, round(altura * 20) / 20))
    f.widow_control = False


def converter(entrada: str, saida: str) -> None:
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(0)
    with pymupdf.open(entrada) as pdf:
        primeira = pdf[0]
        largura, altura_pag = primeira.rect.width, primeira.rect.height
        paginas = []
        menor_x, maior_x = largura, 0.0
        for pagina in pdf:
            linhas = _linhas_da_pagina(pagina)
            regras = _regras(pagina)
            imagens = _imagens(pagina)
            for ln in linhas:
                menor_x = min(menor_x, ln["itens"][0]["x0"])
                maior_x = max(maior_x, max(p["x1"] for p in ln["itens"]))
            for g in regras:
                menor_x = min(menor_x, g[1])
            for im in imagens:
                menor_x = min(menor_x, im["x0"])
            paginas.append((pagina.rect, linhas, regras, imagens))
        esquerda = max(MARGEM_MIN, min(menor_x - 1, 72))
        sec = doc.sections[0]
        sec.page_width, sec.page_height = int(largura * PT), int(altura_pag * PT)
        sec.left_margin = int(esquerda * PT)
        # à direita, sobra o bastante para o texto que encosta na borda não quebrar de linha
        direita = max(2.0, min(MARGEM_MIN, largura - maior_x - 6))
        sec.right_margin = int(direita * PT)
        sec.top_margin = int(MARGEM_MIN * PT)
        sec.bottom_margin = int(6 * PT)     # folga: um arredondamento a mais não joga a última linha para outra página
        sec.header_distance = sec.footer_distance = 0
        util = largura - esquerda - direita

        primeiro_par_da_pagina = True
        for n, (rect, linhas, regras, imagens) in enumerate(paginas):
            # tudo o que vai na página, de cima para baixo
            coisas = [("linha", ln["topo"], ln) for ln in linhas]
            coisas += [("regra", g[0], g) for g in regras]
            coisas += [("imagem", im["topo"], im) for im in imagens]
            coisas.sort(key=lambda c: c[1])
            y = MARGEM_MIN                       # onde o Word está na página (em pontos, desde o topo)
            limite = rect.height - MARGEM_MIN
            primeiro_par_da_pagina = True
            ultimo = None
            for tipo, topo, item in coisas:
                if tipo == "regra":
                    yy, x0, x1, esp, cor = item
                    if ultimo is not None and ultimo[0] == "linha" and yy - y < 3:
                        _borda_inferior(ultimo[1], esp, cor)       # a linha logo abaixo do texto vira borda dele
                        y += _altura_borda(esp)
                        continue
                    par = doc.add_paragraph()
                    alt = 1.0
                    antes = max(0, yy - y - alt)
                    if y + antes + alt > limite:
                        continue
                    _formato(par, antes, alt)
                    par.paragraph_format.left_indent = Pt(max(0, x0 - esquerda))
                    par.paragraph_format.right_indent = Pt(max(0, util - (x1 - esquerda)))
                    _borda_inferior(par, esp, cor)
                    y += _altura_borda(esp)
                    if primeiro_par_da_pagina and n > 0:
                        par.paragraph_format.page_break_before = True
                    primeiro_par_da_pagina = False
                    y += antes + alt
                    ultimo = ("regra", par)
                    continue
                if tipo == "imagem":
                    alt = item["fundo"] - item["topo"]
                    larg = min(util, item["x1"] - item["x0"])
                    if y + alt > limite + 2:
                        alt = max(8, limite - y)
                    par = doc.add_paragraph()
                    _formato(par, max(0, item["topo"] - y), alt)
                    par.paragraph_format.left_indent = Pt(max(0, item["x0"] - esquerda))
                    try:
                        par.add_run().add_picture(io.BytesIO(item["img"]), width=Pt(larg), height=Pt(alt))
                    except Exception:  # noqa: BLE001 — formato de imagem que o Word não aceita: segue sem ela
                        pass
                    if primeiro_par_da_pagina and n > 0:
                        par.paragraph_format.page_break_before = True
                    primeiro_par_da_pagina = False
                    y = max(y, item["topo"]) + alt
                    ultimo = ("imagem", par)
                    continue
                ln = item
                tam = max(s["s"].get("size", 10) for s in ln["itens"])
                alt = max(ln["fundo"] - ln["topo"], tam * 1.15)
                antes = max(0, ln["topo"] - y)
                if y + antes + alt > limite + 1:              # não cabe: encolhe o espaço antes para não pular página
                    antes = max(0, limite - y - alt)
                par = doc.add_paragraph()
                _formato(par, antes, alt)
                if primeiro_par_da_pagina and n > 0:
                    par.paragraph_format.page_break_before = True
                primeiro_par_da_pagina = False
                paradas = par.paragraph_format.tab_stops
                fim_anterior = esquerda
                for i, p in enumerate(ln["itens"]):
                    s = p["s"]
                    texto = s["text"]
                    pos = p["x0"] - esquerda
                    if i == 0 and pos > 0.5:
                        par.paragraph_format.left_indent = Pt(0)
                    if i > 0 or pos > 0.5:
                        if p["x0"] <= fim_anterior + 1.5:      # colado no anterior: sem tab, só um espaço se havia
                            if p["x0"] - fim_anterior > 0.1 and not texto.startswith(" "):
                                texto = " " + texto
                        else:
                            paradas.add_tab_stop(Pt(round(pos * 20) / 20), WD_TAB_ALIGNMENT.LEFT)
                            texto = "\t" + texto
                    r = par.add_run(texto)
                    r.font.size = Pt(max(4, min(96, round(s.get("size", 10) * 2) / 2)))
                    r.font.name = nome_fonte(s.get("font", ""))
                    r.bold, r.italic = _negrito(s), _italico(s)
                    c = s.get("color", 0)
                    if c:
                        r.font.color.rgb = RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)
                    fim_anterior = max(fim_anterior, p["x1"])
                y += antes + alt
                ultimo = ("linha", par)
            if primeiro_par_da_pagina and n > 0:              # página sem nada: mantém a página em branco
                par = doc.add_paragraph()
                par.paragraph_format.page_break_before = True
    doc.save(saida)


if __name__ == "__main__":
    import sys
    converter(sys.argv[1], sys.argv[2])
