"""
Converteme — PDF para Excel (.xlsx).

Duas situações:
  1. Tabela com grade (linhas e colunas desenhadas): o PyMuPDF acha a tabela
     (page.find_tables) e cada célula vira uma célula do Excel.
  2. Relatório em colunas, sem grade (o comum em sistema de RH, estoque, contas):
     cada linha do PDF vira uma linha da planilha, e as colunas saem do alinhamento:
     textos que começam no mesmo x (ou números que terminam no mesmo x) ficam na
     mesma coluna, em todas as páginas.

Números, valores em R$, porcentagens e datas no formato brasileiro viram número e
data de verdade (dá para somar e ordenar). Códigos com zero à esquerda, CPF, CNPJ e
números com 11 dígitos ou mais continuam como texto, para não perder nada.
"""
import datetime
import re

import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import layout

TOL = 3.0   # pontos: começos de texto a até 3 pt um do outro são a mesma coluna

R_NUM = re.compile(r"^(-?)(?:R\$\s*)?(-?)(\d{1,3}(?:\.\d{3})+|\d+)(?:,(\d+))?(%?)$")
R_DATA = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")


def valor(texto: str):
    """Texto -> (valor, formato do Excel ou None)."""
    t = texto.strip()
    m = R_DATA.match(t)
    if m:
        d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return datetime.datetime(a, mes, d), "DD/MM/YYYY"
        except ValueError:
            return t, None
    m = R_NUM.match(t.replace(" ", " "))
    if m:
        sinal = "-" if (m.group(1) or m.group(2)) else ""
        inteiro, dec, pct = m.group(3), m.group(4), m.group(5)
        so_digitos = inteiro.replace(".", "")
        if "." not in inteiro and not dec and len(so_digitos) > 1 and so_digitos.startswith("0"):
            return t, None          # código com zero à esquerda (008897)
        if len(so_digitos) >= 11:
            return t, None          # CPF, CNPJ, telefone, conta: texto
        n = float(sinal + so_digitos + ("." + dec if dec else ""))
        if pct:
            return n / 100, "0.0%" if dec else "0%"
        if "R$" in t:
            return n, '"R$" #,##0.00'
        if dec:
            return n, "#,##0." + "0" * min(len(dec), 6)
        if "." in inteiro:
            return int(n), "#,##0"
        return int(n), None
    return t, None


def _numerico(texto: str) -> bool:
    v, _ = valor(texto)
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _partes(itens):
    """Um pedaço de texto com 2 ou mais espaços seguidos ("000123  NOME") são duas colunas:
    separa, estimando a posição de cada parte pela proporção de letras."""
    for p in itens:
        texto = p["s"]["text"]
        if not re.search(r"\S {2,}\S", texto):
            yield p
            continue
        largura = (p["x1"] - p["x0"]) / max(1, len(texto))
        for m in re.finditer(r"\S+(?: \S+)*", texto):
            x0 = p["x0"] + m.start() * largura
            yield {**p, "x0": x0, "x1": x0 + len(m.group()) * largura, "s": {**p["s"], "text": m.group()}}


def _celulas(linha):
    """Junta as palavras de uma linha em células: palavras próximas são a mesma célula."""
    cels = []
    for p in _partes(linha["itens"]):
        s = p["s"]
        tam = s.get("size", 10)
        if cels and p["x0"] - cels[-1]["x1"] <= max(1.5, tam * 0.35):
            c = cels[-1]
            c["texto"] += ("" if c["texto"].endswith(" ") or s["text"].startswith(" ") else " ") + s["text"]
            c["x1"] = max(c["x1"], p["x1"])
            c["negrito"] = c["negrito"] and layout._negrito(s)
        else:
            cels.append({"x0": p["x0"], "x1": p["x1"], "texto": s["text"], "negrito": layout._negrito(s)})
    for c in cels:
        c["texto"] = re.sub(r"\s+", " ", c["texto"]).strip()
        c["num"] = _numerico(c["texto"])
    return [c for c in cels if c["texto"]]


def _agrupar(valores):
    """Agrupa posições próximas (até TOL). Devolve listas de posições."""
    grupos = []
    for v in sorted(valores):
        if grupos and v - grupos[-1][-1] <= TOL:
            grupos[-1].append(v)
        else:
            grupos.append([v])
    return grupos


def _colunas(linhas_cels):
    """Âncoras de coluna para todas as linhas: texto pelo começo (x0), número pelo fim (x1)."""
    esquerdas = [c["x0"] for cels in linhas_cels for c in cels if not c["num"]]
    direitas = [c["x1"] for cels in linhas_cels for c in cels if c["num"]]
    ancoras = []
    for g in _agrupar(esquerdas):
        ancoras.append({"tipo": "E", "min": g[0], "max": g[-1], "pos": sum(g) / len(g), "n": len(g)})
    for g in _agrupar(direitas):
        ancoras.append({"tipo": "D", "min": g[0], "max": g[-1], "pos": sum(g) / len(g), "n": len(g)})
    return ancoras


def _ancora_de(c, ancoras):
    chave = c["x1"] if c["num"] else c["x0"]
    tipo = "D" if c["num"] else "E"
    for i, a in enumerate(ancoras):
        if a["tipo"] == tipo and a["min"] - 0.01 <= chave <= a["max"] + 0.01:
            return i
    return None


JUNTAR = 14.0   # pontos: colunas vizinhas que nunca aparecem na mesma linha viram uma só


def _ordem_das_colunas(ancoras, linhas_cels):
    """Ordena as âncoras da esquerda para a direita (posição média do começo das células) e junta
    numa coluna só as vizinhas que nunca dividem uma linha (relatório com tipos de linha diferentes,
    como "funcionário" e "período", fica com menos colunas vazias). Devolve {âncora: coluna}."""
    soma = [[0.0, 0] for _ in ancoras]
    linhas_de = [set() for _ in ancoras]
    for r, cels in enumerate(linhas_cels):
        for c in cels:
            i = _ancora_de(c, ancoras)
            if i is not None:
                soma[i][0] += c["x0"]
                soma[i][1] += 1
                linhas_de[i].add(r)
    centro = [s / n if n else a["pos"] for (s, n), a in zip(soma, ancoras)]
    ordem = sorted(range(len(ancoras)), key=lambda i: centro[i])
    coluna, grupos = {}, []          # grupos: [centro do último, linhas usadas]
    for i in ordem:
        g = grupos[-1] if grupos else None
        if g is not None and centro[i] - g[0] <= JUNTAR and not (g[1] & linhas_de[i]):
            g[0] = centro[i]
            g[1] |= linhas_de[i]
        else:
            grupos.append([centro[i], set(linhas_de[i])])
        coluna[i] = len(grupos) - 1
    return coluna


def _verticais(pagina):
    v = 0
    for d in pagina.get_drawings():
        for it in d["items"]:
            if it[0] == "l" and abs(it[1].x - it[2].x) < 1 and abs(it[1].y - it[2].y) > 5:
                v += 1
            elif it[0] == "re" and it[1].width < 2 and it[1].height > 5:
                v += 1
    return v


def converter(entrada: str, saida: str) -> None:
    linhas_saida = []   # cada item: ("tabela", [[texto...]]) | ("linha", cels) | ("pagina", None)
    with pymupdf.open(entrada) as pdf:
        for n, pagina in enumerate(pdf):
            if n > 0:
                linhas_saida.append(("pagina", None, 0))
            areas, blocos = [], []
            if _verticais(pagina) >= 4:
                try:
                    for tab in pagina.find_tables().tables:
                        dados = tab.extract()
                        if dados and len(dados) >= 2 and max(len(r) for r in dados) >= 2:
                            areas.append(pymupdf.Rect(tab.bbox))
                            blocos.append(("tabela", dados, tab.bbox[1]))
                except Exception:  # noqa: BLE001 — se não achar a tabela, vai pelo alinhamento
                    areas, blocos = [], []
            for ln in layout._linhas_da_pagina(pagina):
                meio = pymupdf.Point((ln["itens"][0]["x0"] + ln["itens"][-1]["x1"]) / 2, (ln["topo"] + ln["fundo"]) / 2)
                if any(a.contains(meio) for a in areas):
                    continue
                cels = _celulas(ln)
                if cels:
                    blocos.append(("linha", cels, ln["topo"]))
            blocos.sort(key=lambda b: b[2])
            linhas_saida.extend(blocos)

    linhas_cels = [b[1] for b in linhas_saida if b[0] == "linha"]
    ancoras = _colunas(linhas_cels)
    ordem = _ordem_das_colunas(ancoras, linhas_cels)

    wb = Workbook()
    ws = wb.active
    ws.title = "PDF"
    larguras = {}
    r = 0
    negrito = Font(bold=True)

    def escrever(linha, coluna, texto, em_negrito=False):
        v, formato = valor(texto)
        cel = ws.cell(row=linha, column=coluna, value=v)
        if formato:
            cel.number_format = formato
        if em_negrito:
            cel.font = negrito
        larguras[coluna] = max(larguras.get(coluna, 0), 10 if formato == "DD/MM/YYYY" else len(str(texto)))

    for tipo, dados, _ in linhas_saida:
        if tipo == "pagina":
            continue                    # as páginas seguem na mesma planilha, uma linha depois da outra
        r += 1
        if tipo == "tabela":
            for i, linha in enumerate(dados):
                if i:
                    r += 1
                for j, texto in enumerate(linha):
                    if texto is not None and str(texto).strip():
                        escrever(r, j + 1, re.sub(r"\s+", " ", str(texto)).strip(), em_negrito=(i == 0))
            continue
        usadas = {}
        for c in dados:
            i = _ancora_de(c, ancoras)
            col = (ordem[i] if i is not None else 0) + 1
            if col in usadas:            # duas células caíram na mesma coluna: junta o texto
                anterior = ws.cell(row=r, column=col)
                anterior.value = f"{usadas[col]} {c['texto']}"
                anterior.number_format = "General"
                usadas[col] = anterior.value
                continue
            usadas[col] = c["texto"]
            escrever(r, col, c["texto"], c["negrito"])
    for col, w in larguras.items():
        ws.column_dimensions[get_column_letter(col)].width = max(6, min(60, w + 2))
    if r == 0:
        ws.cell(row=1, column=1, value="(o PDF não tem texto para levar à planilha)")
    wb.save(saida)


if __name__ == "__main__":
    import sys
    converter(sys.argv[1], sys.argv[2])
