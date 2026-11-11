"""Testes do servidor de conversão: sobe o app.py de verdade numa porta livre."""
import io, json, os, socket, subprocess, sys, time, unittest, urllib.request, urllib.error, zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(AQUI, "..", "app.py")
CHAVE = "chave-so-de-teste"


def porta_livre():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p


def docx_exemplo():
    from docx import Document
    d = Document()
    d.add_heading("Relatório de exemplo", 1)
    d.add_paragraph("Ação, coração e informação: acentos do português.")
    t = d.add_table(rows=2, cols=2); t.style = "Table Grid"   # com bordas, como quase toda tabela de documento
    t.cell(0, 0).text, t.cell(0, 1).text, t.cell(1, 0).text, t.cell(1, 1).text = "Item", "Valor", "Serviço", "R$ 1.234,56"
    b = io.BytesIO(); d.save(b); return b.getvalue()


def xlsx_exemplo():
    """Planilha mínima feita à mão (sem biblioteca)."""
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Planilha1" sheetId="1" r:id="r1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Produto inventado</t></is></c><c r="D1"><v>42</v></c></row></sheetData></worksheet>')
    return b.getvalue()


def relatorio_exemplo(paginas=3):
    """Relatório como os de sistema de RH, com dados inventados: cabeçalho, linhas horizontais,
    e cada campo escrito numa posição fixa (sem grade de tabela)."""
    import pymupdf
    d = pymupdf.open()
    n = 0
    for pg in range(paginas):
        p = d.new_page(width=595, height=839)
        p.insert_text((20, 30), "Empresa:  EMPRESA FICTICIA DE TESTE LTDA", fontsize=9)
        p.insert_text((500, 30), f"Página: {pg + 1:04d}", fontsize=8)
        p.draw_line((20, 42), (584, 42))
        p.insert_text((24, 92), "Código  Nome", fontsize=7)
        for x, t in ((270, "Admissão"), (330, "Vcto. Férias"), (385, "Conceder até")):
            p.insert_text((x, 92), t, fontsize=7)
        p.draw_line((20, 97), (584, 97))
        y = 120
        while y < 790:
            n += 1
            p.insert_text((24, y), f"{n:06d}  FUNCIONARIO INVENTADO {n:03d}", fontsize=7)
            p.insert_text((270, y), f"{(n % 28) + 1:02d}/03/2015", fontsize=7)
            p.insert_text((330, y), f"{(n % 28) + 1:02d}/03/2016", fontsize=7)
            p.insert_text((385, y), f"{(n % 28) + 1:02d}/02/2017", fontsize=7)
            for i in range(12):
                p.insert_text((450 + i * 11, y), "1", fontsize=6)
            p.insert_text((24, y + 17), "Período aquisitivo", fontsize=7, fontname="hebo")
            for x, t in ((110, "Faltas"), (220, "Dias gozados"), (280, "Dias de direito")):
                p.insert_text((x, y + 17), t, fontsize=7, fontname="hebo")
            p.insert_text((24, y + 26), "INVENTADO", fontsize=7)
            for x, t in ((130, "0"), (240, "0"), (300, "30")):
                p.insert_text((x, y + 26), t, fontsize=7)
            p.draw_line((24, y + 40), (582, y + 40))
            y += 60
    return d.tobytes()


class TesteConversor(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.porta = porta_livre()
        cls.proc = subprocess.Popen([sys.executable, APP], env={**os.environ, "PORT": str(cls.porta), "CHAVE": CHAVE, "MAX_MB": "5"}, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        for _ in range(100):
            try:
                urllib.request.urlopen(f"http://127.0.0.1:{cls.porta}/saude", timeout=1); break
            except Exception:
                time.sleep(0.1)

    @classmethod
    def tearDownClass(cls):
        cls.proc.terminate(); cls.proc.wait(5)

    def pedir(self, corpo, para, chave=CHAVE, ext=""):
        req = urllib.request.Request(f"http://127.0.0.1:{self.porta}/converter?para={para}&ext={ext}", data=corpo, method="POST", headers={"X-Chave": chave} if chave is not None else {})
        try:
            with urllib.request.urlopen(req, timeout=180) as r:
                return r.status, r.read(), r.headers
        except urllib.error.HTTPError as e:
            return e.code, e.read(), e.headers

    def test_saude(self):
        with urllib.request.urlopen(f"http://127.0.0.1:{self.porta}/saude") as r:
            self.assertEqual(r.read(), b"ok")

    def test_chave(self):
        st, corpo, _ = self.pedir(b"%PDF-1.4", "docx", chave="errada")
        self.assertEqual(st, 403)
        st, _, _ = self.pedir(b"%PDF-1.4", "docx", chave=None)
        self.assertEqual(st, 403)

    def test_tamanho_e_tipo(self):
        try:
            st, corpo, _ = self.pedir(b"x" * (6 * 1024 * 1024), "pdf")
            self.assertEqual(st, 413)
        except urllib.error.URLError:
            pass   # o servidor recusa sem ler o corpo: a conexão fecha (também é recusa)
        st, corpo, _ = self.pedir(b"isto nao e documento", "pdf")
        self.assertEqual(st, 415)
        self.assertIn("Word", json.loads(corpo)["mensagem"])
        st, corpo, _ = self.pedir(docx_exemplo(), "docx")
        self.assertEqual(st, 415)

    def test_word_para_pdf_e_volta(self):
        st, pdf, h = self.pedir(docx_exemplo(), "pdf", ext=".docx")
        self.assertEqual(st, 200, pdf[:300])
        self.assertEqual(h["Content-Type"], "application/pdf")
        self.assertEqual(h["Cache-Control"], "no-store")
        import pymupdf
        with pymupdf.open(stream=pdf, filetype="pdf") as d:
            texto = "".join(p.get_text() for p in d)
        self.assertIn("Ação, coração e informação", texto)
        self.assertIn("R$ 1.234,56", texto)
        # e de volta para Word
        st, docx, h = self.pedir(pdf, "docx")
        self.assertEqual(st, 200, docx[:300])
        self.assertEqual(h["X-Modo"], "completo")
        from docx import Document
        d = Document(io.BytesIO(docx))
        tudo = "\n".join(p.text for p in d.paragraphs) + "\n" + "\n".join(c.text for t in d.tables for r in t.rows for c in r.cells)
        self.assertIn("coração", tudo)
        self.assertIn("1.234,56", tudo)
        self.assertTrue(d.tables, "a tabela continua tabela")

    def test_word_antigo_doc_para_pdf(self):
        """.doc (Word 97-2003): o LibreOffice abre o formato antigo e gera o PDF."""
        import shutil, subprocess, tempfile, pymupdf
        with tempfile.TemporaryDirectory() as t:
            with open(os.path.join(t, "a.docx"), "wb") as f:
                f.write(docx_exemplo())
            subprocess.run([shutil.which("soffice") or "soffice", f"-env:UserInstallation=file://{t}/perfil", "--headless",
                            "--convert-to", "doc", "--outdir", t, os.path.join(t, "a.docx")], check=True, capture_output=True, timeout=120)
            with open(os.path.join(t, "a.doc"), "rb") as f:
                doc = f.read()
        self.assertEqual(doc[:4], bytes.fromhex("d0cf11e0"), "é um .doc de verdade (OLE)")
        st, pdf, h = self.pedir(doc, "pdf", ext=".doc")
        self.assertEqual(st, 200, pdf[:300])
        with pymupdf.open(stream=pdf, filetype="pdf") as d:
            self.assertIn("coração", "".join(p.get_text() for p in d))

    def test_plano_b_quando_a_detalhada_nao_termina(self):
        """Com prazo de 0 segundo a conversão detalhada "trava": tem de sair a "no lugar", com texto, negrito e imagem."""
        import pymupdf
        from PIL import Image
        porta = porta_livre()
        proc = subprocess.Popen([sys.executable, APP], env={**os.environ, "PORT": str(porta), "CHAVE": CHAVE, "PRAZO_DOCX_SEGUNDOS": "0"}, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        try:
            for _ in range(100):
                try:
                    urllib.request.urlopen(f"http://127.0.0.1:{porta}/saude", timeout=1); break
                except Exception:
                    time.sleep(0.1)
            d = pymupdf.open(); pg = d.new_page()
            pg.insert_text((72, 72), "Relatório complicado", fontsize=18, fontname="hebo")
            pg.insert_text((72, 110), "Primeira linha do parágrafo com ação.", fontsize=11)
            b = io.BytesIO(); Image.new("RGB", (120, 60), (20, 90, 200)).save(b, "PNG")
            pg.insert_image(pymupdf.Rect(72, 150, 192, 210), stream=b.getvalue())
            d.new_page().insert_text((72, 72), "Segunda página", fontsize=11)
            req = urllib.request.Request(f"http://127.0.0.1:{porta}/converter?para=docx", data=d.tobytes(), method="POST", headers={"X-Chave": CHAVE})
            with urllib.request.urlopen(req, timeout=60) as r:
                self.assertEqual(r.headers["X-Modo"], "posicao")
                corpo = r.read()
            from docx import Document
            doc = Document(io.BytesIO(corpo))
            textos = [p.text for p in doc.paragraphs if p.text.strip()]
            self.assertIn("Relatório complicado", textos[0])
            self.assertTrue(doc.paragraphs[0].runs[0].bold or any(r.bold for p in doc.paragraphs for r in p.runs if "Relatório" in r.text), "negrito mantido")
            self.assertTrue(any("ação" in t for t in textos))
            self.assertTrue(any("Segunda página" in t for t in textos))
            self.assertEqual(len(doc.inline_shapes), 1, "a imagem veio junto")
        finally:
            proc.terminate(); proc.wait(5)

    def test_plano_b_quando_a_detalhada_estoura_a_memoria(self):
        """Com teto de memória minúsculo a conversão detalhada morre: o servidor continua de pé e entrega a "no lugar"."""
        import pymupdf
        porta = porta_livre()
        proc = subprocess.Popen([sys.executable, APP], env={**os.environ, "PORT": str(porta), "CHAVE": CHAVE, "MEMORIA_DOCX_MB": "40"}, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        try:
            for _ in range(100):
                try:
                    urllib.request.urlopen(f"http://127.0.0.1:{porta}/saude", timeout=1); break
                except Exception:
                    time.sleep(0.1)
            d = pymupdf.open(); d.new_page().insert_text((72, 72), "Texto inventado para o teste", fontsize=12)
            req = urllib.request.Request(f"http://127.0.0.1:{porta}/converter?para=docx", data=d.tobytes(), method="POST", headers={"X-Chave": CHAVE})
            with urllib.request.urlopen(req, timeout=60) as r:
                self.assertEqual(r.headers["X-Modo"], "posicao")
                from docx import Document
                self.assertIn("Texto inventado", "\n".join(p.text for p in Document(io.BytesIO(r.read())).paragraphs))
            urllib.request.urlopen(f"http://127.0.0.1:{porta}/saude", timeout=2)   # ainda de pé
        finally:
            proc.terminate(); proc.wait(5)

    def test_relatorio_em_colunas_fica_no_lugar(self):
        """Relatório de sistema (inventado): vai direto para a conversão "no lugar", rápida,
        com as mesmas páginas e cada linha na mesma altura e coluna do PDF."""
        import pymupdf, re, shutil, tempfile
        pdf = relatorio_exemplo()
        inicio = time.time()
        st, docx, h = self.pedir(pdf, "docx")
        self.assertEqual(st, 200, docx[:300])
        self.assertEqual(h["X-Modo"], "posicao")
        self.assertLess(time.time() - inicio, 20, "sai em segundos")
        from docx import Document
        d = Document(io.BytesIO(docx))
        texto = "\n".join(p.text for p in d.paragraphs)
        self.assertIn("FUNCIONARIO INVENTADO 007", texto)
        self.assertIn("Período aquisitivo", texto)
        self.assertTrue(any(p.paragraph_format.tab_stops and len(p.paragraph_format.tab_stops) >= 3 for p in d.paragraphs), "colunas por tabulação")
        # abre o Word no LibreOffice e confere páginas e posições
        pasta = tempfile.mkdtemp()
        try:
            with open(os.path.join(pasta, "r.docx"), "wb") as f:
                f.write(docx)
            subprocess.run(["soffice", f"-env:UserInstallation=file://{pasta}/perfil", "--headless", "--convert-to", "pdf", "--outdir", pasta, os.path.join(pasta, "r.docx")], capture_output=True, timeout=120)
            with pymupdf.open(stream=pdf, filetype="pdf") as orig, pymupdf.open(os.path.join(pasta, "r.pdf")) as volta:
                self.assertEqual(volta.page_count, orig.page_count, "mesmo número de páginas")
                for n in range(orig.page_count):
                    def posicoes(pg):
                        return {w[4]: (w[0], w[1]) for w in pg.get_text("words") if re.fullmatch(r"\d{2}/\d{2}/\d{4}|INVENTADO", w[4])}
                    a, b = posicoes(orig[n]), posicoes(volta[n])
                    comuns = set(a) & set(b)
                    self.assertGreater(len(comuns), 5)
                    for k in comuns:
                        self.assertAlmostEqual(a[k][0], b[k][0], delta=3, msg=f"coluna de {k} na página {n + 1}")
                        self.assertAlmostEqual(a[k][1], b[k][1], delta=4, msg=f"altura de {k} na página {n + 1}")
        finally:
            shutil.rmtree(pasta, ignore_errors=True)

    def test_pdf_para_excel_relatorio(self):
        """Relatório em colunas (inventado) vira planilha: cada campo na sua coluna, em todas as
        páginas; data é data, número é número, código com zero à esquerda continua texto."""
        import datetime
        from openpyxl import load_workbook
        st, xlsx, h = self.pedir(relatorio_exemplo(), "xlsx")
        self.assertEqual(st, 200, xlsx[:300])
        self.assertEqual(h["Content-Type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.assertEqual(h["X-Modo"], "planilha")
        ws = load_workbook(io.BytesIO(xlsx)).active
        linhas = [r for r in ws.iter_rows(values_only=True)]
        func = [r for r in linhas if any(isinstance(v, str) and v.startswith("FUNCIONARIO INVENTADO") for v in r)]
        self.assertGreater(len(func), 30, "todas as páginas")
        col_nome = {r.index(next(v for v in r if isinstance(v, str) and v.startswith("FUNCIONARIO"))) for r in func}
        self.assertEqual(len(col_nome), 1, "o nome fica sempre na mesma coluna")
        primeira = func[0]
        self.assertEqual(primeira[0], "000001", "código com zero à esquerda é texto")
        datas = [v for v in primeira if isinstance(v, datetime.datetime)]
        self.assertEqual(len(datas), 3, "as três datas viram data")
        col_data = {r.index(next(v for v in r if isinstance(v, datetime.datetime))) for r in func}
        self.assertEqual(len(col_data), 1, "a primeira data fica sempre na mesma coluna")
        self.assertIn(30, [v for r in linhas for v in r], "número vira número")

    def test_pdf_para_excel_tabela_com_grade(self):
        """Tabela com bordas (Word -> PDF -> Excel): as células da tabela viram células; R$ vira número."""
        from openpyxl import load_workbook
        st, pdf, _ = self.pedir(docx_exemplo(), "pdf", ext=".docx")
        self.assertEqual(st, 200)
        st, xlsx, h = self.pedir(pdf, "xlsx")
        self.assertEqual(st, 200, xlsx[:300])
        wb = load_workbook(io.BytesIO(xlsx))
        ws = wb.active
        celulas = {(c.row, c.column): c for r in ws.iter_rows() for c in r if c.value is not None}
        servico = next(c for c in celulas.values() if c.value == "Serviço")
        valor = celulas[(servico.row, servico.column + 1)]
        self.assertEqual(valor.value, 1234.56)
        self.assertIn("R$", valor.number_format)
        item = next(c for c in celulas.values() if c.value == "Item")
        self.assertEqual(celulas[(item.row, item.column + 1)].value, "Valor")
        self.assertEqual(item.column, servico.column, "mesma coluna da tabela")

    def test_pdf_para_excel_numero_embaixo_do_titulo(self):
        """Título alinhado à esquerda e número alinhado à direita, dentro da largura do título
        (como no relatório de férias): o número fica na coluna do título, e não na vizinha.
        Um texto largo de cabeçalho por perto ("Documento:") não pode puxar os números."""
        import pymupdf
        from openpyxl import load_workbook
        d = pymupdf.open()
        f = pymupdf.Font("helv")
        meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun"]
        for pg in range(2):
            p = d.new_page(width=595, height=842)
            p.insert_text((330, 40), "Documento:", fontsize=9)
            for i, m in enumerate(meses):
                p.insert_text((340 + i * 15, 60), m, fontsize=6)
            y = 90
            for n in range(12):
                p.insert_text((24, y), "Período inventado", fontsize=7, fontname="hebo")
                for x, t in ((110, "Faltas no período"), (200, "Dias gozados"), (260, "Dias de direito")):
                    p.insert_text((x, y), t, fontsize=7, fontname="hebo")
                for x_fim, v in ((160, str(n * 7)), (236, str(n % 3)), (300, str(30 - n))):
                    p.insert_text((x_fim - f.text_length(v, 7), y + 10), v, fontsize=7)       # números alinhados à direita
                for i in range(len(meses)):
                    v = str(n % 2 + 1)
                    p.insert_text((350 + i * 15 - f.text_length(v, 6), y + 10), v, fontsize=6)
                y += 40
        st, xlsx, _ = self.pedir(d.tobytes(), "xlsx")
        self.assertEqual(st, 200, xlsx[:300])
        ws = load_workbook(io.BytesIO(xlsx)).active
        col = {}
        for r in ws.iter_rows():
            for c in r:
                if isinstance(c.value, str):
                    col.setdefault(c.value, set()).add(c.column)
        numeros = {}
        for r in ws.iter_rows():
            vals = [c for c in r if isinstance(c.value, int)]
            if len(vals) == 3 + len(meses):
                for k, c in enumerate(vals):
                    numeros.setdefault(k, set()).add(c.column)
        self.assertEqual(len(numeros), 3 + len(meses), "achou as linhas de números")
        titulos = ["Faltas no período", "Dias gozados", "Dias de direito"] + meses
        for k, t in enumerate(titulos):
            self.assertEqual(numeros[k], col[t], f"números de {t!r} na coluna do título")

    def _apresentacao(self):
        import pymupdf
        d = pymupdf.open()
        for i in range(3):
            p = d.new_page(width=842, height=595)
            p.insert_text((60, 80), f"Apresentação inventada, slide {i + 1}", fontsize=28)
            p.insert_text((60, 140), "Tópico com acentuação: ação e coração", fontsize=18)
            p.draw_rect(pymupdf.Rect(500, 250, 780, 500), color=(0.2, 0.4, 0.8), fill=(0.85, 0.9, 1))
        return d.tobytes()

    def test_pdf_para_powerpoint(self):
        """Cada página vira um slide, com o texto editável (caixas de texto) e o desenho."""
        from pptx import Presentation
        st, pptx, h = self.pedir(self._apresentacao(), "pptx")
        self.assertEqual(st, 200, pptx[:300])
        self.assertEqual(h["Content-Type"], "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        apr = Presentation(io.BytesIO(pptx))
        self.assertEqual(len(apr.slides), 3)
        textos = [sh.text_frame.text for sh in apr.slides[1].shapes if sh.has_text_frame]
        self.assertTrue(any("slide 2" in t for t in textos), textos)
        self.assertTrue(any("coração" in t for t in textos), textos)
        self.assertGreater(apr.slide_width, apr.slide_height, "paisagem, como o PDF")

    def test_pdf_para_pdfa(self):
        """PDF/A-2b (padrão) e PDF/A-1b: identificação no XMP, perfil de cor (OutputIntent) e todas as fontes embutidas."""
        import pikepdf, pymupdf
        for nivel, esperado in (("2", "2B"), ("1", "1B")):
            req = urllib.request.Request(f"http://127.0.0.1:{self.porta}/converter?para=pdfa&nivel={nivel}", data=self._apresentacao(), method="POST", headers={"X-Chave": CHAVE})
            with urllib.request.urlopen(req, timeout=120) as r:
                self.assertEqual(r.headers["Content-Type"], "application/pdf")
                self.assertEqual(r.headers["X-Modo"], "pdfa-" + nivel)
                corpo = r.read()
            with pikepdf.open(io.BytesIO(corpo)) as p:
                self.assertEqual(p.open_metadata().pdfa_status, esperado)
                self.assertIn("/OutputIntents", p.Root)
                self.assertEqual(str(p.Root.OutputIntents[0].S), "/GTS_PDFA1")
            with pymupdf.open(stream=corpo, filetype="pdf") as d:
                self.assertEqual(d.page_count, 3)
                fontes = {f[3]: f[1] for pg in d for f in pg.get_fonts(full=True)}
                self.assertTrue(fontes and all(ext != "n/a" for ext in fontes.values()), f"fontes embutidas: {fontes}")
                self.assertIn("coração", d[0].get_text())

    def test_pdfa_tira_o_que_o_padrao_proibe(self):
        """JavaScript e anexos saem; no PDF/A-1b, a página com transparência vira imagem e o servidor avisa."""
        import pymupdf
        d = pymupdf.open()
        p = d.new_page()
        p.insert_text((50, 50), "Página com transparência")
        p.draw_rect(pymupdf.Rect(50, 150, 200, 250), fill=(1, 0, 0), fill_opacity=0.5)
        d.new_page().insert_text((50, 50), "Página comum")
        d.embfile_add("anexo.txt", b"conteudo inventado")
        d.xref_set_key(d.pdf_catalog(), "OpenAction", "<< /S /JavaScript /JS (app.alert(1)) >>")
        original = d.tobytes()
        for nivel, modo in (("2", "pdfa-2"), ("1", "pdfa-1-imagem-1")):
            req = urllib.request.Request(f"http://127.0.0.1:{self.porta}/converter?para=pdfa&nivel={nivel}", data=original, method="POST", headers={"X-Chave": CHAVE})
            with urllib.request.urlopen(req, timeout=120) as r:
                self.assertEqual(r.headers["X-Modo"], modo)
                corpo = r.read()
            with pymupdf.open(stream=corpo, filetype="pdf") as o:
                self.assertNotIn("JavaScript", o.xref_object(o.pdf_catalog()))
                self.assertEqual(o.embfile_count(), 0)
                self.assertIn("comum", o[1].get_text())

    def test_valores_brasileiros(self):
        sys.path.insert(0, os.path.join(AQUI, ".."))
        import datetime
        import planilha
        self.assertEqual(planilha.valor("008897"), ("008897", None))
        self.assertEqual(planilha.valor("40170029000136")[0], "40170029000136")
        self.assertEqual(planilha.valor("123.456.789-09")[0], "123.456.789-09")
        self.assertEqual(planilha.valor("1.234,56"), (1234.56, "#,##0.00"))
        self.assertEqual(planilha.valor("R$ 1.234,56")[0], 1234.56)
        self.assertEqual(planilha.valor("-12,5%")[0], -0.125)
        self.assertEqual(planilha.valor("0"), (0, None))
        self.assertEqual(planilha.valor("30"), (30, None))
        self.assertEqual(planilha.valor("01/10/2020")[0], datetime.datetime(2020, 10, 1))
        self.assertEqual(planilha.valor("31/02/2020")[0], "31/02/2020", "data que não existe fica texto")

    def test_excel_para_pdf(self):
        st, pdf, _ = self.pedir(xlsx_exemplo(), "pdf", ext=".xlsx")
        self.assertEqual(st, 200, pdf[:300])
        import pymupdf
        with pymupdf.open(stream=pdf, filetype="pdf") as d:
            self.assertIn("Produto inventado", d[0].get_text())

    def test_pdf_com_senha(self):
        import pymupdf
        d = pymupdf.open(); d.new_page().insert_text((72, 72), "segredo"); b = d.tobytes(encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="123", owner_pw="456")
        st, corpo, _ = self.pedir(b, "docx")
        self.assertEqual(st, 422)
        self.assertEqual(json.loads(corpo)["erro"], "senha")

    def test_nao_deixa_arquivo(self):
        import tempfile, glob
        antes = set(glob.glob(os.path.join(tempfile.gettempdir(), "cm-*")))
        self.pedir(docx_exemplo(), "pdf", ext=".docx")
        depois = set(glob.glob(os.path.join(tempfile.gettempdir(), "cm-*")))
        self.assertEqual(depois - antes, set(), "a pasta temporária foi apagada")


if __name__ == "__main__":
    unittest.main(verbosity=2)
