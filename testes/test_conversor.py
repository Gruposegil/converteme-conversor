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

    def test_plano_b_quando_a_detalhada_nao_termina(self):
        """Com prazo de 0 segundo a conversão detalhada "trava": tem de sair a simples, com texto, negrito e imagem."""
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
                self.assertEqual(r.headers["X-Modo"], "simples")
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
        """Com teto de memória minúsculo a conversão detalhada morre: o servidor continua de pé e entrega a simples."""
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
                self.assertEqual(r.headers["X-Modo"], "simples")
                from docx import Document
                self.assertIn("Texto inventado", "\n".join(p.text for p in Document(io.BytesIO(r.read())).paragraphs))
            urllib.request.urlopen(f"http://127.0.0.1:{porta}/saude", timeout=2)   # ainda de pé
        finally:
            proc.terminate(); proc.wait(5)

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
