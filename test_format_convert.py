"""Tests fuer dashboard/format_convert.py (offline, temporaere Dateien)."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CONV = ROOT / "format_convert.py"


def run(*args):
    p = subprocess.run([sys.executable, str(CONV), *args], text=True,
                       encoding="utf-8", capture_output=True, timeout=60)
    return p


class FormatConvertTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_list_matrix(self):
        p = run("--list")
        self.assertEqual(p.returncode, 0)
        m = json.loads(p.stdout)
        self.assertIn(".md", m["inputs"])
        self.assertIn(".pdf", m["inputs"][".md"])

    def test_md_to_pdf(self):
        src = self.d / "a.md"
        src.write_text("# Titel\n\nHallo Welt\n", encoding="utf-8")
        out = self.d / "a.pdf"
        p = run(str(src), "-o", str(out))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertTrue(out.read_bytes().startswith(b"%PDF"))

    def test_txt_to_docx_back(self):
        src = self.d / "b.txt"
        src.write_text("# Kopf\n\nText eins\n", encoding="utf-8")
        docx = self.d / "b.docx"
        self.assertEqual(run(str(src), "-o", str(docx)).returncode, 0)
        back = self.d / "b2.txt"
        p = run(str(docx), "-o", str(back))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertIn("Text eins", back.read_text(encoding="utf-8"))

    def test_csv_xlsx_roundtrip_and_md(self):
        src = self.d / "c.csv"
        src.write_text("name,wert\na,1\nb,2\n", encoding="utf-8")
        xlsx = self.d / "c.xlsx"
        self.assertEqual(run(str(src), "-o", str(xlsx)).returncode, 0)
        back = self.d / "c2.csv"
        self.assertEqual(run(str(xlsx), "-o", str(back)).returncode, 0)
        md = self.d / "c.md"
        self.assertEqual(run(str(src), "-o", str(md)).returncode, 0)
        self.assertIn("| name | wert |", md.read_text(encoding="utf-8"))

    def test_json_to_xlsx(self):
        src = self.d / "d.json"
        src.write_text('[{"a": 1, "b": "x"}]', encoding="utf-8")
        out = self.d / "d.xlsx"
        self.assertEqual(run(str(src), "-o", str(out)).returncode, 0)

    def test_image_to_pdf(self):
        from PIL import Image
        img = self.d / "e.png"
        Image.new("RGB", (40, 40), "red").save(img)
        out = self.d / "e.pdf"
        p = run(str(img), "-o", str(out))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertTrue(out.read_bytes().startswith(b"%PDF"))

    def test_pdf_merge(self):
        from PIL import Image
        outs = []
        for i in range(2):
            img = self.d / f"m{i}.png"
            Image.new("RGB", (20, 20), "blue").save(img)
            pdf = self.d / f"m{i}.pdf"
            self.assertEqual(run(str(img), "-o", str(pdf)).returncode, 0)
            outs.append(str(pdf))
        merged = self.d / "merged.pdf"
        p = run(*outs, "-o", str(merged))
        self.assertEqual(p.returncode, 0, p.stderr)
        from pypdf import PdfReader
        self.assertEqual(len(PdfReader(str(merged)).pages), 2)

    def test_no_overwrite_without_flag(self):
        src = self.d / "f.md"
        src.write_text("x\n", encoding="utf-8")
        out = self.d / "f.pdf"
        self.assertEqual(run(str(src), "-o", str(out)).returncode, 0)
        p = run(str(src), "-o", str(out))
        self.assertNotEqual(p.returncode, 0)

    def test_unsupported_refused(self):
        src = self.d / "g.mp4"
        src.write_bytes(b"fake")
        p = run(str(src), "-o", str(self.d / "g.pdf"))
        self.assertNotEqual(p.returncode, 0)

    def test_bad_route_refused(self):
        src = self.d / "h.png"
        from PIL import Image
        Image.new("RGB", (10, 10), "green").save(src)
        p = run(str(src), "-o", str(self.d / "h.xlsx"))
        self.assertNotEqual(p.returncode, 0)

    def test_html_to_pdf_and_txt(self):
        src = self.d / "i.html"
        src.write_text("<html><body><h1>Hi</h1><p>A &amp; B</p></body></html>", encoding="utf-8")
        pdf = self.d / "i.pdf"
        self.assertEqual(run(str(src), "-o", str(pdf)).returncode, 0)
        self.assertTrue(pdf.read_bytes().startswith(b"%PDF"))
        txt = self.d / "i.txt"
        self.assertEqual(run(str(src), "-o", str(txt)).returncode, 0)
        content = txt.read_text(encoding="utf-8")
        self.assertIn("Hi", content)
        self.assertIn("A & B", content)

    def test_xlsx_to_pdf_and_md(self):
        import openpyxl
        src = self.d / "j.xlsx"
        wb = openpyxl.Workbook()
        wb.active.append(["n", "w"])
        wb.active.append(["a", 1])
        wb.save(str(src))
        pdf = self.d / "j.pdf"
        self.assertEqual(run(str(src), "-o", str(pdf)).returncode, 0)
        self.assertTrue(pdf.read_bytes().startswith(b"%PDF"))
        md = self.d / "j.md"
        self.assertEqual(run(str(src), "-o", str(md)).returncode, 0)
        self.assertIn("| n | w |", md.read_text(encoding="utf-8"))

    def test_docx_to_pdf(self):
        import docx
        src = self.d / "k.docx"
        doc = docx.Document()
        doc.add_heading("Titel", level=1)
        doc.add_paragraph("Inhalt")
        doc.save(str(src))
        out = self.d / "k.pdf"
        p = run(str(src), "-o", str(out))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertTrue(out.read_bytes().startswith(b"%PDF"))

    def test_single_pdf_copy(self):
        from PIL import Image
        img = self.d / "l.png"
        Image.new("RGB", (20, 20), "red").save(img)
        one = self.d / "l.pdf"
        self.assertEqual(run(str(img), "-o", str(one)).returncode, 0)
        copy = self.d / "l-copy.pdf"
        p = run(str(one), "-o", str(copy))
        self.assertEqual(p.returncode, 0, p.stderr)
        from pypdf import PdfReader
        self.assertEqual(len(PdfReader(str(copy)).pages), 1)

    def test_sheet_select(self):
        import openpyxl
        src = self.d / "m.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "Erstes"
        wb.active.append(["x"])
        ws2 = wb.create_sheet("Zweites")
        ws2.append(["n", "w"])
        ws2.append(["a", 1])
        wb.save(str(src))
        out = self.d / "m.md"
        p = run(str(src), "-o", str(out), "--sheet", "Zweites")
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertIn("| n | w |", out.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
