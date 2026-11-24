"""Tests fuer dashboard/format_convert.py (offline, temporaere Dateien)."""
import csv
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import format_convert

ROOT = Path(__file__).resolve().parent
CONV = ROOT / "format_convert.py"


def run(*args):
    p = subprocess.run([sys.executable, str(CONV), *args], text=True,
                       encoding="utf-8", capture_output=True, timeout=60, check=False)
    return p


class FormatConvertTests(unittest.TestCase):
    _directory: Path | None = None

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self._directory = Path(tmp.name)

    @property
    def d(self) -> Path:
        directory = self._directory
        assert directory is not None
        return directory

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
        src.write_text('name,wert\na,1\n"b, inc.",2\n"c\nmehrzeilig",\n', encoding="utf-8")
        xlsx = self.d / "c.xlsx"
        self.assertEqual(run(str(src), "-o", str(xlsx)).returncode, 0)
        back = self.d / "c2.csv"
        self.assertEqual(run(str(xlsx), "-o", str(back)).returncode, 0)
        with (
            src.open(encoding="utf-8", newline="") as original,
            back.open(encoding="utf-8", newline="") as converted,
        ):
            self.assertEqual(list(csv.reader(converted)), list(csv.reader(original)))
        md = self.d / "c.md"
        self.assertEqual(run(str(src), "-o", str(md)).returncode, 0)
        self.assertIn("| name | wert |", md.read_text(encoding="utf-8"))

    def test_json_to_xlsx(self):
        src = self.d / "d.json"
        src.write_text('[{"a": 1, "b": "x"}]', encoding="utf-8")
        out = self.d / "d.xlsx"
        self.assertEqual(run(str(src), "-o", str(out)).returncode, 0)

    def test_xlsx_export_rejects_oversized_text_without_replacing_output(self) -> None:
        oversized = "x" * 32_768
        cases = (
            ("csv", f"{oversized}\nvalue\n", "A1"),
            ("json", json.dumps({"name": "example", "notes": oversized}), "B2"),
        )
        out = self.d / "existing.xlsx"
        _ = out.write_bytes(b"original destination")
        for suffix, content, cell in cases:
            with self.subTest(suffix=suffix, cell=cell):
                src = self.d / f"too-long.{suffix}"
                _ = src.write_text(content, encoding="utf-8")
                result = run(str(src), "-o", str(out), "--overwrite")
                self.assertEqual(result.returncode, 2, result.stderr)
                self.assertIn(cell, result.stderr)
                self.assertIn("32767", result.stderr)
                self.assertEqual(out.read_bytes(), b"original destination")

    def test_xlsx_export_preserves_text_at_the_length_limit(self) -> None:
        value = "x" * 32_766 + "z"
        src = self.d / "limit.csv"
        _ = src.write_text(f"name\n{value}\n", encoding="utf-8")
        xlsx = self.d / "limit.xlsx"
        result = run(str(src), "-o", str(xlsx))
        self.assertEqual(result.returncode, 0, result.stderr)
        back = self.d / "roundtrip.csv"
        result = run(str(xlsx), "-o", str(back))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(back.read_text(encoding="utf-8"), f"name\n{value}\n")

    def test_markdown_keeps_pipes_and_line_breaks_inside_cells(self) -> None:
        src = self.d / "cells.csv"
        _ = src.write_bytes(
            b'"Heading|one",Note\r\n"a\\|b","line 1\r\nline 2\rline 3\n<text> &"\r\n'
        )
        md = self.d / "cells.md"
        result = run(str(src), "-o", str(md))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(
            md.read_text(encoding="utf-8").splitlines(),
            [
                "| Heading&#124;one | Note |",
                "| --- | --- |",
                "| a&#92;&#124;b | line 1<br>line 2<br>line 3<br>&lt;text&gt; &amp; |",
            ],
        )

    def test_plain_text_table_is_unchanged(self) -> None:
        src = self.d / "cells.csv"
        _ = src.write_bytes(b'"Heading|one",Note\n"a|b","line 1\nline 2"\n')
        txt = self.d / "cells.txt"
        result = run(str(src), "-o", str(txt))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(
            txt.read_text(encoding="utf-8"),
            "| Heading|one | Note |\n| --- | --- |\n| a|b | line 1\nline 2 |\n",
        )

    def test_markdown_shows_backticks_literally_inside_a_cell(self) -> None:
        cases = {
            "Prefix `R&D` suffix": "Prefix &#96;R&amp;D&#96; suffix",
            "`a|b` and ``one ` two``": (
                "&#96;a&#124;b&#96; and &#96;&#96;one &#96; two&#96;&#96;"
            ),
            "`line 1\r\nline 2`": "&#96;line 1<br>line 2&#96;",
            # GFM review: an odd number of backslashes escapes only the first
            # backtick, so the rest of the run still opened a code span that
            # displayed the HTML entities verbatim.
            "\\```a&b``": "&#92;&#96;&#96;&#96;a&amp;b&#96;&#96;",
            "x \\```a|b`` y": "x &#92;&#96;&#96;&#96;a&#124;b&#96;&#96; y",
            "\\\\\\```a<b``": "&#92;&#92;&#92;&#96;&#96;&#96;a&lt;b&#96;&#96;",
            "\\\\```a&b``": "&#92;&#92;&#96;&#96;&#96;a&amp;b&#96;&#96;",
            "C:\\Temp \\\\server": "C:\\Temp &#92;\\server",
        }
        for content, expected in cases.items():
            with self.subTest(content=content):
                self.assertEqual(format_convert.markdown_cell(content), expected)

    def test_markdown_cell_never_opens_a_code_span(self) -> None:
        for slashes in range(5):
            for opener in range(1, 4):
                for closer in range(1, 4):
                    content = "\\" * slashes + "`" * opener + "a&b|c" + "`" * closer
                    with self.subTest(content=content):
                        cell = format_convert.markdown_cell(content)
                        self.assertNotIn("`", cell)
                        self.assertNotIn("\\", cell)
                        self.assertNotIn("|", cell)

    def test_markdown_cell_keeps_emphasis_links_and_strikethrough_literal(self) -> None:
        # GFM review (markdown-it-py 4.2.0, table + strikethrough): these
        # rendered as <em>, <strong>, <s>, <a> and <img> instead of text.
        cases = {
            "*a* _b_": "&#42;a&#42; &#95;b&#95;",
            "**www.x.de**": "&#42;&#42;www.x.de&#42;&#42;",
            "~~old~~": "&#126;&#126;old&#126;&#126;",
            "[x](u) ![i](u)": "&#91;x&#93;(u) !&#91;i&#93;(u)",
            "\\*a\\*": "&#92;&#42;a&#92;&#42;",
            "snake_case :smile:": "snake&#95;case :smile:",
        }
        for content, expected in cases.items():
            with self.subTest(content=content):
                self.assertEqual(format_convert.markdown_cell(content), expected)

    def test_markdown_cell_property_inert_and_round_trips(self) -> None:
        # Seeded fuzzing without a renderer: every Markdown-active character
        # is an entity, a remaining backslash escapes nothing, and decoding
        # the entities gives back the cell text.
        import html
        import random
        import re

        alphabet = list("ab 1*_[]`|\\<>&~!():#-") + [
            ":smile:", "\n", "\r\n", "\r", "&amp;", "&#42;", "](u)", "![",
            "**", "__", "~~", "www.x.de", "\U0001F600",
        ]
        rnd = random.Random(20261123)
        for _ in range(9000):
            content = "".join(rnd.choice(alphabet) for _ in range(rnd.randint(0, 10)))
            cell = format_convert.markdown_cell(content)
            for char in "*_~[]`|\n\r":
                self.assertNotIn(char, cell, content)
            self.assertIsNone(re.search(r"\\[!-/:-@\[-`{-~]", cell), content)
            expected = content.replace("\r\n", "\n").replace("\r", "\n")
            self.assertEqual(html.unescape(cell.replace("<br>", "\n")), expected)

    def test_xlsx_export_accepts_brackets_in_the_input_filename(self):
        from openpyxl import load_workbook

        src = self.d / "Report [2026].csv"
        _ = src.write_text("Name;Betrag\nAda;12,50\n", encoding="utf-8")
        out = self.d / "report.xlsx"
        result = run(str(src), "-o", str(out))
        self.assertEqual(result.returncode, 0, result.stderr)
        book = load_workbook(out, read_only=True)
        try:
            self.assertEqual(book.sheetnames, ["Report _2026_"])
            self.assertEqual(
                list(book["Report _2026_"].values),
                [("Name", "Betrag"), ("Ada", "12,50")],
            )
        finally:
            book.close()

    def test_xlsx_keeps_imported_strings_literal_and_json_numbers_numeric(self):
        from openpyxl import load_workbook

        for suffix, content, number in (
            ("csv", '=1+2,Value\n=3+4,0042\n#N/A,42\n', "0042"),
            ("json", '[{"=1+2":"=3+4","Value":42},{"=1+2":"#N/A","Value":42}]', 42),
        ):
            with self.subTest(suffix=suffix):
                src = self.d / f"table.{suffix}"
                src.write_text(content, encoding="utf-8")
                out = self.d / f"{suffix}.xlsx"
                result = run(str(src), "-o", str(out))
                self.assertEqual(result.returncode, 0, result.stderr)
                book = load_workbook(out, data_only=False)
                try:
                    sheet = book["table"]
                    for position, value in (("A1", "=1+2"), ("A2", "=3+4"), ("A3", "#N/A")):
                        self.assertEqual(sheet[position].data_type, "s")
                        self.assertEqual(sheet[position].value, value)
                    self.assertEqual(sheet["B2"].value, number)
                finally:
                    book.close()

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

    def test_target_created_during_conversion_is_preserved(self):
        out = self.d / "result.txt"

        def write_conversion(_inputs, temporary_output, _sheet):
            temporary_output.write_bytes(b"converted")
            out.write_bytes(b"concurrent")

        with (
            patch.object(format_convert, "_write_conversion", write_conversion),
            self.assertRaises(SystemExit) as error,
        ):
            format_convert.convert([], out)
        self.assertEqual(error.exception.code, 2)
        self.assertEqual(out.read_bytes(), b"concurrent")
        self.assertEqual(list(self.d.iterdir()), [out])

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

    def test_html_and_htm_to_docx(self):
        from docx import Document

        listing = run("--list")
        self.assertEqual(listing.returncode, 0, listing.stderr)
        matrix = json.loads(listing.stdout)["inputs"]
        for suffix in (".html", ".htm"):
            with self.subTest(suffix=suffix):
                self.assertIn(".docx", matrix[suffix])
                src = self.d / f"input{suffix}"
                src.write_text("<h1>Übersicht</h1><p>A &amp; B</p>", encoding="utf-8")
                out = self.d / f"{suffix[1:]}.docx"
                result = run(str(src), "-o", str(out))
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(
                    [paragraph.text for paragraph in Document(str(out)).paragraphs],
                    ["Übersicht", "A & B"],
                )

    def test_xlsx_to_pdf_and_md(self):
        import openpyxl
        src = self.d / "j.xlsx"
        wb = openpyxl.Workbook()
        sheet = wb.active
        assert sheet is not None
        sheet.append(["n", "w"])
        sheet.append(["a", 1])
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

    def test_sheet_on_non_xlsx_does_not_write_output(self) -> None:
        src = self.d / "source.csv"
        _ = src.write_text("name,score\nAda,10\n", encoding="utf-8")
        out = self.d / "result.xlsx"
        result = run(str(src), "-o", str(out), "--sheet", "Report")
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertIn("--sheet", result.stderr)
        self.assertFalse(out.exists())

    def test_sheet_select(self):
        import openpyxl
        src = self.d / "m.xlsx"
        wb = openpyxl.Workbook()
        sheet = wb.active
        assert sheet is not None
        sheet.title = "Erstes"
        sheet.append(["x"])
        ws2 = wb.create_sheet("Zweites")
        ws2.append(["n", "w"])
        ws2.append(["a", 1])
        wb.save(str(src))
        out = self.d / "m.md"
        p = run(str(src), "-o", str(out), "--sheet", "Zweites")
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertIn("| n | w |", out.read_text(encoding="utf-8"))

    def test_list_sheets_and_export_selected_sheet(self):
        import openpyxl
        from openpyxl.chart import BarChart, Reference

        src = self.d / "sheets.xlsx"
        wb = openpyxl.Workbook()
        overview = wb.active
        assert overview is not None
        overview.title = "Übersicht"
        detail = wb.create_sheet("Daten 2026")
        detail.append(["Name", "Notiz"])
        detail.append(["Ada", "Komma, und\nZeilenumbruch"])
        detail.append([None, 12.5])
        chart = BarChart()
        chart.add_data(Reference(detail, min_col=2, min_row=3, max_row=3))
        wb.create_chartsheet("Diagramm").add_chart(chart)
        wb.save(src)

        listing = run(str(src), "--list-sheets")
        self.assertEqual(listing.returncode, 0, listing.stderr)
        self.assertEqual(json.loads(listing.stdout), ["Übersicht", "Daten 2026"])
        out = self.d / "sheets.csv"
        export = run(str(src), "-o", str(out), "--sheet", "Daten 2026")
        self.assertEqual(export.returncode, 0, export.stderr)
        with out.open(encoding="utf-8", newline="") as stream:
            self.assertEqual(list(csv.reader(stream)), [
                ["Name", "Notiz"], ["Ada", "Komma, und\nZeilenumbruch"], ["", "12.5"],
            ])

    def test_missing_sheet_preserves_existing_destination(self):
        import openpyxl
        src = self.d / "sheets.xlsx"
        openpyxl.Workbook().save(src)
        out = self.d / "sheets.csv"
        out.write_text("original", encoding="utf-8")
        result = run(str(src), "-o", str(out), "--overwrite", "--sheet", "Fehlt")
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertIn("Fehlt", result.stderr)
        self.assertNotIn("Traceback", result.stderr)
        self.assertEqual(out.read_text(encoding="utf-8"), "original")

    # German Excel "CSV (Trennzeichen-getrennt)": ';' as separator, ',' as decimal mark, often Windows-1252.
    def test_csv_excel_de_semicolon_utf8(self):
        src = self.d / "de.csv"
        src.write_bytes("﻿Name;Betrag;Ort\nMüller;12,50;Köln\n".encode())
        out = self.d / "de.md"
        p = run(str(src), "-o", str(out))
        self.assertEqual(p.returncode, 0, p.stderr)
        md = out.read_text(encoding="utf-8")
        self.assertIn("| Name | Betrag | Ort |", md)
        self.assertIn("| Müller | 12,50 | Köln |", md)

    def test_csv_excel_de_cp1252(self):
        src = self.d / "ansi.csv"
        src.write_bytes("Name;Ort\r\nSchäfer;Gießen\r\n".encode("cp1252"))
        out = self.d / "ansi.md"
        p = run(str(src), "-o", str(out))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertIn("| Schäfer | Gießen |", out.read_text(encoding="utf-8"))

    def test_csv_utf16_refused_cleanly(self):
        src = self.d / "u16.csv"
        src.write_bytes("Name;Ort\r\nA;B\r\n".encode("utf-16"))
        p = run(str(src), "-o", str(self.d / "u16.md"))
        self.assertEqual(p.returncode, 2)
        self.assertIn("FEHLER", p.stderr)
        self.assertNotIn("Traceback", p.stderr)
        self.assertFalse((self.d / "u16.md").exists())

    def test_csv_comma_with_quoted_semicolon_stays_comma(self):
        src = self.d / "q.csv"
        src.write_text('"Name; Vorname","Notiz; kurz; lang",Ort\n"Do; Trung","x; y",Berlin\n', encoding="utf-8")
        out = self.d / "q.md"
        p = run(str(src), "-o", str(out))
        self.assertEqual(p.returncode, 0, p.stderr)
        md = out.read_text(encoding="utf-8")
        self.assertIn("| Name; Vorname | Notiz; kurz; lang | Ort |", md)
        self.assertIn("| Do; Trung | x; y | Berlin |", md)

    def test_csv_delimiter_uses_first_logical_record(self):
        from format_convert import read_table

        cases = [
            ('"Name\ncontinued";Amount\nAlice;12,50\n', ["Name\ncontinued", "Amount"], ["Alice", "12,50"]),
            ('"Name\r\ncontinued"\tAmount\r\nAlice\t12,50\r\n', ["Name\r\ncontinued", "Amount"], ["Alice", "12,50"]),
            ('Name;Amount\rAlice;12,50, EUR\r', ["Name", "Amount"], ["Alice", "12,50, EUR"]),
        ]
        for contents, header, row in cases:
            with self.subTest(contents=contents):
                src = self.d / "records.csv"
                src.write_bytes(contents.encode("utf-8"))
                self.assertEqual(read_table(src, ".csv"), (header, [row]))


    def test_csv_ignores_blank_preamble_but_preserves_later_empty_rows(self) -> None:
        src = self.d / "preamble.csv"
        _ = src.write_bytes(
            b'\xef\xbb\xbf\r\n\n"Name\r\ncontinued";Amount\r\n\r\nAda;12,50\r\n'
        )
        self.assertEqual(
            format_convert.read_table(src, ".csv"),
            (["Name\r\ncontinued", "Amount"], [[], ["Ada", "12,50"]]),
        )

    def test_csv_preserves_quoted_empty_record_and_handles_only_blank_lines(self) -> None:
        src = self.d / "empty.csv"
        _ = src.write_bytes(b'\r\n""\r\nvalue\r\n')
        self.assertEqual(format_convert.read_table(src, ".csv"), ([""], [["value"]]))
        _ = src.write_bytes(b'\r\n\n\r')
        self.assertEqual(format_convert.read_table(src, ".csv"), ([], []))

    def test_csv_quote_only_opens_at_start_of_field(self):
        from format_convert import read_table

        cases = [
            ('Display "name;City\nAlice;Berlin\n', ['Display "name', "City"]),
            ('Id;"Label ""with"" commas, a, b, c";City\r\n1;Plain;Berlin\r\n',
             ["Id", 'Label "with" commas, a, b, c', "City"]),
        ]
        for contents, header in cases:
            with self.subTest(contents=contents):
                src = self.d / "quoted-header.csv"
                src.write_bytes(contents.encode("utf-8"))
                self.assertEqual(read_table(src, ".csv")[0], header)

    def test_csv_single_column_keeps_semicolons_in_data(self):
        src = self.d / "one.csv"
        src.write_text("Notiz\nerst a; dann b\n", encoding="utf-8")
        out = self.d / "one.md"
        p = run(str(src), "-o", str(out))
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertIn("| erst a; dann b |", out.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
