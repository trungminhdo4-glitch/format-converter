"""Universeller Datei-Formatkonverter (offline, fail-closed).

Verwendung durch Agenten und Dashboard:
  python dashboard/format_convert.py EINGABE [-o AUSGABE] [--overwrite] [--sheet BLATT]
  python dashboard/format_convert.py --list
  python dashboard/format_convert.py A.pdf B.pdf -o GESAMT.pdf   (PDF-Merge)

Unterstuetzte Matrix (--list zeigt sie als JSON):
  .md/.txt/.html -> .pdf / .docx / .txt / .md
  .csv/.json     -> .xlsx / .md / .txt
  .xlsx          -> .csv / .md / .txt / .pdf
  .docx          -> .txt / .md / .pdf
  .png/.jpg/.jpeg-> .pdf (ein Bild pro Seite)
  .pdf + .pdf    -> .pdf (Merge, pypdf)

Absichtlich NICHT enthalten: Audio/Video/Archive/Ausfuehrbares/Office-Makros.
Unbekannte Endungen und fehlende Abhaengigkeiten brechen mit klarer
Fehlermeldung ab (Exit != 0), statt Teilergebnisse zu liefern.
Keine Netzwerkzugriffe, keine Secrets, atomares Schreiben.
"""
import argparse
import csv
import io
import json
import os
import re
import sys
import tempfile
from contextlib import closing, contextmanager
from html import escape
from html.parser import HTMLParser
from itertools import chain
from pathlib import Path
from typing import NoReturn
from zipfile import BadZipFile

DEPS = {
    "pdf": ("reportlab", "reportlab"),
    "xlsx": ("openpyxl", "openpyxl"),
    "docx": ("python-docx", "docx"),
    "image": ("Pillow", "PIL"),
    "pdfmerge": ("pypdf", "pypdf"),
}

MATRIX = {
    ".md": [".pdf", ".docx", ".txt", ".md"],
    ".txt": [".pdf", ".docx", ".txt", ".md"],
    ".html": [".pdf", ".docx", ".txt", ".md"],
    ".htm": [".pdf", ".docx", ".txt", ".md"],
    ".csv": [".xlsx", ".md", ".txt"],
    ".json": [".xlsx", ".md", ".txt"],
    ".xlsx": [".csv", ".md", ".txt", ".pdf"],
    ".docx": [".txt", ".md", ".pdf"],
    ".png": [".pdf"],
    ".jpg": [".pdf"],
    ".jpeg": [".pdf"],
    ".pdf": [".pdf"],
}

REFUSED = {".exe", ".msi", ".bat", ".ps1", ".zip", ".7z", ".rar",
           ".mp3", ".mp4", ".avi", ".mkv", ".wav", ".mov"}


XLSX_TITLE_REPLACEMENTS = str.maketrans(dict.fromkeys("[]:*?/\\", "_"))


def fail(msg: str) -> NoReturn:
    print(f"FEHLER: {msg}", file=sys.stderr)
    raise SystemExit(2)


def need(mod):
    try:
        return __import__(mod)
    except ImportError:
        pkg = next((p for p, m in DEPS.items() if m[1] == mod), mod)
        fail(f"Abhaengigkeit fehlt: {pkg} (pip install {pkg})")


class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        text = data.strip()
        if text:
            self.parts.append(text)


def read_text_input(path, suffix):
    if suffix in (".md", ".txt"):
        return path.read_text(encoding="utf-8")
    if suffix in (".html", ".htm"):
        parser = _TextExtractor()
        parser.feed(path.read_text(encoding="utf-8"))
        return "\n".join(parser.parts)
    if suffix == ".docx":
        docx = need("docx")
        doc = docx.Document(str(path))
        heads = []
        for p in doc.paragraphs:
            t = p.text.strip()
            if t:
                heads.append(("# " if p.style.name.startswith("Heading 1") else "") + t)
        return "\n\n".join(heads)
    fail(f"kein Text-Import fuer {suffix}")


def read_csv_text(path: Path) -> str:
    """UTF-8 (with/without BOM), else Windows-1252 (German Excel 'CSV (Trennzeichen-getrennt)').
    NUL bytes mean UTF-16 ('Unicode-Text') or binary: refused instead of guessed."""
    raw = path.read_bytes()
    if b"\x00" in raw:
        fail(f"{path.name}: keine UTF-8/ANSI-CSV (UTF-16 oder binaer?) - in Excel als 'CSV UTF-8' speichern")
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    fail(f"{path.name}: Zeichenkodierung weder UTF-8 noch Windows-1252")


def csv_delimiter(text):
    """Separator of the first CSV record: the most frequent of , ; TAB outside double quotes (tie -> ',')."""
    counts = {",": 0, ";": 0, "\t": 0}
    quoted = False
    at_field_start = True
    characters = iter(enumerate(text))
    for index, ch in characters:
        if quoted:
            if ch == '"':
                if text[index + 1:index + 2] == '"':
                    next(characters)
                else:
                    quoted = False
            continue
        if ch in "\r\n":
            break
        if ch == '"' and at_field_start:
            quoted = True
        at_field_start = ch in counts
        if at_field_start:
            counts[ch] += 1
    return max(counts, key=lambda c: (counts[c], c == ","))


def open_workbook(path: Path):
    oxl = need("openpyxl")
    from openpyxl.utils.exceptions import InvalidFileException

    try:
        return closing(oxl.load_workbook(str(path), read_only=True, data_only=True))
    except (OSError, BadZipFile, InvalidFileException) as exc:
        fail(f"XLSX kann nicht geöffnet werden: {path} ({exc})")


@contextmanager
def worksheet_rows(path: Path, sheet: str | None):
    with open_workbook(path) as workbook:
        worksheet = workbook.active
        if sheet:
            if sheet not in [ws.title for ws in workbook.worksheets]:
                fail(f"Tabellenblatt fehlt: {sheet!r} (--list-sheets zeigt die Namen)")
            worksheet = workbook[sheet]
        if worksheet not in workbook.worksheets:
            fail("Kein aktives Tabellenblatt (--list-sheets zeigt die Namen)")
        yield worksheet.iter_rows(values_only=True)


def read_table(path: Path, suffix, sheet=None):
    if suffix == ".csv":
        # Drop only the empty physical lines before the first record. Later
        # blank records and line breaks inside quoted fields are data.
        text = read_csv_text(path).lstrip("\r\n")
        rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=csv_delimiter(text)))
        return rows[0] if rows else [], rows[1:]
    if suffix == ".json":
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except UnicodeDecodeError:
            fail(f"{path.name}: JSON-Eingabe muss UTF-8 sein")
        except json.JSONDecodeError as exc:
            fail(f"{path.name}: fehlerhaftes JSON (Zeile {exc.lineno}, Spalte {exc.colno})")
        items = data if isinstance(data, list) else [data]
        keys = []
        for item in items:
            if isinstance(item, dict):
                for k in item:
                    if k not in keys:
                        keys.append(k)
        return keys, [[r.get(k, "") if isinstance(r, dict) else r for k in keys] for r in items]
    if suffix == ".xlsx":
        with worksheet_rows(path, sheet) as source_rows:
            rows = [
                [value if value is not None else "" for value in row]
                for row in source_rows
            ]
        return (rows[0] if rows else []), rows[1:]
    fail(f"kein Tabellen-Import fuer {suffix}")


# A backslash only escapes ASCII punctuation (CommonMark); "\n" becomes "<br>".
_ESCAPING_BACKSLASH = re.compile(r"\\(?=[!-/:-@\[-`{-~\n])")
# Characters that open inline Markdown (code, emphasis, strikethrough, links).
_INLINE_ENTITIES = str.maketrans({
    "|": "&#124;", "`": "&#96;", "*": "&#42;", "_": "&#95;",
    "~": "&#126;", "[": "&#91;", "]": "&#93;", "\n": "<br>",
})


def markdown_cell(value: object) -> str:
    """Render a cell value as literal text inside a GFM table cell.

    Inline delimiters become entities, so no code span, emphasis,
    strikethrough or link can open; a backslash that would escape the next
    character becomes an entity, so it stays visible.
    """
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return _ESCAPING_BACKSLASH.sub("&#92;", escape(text, quote=False)).translate(
        _INLINE_ENTITIES
    )


def table_to_markdown(header, rows, *, escape_cells: bool = True):
    cell_text = markdown_cell if escape_cells else str
    out = [
        "| " + " | ".join(map(cell_text, header)) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    out += ["| " + " | ".join(map(cell_text, r)) + " |" for r in rows]
    return "\n".join(out) + "\n"


def write_pdf_text(target, title, text):
    rl = need("reportlab")
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(target), pagesize=A4)
    c.setTitle(title[:100])
    _, height = A4
    y = height - 20 * rl.lib.units.mm
    c.setFont("Helvetica", 10)
    for raw in text.splitlines() or [""]:
        line = raw.encode("latin-1", "replace").decode("latin-1")
        while len(line) > 95:
            if y < 25 * rl.lib.units.mm:
                c.showPage()
                c.setFont("Helvetica", 10)
                y = height - 20 * rl.lib.units.mm
            c.drawString(20 * rl.lib.units.mm, y, line[:95])
            y -= 5 * rl.lib.units.mm
            line = line[95:]
        if y < 25 * rl.lib.units.mm:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = height - 20 * rl.lib.units.mm
        if line.startswith("# "):
            c.setFont("Helvetica-Bold", 14)
            c.drawString(20 * rl.lib.units.mm, y, line[2:110])
            c.setFont("Helvetica", 10)
        else:
            c.drawString(20 * rl.lib.units.mm, y, line)
        y -= 5 * rl.lib.units.mm
    c.save()


def write_pdf_table(target, title, header, rows):
    rl = need("reportlab")
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas
    page = landscape(A4)
    c = canvas.Canvas(str(target), pagesize=page)
    c.setTitle(title[:100])
    width, height = page
    y = height - 15 * rl.lib.units.mm
    c.setFont("Helvetica-Bold", 11)
    c.drawString(15 * rl.lib.units.mm, y, title[:110].encode("latin-1", "replace").decode("latin-1"))
    y -= 7 * rl.lib.units.mm
    c.setFont("Helvetica", 8)
    cols = max(len(header), 1)
    col_w = (width - 30 * rl.lib.units.mm) / cols
    for row in [header] + list(rows):
        if y < 15 * rl.lib.units.mm:
            c.showPage()
            c.setFont("Helvetica", 8)
            y = height - 15 * rl.lib.units.mm
        for i, cell in enumerate(list(row) + [""] * (cols - len(row))):
            c.drawString(15 * rl.lib.units.mm + i * col_w, y,
                         str(cell)[:40].encode("latin-1", "replace").decode("latin-1"))
        y -= 4.5 * rl.lib.units.mm
    c.save()


def write_pdf_images(target, title, images):
    need("PIL")
    need("reportlab")
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(target), pagesize=A4)
    c.setTitle(title[:100])
    width, height = A4
    for img_path in images:
        img = ImageReader(str(img_path))
        iw, ih = img.getSize()
        scale = min(width / iw, height / ih)
        w, h = iw * scale, ih * scale
        c.drawImage(img, (width - w) / 2, (height - h) / 2, w, h, preserveAspectRatio=True)
        c.showPage()
    c.save()


def write_docx(target, text):
    docx = need("docx")
    doc = docx.Document()
    for raw in text.splitlines():
        s = raw.strip()
        if s.startswith("# "):
            doc.add_heading(s[2:], level=1)
        elif s.startswith("## "):
            doc.add_heading(s[3:], level=2)
        elif s:
            doc.add_paragraph(raw)
    doc.save(str(target))


def convert(inputs, output, overwrite=False, sheet=None):
    out = Path(output)
    if out.exists() and not overwrite:
        fail(f"Ziel existiert (ohne --overwrite kein Ueberschreiben): {out}")
    with tempfile.TemporaryDirectory(dir=out.parent) as directory:
        tmp = Path(directory) / out.name
        _write_conversion(inputs, tmp, sheet)
        try:
            if overwrite:
                tmp.replace(out)
            elif os.name == "nt":
                tmp.rename(out)
            else:
                out.hardlink_to(tmp)
        except FileExistsError:
            fail(f"Ziel existiert (ohne --overwrite kein Ueberschreiben): {out}")
    return str(out)


def _write_conversion(inputs, tmp: Path, sheet):
    srcs = [Path(p) for p in inputs]
    for s in srcs:
        if not s.is_file():
            fail(f"Eingabe fehlt: {s}")
        if s.suffix.lower() in REFUSED:
            fail(f"Typ ausgenommen (kein Audio/Video/Archiv/Ausfuehrbar): {s.suffix}")
        if s.suffix.lower() not in MATRIX:
            fail(f"Quelltyp nicht unterstuetzt: {s.suffix} (--list zeigt die Matrix)")
    in_suf = [s.suffix.lower() for s in srcs]
    if sheet is not None and in_suf != [".xlsx"]:
        fail("--sheet erwartet genau eine XLSX-Eingabedatei")
    out_suf = tmp.suffix.lower()
    if out_suf == ".pdf" and all(s == ".pdf" for s in in_suf):
        pypdf = need("pypdf")
        writer = pypdf.PdfWriter()
        for s in srcs:
            for page in pypdf.PdfReader(str(s)).pages:
                writer.add_page(page)
        with tmp.open("wb") as f:
            writer.write(f)
        return
    if len(srcs) > 1:
        fail("Mehrere Eingaben nur als PDF-Merge (.pdf -> .pdf)")
    src, suf = srcs[0], in_suf[0]
    if out_suf not in MATRIX[suf]:
        fail(f"{suf} -> {out_suf} nicht unterstuetzt (--list zeigt die Matrix)")
    if out_suf == ".pdf":
        if suf in (".png", ".jpg", ".jpeg"):
            write_pdf_images(tmp, src.stem, [src])
        elif suf in (".csv", ".json", ".xlsx"):
            header, rows = read_table(src, suf, sheet)
            write_pdf_table(tmp, src.stem, [str(h) for h in header], rows)
        else:
            write_pdf_text(tmp, src.stem, read_text_input(src, suf))
    elif out_suf == ".xlsx":
        oxl = need("openpyxl")
        header, rows = read_table(src, suf, sheet)
        wb = oxl.Workbook()
        ws = wb.active
        ws.title = src.stem.translate(XLSX_TITLE_REPLACEMENTS)[:31]
        for row_number, row in enumerate(chain(([str(h) for h in header],), rows), 1):
            for column, value in enumerate(row, 1):
                cell = ws.cell(row=row_number, column=column, value=value)
                if isinstance(value, str):
                    cell.data_type = "s"
        wb.save(str(tmp))
    elif out_suf == ".docx":
        write_docx(tmp, read_text_input(src, suf))
    elif out_suf == ".csv":
        with (
            worksheet_rows(src, sheet) as rows,
            tmp.open("w", encoding="utf-8", newline="") as stream,
        ):
            csv.writer(stream).writerows(rows)
    else:
        if suf in (".csv", ".json", ".xlsx"):
            header, rows = read_table(src, suf, sheet)
            _ = tmp.write_text(
                table_to_markdown(header, rows, escape_cells=out_suf == ".md"),
                encoding="utf-8",
            )
        else:
            _ = tmp.write_text(read_text_input(src, suf), encoding="utf-8")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Offline Formatkonverter (fail-closed).")
    ap.add_argument("inputs", nargs="*", help="Eingabedatei(en)")
    ap.add_argument("-o", "--output", help="Zieldatei")
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--sheet", help="XLSX-Blattname (Standard: aktives Blatt)")
    modes = ap.add_mutually_exclusive_group()
    modes.add_argument("--list", action="store_true", help="Matrix als JSON zeigen")
    modes.add_argument(
        "--list-sheets", action="store_true", help="XLSX-Tabellenblätter als JSON zeigen"
    )
    args = ap.parse_args(argv)
    if args.list:
        print(json.dumps({"inputs": MATRIX, "refused": sorted(REFUSED),
                          "deps": {k: v[0] for k, v in DEPS.items()}},
                         indent=2, sort_keys=True))
        return 0
    if args.list_sheets:
        if len(args.inputs) != 1 or args.output is not None or args.sheet is not None or args.overwrite:
            ap.error("--list-sheets erwartet genau eine XLSX-Datei ohne Konvertierungsoptionen")
        source = Path(args.inputs[0])
        if source.suffix.lower() != ".xlsx":
            ap.error("--list-sheets erwartet eine XLSX-Datei")
        with open_workbook(source) as workbook:
            print(json.dumps([ws.title for ws in workbook.worksheets]))
        return 0
    if not args.inputs or not args.output:
        fail("EINGABE und -o AUSGABE angeben (oder --list)")
    print(convert(args.inputs, args.output, args.overwrite, args.sheet))
    return 0


if __name__ == "__main__":
    main()
