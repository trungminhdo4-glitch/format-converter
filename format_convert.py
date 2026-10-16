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
import json
import sys
from html.parser import HTMLParser
from pathlib import Path

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
    ".html": [".pdf", ".txt", ".md"],
    ".htm": [".pdf", ".txt", ".md"],
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


def fail(msg):
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


def read_table(path, suffix, sheet=None):
    if suffix == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        return rows[0] if rows else [], rows[1:]
    if suffix == ".json":
        data = json.loads(path.read_text(encoding="utf-8"))
        items = data if isinstance(data, list) else [data]
        keys = []
        for item in items:
            if isinstance(item, dict):
                for k in item:
                    if k not in keys:
                        keys.append(k)
        return keys, [[r.get(k, "") if isinstance(r, dict) else r for k in keys] for r in items]
    if suffix == ".xlsx":
        oxl = need("openpyxl")
        wb = oxl.load_workbook(str(path), read_only=True, data_only=True)
        ws = wb[sheet] if sheet else wb.active
        rows = [[c.value if c.value is not None else "" for c in row] for row in ws.rows]
        return (rows[0] if rows else []), rows[1:]
    fail(f"kein Tabellen-Import fuer {suffix}")


def table_to_markdown(header, rows):
    out = ["| " + " | ".join(map(str, header)) + " |",
           "| " + " | ".join("---" for _ in header) + " |"]
    out += ["| " + " | ".join(map(str, r)) + " |" for r in rows]
    return "\n".join(out) + "\n"


def write_pdf_text(target, title, text):
    rl = need("reportlab")
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(target), pagesize=A4)
    c.setTitle(title[:100])
    width, height = A4
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
    rl = need("reportlab")
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
    srcs = [Path(p) for p in inputs]
    for s in srcs:
        if not s.is_file():
            fail(f"Eingabe fehlt: {s}")
        if s.suffix.lower() in REFUSED:
            fail(f"Typ ausgenommen (kein Audio/Video/Archiv/Ausfuehrbar): {s.suffix}")
        if s.suffix.lower() not in MATRIX:
            fail(f"Quelltyp nicht unterstuetzt: {s.suffix} (--list zeigt die Matrix)")
    in_suf = [s.suffix.lower() for s in srcs]
    out_suf = out.suffix.lower()
    if out_suf == ".pdf" and all(s == ".pdf" for s in in_suf):
        pypdf = need("pypdf")
        writer = pypdf.PdfWriter()
        for s in srcs:
            for page in pypdf.PdfReader(str(s)).pages:
                writer.add_page(page)
        tmp = out.with_suffix(".tmp")
        with tmp.open("wb") as f:
            writer.write(f)
        tmp.replace(out)
        return str(out)
    if len(srcs) > 1:
        fail("Mehrere Eingaben nur als PDF-Merge (.pdf -> .pdf)")
    src, suf = srcs[0], in_suf[0]
    if out_suf not in MATRIX[suf]:
        fail(f"{suf} -> {out_suf} nicht unterstuetzt (--list zeigt die Matrix)")
    tmp = out.with_suffix(".tmp")
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
        ws.title = src.stem[:31]
        ws.append([str(h) for h in header])
        for r in rows:
            ws.append(list(r))
        wb.save(str(tmp))
    elif out_suf == ".docx":
        write_docx(tmp, read_text_input(src, suf))
    else:
        if suf in (".csv", ".json", ".xlsx"):
            header, rows = read_table(src, suf, sheet)
            tmp.write_text(table_to_markdown([str(h) for h in header], rows), encoding="utf-8")
        else:
            tmp.write_text(read_text_input(src, suf), encoding="utf-8")
    tmp.replace(out)
    return str(out)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Offline Formatkonverter (fail-closed).")
    ap.add_argument("inputs", nargs="*", help="Eingabedatei(en)")
    ap.add_argument("-o", "--output", help="Zieldatei")
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--sheet", help="XLSX-Blattname (Standard: aktives Blatt)")
    ap.add_argument("--list", action="store_true", help="Matrix als JSON zeigen")
    args = ap.parse_args(argv)
    if args.list:
        print(json.dumps({"inputs": MATRIX, "refused": sorted(REFUSED),
                          "deps": {k: v[0] for k, v in DEPS.items()}},
                         indent=2, sort_keys=True))
        return 0
    if not args.inputs or not args.output:
        fail("EINGABE und -o AUSGABE angeben (oder --list)")
    print(convert(args.inputs, args.output, args.overwrite, args.sheet))
    return 0


if __name__ == "__main__":
    main()
