# format-converter

Offline, fail-closed file format converter. One CLI for the conversions
school and office work actually needs — nothing more.

```bash
python format_convert.py EINGABE -o AUSGABE [--overwrite] [--sheet BLATT]
python format_convert.py --list   # supported matrix as JSON
python format_convert.py Mappe.xlsx --list-sheets
python format_convert.py Mappe.xlsx -o Daten.csv --sheet "Daten 2026"
```

## Matrix

| Input | Output |
|---|---|
| `.md` / `.txt` | `.pdf`, `.docx`, `.txt`, `.md` |
| `.html` / `.htm` | `.pdf`, `.docx`, `.txt`, `.md` (text extraction) |
| `.csv` / `.json` | `.xlsx`, `.md`, `.txt` |
| `.xlsx` | `.csv`, `.md`, `.txt`, `.pdf` (`--sheet` selects a sheet, default: active) |
| `.docx` | `.txt`, `.md`, `.pdf` (text + headings; complex layout is not preserved) |
| `.png` / `.jpg` | `.pdf` (one image per page) |
| `.pdf` + `.pdf` | `.pdf` (merge; single `.pdf` → `.pdf` copies) |

CSV input: UTF-8 (with or without BOM) or Windows-1252, as German Excel saves
"CSV (Trennzeichen-getrennt)"; the separator (`,` `;` or TAB) is the most frequent one
outside quotes in the first record after any empty leading lines (tie: `,`).
UTF-16 CSV ("Unicode-Text") is refused with
exit code 2 — save as "CSV UTF-8". Cell values stay text (`12,50` is not converted).
In XLSX output, strings from CSV and JSON are stored as literal text, including
values such as `=1+2` and `#N/A`. JSON numbers remain numeric.

`--list-sheets` prints a JSON array of worksheet names in workbook order.
Pass a name to `--sheet` to select that worksheet. XLSX-to-CSV export writes
rows as it reads them, so it can export large worksheets without holding all
their rows in memory.

Markdown tables keep literal `|` characters and embedded line breaks inside
their original cells using HTML entities and `<br>`; `.txt` export is unchanged.

Deliberately refused with exit code 2: audio, video, archives,
executables, and any unknown extension. Existing output files are never
overwritten without `--overwrite`. Writes are atomic (temp + replace).

## Requirements

Measured versions: Python 3.12, reportlab 5.0.1, openpyxl 3.1.5,
python-docx 1.2.0, Pillow 12.3.0, pypdf 6.20.1. No network access.

## Tests

```bash
python -m unittest discover -s . -p "test_*.py"   # offline
```

## License

MIT — see LICENSE.
