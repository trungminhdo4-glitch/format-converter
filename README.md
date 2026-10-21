# format-converter

Offline, fail-closed file format converter. One CLI for the conversions
school and office work actually needs — nothing more.

```bash
python format_convert.py EINGABE -o AUSGABE [--overwrite] [--sheet BLATT]
python format_convert.py --list   # supported matrix as JSON
```

## Matrix

| Input | Output |
|---|---|
| `.md` / `.txt` / `.html` | `.pdf`, `.docx`, `.txt`, `.md` |
| `.csv` / `.json` | `.xlsx`, `.md`, `.txt` |
| `.xlsx` | `.csv`, `.md`, `.txt`, `.pdf` (`--sheet` selects a sheet, default: active) |
| `.docx` | `.txt`, `.md`, `.pdf` (text + headings; complex layout is not preserved) |
| `.png` / `.jpg` | `.pdf` (one image per page) |
| `.pdf` + `.pdf` | `.pdf` (merge; single `.pdf` → `.pdf` copies) |

CSV input: UTF-8 (with or without BOM) or Windows-1252, as German Excel saves
"CSV (Trennzeichen-getrennt)"; the separator (`,` `;` or TAB) is the most frequent one in
the header line outside quotes (tie: `,`). UTF-16 CSV ("Unicode-Text") is refused with
exit code 2 — save as "CSV UTF-8". Cell values stay text (`12,50` is not converted).

Deliberately refused with exit code 2: audio, video, archives,
executables, and any unknown extension. Existing output files are never
overwritten without `--overwrite`. Writes are atomic (temp + replace).

## Requirements

Measured versions: Python 3.12, reportlab 5.0.1, openpyxl 3.1.5,
python-docx 1.2.0, Pillow 12.3.0, pypdf 6.20.1. No network access.

## Tests

```bash
python -m unittest discover -s . -p "test_*.py"   # 20 tests, offline
```

## License

MIT — see LICENSE.
