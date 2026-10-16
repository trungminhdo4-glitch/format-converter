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

Deliberately refused with exit code 2: audio, video, archives,
executables, and any unknown extension. Existing output files are never
overwritten without `--overwrite`. Writes are atomic (temp + replace).

## Requirements

Measured versions: Python 3.12, reportlab 5.0.1, openpyxl 3.1.5,
python-docx 1.2.0, Pillow 12.3.0, pypdf 6.20.1. No network access.

## Tests

```bash
python -m unittest discover -s . -p "test_*.py"   # 15 tests, offline
```

## License

MIT — see LICENSE.
