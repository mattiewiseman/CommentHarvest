# Word Comments ➜ Excel (macOS-friendly)

Extract the **commented (highlighted) text** and its **associated Word comment** from a `.docx` file and export them to an Excel spreadsheet with two columns:

1. **Commented Text** (the text range the comment is attached to — what Word visually highlights for the comment)
2. **Comment** (the reviewer’s comment text)

> ✅ Works on macOS (and Windows/Linux).  
> ✅ Supports modern `.docx` files created by Microsoft Word.  
> ⚠️ Does **not** support legacy `.doc` files.

---

## Quick Start

```bash
# 1) Create and activate a virtual environment (macOS / zsh)
python3 -m venv .venv
source .venv/bin/activate

# 2) Install dependencies
pip install -r requirements.txt

# 3) Run
python -m src.extract_docx_comments your_file.docx -o output.xlsx
```

This will create `output.xlsx` with two columns: **Commented Text** and **Comment**.

---

## How it works

- A `.docx` is a zip archive. We read:
  - `word/document.xml` for the **commented text ranges** (between `w:commentRangeStart` and `w:commentRangeEnd`).
  - `word/comments.xml` for the **comment bodies** (paragraph runs inside each comment node).
- We match ranges by `w:id` and export pairs.

### Notes & Edge Cases

- **Tracked changes**: This script grabs what’s in the document stream; deleted/accepted text behavior depends on the saved state of the doc. If something looks off, accept/reject changes in Word first.
- **Images/shapes** inside ranges are ignored (only text is exported).
- **Multiple comments per paragraph** work fine. Nested comment ranges are rare in Word; if present, text will be attributed to all open ranges.
- If your file has **protected** or **encrypted** content, remove protection first.

---

## CLI Usage

```bash
python -m src.extract_docx_comments input.docx -o output.xlsx
```

Options:
- `-o / --output` — path to the Excel file to write (default: `<input_basename>.xlsx`)
- `--author` — include author column
- `--date` — include date column (as written in the comment metadata)
- `--keep-empty` — include rows where the commented text is empty (default drops empties)

Example (with author/date):
```bash
python -m src.extract_docx_comments "My Document.docx" -o comments.xlsx --author --date
```

---

## Tested Environments

- macOS 12+ with Python 3.9–3.12
- Microsoft Word `.docx` files (Office 2016+ / 365)

---

## License

MIT
