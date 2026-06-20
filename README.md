# Letter of Employment — Mass Generator

https://loe-generator-cjzsuua9glifuch79tb7be.streamlit.app

A laptop-hosted demo app: upload a Word template + an Excel of candidates, get one
Letter of Employment (`.docx`) per candidate, downloadable as a zip.

## Setup (one-time)

```powershell
cd C:\Users\eug_l\Documents\loe-generator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
.\.venv\Scripts\Activate.ps1   # if not already active
streamlit run app.py
```

It opens in your browser at http://localhost:8501.

**Demo flow:** upload `samples/LOE_template.docx` → upload `samples/candidates.xlsx`
→ review the detected fields → **Generate letters** → download the zip.

## How it works

- The template is a Word doc with Jinja2 placeholders, e.g. `Dear {{ first_name }},`.
  Placeholders are highlighted yellow so they're easy to see.
- The Excel's **column headers must match the placeholder names exactly**
  (case-sensitive). Run `python make_template.py` to see/print the full list.
- The app fills the template once per Excel row using `docxtpl` (which preserves all
  Word formatting) and zips the results as `LOE - <Name>.docx`.

## Files

| File | Purpose |
|------|---------|
| `app.py` | The Streamlit app. |
| `make_template.py` | Regenerates `samples/LOE_template.docx` (edit clauses here). |
| `make_samples.py` | Regenerates `samples/candidates.xlsx` (3 dummy candidates). |
| `requirements.txt` | Dependencies. |
| `BUILD_PLAN.md` | Full design spec and rationale. |
| `samples/` | Ready-to-use template + candidate data. |

## Notes

- Output is `.docx` only (no PDF) — pure Python, no Microsoft Word needed to run.
- Yellow placeholder highlights remain on the merged text. Fine for a demo; for clean
  output, remove/recolour the highlight in `make_template.py`.
- Numbers/dates render exactly as typed in the Excel (the app reads everything as text),
  so format them the way you want them to appear in the letter.
