# Letter of Employment — Mass Generator (Build Plan)

A small laptop-hosted demo app: upload a Word template + an Excel of candidates, get one
Letter of Employment (`.docx`) per candidate, downloadable as a zip.

**Status:** spec only — code below is ready to build later.
**Output format:** Word `.docx` only (no PDF). Pure Python, no Microsoft Word / LibreOffice needed.

---

## 1. How it works (the mental model)

This is a **mail merge**, done properly:

1. Take your real Letter of Employment `.docx` and replace the variable text with
   **placeholders** in Jinja2 syntax: `Dear {{ name }}`, `NRIC: {{ nric }}`.
2. Each placeholder name maps to a **column header** in your Excel
   (`name`, `nric`, `salary`, …).
3. The app loops over every Excel row, fills the template, and writes
   `LOE - <Name>.docx` for each. All letters are zipped for one-click download.

The engine is **`docxtpl`** (DocxTemplate). Unlike `python-docx`, it preserves the
original document's fonts, spacing, headers/footers, logos, and tables — it only swaps
the placeholder text. This is what makes the output look like a real letter, not a
regenerated one.

---

## 2. Tech stack & why

| Choice | Why |
|--------|-----|
| **Python 3.10+** | Best mail-merge libraries; trivial to run on a laptop. |
| **Streamlit** | Browser upload UI in ~100 lines. `streamlit run app.py` → opens in browser. No HTML/JS. |
| **docxtpl** | Templating built for Word — keeps all formatting intact. |
| **pandas + openpyxl** | Read the Excel into rows/columns reliably (`.xlsx`). |

**Rejected alternatives:** Word's built-in Mail Merge (no demo UI, can't easily emit
separate files per person), `python-docx` alone (mangles complex formatting), a full
Flask/React app (overkill for a laptop demo).

---

## 3. Project structure

```
loe-generator/
├── BUILD_PLAN.md          ← this file
├── app.py                 ← the Streamlit app (full code in §6)
├── requirements.txt       ← dependencies (§5)
├── README.md              ← short run instructions (§7 content)
└── samples/
    ├── LOE_template.docx  ← your template with {{ placeholders }}
    └── candidates.xlsx    ← columns matching the placeholders
```

---

## 4. The two inputs (conventions)

### 4a. The Word template (`.docx`)
- Open your real Letter of Employment.
- Replace each variable with `{{ placeholder_name }}`. Placeholder names should be
  lowercase, no spaces (use underscores): `{{ name }}`, `{{ nric }}`, `{{ salary }}`,
  `{{ job_title }}`, `{{ start_date }}`, `{{ benefits }}`.
- **Tip:** type the placeholder in one go. If Word's spellcheck splits `{{name}}` across
  runs, docxtpl can miss it — easiest fix is to type it in Notepad and paste, or disable
  autocorrect for that line.
- Conditional / optional clauses are supported by Jinja2, e.g.
  `{% if benefits %}Benefits: {{ benefits }}{% endif %}`.

### 4b. The Excel (`.xlsx`)
- **Row 1 = headers**, and the headers must match the placeholder names exactly
  (case-sensitive): `name`, `nric`, `salary`, `job_title`, `start_date`, `benefits`.
- One candidate per row.
- Example:

| name | nric | job_title | salary | start_date | benefits |
|------|------|-----------|--------|-----------|----------|
| John Tan | S1234567A | Software Engineer | 6,500 | 2026-07-01 | 14 days annual leave, medical |
| Mary Lim | S7654321B | Product Manager | 8,200 | 2026-07-15 | 18 days annual leave, medical, dental |

---

## 5. `requirements.txt`

```
streamlit>=1.30
docxtpl>=0.16
pandas>=2.0
openpyxl>=3.1
```

---

## 6. `app.py` (full code)

```python
import io
import re
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

st.set_page_config(page_title="LOE Mass Generator", page_icon="📄", layout="centered")
st.title("📄 Letter of Employment — Mass Generator")
st.caption("Upload a Word template with {{ placeholders }} and an Excel of candidates.")


def safe_filename(value: str) -> str:
    """Strip characters that are illegal in Windows filenames."""
    cleaned = re.sub(r'[<>:"/\\|?*]', "", str(value)).strip()
    return cleaned or "candidate"


# --- Step 1: uploads ---------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("1. Word template (.docx)", type=["docx"])
with col2:
    excel_file = st.file_uploader("2. Candidate data (.xlsx)", type=["xlsx"])

if not (template_file and excel_file):
    st.info("Upload both files to continue.")
    st.stop()

# --- Step 2: read & validate -------------------------------------------------
# Read template bytes once; DocxTemplate is re-created per row from these bytes.
template_bytes = template_file.read()

try:
    df = pd.read_excel(excel_file, dtype=str).fillna("")
except Exception as e:
    st.error(f"Could not read the Excel file: {e}")
    st.stop()

# Detect placeholders declared in the template.
placeholders = DocxTemplate(io.BytesIO(template_bytes)).get_undeclared_template_variables()
excel_columns = set(df.columns)

st.subheader("Detected fields")
st.write(f"**Template placeholders:** {', '.join(sorted(placeholders)) or '(none found)'}")
st.write(f"**Excel columns:** {', '.join(df.columns)}")

missing = placeholders - excel_columns
if missing:
    st.warning(
        f"These placeholders have no matching Excel column and will render blank: "
        f"{', '.join(sorted(missing))}"
    )

# Which column to use for output filenames.
name_col = st.selectbox(
    "Column to use in each filename",
    options=list(df.columns),
    index=list(df.columns).index("name") if "name" in df.columns else 0,
)

st.dataframe(df, use_container_width=True)
st.write(f"**{len(df)} candidate(s)** ready.")

# --- Step 3: generate --------------------------------------------------------
if st.button("Generate letters", type="primary"):
    zip_buffer = io.BytesIO()
    errors = []

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        used_names = {}
        for i, row in df.iterrows():
            context = row.to_dict()
            try:
                doc = DocxTemplate(io.BytesIO(template_bytes))
                doc.render(context)

                base = f"LOE - {safe_filename(row[name_col])}"
                # de-duplicate identical names
                used_names[base] = used_names.get(base, 0) + 1
                if used_names[base] > 1:
                    base = f"{base} ({used_names[base]})"

                out = io.BytesIO()
                doc.save(out)
                zf.writestr(f"{base}.docx", out.getvalue())
            except Exception as e:
                errors.append(f"Row {i + 2} ({row.get(name_col, '?')}): {e}")

    if errors:
        st.error("Some rows failed:\n\n" + "\n\n".join(errors))

    if zip_buffer.getbuffer().nbytes > 0:
        st.success(f"Generated {len(df) - len(errors)} letter(s).")
        st.download_button(
            "⬇️ Download all letters (.zip)",
            data=zip_buffer.getvalue(),
            file_name="letters_of_employment.zip",
            mime="application/zip",
        )
```

### Key design points in the code
- **Template read once, re-instantiated per row** — `DocxTemplate` is stateful after
  `.render()`, so each candidate gets a fresh copy from the original bytes.
- **`dtype=str` + `.fillna("")`** — keeps NRICs, salaries, and dates exactly as typed
  (no `6500.0`, no `NaN`, no Excel date serial numbers). Format numbers/dates the way you
  want them to appear *in the Excel itself*.
- **Placeholder auto-detection** — `get_undeclared_template_variables()` shows the user
  which fields the template expects and warns about mismatches before generating.
- **Filename safety + de-duplication** — strips illegal Windows characters and appends
  `(2)`, `(3)` for duplicate names so nothing is silently overwritten in the zip.
- **Per-row error capture** — one bad row doesn't kill the whole batch.

---

## 7. Setup & run (README content)

```powershell
# from the loe-generator folder, one-time setup:
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# run the app (opens in your browser at http://localhost:8501):
streamlit run app.py
```

Demo flow: upload `samples/LOE_template.docx` → upload `samples/candidates.xlsx` →
review detected fields → **Generate letters** → download the zip.

---

## 8. Edge cases handled / to watch

| Case | Handling |
|------|----------|
| Placeholder split across Word runs | Type placeholders in one go / paste from Notepad. |
| Numbers showing as `6500.0` | `dtype=str` on read; format in Excel. |
| Empty cells → `NaN` in doc | `.fillna("")` renders blank instead. |
| Duplicate candidate names | Filenames de-duplicated with `(2)`, `(3)`. |
| Illegal filename characters | Stripped by `safe_filename()`. |
| Placeholder with no Excel column | Warned before generation; renders blank. |
| One malformed row | Caught per-row; rest of the batch still completes. |

---

## 9. Optional enhancements (later, if the demo lands)

- **PDF output** — add a "Convert to PDF" toggle using `docx2pdf` (needs Word installed)
  or LibreOffice headless (`soffice --headless --convert-to pdf`). Deliberately left out
  to keep this dependency-free.
- **Per-letter preview** — render the first row and show it before bulk generation.
- **Email merge** — send each letter to the candidate via SMTP.
- **Column ↔ placeholder mapping UI** — let the user map mismatched names instead of
  requiring exact matches.
- **Bulk PDF + digital signature block** for a more "final" look.

---

## 10. Build checklist

- [ ] Create `requirements.txt` (§5) and install into a venv.
- [ ] Save `app.py` (§6).
- [ ] Make `samples/LOE_template.docx` from your real letter with `{{ placeholders }}`.
- [ ] Make `samples/candidates.xlsx` with matching column headers (§4b).
- [ ] `streamlit run app.py` and walk the demo flow.
- [ ] (Optional) Write `README.md` from §7.
