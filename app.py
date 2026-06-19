import io
import re
import zipfile

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
template_bytes = template_file.read()

try:
    df = pd.read_excel(excel_file, dtype=str).fillna("")
except Exception as e:
    st.error(f"Could not read the Excel file: {e}")
    st.stop()

if df.empty:
    st.error("The Excel file has no candidate rows.")
    st.stop()

# Detect placeholders declared in the template.
placeholders = DocxTemplate(io.BytesIO(template_bytes)).get_undeclared_template_variables()
excel_columns = set(df.columns)

st.subheader("Detected fields")
c1, c2 = st.columns(2)
c1.metric("Template placeholders", len(placeholders))
c2.metric("Candidate rows", len(df))

with st.expander("Field details", expanded=bool(placeholders - excel_columns)):
    st.write(f"**Template placeholders:** {', '.join(sorted(placeholders)) or '(none found)'}")
    st.write(f"**Excel columns:** {', '.join(df.columns)}")
    missing = placeholders - excel_columns
    extra = excel_columns - placeholders
    if missing:
        st.warning(
            "These placeholders have no matching Excel column and will render blank: "
            + ", ".join(sorted(missing))
        )
    if extra:
        st.info("These Excel columns are not used by the template: " + ", ".join(sorted(extra)))
    if not missing:
        st.success("Every template placeholder has a matching Excel column. ✅")

# Which column to use for output filenames.
name_col = st.selectbox(
    "Column to use in each filename",
    options=list(df.columns),
    index=list(df.columns).index("name") if "name" in df.columns else 0,
)

st.dataframe(df, use_container_width=True)

# --- Step 3: generate --------------------------------------------------------
if st.button("Generate letters", type="primary"):
    zip_buffer = io.BytesIO()
    errors = []
    generated = 0

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        used_names = {}
        progress = st.progress(0.0, text="Generating…")
        for pos, (i, row) in enumerate(df.iterrows()):
            context = row.to_dict()
            try:
                doc = DocxTemplate(io.BytesIO(template_bytes))
                doc.render(context)

                base = f"LOE - {safe_filename(row[name_col])}"
                used_names[base] = used_names.get(base, 0) + 1
                if used_names[base] > 1:
                    base = f"{base} ({used_names[base]})"

                out = io.BytesIO()
                doc.save(out)
                zf.writestr(f"{base}.docx", out.getvalue())
                generated += 1
            except Exception as e:
                errors.append(f"Row {i + 2} ({row.get(name_col, '?')}): {e}")
            progress.progress((pos + 1) / len(df), text=f"Generating… {pos + 1}/{len(df)}")

    if errors:
        st.error("Some rows failed:\n\n" + "\n\n".join(errors))

    if generated:
        st.success(f"Generated {generated} letter(s).")
        st.download_button(
            "⬇️ Download all letters (.zip)",
            data=zip_buffer.getvalue(),
            file_name="letters_of_employment.zip",
            mime="application/zip",
        )
