import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
import zipfile
import uuid
import re
from collections import defaultdict

# App config
st.set_page_config(page_title="PDF Splitter + Excel Naming", layout="wide")
st.title("📄 PDF Splitter + 📊 Excel-Based Naming")

# Upload files
pdf_file = st.file_uploader("Upload a PDF file", type="pdf")
excel_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

output_files = []
filename_preview = []

# Excel handling
if excel_file:
    df = pd.read_excel(excel_file)
    st.subheader("Excel Preview")
    st.dataframe(df.head())

    excel_columns = df.columns.tolist()
    selected_columns = st.multiselect("Select column(s) for naming PDF files", excel_columns)

    delimiter = "-"
    if len(selected_columns) > 1:
        delimiter = st.text_input("Delimiter for joining column values", "-")

# PDF split config
split_mode = st.radio("Choose PDF Split Mode", ["Fixed pages per file", "Custom page ranges"])
page_ranges = []

if pdf_file:
    reader = PdfReader(pdf_file)
    total_pages = len(reader.pages)
    st.write(f"Total PDF pages: {total_pages}")

    if split_mode == "Fixed pages per file":
        pages_per_file = st.number_input("Pages per file", min_value=1, max_value=total_pages, value=1)
        page_ranges = [(i, min(i + pages_per_file - 1, total_pages - 1)) for i in range(0, total_pages, pages_per_file)]
    else:
        ranges_input = st.text_area("Enter page ranges (e.g. 1-5,6-10,11-12)")
        if ranges_input:
            try:
                for part in ranges_input.split(","):
                    start, end = map(int, part.strip().split("-"))
                    page_ranges.append((start - 1, end - 1))
            except Exception as e:
                st.error(f"Invalid range format: {e}")

# Function to resolve duplicate names
def resolve_duplicate_filenames(names):
    name_count = defaultdict(int)
    final_names = []
    for name in names:
        base, ext = name.rsplit(".", 1)
        count = name_count[base]
        if count == 0:
            final_names.append(f"{base}.{ext}")
        else:
            final_names.append(f"{base}_{count}.{ext}")
        name_count[base] += 1
    return final_names

# Filename preview
if pdf_file and excel_file and selected_columns and page_ranges:
    st.subheader("🔍 Preview Filenames")
    raw_names = []
    for i in range(min(len(page_ranges), len(df))):
        row = df.iloc[i]
        try:
            name_parts = [str(row[col]) if pd.notna(row[col]) else "NA" for col in selected_columns]
        except Exception:
            name_parts = [f"row_{i+1}"]
        filename = delimiter.join(name_parts).strip().replace(" ", "_")
        filename = re.sub(r"[^\w\-_.]", "_", filename) + ".pdf"
        raw_names.append(filename)

    resolved_names = resolve_duplicate_filenames(raw_names)
    filename_preview = resolved_names
    st.table(pd.DataFrame({"Split #": range(1, len(resolved_names)+1), "Filename": resolved_names}))

# Generate PDFs
if st.button("Generate Split PDFs"):
    if not (pdf_file and excel_file and selected_columns):
        st.error("Please upload files and select columns.")
    elif len(page_ranges) > len(df):
        st.warning("More PDF splits than Excel rows. Extra PDFs will be skipped.")
    else:
        reader = PdfReader(pdf_file)
        raw_names = []
        for i in range(min(len(page_ranges), len(df))):
            row = df.iloc[i]
            try:
                name_parts = [str(row[col]) if pd.notna(row[col]) else "NA" for col in selected_columns]
            except Exception:
                name_parts = [f"row_{i+1}"]
            filename = delimiter.join(name_parts).strip().replace(" ", "_")
            filename = re.sub(r"[^\w\-_.]", "_", filename) + ".pdf"
            raw_names.append(filename)

        resolved_names = resolve_duplicate_filenames(raw_names)

        for i, (start, end) in enumerate(page_ranges[:len(resolved_names)]):
            writer = PdfWriter()
            for j in range(start, end + 1):
                writer.add_page(reader.pages[j])

            pdf_bytes = BytesIO()
            writer.write(pdf_bytes)
            pdf_bytes.seek(0)
            output_files.append((resolved_names[i], pdf_bytes))

        st.success("✅ PDFs generated successfully!")

# Download Section
if output_files:
    st.subheader("📥 Download Options")

    selected_filenames = st.multiselect(
        "Select files to download individually",
        [fname for fname, _ in output_files],
        default=[fname for fname, _ in output_files]
    )

    for fname, data in output_files:
        if fname not in selected_filenames:
            continue
        with st.container():
            data.seek(0)
            st.download_button(
                label=f"Download {fname}",
                data=data,
                file_name=fname,
                mime="application/pdf",
                key=f"{fname}_{uuid.uuid4()}"
            )

    # ZIP all selected
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for fname, data in output_files:
            if fname in selected_filenames:
                data.seek(0)
                zipf.writestr(fname, data.read())
    zip_buffer.seek(0)

    st.download_button(
        label="📦 Download Selected as ZIP",
        data=zip_buffer,
        file_name="selected_pdfs.zip",
        mime="application/zip",
        key=f"zip_download_{uuid.uuid4()}"
    )
