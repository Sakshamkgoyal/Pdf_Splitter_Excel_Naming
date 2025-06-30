import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
import zipfile
import uuid
import re
from collections import defaultdict

st.set_page_config(page_title="PDF Splitter + Excel Naming", layout="wide")
st.title("📄 PDF Splitter + Excel-Based Naming")

pdf_file = st.file_uploader("Upload a PDF file", type="pdf")
excel_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

output_files = []

# Excel Handling
if excel_file:
    df = pd.read_excel(excel_file)
    st.subheader("📊 Excel Preview")
    st.dataframe(df.head())

    columns = df.columns.tolist()
    selected_columns = st.multiselect("Select Excel column(s) to name PDF files", columns)

    delimiter = "-"
    if len(selected_columns) > 1:
        delimiter = st.text_input("Delimiter between values", "-")

# PDF Split Configuration
split_mode = st.radio("Choose how to split the PDF", ["Fixed pages per file", "Custom page ranges"])
page_ranges = []

if pdf_file:
    reader = PdfReader(pdf_file)
    total_pages = len(reader.pages)
    st.write(f"📄 Total PDF pages: {total_pages}")

    if split_mode == "Fixed pages per file":
        pages_per_file = st.number_input("Pages per split", min_value=1, max_value=total_pages, value=1)
        page_ranges = [(i, min(i + pages_per_file - 1, total_pages - 1)) for i in range(0, total_pages, pages_per_file)]
    else:
        input_range = st.text_area("Enter custom ranges (e.g. 1-5,6-10,11-12):")
        if input_range:
            try:
                for part in input_range.split(","):
                    start, end = map(int, part.strip().split("-"))
                    page_ranges.append((start - 1, end - 1))
            except Exception as e:
                st.error(f"Invalid input format: {e}")

# Handle Duplicates
def resolve_duplicate_names(base_names):
    counts = defaultdict(int)
    final_names = []
    for name in base_names:
        if counts[name] == 0:
            final_names.append(f"{name}.pdf")
        else:
            final_names.append(f"{name}_{counts[name]}.pdf")
        counts[name] += 1
    return final_names

# Filename Preview
if pdf_file and excel_file and selected_columns and page_ranges:
    st.subheader("🔍 PDF Split Filename Preview")

    raw_names = []
    for i in range(min(len(page_ranges), len(df))):
        row = df.iloc[i]
        name_parts = [str(row[col]) if pd.notna(row[col]) else "NA" for col in selected_columns]
        base = delimiter.join(name_parts).strip().replace(" ", "_")
        base = re.sub(r"[^\w\-_.]", "_", base)  # Sanitize / and other symbols
        raw_names.append(base)

    final_filenames = resolve_duplicate_names(raw_names)
    st.table(pd.DataFrame({"Split #": list(range(1, len(final_filenames)+1)), "Filename": final_filenames}))

# Generate PDFs
if st.button("Generate Split PDFs"):
    if not (pdf_file and excel_file and selected_columns):
        st.error("Please upload both files and select column(s) for naming.")
    else:
        reader = PdfReader(pdf_file)
        raw_names = []
        for i in range(min(len(page_ranges), len(df))):
            row = df.iloc[i]
            name_parts = [str(row[col]) if pd.notna(row[col]) else "NA" for col in selected_columns]
            base = delimiter.join(name_parts).strip().replace(" ", "_")
            base = re.sub(r"[^\w\-_.]", "_", base)
            raw_names.append(base)

        final_filenames = resolve_duplicate_names(raw_names)

        for i, (start, end) in enumerate(page_ranges[:len(final_filenames)]):
            writer = PdfWriter()
            for j in range(start, end + 1):
                writer.add_page(reader.pages[j])
            pdf_bytes = BytesIO()
            writer.write(pdf_bytes)
            pdf_bytes.seek(0)
            output_files.append((final_filenames[i], pdf_bytes))

        st.success("✅ PDF files generated successfully!")

# Download Buttons
if output_files:
    st.subheader("📥 Download Split PDFs")

    selected_names = st.multiselect(
        "Select PDFs to download individually:",
        [fname for fname, _ in output_files],
        default=[fname for fname, _ in output_files]
    )

    for idx, (fname, data) in enumerate(output_files):
        if fname in selected_names:
            data.seek(0)
            st.download_button(
                label=f"Download {fname}",
                data=data,
                file_name=fname,
                mime="application/pdf",
                key=f"{fname}_{uuid.uuid4()}"
            )

    # ZIP download
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for fname, data in output_files:
            if fname in selected_names:
                data.seek(0)
                zipf.writestr(fname, data.read())
    zip_buffer.seek(0)

    st.download_button(
        label="📦 Download All Selected as ZIP",
        data=zip_buffer,
        file_name="split_pdfs.zip",
        mime="application/zip",
        key=f"zip_all_{uuid.uuid4()}"
    )
