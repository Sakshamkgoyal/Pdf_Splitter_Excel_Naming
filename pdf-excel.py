import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
import zipfile
import uuid
import re
from collections import defaultdict

st.set_page_config(page_title="PDF Splitter + Excel-Based Naming", layout="wide")
st.title("📄 PDF Splitter + Excel-Based Naming")

# Initialize session state
if "output_files" not in st.session_state:
    st.session_state.output_files = []

# Upload files
pdf_file = st.file_uploader("Upload a PDF file", type="pdf", key="pdf_upload")
excel_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"], key="excel_upload")

if excel_file:
    df = pd.read_excel(excel_file)
    st.subheader("📊 Excel Preview")
    st.dataframe(df.head())

    columns = df.columns.tolist()
    selected_columns = st.multiselect("Select Excel column(s) to name PDF files", columns, key="column_select")

    delimiter = "-"
    if len(selected_columns) > 1:
        delimiter = st.text_input("Delimiter between values", "-", key="delimiter_input")

# PDF split settings
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

# Filename generation
def generate_filenames(df, columns, delimiter, count):
    raw_names = []
    for i in range(min(count, len(df))):
        row = df.iloc[i]
        parts = [str(row[col]) if pd.notna(row[col]) else "NA" for col in columns]
        base = delimiter.join(parts).replace(" ", "_")
        base = re.sub(r"[^\w\-_.]", "_", base)
        raw_names.append(base)

    final = []
    counts = defaultdict(int)
    for name in raw_names:
        if counts[name] == 0:
            final.append(f"{name}.pdf")
        else:
            final.append(f"{name}_{counts[name]}.pdf")
        counts[name] += 1
    return final

# Preview filenames
if pdf_file and excel_file and selected_columns and page_ranges:
    usable_count = min(len(page_ranges), len(df))
    final_filenames = generate_filenames(df, selected_columns, delimiter, usable_count)

    st.subheader("🔍 Filename Preview")
    st.dataframe(pd.DataFrame({
        "Split #": list(range(1, usable_count + 1)),
        "Filename": final_filenames
    }))

# Generate PDFs
if st.button("Generate Split PDFs"):
    if not (pdf_file and excel_file and selected_columns and page_ranges):
        st.error("Please upload both files and select naming columns.")
    else:
        st.session_state.output_files = []  # Clear previous outputs

        reader = PdfReader(pdf_file)
        usable_count = min(len(page_ranges), len(df))
        final_filenames = generate_filenames(df, selected_columns, delimiter, usable_count)

        for i, (start, end) in enumerate(page_ranges[:usable_count]):
            writer = PdfWriter()
            for j in range(start, end + 1):
                writer.add_page(reader.pages[j])

            buffer = BytesIO()
            writer.write(buffer)
            buffer.seek(0)
            st.session_state.output_files.append((final_filenames[i], buffer))

        st.success("✅ PDFs generated successfully!")

# Download section
if st.session_state.output_files:
    st.subheader("📥 Download PDFs")
    selected_names = st.multiselect(
        "Select PDFs to download:",
        [fname for fname, _ in st.session_state.output_files],
        default=[fname for fname, _ in st.session_state.output_files]
    )

    for i, (fname, data) in enumerate(st.session_state.output_files):
        if fname in selected_names:
            st.download_button(
                label=f"Download {fname}",
                data=data.getvalue(),
                file_name=fname,
                mime="application/pdf",
                key=f"dl_button_{i}_{uuid.uuid4()}"
            )

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for fname, data in st.session_state.output_files:
            if fname in selected_names:
                data.seek(0)
                zipf.writestr(fname, data.read())
    zip_buffer.seek(0)

    st.download_button(
        label="📦 Download Selected as ZIP",
        data=zip_buffer,
        file_name="split_pdfs.zip",
        mime="application/zip",
        key=f"zip_download_{uuid.uuid4()}"
    )
