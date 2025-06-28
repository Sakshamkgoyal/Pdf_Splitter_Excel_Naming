import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO

st.set_page_config(page_title="PDF Splitter & Excel Mapper", layout="wide")
st.title("📄 PDF Splitter + 📊 Excel-Based Naming")

# File upload
pdf_file = st.file_uploader("Upload a PDF file", type="pdf")
excel_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

# Read Excel and let user choose columns
if excel_file:
    df = pd.read_excel(excel_file)
    st.subheader("Excel Preview")
    st.dataframe(df.head())

    excel_columns = df.columns.tolist()
    selected_columns = st.multiselect("Select columns for naming PDF files", excel_columns)

    delimiter = "-"
    if len(selected_columns) > 1:
        delimiter = st.text_input("Enter delimiter for combining column values", value="-")

# PDF splitting configuration
split_mode = st.radio("Choose PDF Split Mode", ["Fixed pages per file", "Custom page ranges"])

page_ranges = []
pages_per_file = 1

if pdf_file:
    reader = PdfReader(pdf_file)
    total_pages = len(reader.pages)
    st.write(f"Total PDF pages: {total_pages}")

    if split_mode == "Fixed pages per file":
        pages_per_file = st.number_input("Pages per file", min_value=1, max_value=total_pages, value=1)
        page_ranges = [(i, min(i + pages_per_file - 1, total_pages - 1)) for i in range(0, total_pages, pages_per_file)]
    else:
        ranges_input = st.text_area("Enter custom page ranges (e.g., 1-5,6-10,11-12)")
        if ranges_input:
            try:
                for part in ranges_input.split(","):
                    start, end = map(int, part.split("-"))
                    page_ranges.append((start - 1, end - 1))
            except:
                st.error("Invalid page range format")

# Generate PDFs
if st.button("Generate Split PDFs"):
    if not pdf_file or not excel_file:
        st.error("Please upload both PDF and Excel files.")
    elif not selected_columns:
        st.error("Please select at least one column for naming.")
    elif len(page_ranges) > len(df):
        st.error("More PDF splits than Excel rows. Ensure mapping makes sense.")
    else:
        output_files = []
        reader = PdfReader(pdf_file)

        for i, (start, end) in enumerate(page_ranges):
            writer = PdfWriter()
            for j in range(start, end + 1):
                writer.add_page(reader.pages[j])

            row = df.iloc[i]
            name_parts = [str(row[col]) for col in selected_columns]
            filename = delimiter.join(name_parts).replace(" ", "_") + ".pdf"

            pdf_bytes = BytesIO()
            writer.write(pdf_bytes)
            pdf_bytes.seek(0)
            output_files.append((filename, pdf_bytes))

        st.success("PDFs generated successfully!")

        for fname, data in output_files:
            st.download_button(label=f"Download {fname}", data=data, file_name=fname, mime="application/pdf")
