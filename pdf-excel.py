import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
import zipfile

st.set_page_config(page_title="PDF Splitter + Excel Naming", layout="wide")
st.title("📄 PDF Splitter + 📊 Excel-Based Naming")

# Upload files
pdf_file = st.file_uploader("Upload a PDF file", type="pdf")
excel_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

# Show Excel and select columns
if excel_file:
    df = pd.read_excel(excel_file)
    st.subheader("Excel Preview")
    st.dataframe(df.head())

    excel_columns = df.columns.tolist()
    selected_columns = st.multiselect("Select column(s) for naming the PDF files", excel_columns)

    delimiter = "-"
    if len(selected_columns) > 1:
        delimiter = st.text_input("Delimiter for joining column values", "-")

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
        ranges_input = st.text_area("Enter page ranges (e.g. 1-5,6-8,9-12)")
        if ranges_input:
            try:
                for part in ranges_input.split(","):
                    start, end = map(int, part.strip().split("-"))
                    page_ranges.append((start - 1, end - 1))
            except Exception as e:
                st.error(f"Invalid range format: {e}")

# Generate and download PDFs
if st.button("Generate and Download PDFs"):
    if not pdf_file or not excel_file:
        st.error("Please upload both a PDF and an Excel file.")
    elif not selected_columns:
        st.error("Please select column(s) from Excel for naming the files.")
    elif len(page_ranges) > len(df):
        st.warning("You have more split segments than Excel rows. Extra PDFs will be skipped.")
    else:
        output_files = []
        reader = PdfReader(pdf_file)

        for i, (start, end) in enumerate(page_ranges):
            if i >= len(df):
                break

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

        # Download all as ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for fname, data in output_files:
                zipf.writestr(fname, data.read())
        zip_buffer.seek(0)

        st.download_button(
            label="📦 Download All as ZIP",
            data=zip_buffer,
            file_name="split_pdfs.zip",
            mime="application/zip"
        )

        st.markdown("---")
        st.subheader("📥 Download Individual Files")

        for idx, (fname, data) in enumerate(output_files):
            safe_key = f"download_{idx}_{fname}".replace(" ", "_").replace(".", "_")
            st.download_button(
                label=f"Download {fname}",
                data=data,
                file_name=fname,
                mime="application/pdf",
                key=safe_key  # ✅ Guaranteed unique key
            )
