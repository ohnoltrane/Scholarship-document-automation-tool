import os
from PyPDF2 import PdfReader, PdfWriter

folder = r"C:\Users\gonza\NUS Dropbox\NCL-Team\NCL Admin 2.0\NCLS scholarship docs\2025-2026 Sem 1 (Aug) Intake\Templates\Aug-25\LoA Aug-25"  # Change to your folder
append_pdf_path = r"C:\Users\gonza\NUS Dropbox\NCL-Team\NCL Admin 2.0\NCLS scholarship docs\2025-2026 Sem 1 (Aug) Intake\Templates\Aug-25\LoA Aug-25\LoA_Aug25 2.pdf"  # Path to the PDF to append

for filename in os.listdir(folder):
    if filename.lower().endswith(".pdf") and filename != os.path.basename(append_pdf_path):
        pdf_path = os.path.join(folder, filename)
        output_path = os.path.join(folder, os.path.splitext(filename)[0] + "_appended.pdf")

        writer = PdfWriter()

        # Add original PDF pages
        with open(pdf_path, "rb") as orig_file:
            reader = PdfReader(orig_file)
            for page in reader.pages:
                writer.add_page(page)

        # Add appendix pages (open file each time)
        with open(append_pdf_path, "rb") as f:
            appendix_reader = PdfReader(f)
            for page in appendix_reader.pages:
                writer.add_page(page)

        # Write the new PDF
        with open(output_path, "wb") as out_file:
            writer.write(out_file)

        print(f"Appended {append_pdf_path} to {pdf_path} -> {output_path}")