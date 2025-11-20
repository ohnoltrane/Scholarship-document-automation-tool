# Converts .doc and .docx files to PDF using Microsoft Word (Windows only)
import os
import win32com.client

folder = r"<folder path>"  # Change this to your folder path

print(f"Looking in: {folder}")
print("Files found:", os.listdir(folder))

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

for filename in os.listdir(folder):
    print(f"Checking: {filename}")
    if filename.lower().endswith((".doc", ".docx")):
        doc_path = os.path.join(folder, filename)
        pdf_path = os.path.join(folder, os.path.splitext(filename)[0] + ".pdf")
        print(f"Converting: {doc_path} -> {pdf_path}")
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
        print(f"Converted: {doc_path} -> {pdf_path}")

word.Quit()

print("All conversions done.")
