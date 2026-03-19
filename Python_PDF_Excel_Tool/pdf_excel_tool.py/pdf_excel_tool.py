# pdf_excel_tool.py
# Python Project: PDF & Excel Automation Tool
# Author: Hiba

import os
import pandas as pd
from PyPDF2 import PdfReader

# مجلد الملفات
folder = "files_to_merge"  # ضع ملفات PDF وExcel هنا

# دمج ملفات Excel
excel_files = [f for f in os.listdir(folder) if f.endswith('.xlsx')]
all_data = pd.DataFrame()

for file in excel_files:
    df = pd.read_excel(os.path.join(folder, file))
    all_data = pd.concat([all_data, df], ignore_index=True)

# حفظ ملف Excel موحد
all_data.to_excel("merged_excel.xlsx", index=False)
print("✅ Excel files merged into merged_excel.xlsx")

# استخراج النصوص من ملفات PDF
pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]
pdf_texts = []

for file in pdf_files:
    reader = PdfReader(os.path.join(folder, file))
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    pdf_texts.append({"file": file, "content": text})

# حفظ النصوص في CSV
pdf_df = pd.DataFrame(pdf_texts)
pdf_df.to_csv("extracted_pdf.csv", index=False)
print("✅ PDF texts extracted into extracted_pdf.csv")