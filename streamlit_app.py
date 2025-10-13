import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import os

st.set_page_config(page_title="פיצול חשבוניות ושינוי שמות", layout="centered")

st.title("📄 פיצול חשבוניות ושינוי שמות קובצי PDF לפי לקוח")

pdf_file = st.file_uploader("בחר קובץ PDF:", type=["pdf"])
excel_file = st.file_uploader("בחר קובץ Excel עם שמות לקוחות:", type=["xlsx"])
output_dir = st.text_input("📁 תיקיית פלט (למשל C:\\Users\\user110\\Desktop\\output):")

if st.button("התחל פיצול"):
    if not pdf_file or not excel_file or not output_dir:
        st.error("❗ יש לבחור את כל הקבצים לפני תחילת הפעולה.")
    else:
        try:
            df = pd.read_excel(excel_file)
            if not {"חשבונית", "שם לקוח"}.issubset(df.columns):
                st.error("❌ הקובץ Excel חייב להכיל עמודות בשם 'חשבונית' ו-'שם לקוח'.")
            else:
                pdf = PdfReader(pdf_file)
                for i, row in df.iterrows():
                    invoice = str(row["חשבונית"]).strip()
                    name = str(row["שם לקוח"]).strip().replace("/", "_")
                    writer = PdfWriter()
                    if i < len(pdf.pages):
                        writer.add_page(pdf.pages[i])
                        os.makedirs(output_dir, exist_ok=True)
                        output_path = os.path.join(output_dir, f"{invoice}_{name}.pdf")
                        with open(output_path, "wb") as f:
                            writer.write(f)
                st.success("✅ כל הקבצים נשמרו בהצלחה!")
        except Exception as e:
            st.error(f"שגיאה: {e}")
