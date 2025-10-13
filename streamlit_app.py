import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import io
import zipfile
import re

st.set_page_config(page_title="פיצול חשבוניות ושינוי שמות", layout="centered")

st.title("📄 פיצול חשבוניות ושינוי שמות קובצי PDF לפי לקוח")
st.caption("בסיום התהליך יופיע כפתור להורדת כל הקבצים בקובץ ZIP אחד.")

# העלאות קבצים
pdf_file = st.file_uploader("בחר קובץ PDF:", type=["pdf"])
excel_file = st.file_uploader("בחר קובץ Excel עם שמות לקוחות:", type=["xlsx"])

# פונקציה לסניטציה של שם קובץ (הסרת תווים אסורים)
_illegal = r'[<>:"/\\|?*\n\r\t]'
def sanitize_filename(name: str) -> str:
    name = re.sub(_illegal, "_", str(name))
    return name.strip().strip(" .")[:200] or "unnamed"

# ניסיון לזהות שמות עמודות גם אם יש וריאציות
def resolve_columns(df: pd.DataFrame):
    cols = {c.strip(): c for c in df.columns if isinstance(c, str)}
    # שמות מקובלים
    invoice_keys = ["חשבונית", "מספר חשבונית", "מספר חשבונית/מסמ"]
    name_keys    = ["שם לקוח", "לקוח", "שם הלקוח"]

    inv_col = next((cols[k] for k in invoice_keys if k in cols), None)
    name_col = next((cols[k] for k in name_keys if k in cols), None)
    return inv_col, name_col

if st.button("התחל פיצול"):
    if not pdf_file or not excel_file:
        st.error("❗ יש לבחור את שני הקבצים (PDF ו-Excel) לפני תחילת הפעולה.")
        st.stop()

    try:
        # קריאת אקסל
        df = pd.read_excel(excel_file, engine="openpyxl")
        inv_col, name_col = resolve_columns(df)

        if not inv_col or not name_col:
            st.error("❌ קובץ ה-Excel חייב להכיל עמודות בשם 'חשבונית' ו-'שם לקוח' (או שמות שקולים).")
            st.stop()

        # קריאת PDF
        pdf_reader = PdfReader(pdf_file)

        if df.empty or len(pdf_reader.pages) == 0:
            st.error("❌ לא נמצאו נתונים באקסל או עמודים ב-PDF.")
            st.stop()

        # נפיק קבצי PDF בזיכרון לפי מיפוי: שורה i -> עמוד i
        n = min(len(df), len(pdf_reader.pages))
        if len(df) != len(pdf_reader.pages):
            st.info(f"ℹ️ מספר שורות האקסל ({len(df)}) שונה ממספר העמודים ב-PDF ({len(pdf_reader.pages)}). "
                    f"יעובדו {n} העמודים/השורות הראשונים.")

        in_memory_files = []  # [(filename, bytes), ...]

        for i in range(n):
            inv = sanitize_filename(df.iloc[i][inv_col])
            cust = sanitize_filename(df.iloc[i][name_col])
            fname = f"{inv}_{cust}.pdf"

            writer = PdfWriter()
            writer.add_page(pdf_reader.pages[i])

            buf = io.BytesIO()
            writer.write(buf)
            buf.seek(0)

            in_memory_files.append((fname, buf.getvalue()))

        # יצירת ZIP בזיכרון
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for fname, data in in_memory_files:
                zf.writestr(fname, data)
        zip_buf.seek(0)

        st.success("✅ כל הקבצים הוכנו בהצלחה!")
        st.download_button(
            label="📦 הורד קבצים (ZIP)",
            data=zip_buf,
            file_name="split_pdfs.zip",
            mime="application/zip",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"❌ שגיאה: {e}")
