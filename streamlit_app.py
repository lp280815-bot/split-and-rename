import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import io
import zipfile
import re
import unicodedata

st.set_page_config(page_title="פיצול חשבוניות ושינוי שמות", layout="centered")
st.title("📄 פיצול חשבוניות ושינוי שמות קובצי PDF לפי לקוח")
st.caption("המערכת מזהה את מספר החשבונית בכל עמוד PDF ומצליבה לשם הלקוח מתוך קובץ ה-Excel. התוצאה תורד כ-ZIP.")

# ----------------- Utilities -----------------

def normalize_text(s: str) -> str:
    """נירמול טקסט: הסרת ניקוד/תווים לא נצרכים + המרה לאותיות גדולות."""
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    # NFC של יוניקוד + המרה לאותיות גדולות
    s = unicodedata.normalize("NFC", s)
    return s.upper()

def sanitize_filename(name: str) -> str:
    """ניקוי שם קובץ מתווים אסורים במערכות קבצים."""
    if not isinstance(name, str):
        name = str(name) if name is not None else ""
    name = name.strip()
    # מחליפים תווים אסורים בקובץ (Windows)
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)
    # למקרה של רווחים מיותרים
    name = re.sub(r"\s+", " ", name).strip()
    return name[:180]  # שלא יתפוצץ על שמות ארוכים

def build_invoice_map(df: pd.DataFrame):
    """
    בונה מיפוי: מפתח נורמלי של 'חשבונית' -> (חשבונית מקורית, שם לקוח מנוקה).
    דורש עמודות: 'חשבונית', 'שם לקוח'
    """
    required = {"חשבונית", "שם לקוח"}
    if not required.issubset(df.columns):
        raise ValueError("קובץ ה-Excel חייב להכיל עמודות בשם 'חשבונית' ו-'שם לקוח'.")

    inv_map = {}
    for _, row in df.iterrows():
        inv_raw = str(row["חשבונית"]).strip()
        cust_raw = str(row["שם לקוח"]).strip()
        if not inv_raw:
            continue
        key = normalize_text(inv_raw)
        inv_map[key] = (inv_raw, sanitize_filename(cust_raw))
    return inv_map

# Regex לזיהוי קוד חשבונית אופייני (אותיות + ספרות. לדוגמה: OV255004935)
INVOICE_CANDIDATE_RE = re.compile(r"[A-Z]{1,4}\d{5,}")

def find_invoice_in_page_text(text: str, invoice_map_keys):
    """
    מחפש בעמוד את קוד החשבונית. מחזיר את ה-Key הנורמלי שמצאנו במפה (או None).
    אסטרטגיה:
    1) לזהות מועמדים עם Regex (אותיות+ספרות).
    2) לנרמל ולבדוק אם נמצא במפה.
    """
    if not text:
        return None
    text_norm = normalize_text(text)

    # קודם מנסים לזהות מועמדים עם Regex
    for cand in INVOICE_CANDIDATE_RE.findall(text_norm):
        if cand in invoice_map_keys:
            return cand

    # אם לא נמצא, ננסה חיפוש ישיר של כל מפתח במפה בתוך הטקסט (יקר יותר)
    # אבל טוב למקרים חריגים.
    for key in invoice_map_keys:
        if key in text_norm:
            return key

    return None

# ----------------- UI -----------------

pdf_file = st.file_uploader("בחר קובץ PDF:", type=["pdf"])
excel_file = st.file_uploader("בחר קובץ Excel עם שמות לקוחות:", type=["xlsx"])
run = st.button("🚀 התחל פיצול")

if run:
    if not pdf_file or not excel_file:
        st.error("❗ יש לבחור גם PDF וגם Excel לפני תחילת הפעולה.")
        st.stop()

    try:
        # קורא את האקסל
        df = pd.read_excel(excel_file)
        invoice_map = build_invoice_map(df)  # key=חשבונית מנורמלת -> (מקורית, לקוח)
        if not invoice_map:
            st.error("לא נמצאו חשבוניות תקינות באקסל.")
            st.stop()

        # קורא את ה-PDF
        reader = PdfReader(pdf_file)

        results = []            # לטבלת סיכום
        used_names = set()      # למניעת כפילויות שמות
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for page_idx in range(len(reader.pages)):
                page = reader.pages[page_idx]
                text = page.extract_text() or ""

                found_key = find_invoice_in_page_text(text, invoice_map.keys())

                if found_key:
                    inv_orig, cust_name = invoice_map[found_key]
                    base_name = f"{inv_orig}_{cust_name}"
                    file_name = sanitize_filename(base_name)
                else:
                    # לא נמצא קוד חשבונית — שם ברירת-מחדל
                    inv_orig, cust_name = "", ""
                    base_name = f"UNMATCHED_page_{page_idx + 1}"
                    file_name = sanitize_filename(base_name)

                # ודא ייחודיות
                final_name = file_name
                counter = 2
                while final_name in used_names:
                    final_name = sanitize_filename(f"{file_name}_{counter}")
                    counter += 1
                used_names.add(final_name)

                # יצירת PDF לעמוד זה
                writer = PdfWriter()
                writer.add_page(page)
                buf = io.BytesIO()
                writer.write(buf)
                buf.seek(0)

                # כתיבה ל-ZIP
                zf.writestr(f"{final_name}.pdf", buf.getvalue())

                results.append({
                    "עמוד": page_idx + 1,
                    "חשבונית שנמצאה": inv_orig if inv_orig else "—",
                    "שם לקוח": cust_name if cust_name else "—",
                    "שם קובץ": f"{final_name}.pdf",
                    "סטטוס": "הותאם" if found_key else "לא נמצא קוד חשבונית"
                })

        # תכולת ה-ZIP
        zip_buffer.seek(0)
        st.success("✅ הפיצול הושלם! ניתן להוריד כעת את הקבצים כ-ZIP.")
        st.download_button(
            label="⬇️ הורדת ZIP",
            data=zip_buffer.getvalue(),
            file_name="split_invoices.zip",
            mime="application/zip"
        )

        # טבלת סיכום
        st.write("### סיכום התאמות")
        st.dataframe(pd.DataFrame(results))

    except Exception as e:
        st.error(f"❌ שגיאה במהלך העיבוד: {e}")
