import re
import os
import io
import time
import pandas as pd
import streamlit as st
from PyPDF2 import PdfReader, PdfWriter

# =============== THEME / CSS ===============
st.set_page_config(
    page_title="פיצול חשבוניות ושינוי שמות",
    page_icon="📄",
    layout="centered",
)

PRIMARY = "#2f9e9a"     # טורקיז בסגנון RISE
PRIMARY_DARK = "#257c79"
MUTED = "#6b7d87"
BG_SOFT = "#f7fbfb"
BORDER = "#e6f2f1"

st.markdown(
    f"""
    <style>
    html, body, [class*="css"]  {{ font-family: Heebo, Rubik, Alef, Arial, sans-serif; }}
    .appview-container {{
        background: white;
    }}

    /* כותרת עליונה */
    .rise-hero {{
        background: linear-gradient(180deg, {BG_SOFT}, #ffffff 60%);
        border-bottom: 1px solid {BORDER};
        padding: 24px 18px 14px;
        text-align: center;
    }}
    .rise-title {{
        color: #0d3c40;
        font-weight: 800;
        font-size: 38px;
        letter-spacing: 0.2px;
        margin: 0 0 6px 0;
    }}
    .rise-sub {{
        color: {MUTED};
        font-size: 16px;
        margin-top: -6px;
    }}

    /* "תכונות" בסגנון צ'יפים */
    .feature-card {{
        background: white;
        border: 1px solid {BORDER};
        border-radius: 18px;
        padding: 18px 14px;
        text-align: center;
        box-shadow: 0 6px 14px rgba(47,158,154,0.06);
        transition: transform .15s ease;
    }}
    .feature-card:hover {{ transform: translateY(-2px); }}
    .feature-emoji {{
        font-size: 32px;
        line-height: 32px;
        display: inline-block;
        margin-bottom: 8px;
        color: {PRIMARY};
    }}
    .feature-title {{
        font-size: 18px;
        font-weight: 700;
        color: #0f5257;
        margin: 0;
    }}

    /* תיבות קלט */
    .stTextInput > div > div > input,
    .stTextArea textarea {{
        border-radius: 12px !important;
        border: 1px solid {BORDER};
        background: #ffffff;
    }}

    /* מעלה קבצים */
    .uploadedFile, .stFileUploaderDiv, .stFileUploader {{
        border-radius: 14px !important;
    }}
    .st-emotion-cache-1dv6l7z, .stFileUploader {{
        background: #ffffff !important;
        border: 1px dashed {BORDER} !important;
    }}

    /* כפתור ראשי */
    .stButton > button {{
        background: linear-gradient(180deg, {PRIMARY}, {PRIMARY_DARK});
        color: white;
        border: none;
        padding: 12px 22px;
        font-weight: 700;
        border-radius: 14px;
        box-shadow: 0 10px 18px rgba(47,158,154,0.22);
        transition: all .15s ease;
    }}
    .stButton > button:hover {{
        filter: brightness(1.05);
        transform: translateY(-1px);
        box-shadow: 0 12px 22px rgba(47,158,154,0.28);
    }}

    /* התראות */
    .stAlert {{
        border-radius: 14px;
        border: 1px solid {BORDER};
    }}

    /* פוטר / קרדיט */
    .rise-footer {{
        border-top: 1px solid {BORDER};
        margin-top: 26px;
        padding-top: 14px;
        text-align: center;
        color: {MUTED};
        font-size: 14px;
    }}
    .heart {{
        color: #e25b73;
        font-weight: 800;
        padding: 0 2px;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# =============== HEADER ===============
st.markdown(
    """
    <div class="rise-hero">
        <h1 class="rise-title">פיצול חשבוניות ושינוי שמות PDF</h1>
        <div class="rise-sub">מזהים את מספר החשבונית מתוך ה-PDF, משייכים שם לקוח מהאקסל, ושומרים אוטומטית בעיצוב נקי</div>
    </div>
    """,
    unsafe_allow_html=True
)

# =============== FEATURES ROW ===============
c1, c2, c3 = st.columns(3)
with c1:
    st.markdown('<div class="feature-card"><div class="feature-emoji">✅</div><p class="feature-title">אמינות</p></div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="feature-card"><div class="feature-emoji">👩‍💻</div><p class="feature-title">צוות מקצועי</p></div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="feature-card"><div class="feature-emoji">💎</div><p class="feature-title">שירותים איכותיים</p></div>', unsafe_allow_html=True)

st.markdown("")

# =============== HELPERS ===============
def sanitize_filename(name: str) -> str:
    """מנקה תווים לא חוקיים משמות קבצים"""
    return re.sub(r'[\\\\/:*?"<>|]+', "_", name).strip()

def find_invoice_number(text: str) -> str | None:
    """מאתר מספר חשבונית לפי תבניות נפוצות"""
    if not text:
        return None
    for pat in (r'(OV\d{5,})', r'(חשבונית[:\s\-]*OV\d{5,})'):
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            return re.search(r'(OV\d{5,})', m.group(0), re.IGNORECASE).group(1)
    return None

# =============== FORM / UI ===============
with st.container():
    st.subheader("טעינת קבצים")

    pdf_file = st.file_uploader("בחר קובץ PDF שמכיל כמה חשבוניות בעמודים:", type=["pdf"])
    excel_file = st.file_uploader("בחר קובץ Excel עם מיפוי עמודות: 'חשבונית' ו-'שם לקוח':", type=["xlsx"])

    output_dir = st.text_input(
        "📁 תיקיית פלט (למשל ‎C:\\Users\\user\\Desktop\\output‎ – לשמירה מקומית):",
        help="כשמריצים מקומית – האפליקציה תשמור לשם. ב-Cloud אין גישה לתיקיות מקומיות."
    )

    st.caption("טיפ: ודאי שבגיליון האקסל מופיעות בדיוק הכותרות: **חשבונית**, **שם לקוח** (כולל עברית מלאה).")

# =============== ACTION ===============
run = st.button("התחל פיצול ✂️")
log = st.empty()

if run:
    if not pdf_file or not excel_file:
        st.error("❗ חובה לבחור גם PDF וגם Excel.")
        st.stop()

    try:
        # קורא אקסל + בונה מיפוי חשבונית→שם
        df = pd.read_excel(excel_file)
        need_cols = {"חשבונית", "שם לקוח"}
        if not need_cols.issubset(df.columns):
            st.error("❌ קובץ האקסל חייב להכיל עמודות בשם: 'חשבונית' ו-'שם לקוח'.")
            st.stop()

        # מיפוי כסטנדרט (מפתח: OVxxxxx)
        map_dict = {str(row["חשבונית"]).strip(): str(row["שם לקוח"]).strip()
                    for _, row in df.iterrows()}

        # קורא PDF מה־UploadedFile
        pdf_bytes = io.BytesIO(pdf_file.read())
        reader = PdfReader(pdf_bytes)

        # יוצר תיקייה אם נדרש (רק כשמריצים מקומית)
        save_local = bool(output_dir.strip())
        if save_local:
            os.makedirs(output_dir, exist_ok=True)

        saved = 0
        progress = st.progress(0.0, text="מתחיל בפיצול...")

        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            inv = find_invoice_number(text)

            # השמטת עמודים ללא זיהוי חשבונית
            if not inv:
                log.info(f"עמוד {i+1}: לא זוהה מספר חשבונית – דילוג.")
                continue

            customer = map_dict.get(inv, "").strip()
            if not customer:
                customer = "ללא_שם"

            # שם קובץ
            filename = sanitize_filename(f"{inv}_{customer}.pdf")

            # כותב עמוד בודד
            writer = PdfWriter()
            writer.add_page(page)

            if save_local:
                path = os.path.join(output_dir, filename)
                with open(path, "wb") as f:
                    writer.write(f)
            else:
                # במצב ללא תיקייה לוקאלית – מספק הורדה מיידית לעמוד-עמוד
                buf = io.BytesIO()
                writer.write(buf)
                st.download_button(
                    label=f"⬇️ הורדה: {filename}",
                    data=buf.getvalue(),
                    file_name=filename,
                    mime="application/pdf",
                    key=f"dl_{i}_{time.time()}"
                )

            saved += 1
            progress.progress((i + 1) / len(reader.pages), text=f"מייצא עמוד {i+1} מתוך {len(reader.pages)}")

        if saved == 0:
            st.warning("לא נשמרו קבצים. ייתכן שמספרי החשבוניות לא זוהו או שאין התאמות בגיליון.")
        else:
            if save_local:
                st.success(f"✅ בוצע! {saved} קבצים נשמרו אל: {output_dir}")
            else:
                st.success(f"✅ בוצע! {saved} קבצים זמינים להורדה כאן בעמוד.")

    except Exception as e:
        st.error(f"שגיאה: {e}")

# =============== FOOTER CREDIT ===============
st.markdown(
    """
    <div class="rise-footer">
        מתוכנן ומעוצב <span class="heart">באהבה</span> על ידי <b>ילנה זמליאנסקי</b>
    </div>
    """,
    unsafe_allow_html=True
)
