import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import io
import zipfile
import re
import unicodedata
from time import sleep

# ----------------- PAGE & THEME -----------------
st.set_page_config(
    page_title="פיצול חשבוניות · שינוי שמות",
    page_icon="🧾",
    layout="centered"
)

# Minimal “glass” style
CUSTOM_CSS = """
<style>
/* Base */
:root{
  --brand:#7C3AED;
  --brand-2:#5B21B6;
  --ok:#10B981;
  --warn:#F59E0B;
  --err:#EF4444;
}
.block-container{
  max-width: 980px !important;
}
/* Title */
.big-title{
  font-size: 2.1rem; font-weight: 800; letter-spacing:.3px;
  background: linear-gradient(90deg,var(--brand),#EC4899);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  margin: .25rem 0 1rem 0;
}
/* Cards */
.glass{
  border-radius: 14px;
  padding: 18px 18px 14px 18px;
  background: rgba(255,255,255,.55);
  border: 1px solid rgba(0,0,0,.06);
  box-shadow: 0 10px 30px rgba(0,0,0,.06);
}
[data-testid="stFileUploader"] > div > div{
  border: 1.5px dashed rgba(0,0,0,.2);
}
.kicker{
  font-size:.85rem; font-weight:700; color:#64748B; letter-spacing:.06em;
  text-transform:uppercase; margin-bottom:.35rem;
}
hr.grad{
  height: 1px; border: none;
  background: linear-gradient(90deg, transparent, rgba(0,0,0,.12), transparent);
  margin: .8rem 0 1.1rem 0;
}
/* Buttons */
.stDownloadButton button, .stButton button{
  border-radius: 12px !important;
  padding: .58rem 1rem !important;
  font-weight: 700 !important;
}
.stButton button:hover{ transform: translateY(-1px); transition:.2s; }
/* Table */
[data-testid="stDataFrame"] div[role="table"]{
  border-radius: 12px; overflow: hidden;
  border:1px solid rgba(0,0,0,.05);
}
/* Badges */
.badge{
  display:inline-flex; align-items:center; gap:.5rem;
  padding:.25rem .6rem; border-radius:999px;
  background:#F1F5F9; color:#0F172A; font-size:.85rem; font-weight:700;
}
.badge.ok{ background:rgba(16,185,129,.12); color:#065F46;}
.badge.warn{ background:rgba(245,158,11,.12); color:#92400E;}
.badge.err{ background:rgba(239,68,68,.12); color:#7F1D1D;}
/* Fade-in */
@keyframes fade {
  from {opacity:0; transform: translateY(4px);}
  to {opacity:1; transform: none;}
}
.fade{ animation: fade .4s ease both; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ----------------- HELPERS -----------------
def normalize_text(s: str) -> str:
    """נירמול טקסט למפתח חיפוש"""
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFC", s)
    return s.upper()

def sanitize_filename(name: str) -> str:
    """ניקוי שם קובץ מתווים אסורים"""
    if not isinstance(name, str):
        name = str(name) if name is not None else ""
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:180]

def build_invoice_map(df: pd.DataFrame):
    required = {"חשבונית", "שם לקוח"}
    if not required.issubset(df.columns):
        raise ValueError("קובץ ה-Excel חייב להכיל עמודות בשם 'חשבונית' ו-'שם לקוח'.")
    mapping = {}
    for _, r in df.iterrows():
        inv_raw = str(r["חשבונית"]).strip()
        cust_raw = str(r["שם לקוח"]).strip()
        if not inv_raw:
            continue
        key = normalize_text(inv_raw)
        mapping[key] = (inv_raw, sanitize_filename(cust_raw))
    return mapping

INVOICE_CANDIDATE_RE = re.compile(r"[A-Z]{1,4}\d{5,}")

def find_invoice_in_page_text(text: str, invoice_map_keys):
    if not text:
        return None
    text_norm = normalize_text(text)
    for cand in INVOICE_CANDIDATE_RE.findall(text_norm):
        if cand in invoice_map_keys:
            return cand
    for key in invoice_map_keys:
        if key in text_norm:
            return key
    return None

# ----------------- SIDEBAR -----------------
with st.sidebar:
    st.markdown("### 🧭 איך זה עובד?")
    st.write(
        "- העלי PDF של החשבוניות\n"
        "- העלי Excel עם העמודות: **'חשבונית'** ו-**'שם לקוח'**\n"
        "- לחצי **התחל פיצול** – נקבל ZIP להורדה"
    )
    st.markdown("#### 🧩 טיפים")
    st.write("אם תבנית מספר החשבונית שונה, עדכני אותי ואכוון את החוקיות (Regex).")

# ----------------- HEADER -----------------
st.markdown('<div class="big-title">פיצול חשבוניות + שינוי שמות</div>', unsafe_allow_html=True)
st.caption("זיהוי אוטומטי של מספר חשבונית בכל עמוד PDF והצלבה לשם לקוח מתוך Excel • תוצאה: קבצים נקיים להורדה.")

# ----------------- INPUT CARDS -----------------
with st.container():
    st.markdown('<div class="glass fade">', unsafe_allow_html=True)
    st.markdown('<div class="kicker">קובצי קלט</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("📄 בחרי קובץ PDF", type=["pdf"])
    with col2:
        excel_file = st.file_uploader("📊 בחרי קובץ Excel (עמודות: 'חשבונית', 'שם לקוח')", type=["xlsx"])
    st.markdown('<hr class="grad">', unsafe_allow_html=True)
    go = st.button("🚀 התחל פיצול")
    st.markdown("</div>", unsafe_allow_html=True)

# ----------------- PROCESS -----------------
if go:
    if not pdf_file or not excel_file:
        st.error("❗ חובה להעלות גם PDF וגם Excel.")
        st.stop()

    with st.spinner("מחפש חשבוניות ומכין קבצים..."):
        sleep(0.3)
        try:
            df = pd.read_excel(excel_file)
            inv_map = build_invoice_map(df)
            if not inv_map:
                st.error("לא נמצאו חשבוניות תקינות בקובץ ה-Excel.")
                st.stop()

            reader = PdfReader(pdf_file)
            results, used_names = [], set()
            zip_buffer = io.BytesIO()

            progress = st.progress(0, text="מעבד עמודים...")
            total = len(reader.pages) if reader.pages else 0

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for i in range(total):
                    page = reader.pages[i]
                    text = page.extract_text() or ""

                    found_key = find_invoice_in_page_text(text, inv_map.keys())
                    if found_key:
                        inv_orig, cust_name = inv_map[found_key]
                        base = f"{inv_orig}_{cust_name}"
                    else:
                        inv_orig, cust_name = "", ""
                        base = f"UNMATCHED_page_{i+1}"

                    file_base = sanitize_filename(base)
                    final = file_base
                    c = 2
                    while final in used_names:
                        final = sanitize_filename(f"{file_base}_{c}")
                        c += 1
                    used_names.add(final)

                    writer = PdfWriter()
                    writer.add_page(page)
                    buf = io.BytesIO()
                    writer.write(buf); buf.seek(0)
                    zf.writestr(f"{final}.pdf", buf.getvalue())

                    results.append({
                        "עמוד": i+1,
                        "חשבונית": inv_orig if inv_orig else "—",
                        "שם לקוח": cust_name if cust_name else "—",
                        "שם קובץ": f"{final}.pdf",
                        "סטטוס": "התאמה נמצאה ✅" if found_key else "לא זוהה קוד 🔎"
                    })

                    progress.progress((i+1)/max(total,1), text=f"מעבד עמוד {i+1} מתוך {total}")

            zip_buffer.seek(0)

        except Exception as e:
            st.error(f"❌ שגיאה: {e}")
            st.stop()

    # ----------------- OUTPUT CARD -----------------
    st.markdown('<div class="glass fade">', unsafe_allow_html=True)
    st.markdown('<div class="kicker">תוצאה</div>', unsafe_allow_html=True)
    st.success("הפיצול הושלם בהצלחה! אפשר להוריד את כל הקבצים כ-ZIP.")
    st.download_button(
        "⬇️ הורדת ZIP",
        data=zip_buffer.getvalue(),
        file_name="split_invoices.zip",
        mime="application/zip"
    )
    st.markdown('<hr class="grad">', unsafe_allow_html=True)
    st.markdown("##### סיכום התאמות")
    st.dataframe(pd.DataFrame(results), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ----------------- FOOTER -----------------
st.caption("נבנה באהבה • אם תרצי לשמור לתיקייה מקומית במקום ZIP – אשדרג גם לזה 😉")
