import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import io, zipfile, re, unicodedata
from time import sleep

# =========================
# Page config (Reiz-like clean UI)
# =========================
st.set_page_config(
    page_title="פיצול חשבוניות · שינוי שמות",
    page_icon="🧾",
    layout="centered"
)

REIZ_CSS = """
<style>
:root{
  --txt:#0F172A;          /* main text */
  --muted:#64748B;        /* secondary */
  --line:#E5E7EB;         /* borders */
  --bg:#FFFFFF;           /* page */
  --card:#FFFFFF;         /* cards */
  --primary:#3B82F6;      /* Reiz-like blue */
  --primary-2:#1D4ED8;
  --ok:#059669;
  --warn:#D97706;
  --err:#DC2626;
}
html, body, [data-testid="stAppViewContainer"]{
  color:var(--txt);
  background:var(--bg);
  font-variant-ligatures:none;
}
.block-container{ max-width: 880px; }

/* Title */
.h-title{
  font-weight: 800; letter-spacing: .2px;
  font-size: 2.0rem; margin:.2rem 0 .7rem 0;
  color:var(--txt);
}
.h-sub{ color:var(--muted); margin:-.2rem 0 1.2rem 0; }

/* Card */
.card{
  border:1px solid var(--line);
  background:var(--card);
  border-radius: 14px; padding: 18px;
}

/* File uploader border */
[data-testid="stFileUploader"] > div > div{
  border: 1.5px dashed var(--line);
}

/* Buttons */
.stButton button, .stDownloadButton button{
  background:var(--primary); color:white; border:none;
  padding:.6rem 1rem; border-radius: 10px; font-weight:700;
}
.stButton button:hover, .stDownloadButton button:hover{
  background:var(--primary-2); transition:.15s;
}

/* Progress */
[data-testid="stProgress"] div[data-testid="stThumbValue"]{ color:var(--muted) !important; }

/* Table */
[data-testid="stDataFrame"] div[role="table"]{
  border:1px solid var(--line); border-radius: 10px; overflow:hidden;
}

/* Footer */
.footer{
  margin-top:1rem; padding-top:.6rem;
  border-top:1px solid var(--line); color:var(--muted);
  font-size:.92rem; text-align:center;
}
</style>
"""
st.markdown(REIZ_CSS, unsafe_allow_html=True)

# =========================
# Helpers (same logic)
# =========================
def normalize_text(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFC", s)
    return s.upper()

def sanitize_filename(name: str) -> str:
    if not isinstance(name, str):
        name = str(name) if name is not None else ""
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:180]

def build_invoice_map(df: pd.DataFrame):
    required = {"חשבונית", "שם לקוח"}
    if not required.issubset(df.columns):
        raise ValueError("קובץ ה-Excel חייב לכלול עמודות בשם 'חשבונית' ו-'שם לקוח'.")
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
    # חיפוש מועמדים אופייניים (OVxxxxx וכו')
    for cand in INVOICE_CANDIDATE_RE.findall(text_norm):
        if cand in invoice_map_keys:
            return cand
    # נפילה "חכמה": אם כל המפתח מופיע בטקסט
    for key in invoice_map_keys:
        if key in text_norm:
            return key
    return None

# =========================
# Sidebar (short)
# =========================
with st.sidebar:
    st.markdown("### ?איך משתמשים")
    st.write(
        "- תעלו קובץ פי די אף-חשבוניות\n"
        "- העלי Excel עם העמודות **'חשבונית'** ו-**'שם לקוח'**\n"
        "- לחצי **התחל פיצול** לקבלת ZIP"
    )

# =========================
# Header
# =========================
st.markdown('<div class="h-title">פיצול חשבוניות + שינוי שמות</div>', unsafe_allow_html=True)
st.markdown('<div class="h-sub">זיהוי אוטומטי של מספר חשבונית בכל עמוד PDF והצלבה לשם לקוח מתוך Excel</div>', unsafe_allow_html=True)

# =========================
# Inputs
# =========================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        pdf_file = st.file_uploader("📄 קובץ PDF", type=["pdf"])
    with c2:
        excel_file = st.file_uploader("📊 קובץ Excel (עמודות: 'חשבונית', 'שם לקוח')", type=["xlsx"])
    st.markdown("</div>", unsafe_allow_html=True)

start = st.button("🚀 התחל פיצול")

# =========================
# Process
# =========================
if start:
    if not pdf_file or not excel_file:
        st.error("חובה להעלות גם PDF וגם Excel.")
        st.stop()

    with st.spinner("מעבד..."):
        sleep(0.2)
        try:
            df = pd.read_excel(excel_file)
            inv_map = build_invoice_map(df)
            if not inv_map:
                st.error("לא נמצאו רשומות תקינות ב-Excel.")
                st.stop()

            reader = PdfReader(pdf_file)
            total = len(reader.pages) if reader.pages else 0
            results, used_names = [], set()
            zip_buffer = io.BytesIO()
            prog = st.progress(0, text="מעבד עמודים...")

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
                        "סטטוס": "התאמה נמצאה" if found_key else "לא זוהה קוד"
                    })
                    prog.progress((i+1)/max(total,1), text=f"מעבד עמוד {i+1} מתוך {total}")

            zip_buffer.seek(0)

        except Exception as e:
            st.error(f"שגיאה: {e}")
            st.stop()

    # =========================
    # Output
    # =========================
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.success("הפיצול הושלם! ניתן להוריד את כל הקבצים כ-ZIP.")
        st.download_button(
            "⬇️ הורדת ZIP",
            data=zip_buffer.getvalue(),
            file_name="split_invoices.zip",
            mime="application/zip"
        )
        st.markdown("#### סיכום התאמות")
        st.dataframe(pd.DataFrame(results), use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

# =========================
# Footer credit
# =========================
st.markdown('<div class="footer">מתוכנן על ידי ילנה זמליאנסקי</div>', unsafe_allow_html=True)


