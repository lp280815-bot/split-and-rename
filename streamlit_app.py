import io
import re
import zipfile
import unicodedata
from collections import defaultdict
import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter

# ---------- Utilities ----------

# תווי כיווניות שמפריעים לטקסטים בעברית
BIDI_CONTROL = dict.fromkeys(map(ord, "\u200e\u200f\u202a\u202b\u202c\u202d\u202e"), None)

def clean_text(s: str) -> str:
    """ניקוי טקסט – רווחים, תווי כיווניות, סימנים מיותרים"""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(BIDI_CONTROL)
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_header(h: str) -> str:
    """ניקוי ושיוך שמות עמודות"""
    h = clean_text(h)
    h_stripped = re.sub(r"[^\w\u0590-\u05FF ]", "", h).strip().lower()

    synonyms = {
        "חשבונית": {
            "חשבונית", "מספר חשבונית", "מס חשבונית", "invoice", "inv", "מספר/חשבונית"
        },
        "שם לקוח": {
            "שם לקוח", "לקוח", "שם הלקוח", "שם לקוחות", "customer", "client name"
        }
    }

    for canon, alts in synonyms.items():
        if h in alts or h_stripped in {clean_text(x).lower() for x in alts}:
            return canon

    if "חשבונית" in h or "invoice" in h.lower():
        return "חשבונית"
    if "לקוח" in h or "customer" in h.lower() or "client" in h.lower():
        return "שם לקוח"
    return h

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {col: normalize_header(col) for col in df.columns}
    df = df.rename(columns=mapping)
    df = df.loc[:, ~df.columns.duplicated()]
    return df

def load_mapping(xlsx_bytes) -> dict:
    df = pd.read_excel(xlsx_bytes, engine="openpyxl")
    df.columns = [clean_text(c) for c in df.columns]
    df = normalize_columns(df)

    if not {"חשבונית", "שם לקוח"}.issubset(df.columns):
        missing = {"חשבונית", "שם לקוח"} - set(df.columns)
        raise ValueError(
            f"❌ חסרות עמודות נדרשות: {', '.join(missing)}.\n"
            f"ודאי שהעמודות נקראות 'חשבונית' ו-'שם לקוח' (בלי רווחים או סימנים)."
        )

    df["חשבונית"] = df["חשבונית"].apply(clean_text)
    df["שם לקוח"] = df["שם לקוח"].apply(clean_text)
    df = df[(df["חשבונית"] != "") & (df["שם לקוח"] != "")]
    return dict(zip(df["חשבונית"], df["שם לקוח"]))


# ---------- PDF Split Logic ----------

INV_REGEX = re.compile(r"(OV\d{6,})")

def extract_invoice_candidates(page_text: str) -> list[str]:
    text = clean_text(page_text)
    return INV_REGEX.findall(text)

def split_pdf_by_mapping(pdf_bytes, inv2name: dict) -> tuple[bytes, list[str]]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    logs = []
    bucket: dict[str, PdfWriter] = defaultdict(PdfWriter)
    unknown_writer = PdfWriter()
    known_invs = set(inv2name.keys())

    for i, page in enumerate(reader.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        found = None
        candidates = extract_invoice_candidates(text)
        for cand in candidates:
            if cand in known_invs:
                found = cand
                break

        if found:
            cust = inv2name[found]
            bucket[cust].add_page(page)
            logs.append(f"✅ עמוד {i}: נמצא {found} → {cust}")
        else:
            unknown_writer.add_page(page)
            logs.append(f"⚠️ עמוד {i}: לא נמצאה חשבונית מתאימה")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for cust, writer in bucket.items():
            pdf_buf = io.BytesIO()
            writer.write(pdf_buf)
            pdf_buf.seek(0)
            safe_name = re.sub(r'[\\/:*?"<>|]', "_", cust)
            zf.writestr(f"{safe_name}.pdf", pdf_buf.read())

        if len(unknown_writer.pages) > 0:
            u_buf = io.BytesIO()
            unknown_writer.write(u_buf)
            u_buf.seek(0)
            zf.writestr("Unknown.pdf", u_buf.read())

    return zip_buf.getvalue(), logs


# ---------- Streamlit UI ----------

st.set_page_config(page_title="פיצול חשבוניות + שינוי שם", page_icon="🧾", layout="centered")
st.title("🧾 פיצול חשבוניות + שינוי שם (גרסה משופרת)")

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("בחר/י קובץ PDF:", type=["pdf"])
with col2:
    xlsx_file = st.file_uploader("בחר/י קובץ Excel עם שמות לקוחות:", type=["xlsx"])

st.markdown("---")
run = st.button("🚀 התחל פיצול", use_container_width=True)
log_box = st.empty()

if run:
    if not pdf_file or not xlsx_file:
        st.error("חובה לבחור גם PDF וגם Excel לפני הפעלה.")
        st.stop()

    try:
        inv2name = load_mapping(xlsx_file)
        st.success(f"נמצאו {len(inv2name)} שורות מיפוי תקינות.")
        st.write(pd.DataFrame(list(inv2name.items())[:5], columns=["חשבונית", "שם לקוח"]))

        zip_bytes, logs = split_pdf_by_mapping(pdf_file.read(), inv2name)
        st.download_button(
            "⬇️ הורדת קבצים (ZIP)",
            data=zip_bytes,
            file_name="Invoices_Splitted.zip",
            mime="application/zip",
            use_container_width=True,
        )
        st.info("📋 יומן פעולות:")
        log_box.code("\n".join(logs), language="text")

    except Exception as e:
        st.error(f"שגיאה: {e}")
