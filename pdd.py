import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import RGBColor, Inches, Pt
from pptx import Presentation
from pptx.util import Inches as PPTInches, Pt as PPTPt
from pptx.dml.color import RGBColor as PPTColor
from pptx.enum.shapes import MSO_SHAPE
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="PDD Automation Suite", page_icon="📋", layout="wide")

# ─── SESSION STATE ───
DEFAULTS = {
    "gids_number": "", "process_name": "", "business_unit": "", "process_owner": "",
    "processing_steps": [], "business_rules": [],
    "aht_minutes": 0, "daily_volume": 0, "monthly_volume": 0,
    "peak_periods": "", "target_sla": "", "out_of_scope": "",
    "applications": [], "as_is_flow": "", "to_be_flow": "",
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─── SIDEBAR ───
with st.sidebar:
    st.title("⚙️ Controls")
    st.info("Default IA-compliant template active")
    st.divider()
    if st.button("🔒 Wipe All Data", use_container_width=True, type="primary"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
    st.caption("🛡️ Session-only. Nothing is saved server-side.")

# ─── HEADER ───
st.title("📋 PDD / HLD / LLD Generator")
st.caption("Fill any tab in any order — your data persists automatically.")

tabs = st.tabs([
    "1️⃣ Metadata", "2️⃣ Steps", "3️⃣ Rules",
    "4️⃣ Metrics", "5️⃣ Apps", "6️⃣ As-Is/To-Be", "🚀 Generate"
])

with tabs[0]:
    st.subheader("Process Metadata")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["gids_number"] = st.text_input("GIDS Number *", st.session_state["gids_number"])
        st.session_state["process_name"] = st.text_input("Process Name *", st.session_state["process_name"])
    with c2:
        st.session_state["business_unit"] = st.text_input("Business Unit *", st.session_state["business_unit"])
        st.session_state["process_owner"] = st.text_input("Process Owner", st.session_state["process_owner"])

with tabs[1]:
    st.subheader("Processing Steps")
    df = pd.DataFrame(st.session_state["processing_steps"]) if st.session_state["processing_steps"] \
        else pd.DataFrame(columns=["Step", "Action", "System", "Output"])
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="steps_ed")
    st.session_state["processing_steps"] = edited.to_dict("records")

with tabs[2]:
    st.subheader("Business Rules")
    df = pd.DataFrame(st.session_state["business_rules"]) if st.session_state["business_rules"] \
        else pd.DataFrame(columns=["Rule ID", "Condition", "Action"])
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="rules_ed")
    st.session_state["business_rules"] = edited.to_dict("records")

with tabs[3]:
    st.subheader("Metrics & Scope")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["aht_minutes"] = st.number_input("AHT (min)", 0, 1000, st.session_state["aht_minutes"])
        st.session_state["daily_volume"] = st.number_input("Daily Volume", 0, 1000000, st.session_state["daily_volume"])
        st.session_state["monthly_volume"] = st.number_input("Monthly Volume", 0, 10000000, st.session_state["monthly_volume"])
    with c2:
        st.session_state["target_sla"] = st.text_input("Target SLA", st.session_state["target_sla"])
        st.session_state["peak_periods"] = st.text_area("Peak Periods", st.session_state["peak_periods"])
    st.session_state["out_of_scope"] = st.text_area("Out-of-Scope", st.session_state["out_of_scope"])

with tabs[4]:
    st.subheader("Applications & Environment")
    df = pd.DataFrame(st.session_state["applications"]) if st.session_state["applications"] \
        else pd.DataFrame(columns=["Application", "Type", "Credentials Profile"])
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="apps_ed")
    st.session_state["applications"] = edited.to_dict("records")

with tabs[5]:
    st.subheader("As-Is vs To-Be")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["as_is_flow"] = st.text_area("As-Is (Manual)", st.session_state["as_is_flow"], height=250)
    with c2:
        st.session_state["to_be_flow"] = st.text_area("To-Be (Automated)", st.session_state["to_be_flow"], height=250)

# ─── GENERATORS ───
def build_pdd():
    doc = Document()
    doc.add_heading("Process Definition Document", 0)
    doc.add_paragraph(f"Process: {st.session_state['process_name']}")
    doc.add_paragraph(f"Generated: {datetime.now():%Y-%m-%d %H:%M}")
    doc.add_page_break()

    doc.add_heading("1. Metadata", 1)
    t = doc.add_table(rows=4, cols=2); t.style = "Light Grid Accent 1"
    meta = [("GIDS", st.session_state["gids_number"]), ("Name", st.session_state["process_name"]),
            ("BU", st.session_state["business_unit"]), ("Owner", st.session_state["process_owner"])]
    for i, (k, v) in enumerate(meta):
        t.cell(i, 0).text = k; t.cell(i, 1).text = str(v)

    doc.add_heading("2. As-Is vs To-Be", 1)
    doc.add_paragraph("As-Is: " + (st.session_state["as_is_flow"] or "N/A"))
    doc.add_paragraph("To-Be: " + (st.session_state["to_be_flow"] or "N/A"))

    for title, data in [("3. Processing Steps", st.session_state["processing_steps"]),
                        ("4. Business Rules", st.session_state["business_rules"]),
                        ("7. Applications", st.session_state["applications"])]:
        doc.add_heading(title, 1)
        if data:
            cols = list(data[0].keys())
            t = doc.add_table(rows=1, cols=len(cols)); t.style = "Light Grid Accent 1"
            for i, c in enumerate(cols): t.rows[0].cells[i].text = str(c)
            for row in data:
                cells = t.add_row().cells
                for i, c in enumerate(cols): cells[i].text = str(row.get(c, ""))
        else:
            doc.add_paragraph("No data.")

    doc.add_heading("5. Metrics & Scope", 1)
    for k, v in [("AHT (min)", st.session_state["aht_minutes"]),
                 ("Daily Volume", st.session_state["daily_volume"]),
                 ("Monthly Volume", st.session_state["monthly_volume"]),
                 ("SLA", st.session_state["target_sla"]),
                 ("Peak Periods", st.session_state["peak_periods"]),
                 ("Out-of-Scope", st.session_state["out_of_scope"])]:
        doc.add_paragraph(f"{k}: {v}")

    buf = BytesIO(); doc.save(buf); buf.seek(0); return buf

def build_exceptions():
    doc = Document()
    doc.add_heading("Exception Handling Matrix", 0)
    doc.add_heading("Business Exceptions", 1)
    t = doc
