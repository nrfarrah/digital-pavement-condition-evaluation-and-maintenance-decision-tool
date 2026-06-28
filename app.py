import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="TCG633 – Pavement Condition Evaluation Tool",
    page_icon="🛣️",
    layout="wide"
)

DEFECT_WEIGHTS = {
    "Alligator (Fatigue) Crack": 2.0,
    "Bleeding/Flushing": 1.0,
    "Block Crack": 1.4,
    "Bumps and Sags": 1.4,
    "Corrugation": 1.5,
    "Depression": 1.6,
    "Edge Crack": 1.1,
    "Lane/Shoulder Drop-off": 1.2,
    "Longitudinal Crack": 1.0,
    "Patching": 1.5,
    "Polished Aggregate": 0.8,
    "Potholes": 2.2,
    "Railroad Crossing": 1.0,
    "Raveling": 1.2,
    "Reflection Crack at Joints": 1.2,
    "Rutting": 1.8,
    "Shoving": 1.8,
    "Spalling of Longitudinal Joint": 1.3,
    "Spalling of Transverse Joints": 1.3,
    "Swells": 1.3,
    "Transverse Crack": 1.1,
    "Weathering": 1.0,
    "Other (Custom)": 1.0,  # fallback weight; user supplies name
}

DEFAULT_CUSTOM_WEIGHT = 1.0  # weight applied to any user-typed custom defect

SEVERITY_FACTORS = {"Low": 0.6, "Medium": 1.0, "High": 1.4}
DEFECT_LIST = list(DEFECT_WEIGHTS.keys())
SEVERITY_LIST = ["Low", "Medium", "High"]

def classify_pci(pci):
    if pci >= 85:
        return "Very Good", "#2ecc71"
    elif pci >= 70:
        return "Good / Satisfactory", "#27ae60"
    elif pci >= 55:
        return "Fair", "#f39c12"
    else:
        return "Poor", "#e74c3c"

def pci_recommendation(pci):
    if pci >= 85:
        return "Routine maintenance (cleaning, grass cutting, minor touch-ups)"
    elif pci >= 70:
        return "Preventive maintenance (crack sealing, local patching)"
    elif pci >= 55:
        return "Surface treatment / Overlay (localized)"
    else:
        return "Major rehabilitation / Reconstruction assessment"

def classify_iri(iri):
    if iri < 2:
        return "Very Good (Smooth)", "#2ecc71"
    elif iri < 3:
        return "Good", "#27ae60"
    elif iri < 4:
        return "Fair", "#f39c12"
    else:
        return "Poor (Rough)", "#e74c3c"

def iri_recommendation(iri):
    if iri < 2:
        return "Routine maintenance"
    elif iri < 3:
        return "Preventive maintenance (localized patching/leveling)"
    elif iri < 4:
        return "Surface treatment / thin overlay"
    else:
        return "Structural overlay / rehabilitation"

CONDITION_RANK = {
    "Very Good": 4, "Very Good (Smooth)": 4,
    "Good / Satisfactory": 3, "Good": 3,
    "Fair": 2,
    "Poor": 1, "Poor (Rough)": 1
}

def hybrid_condition(pci_class, iri_class):
    return pci_class if CONDITION_RANK.get(pci_class, 4) <= CONDITION_RANK.get(iri_class, 4) else iri_class

def hybrid_recommendation(combined):
    rank = CONDITION_RANK.get(combined, 4)
    if rank == 4:
        return "Routine maintenance (cleaning, grass cutting, minor touch-ups)"
    elif rank == 3:
        return "Preventive maintenance (crack sealing, local patching)"
    elif rank == 2:
        return "Surface treatment / Overlay (localized)"
    else:
        return "Major rehabilitation / Reconstruction assessment"

def color_condition(val):
    colors = {
        "Very Good": "background-color: #d5f5e3; color: #1e8449",
        "Good / Satisfactory": "background-color: #d5f5e3; color: #1e8449",
        "Good": "background-color: #d5f5e3; color: #1e8449",
        "Very Good (Smooth)": "background-color: #d5f5e3; color: #1e8449",
        "Fair": "background-color: #fef9e7; color: #d68910",
        "Poor": "background-color: #fadbd8; color: #922b21",
        "Poor (Rough)": "background-color: #fadbd8; color: #922b21"
    }
    return colors.get(val, "")

# ── HEADER ──
st.markdown("""
<div style='background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460);
            padding: 2rem; border-radius: 12px; margin-bottom: 1.5rem;'>
    <h1 style='color: white; margin:0; font-size: 2rem;'>🛣️ Digital Pavement Condition Evaluation Tool</h1>
    <p style='color: #a0aec0; margin: 0.5rem 0 0 0; font-size: 1rem;'>
        TCG633 – Bridge & Road Maintenance | UiTM Cawangan Sarawak | JKR Standard
    </p>
</div>
""", unsafe_allow_html=True)

# ── SIDEBAR ──
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    mode = st.radio("**Analysis Mode**", ["PCI Only", "IRI Only", "Hybrid (PCI + IRI)"], index=2)
    st.markdown("---")
    st.markdown("### 📋 About")
    st.info("This tool evaluates road pavement condition using PCI (visual defect survey) and IRI (roughness measurement) in accordance with JKR and ASTM D6433 standards.")
    st.markdown("---")

    # ── BENCHMARK REFERENCE TABLE IN SIDEBAR ──
    st.markdown("### 📊 Classification Benchmarks")
    st.markdown("**PCI – Pavement Condition Index**")
    pci_bench = pd.DataFrame({
        "Range": ["85 – 100", "70 – 84", "55 – 69", "0 – 54"],
        "Rating": ["Very Good", "Good / Satisfactory", "Fair", "Poor"],
        "Action": ["Routine maintenance", "Preventive maintenance", "Surface treatment", "Rehabilitation"]
    })
    st.dataframe(pci_bench, use_container_width=True, hide_index=True)

    st.markdown("**IRI – International Roughness Index (m/km)**")
    iri_bench = pd.DataFrame({
        "Range": ["< 2.0", "2.0 – 3.0", "3.0 – 4.0", "> 4.0"],
        "Rating": ["Very Good", "Good", "Fair", "Poor (Rough)"],
        "Action": ["Routine maintenance", "Preventive maintenance", "Surface treatment", "Rehabilitation"]
    })
    st.dataframe(iri_bench, use_container_width=True, hide_index=True)

# ── SESSION STATE FOR MANUAL ENTRIES ──
if "manual_sections" not in st.session_state:
    st.session_state.manual_sections = []

# ── TABS ──
if mode == "PCI Only":
    tab_labels = ["📥 Data Input", "📊 PCI Analysis"]
elif mode == "IRI Only":
    tab_labels = ["📥 Data Input", "📈 IRI Analysis"]
else:
    tab_labels = ["📥 Data Input", "📊 PCI Analysis", "📈 IRI Analysis", "🔀 Summary & Hybrid"]

tabs = st.tabs(tab_labels)

# ══════════════════════════════════════════════
# TAB 1 – DATA INPUT
# ══════════════════════════════════════════════
with tabs[0]:
    st.markdown("## 📥 Data Input")

    input_method = st.radio(
        "**Choose Input Method**",
        ["📂 Upload Excel File", "✏️ Enter Data Manually"],
        horizontal=True
    )
    st.markdown("---")

    pci_data = []
    iri_data = []

    # ══════════════════════════════════════════
    # OPTION A: UPLOAD EXCEL
    # ══════════════════════════════════════════
    if input_method == "📂 Upload Excel File":
        st.markdown("### 📂 Upload Your Excel Dataset")
        st.caption("Upload your TCG633 Excel file. The app will read the PCI_Input and IRI_Input sheets automatically.")

        uploaded_file = st.file_uploader(
            "Drop your Excel file here or click to browse",
            type=["xlsx", "xls"],
            help="Must contain sheets named 'PCI_Input' and 'IRI_Input'"
        )

        if uploaded_file is not None:
            try:
                xl = pd.ExcelFile(uploaded_file)
                available_sheets = xl.sheet_names
                st.success(f"✅ File uploaded! Sheets found: {', '.join(available_sheets)}")

                col_pci, col_iri = st.columns(2)

                with col_pci:
                    if "PCI_Input" in available_sheets:
                        df_pci_input = xl.parse("PCI_Input", header=None)
                        header_row = None
                        for idx, row in df_pci_input.iterrows():
                            if any(str(cell).strip() == "Section ID" for cell in row):
                                header_row = idx
                                break
                        if header_row is not None:
                            df_pci_input.columns = df_pci_input.iloc[header_row]
                            df_pci_input = df_pci_input.iloc[header_row+1:].reset_index(drop=True)
                            df_pci_input = df_pci_input.dropna(subset=["Section ID"])
                            df_pci_input = df_pci_input[df_pci_input["Section ID"].astype(str).str.strip().str.match(r'^\d+$')]
                            st.markdown("### 🔍 PCI Data from Excel")
                            st.dataframe(df_pci_input[["Section ID","Defect Type","Severity","Area Affected (%)"]].reset_index(drop=True),
                                         use_container_width=True, hide_index=True)
                            for _, row in df_pci_input.iterrows():
                                try:
                                    section = int(float(str(row["Section ID"]).strip()))
                                    defect = str(row["Defect Type"]).strip()
                                    severity = str(row["Severity"]).strip()
                                    area = float(str(row["Area Affected (%)"]).strip()) if str(row["Area Affected (%)"]).strip() not in ["nan",""] else 0.0
                                    if defect in DEFECT_WEIGHTS and severity in SEVERITY_FACTORS:
                                        pci_data.append({"Section": section, "Defect Type": defect,
                                                         "Severity": severity, "Area (%)": area})
                                except:
                                    continue
                        else:
                            st.warning("Could not find header row in PCI_Input sheet.")
                    else:
                        st.warning("PCI_Input sheet not found in uploaded file.")

                with col_iri:
                    if "IRI_Input" in available_sheets:
                        df_iri_input = xl.parse("IRI_Input", header=None)
                        header_row_iri = None
                        for idx, row in df_iri_input.iterrows():
                            if any(str(cell).strip() == "Section ID" for cell in row):
                                header_row_iri = idx
                                break
                        if header_row_iri is not None:
                            df_iri_input.columns = df_iri_input.iloc[header_row_iri]
                            df_iri_input = df_iri_input.iloc[header_row_iri+1:].reset_index(drop=True)
                            df_iri_input = df_iri_input.dropna(subset=["Section ID"])
                            df_iri_input = df_iri_input[df_iri_input["Section ID"].astype(str).str.strip().str.match(r'^\d+$')]
                            df_iri_input["Section ID"] = df_iri_input["Section ID"].astype(float).astype(int)
                            df_iri_input["IRI (m/km)"] = pd.to_numeric(df_iri_input["IRI (m/km)"], errors="coerce")
                            iri_avg = df_iri_input.groupby("Section ID")["IRI (m/km)"].mean().reset_index()
                            iri_avg.columns = ["Section", "IRI (m/km)"]
                            st.markdown("### 📏 IRI Data from Excel")
                            st.dataframe(iri_avg, use_container_width=True, hide_index=True)
                            for _, row in iri_avg.iterrows():
                                iri_data.append({"Section": int(row["Section"]),
                                                 "IRI (m/km)": float(row["IRI (m/km)"])})
                        else:
                            st.warning("Could not find header row in IRI_Input sheet.")
                    else:
                        st.warning("IRI_Input sheet not found in uploaded file.")

                if pci_data and iri_data:
                    st.success("✅ Data loaded successfully! Navigate to the analysis tabs to view results.")
                else:
                    st.error("❌ Could not load data. Please check your Excel file format.")

            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)}")
        else:
            st.info("👆 Please upload your Excel file to proceed.")

    # ══════════════════════════════════════════
    # OPTION B: MANUAL INPUT — FORM TEMPLATE
    # ══════════════════════════════════════════
    else:
        st.markdown("### ✏️ Add a Section")
        st.caption("Fill in one section at a time, then click **Add Section**. Your entries will appear in the table below.")

        # ── ENTRY FORM ──
        with st.form("section_form", clear_on_submit=True):
            st.markdown("#### 📋 Section Entry Form")
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**🔍 PCI – Defect Info**")
                f_section = st.number_input("Section ID", min_value=1, max_value=999, value=1, step=1,
                                            help="Unique ID for each 100m road section")
                f_defect = st.selectbox("Defect Type", DEFECT_LIST,
                                        help="Select the primary defect observed in this section. Choose 'Other (Custom)' if your defect is not listed.")
                f_custom_defect = ""
                if f_defect == "Other (Custom)":
                    f_custom_defect = st.text_input(
                        "Custom Defect Name",
                        placeholder="e.g. Depression, Delamination, Joint Failure…",
                        help="Type in the exact defect name. A default weight of 1.0 will be applied."
                    )
                    st.caption("ℹ️ Custom defects use a default weighting of **1.0**. The PCI result is an estimate.")
                f_severity = st.selectbox("Severity Level", SEVERITY_LIST,
                                          help="Low = minor, Medium = moderate, High = severe")
                f_area = st.number_input("Area Affected (%)", min_value=0.0, max_value=100.0,
                                         value=5.0, step=0.5,
                                         help="Percentage of the section's surface area affected by the defect")

            with col2:
                st.markdown("**📏 IRI – Roughness Info**")
                f_iri = st.number_input("IRI Value (m/km)", min_value=0.0, max_value=20.0,
                                        value=2.0, step=0.05,
                                        help="International Roughness Index measured for this section. Lower = smoother road.")
                st.markdown(" ")
                st.markdown(" ")

                # Quick reference hint inside form
                st.markdown("""
<div style='background:#f0f4ff; border-left: 4px solid #3498db;
            padding: 0.8rem 1rem; border-radius: 6px; font-size: 0.85rem; margin-top: 0.5rem;'>
<b>Quick Reference</b><br>
🟢 IRI &lt; 2.0 → Very Good<br>
🟡 IRI 2.0–3.0 → Good<br>
🟠 IRI 3.0–4.0 → Fair<br>
🔴 IRI &gt; 4.0 → Poor
</div>
""", unsafe_allow_html=True)

            st.markdown("---")
            submitted = st.form_submit_button("➕ Add Section", use_container_width=True, type="primary")

        if submitted:
            # Resolve the actual defect name
            actual_defect = f_custom_defect.strip() if f_defect == "Other (Custom)" and f_custom_defect.strip() else f_defect
            if f_defect == "Other (Custom)" and not f_custom_defect.strip():
                st.warning("⚠️ Please enter a custom defect name before adding.")
            else:
                # Check for duplicate section ID
                existing_ids = [s["Section"] for s in st.session_state.manual_sections]
                if f_section in existing_ids:
                    st.warning(f"⚠️ Section {f_section} already exists. Please use a different Section ID or delete the existing entry first.")
                else:
                    st.session_state.manual_sections.append({
                        "Section": int(f_section),
                        "Defect Type": actual_defect,
                        "Is Custom": f_defect == "Other (Custom)",
                        "Severity": f_severity,
                        "Area (%)": f_area,
                        "IRI (m/km)": f_iri,
                    })
                    st.success(f"✅ Section {f_section} added! Defect: **{actual_defect}**")

        # ── SECTIONS TABLE ──
        st.markdown("---")
        st.markdown("### 📊 Sections Entered So Far")

        if not st.session_state.manual_sections:
            st.info("No sections added yet. Fill in the form above and click **Add Section** to get started.")
        else:
            df_manual = pd.DataFrame(st.session_state.manual_sections).sort_values("Section").reset_index(drop=True)

            # Compute quick preview columns
            def quick_pci(row):
                w = DEFECT_WEIGHTS.get(row["Defect Type"], DEFAULT_CUSTOM_WEIGHT)
                sf = SEVERITY_FACTORS.get(row["Severity"], 1.0)
                return round(max(0, 100 - row["Area (%)"] * w * sf), 1)

            df_manual["PCI (preview)"] = df_manual.apply(quick_pci, axis=1)
            df_manual["PCI Rating"] = df_manual["PCI (preview)"].apply(lambda x: classify_pci(x)[0])
            df_manual["IRI Rating"] = df_manual["IRI (m/km)"].apply(lambda x: classify_iri(x)[0])

            display_cols = ["Section", "Defect Type", "Severity", "Area (%)", "IRI (m/km)",
                            "PCI (preview)", "PCI Rating", "IRI Rating"]

            styled_manual = df_manual[display_cols].style.map(
                color_condition, subset=["PCI Rating", "IRI Rating"]
            )
            st.dataframe(styled_manual, use_container_width=True, hide_index=True)

            st.caption(f"**{len(df_manual)} section(s) entered.** PCI (preview) is a quick estimate — full analysis is in the Analysis tabs.")

            # Delete a section
            col_del1, col_del2 = st.columns([2, 1])
            with col_del1:
                del_id = st.number_input("Delete Section ID", min_value=1, max_value=999,
                                         step=1, label_visibility="collapsed",
                                         placeholder="Enter Section ID to delete...")
            with col_del2:
                if st.button("🗑️ Delete Section", use_container_width=True):
                    before = len(st.session_state.manual_sections)
                    st.session_state.manual_sections = [
                        s for s in st.session_state.manual_sections if s["Section"] != del_id
                    ]
                    if len(st.session_state.manual_sections) < before:
                        st.success(f"Deleted Section {del_id}.")
                        st.rerun()
                    else:
                        st.warning(f"Section {del_id} not found.")

            if st.button("🗑️ Clear All Sections", type="secondary"):
                st.session_state.manual_sections = []
                st.rerun()

        # Build pci_data / iri_data from session state for analysis tabs
        for s in st.session_state.manual_sections:
            pci_data.append({
                "Section": s["Section"],
                "Defect Type": s["Defect Type"],
                "Severity": s["Severity"],
                "Area (%)": s["Area (%)"],
            })
            iri_data.append({
                "Section": s["Section"],
                "IRI (m/km)": s["IRI (m/km)"],
            })

        if st.session_state.manual_sections:
            st.success("✅ Navigate to the Analysis tabs to view full results.")

# ── COMPUTE ──
pci_results = []
for row in pci_data:
    w = DEFECT_WEIGHTS.get(row["Defect Type"], DEFAULT_CUSTOM_WEIGHT)
    sf = SEVERITY_FACTORS.get(row["Severity"], 1.0)
    deduct = row["Area (%)"] * w * sf
    pci = max(0, 100 - deduct)
    condition, color = classify_pci(pci)
    pci_results.append({
        "Section": row["Section"], "Defect Type": row["Defect Type"],
        "Severity": row["Severity"], "Area (%)": row["Area (%)"],
        "Weighting": w, "Severity Factor": sf, "Deduct Value": round(deduct, 2),
        "PCI": round(pci, 2), "Condition": condition, "Color": color,
        "Recommendation": pci_recommendation(pci)
    })

iri_results = []
for row in iri_data:
    condition, color = classify_iri(row["IRI (m/km)"])
    iri_results.append({
        "Section": row["Section"], "IRI (m/km)": row["IRI (m/km)"],
        "Condition": condition, "Color": color,
        "Recommendation": iri_recommendation(row["IRI (m/km)"])
    })

df_pci = pd.DataFrame(pci_results) if pci_results else pd.DataFrame()
df_iri = pd.DataFrame(iri_results) if iri_results else pd.DataFrame()


def benchmark_table_pci():
    """Render a color-coded PCI benchmark reference card."""
    st.markdown("""
<div style='margin-bottom:1rem;'>
<b>📖 PCI Benchmark Reference</b>
<table style='width:100%; border-collapse:collapse; font-size:0.88rem; margin-top:0.5rem;'>
<tr style='background:#2c3e50; color:white;'>
  <th style='padding:6px 10px; text-align:left;'>PCI Range</th>
  <th style='padding:6px 10px; text-align:left;'>Rating</th>
  <th style='padding:6px 10px; text-align:left;'>What it means</th>
  <th style='padding:6px 10px; text-align:left;'>Recommended Action</th>
</tr>
<tr style='background:#d5f5e3; color:#1e8449;'>
  <td style='padding:6px 10px;'>85 – 100</td><td style='padding:6px 10px;'>🟢 Very Good</td>
  <td style='padding:6px 10px;'>Pavement in excellent condition, minor or no defects</td>
  <td style='padding:6px 10px;'>Routine maintenance (cleaning, grass cutting)</td>
</tr>
<tr style='background:#eafaf1; color:#1e8449;'>
  <td style='padding:6px 10px;'>70 – 84</td><td style='padding:6px 10px;'>🟢 Good / Satisfactory</td>
  <td style='padding:6px 10px;'>Slight defects, still functional with good ride quality</td>
  <td style='padding:6px 10px;'>Preventive maintenance (crack sealing, local patching)</td>
</tr>
<tr style='background:#fef9e7; color:#d68910;'>
  <td style='padding:6px 10px;'>55 – 69</td><td style='padding:6px 10px;'>🟡 Fair</td>
  <td style='padding:6px 10px;'>Visible defects affecting ride comfort, early deterioration</td>
  <td style='padding:6px 10px;'>Surface treatment / localized overlay</td>
</tr>
<tr style='background:#fadbd8; color:#922b21;'>
  <td style='padding:6px 10px;'>0 – 54</td><td style='padding:6px 10px;'>🔴 Poor</td>
  <td style='padding:6px 10px;'>Severe defects, significant structural or surface failure</td>
  <td style='padding:6px 10px;'>Major rehabilitation / reconstruction assessment</td>
</tr>
</table>
</div>
""", unsafe_allow_html=True)


def benchmark_table_iri():
    """Render a color-coded IRI benchmark reference card."""
    st.markdown("""
<div style='margin-bottom:1rem;'>
<b>📖 IRI Benchmark Reference</b>
<table style='width:100%; border-collapse:collapse; font-size:0.88rem; margin-top:0.5rem;'>
<tr style='background:#2c3e50; color:white;'>
  <th style='padding:6px 10px; text-align:left;'>IRI (m/km)</th>
  <th style='padding:6px 10px; text-align:left;'>Rating</th>
  <th style='padding:6px 10px; text-align:left;'>What it means</th>
  <th style='padding:6px 10px; text-align:left;'>Recommended Action</th>
</tr>
<tr style='background:#d5f5e3; color:#1e8449;'>
  <td style='padding:6px 10px;'>&lt; 2.0</td><td style='padding:6px 10px;'>🟢 Very Good</td>
  <td style='padding:6px 10px;'>Very smooth ride, like a new highway surface</td>
  <td style='padding:6px 10px;'>Routine maintenance</td>
</tr>
<tr style='background:#eafaf1; color:#1e8449;'>
  <td style='padding:6px 10px;'>2.0 – 3.0</td><td style='padding:6px 10px;'>🟢 Good</td>
  <td style='padding:6px 10px;'>Slightly rough but comfortable; minor surface variation</td>
  <td style='padding:6px 10px;'>Preventive maintenance (localized patching/leveling)</td>
</tr>
<tr style='background:#fef9e7; color:#d68910;'>
  <td style='padding:6px 10px;'>3.0 – 4.0</td><td style='padding:6px 10px;'>🟡 Fair</td>
  <td style='padding:6px 10px;'>Noticeable roughness, causes driver/passenger discomfort</td>
  <td style='padding:6px 10px;'>Surface treatment / thin overlay</td>
</tr>
<tr style='background:#fadbd8; color:#922b21;'>
  <td style='padding:6px 10px;'>&gt; 4.0</td><td style='padding:6px 10px;'>🔴 Poor (Rough)</td>
  <td style='padding:6px 10px;'>Severely rough, vehicle damage risk, urgent attention needed</td>
  <td style='padding:6px 10px;'>Structural overlay / rehabilitation</td>
</tr>
</table>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# PCI TAB
# ══════════════════════════════════════════════
if mode in ["PCI Only", "Hybrid (PCI + IRI)"]:
    pci_tab = tabs[1]
    with pci_tab:
        st.markdown("## 📊 PCI Analysis Results")
        if df_pci.empty:
            st.warning("⚠️ No PCI data available. Please upload a file or enter data manually in the Data Input tab.")
        else:
            avg_pci = df_pci["PCI"].mean()
            worst = df_pci.loc[df_pci["PCI"].idxmin()]
            best = df_pci.loc[df_pci["PCI"].idxmax()]
            n = len(df_pci)
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Average PCI", f"{avg_pci:.1f}")
            k2.metric("Best Section", f"S{int(best['Section'])} ({best['PCI']:.1f})")
            k3.metric("Worst Section", f"S{int(worst['Section'])} ({worst['PCI']:.1f})")
            k4.metric("Sections Needing Attention", f"{len(df_pci[df_pci['Condition'].isin(['Poor','Fair'])])}/{n}")

            st.markdown("---")

            # Benchmark table
            benchmark_table_pci()

            st.markdown("---")
            col1, col2 = st.columns([3, 2])
            with col1:
                st.markdown("### PCI Score by Section")
                fig_bar = go.Figure()
                for _, row in df_pci.iterrows():
                    fig_bar.add_trace(go.Bar(x=[f"S{int(row['Section'])}"], y=[row["PCI"]],
                        marker_color=row["Color"], showlegend=False,
                        text=[f"{row['PCI']:.1f}"], textposition="outside"))

                # Benchmark bands as horizontal spans
                fig_bar.add_hrect(y0=85, y1=110, fillcolor="#2ecc71", opacity=0.07, line_width=0)
                fig_bar.add_hrect(y0=70, y1=85, fillcolor="#27ae60", opacity=0.07, line_width=0)
                fig_bar.add_hrect(y0=55, y1=70, fillcolor="#f39c12", opacity=0.08, line_width=0)
                fig_bar.add_hrect(y0=0, y1=55, fillcolor="#e74c3c", opacity=0.07, line_width=0)

                fig_bar.add_hline(y=85, line_dash="dash", line_color="#2ecc71",
                                  annotation_text="Very Good (85)", annotation_position="right")
                fig_bar.add_hline(y=70, line_dash="dash", line_color="#27ae60",
                                  annotation_text="Good (70)", annotation_position="right")
                fig_bar.add_hline(y=55, line_dash="dash", line_color="#f39c12",
                                  annotation_text="Fair (55)", annotation_position="right")

                fig_bar.update_layout(yaxis_range=[0, 115], yaxis_title="PCI Score",
                                      xaxis_title="Section", height=420,
                                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_bar, use_container_width=True)

            with col2:
                st.markdown("### Condition Distribution")
                cond_counts = df_pci["Condition"].value_counts().reset_index()
                cond_counts.columns = ["Condition", "Count"]
                fig_pie = px.pie(cond_counts, names="Condition", values="Count", color="Condition",
                                 color_discrete_map={"Very Good": "#2ecc71", "Good / Satisfactory": "#27ae60",
                                                     "Fair": "#f39c12", "Poor": "#e74c3c"}, hole=0.4)
                fig_pie.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_pie, use_container_width=True)

            st.markdown("### 📋 PCI Computation Table")
            display_pci = df_pci[["Section", "Defect Type", "Severity", "Area (%)", "Weighting",
                                   "Severity Factor", "Deduct Value", "PCI", "Condition", "Recommendation"]].copy()
            styled = display_pci.style.map(color_condition, subset=["Condition"])
            st.dataframe(styled, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# IRI TAB
# ══════════════════════════════════════════════
if mode == "IRI Only":
    iri_tab = tabs[1]
elif mode == "Hybrid (PCI + IRI)":
    iri_tab = tabs[2]
else:
    iri_tab = None

if iri_tab:
    with iri_tab:
        st.markdown("## 📈 IRI Analysis Results")
        if df_iri.empty:
            st.warning("⚠️ No IRI data available. Please upload a file or enter data manually in the Data Input tab.")
        else:
            avg_iri = df_iri["IRI (m/km)"].mean()
            worst_iri = df_iri.loc[df_iri["IRI (m/km)"].idxmax()]
            best_iri = df_iri.loc[df_iri["IRI (m/km)"].idxmin()]
            n = len(df_iri)
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Average IRI", f"{avg_iri:.2f} m/km")
            k2.metric("Smoothest Section", f"S{int(best_iri['Section'])} ({best_iri['IRI (m/km)']:.2f})")
            k3.metric("Roughest Section", f"S{int(worst_iri['Section'])} ({worst_iri['IRI (m/km)']:.2f})")
            k4.metric("Sections Needing Attention", f"{len(df_iri[df_iri['Condition'].isin(['Poor (Rough)', 'Fair'])])}/{n}")

            st.markdown("---")

            # Benchmark table
            benchmark_table_iri()

            st.markdown("---")
            col1, col2 = st.columns([3, 2])
            with col1:
                st.markdown("### IRI Profile Along Road")
                fig_line = go.Figure()

                # Benchmark zones
                fig_line.add_hrect(y0=0, y1=2, fillcolor="#2ecc71", opacity=0.10, line_width=0,
                                   annotation_text="Very Good", annotation_position="left")
                fig_line.add_hrect(y0=2, y1=3, fillcolor="#f1c40f", opacity=0.10, line_width=0,
                                   annotation_text="Good", annotation_position="left")
                fig_line.add_hrect(y0=3, y1=4, fillcolor="#e67e22", opacity=0.10, line_width=0,
                                   annotation_text="Fair", annotation_position="left")
                fig_line.add_hrect(y0=4, y1=max(df_iri["IRI (m/km)"].max() + 1, 6),
                                   fillcolor="#e74c3c", opacity=0.10, line_width=0,
                                   annotation_text="Poor", annotation_position="left")

                # Threshold lines
                fig_line.add_hline(y=2, line_dash="dash", line_color="#2ecc71",
                                   annotation_text="Good threshold (2.0)")
                fig_line.add_hline(y=3, line_dash="dash", line_color="#f39c12",
                                   annotation_text="Fair threshold (3.0)")
                fig_line.add_hline(y=4, line_dash="dash", line_color="#e74c3c",
                                   annotation_text="Poor threshold (4.0)")

                fig_line.add_trace(go.Scatter(
                    x=[f"S{int(r['Section'])}" for _, r in df_iri.iterrows()],
                    y=df_iri["IRI (m/km)"], mode="lines+markers+text",
                    text=[f"{v:.2f}" for v in df_iri["IRI (m/km)"]],
                    textposition="top center",
                    marker=dict(color=[r["Color"] for _, r in df_iri.iterrows()], size=12),
                    line=dict(color="#3498db", width=2)))

                fig_line.update_layout(yaxis_title="IRI (m/km)", xaxis_title="Section",
                                       height=420, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_line, use_container_width=True)

            with col2:
                st.markdown("### Condition Distribution")
                iri_cond_counts = df_iri["Condition"].value_counts().reset_index()
                iri_cond_counts.columns = ["Condition", "Count"]
                fig_pie2 = px.pie(iri_cond_counts, names="Condition", values="Count", color="Condition",
                                  color_discrete_map={"Very Good (Smooth)": "#2ecc71", "Good": "#27ae60",
                                                      "Fair": "#f39c12", "Poor (Rough)": "#e74c3c"}, hole=0.4)
                fig_pie2.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_pie2, use_container_width=True)

            st.markdown("### 📋 IRI Results Table")
            display_iri = df_iri[["Section", "IRI (m/km)", "Condition", "Recommendation"]].copy()
            styled_iri = display_iri.style.map(color_condition, subset=["Condition"])
            st.dataframe(styled_iri, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# HYBRID TAB
# ══════════════════════════════════════════════
if mode == "Hybrid (PCI + IRI)":
    with tabs[3]:
        st.markdown("## 🔀 Summary & Hybrid Index")
        if df_pci.empty or df_iri.empty:
            st.warning("⚠️ Both PCI and IRI data are required for Hybrid analysis.")
        else:
            # Match by section ID
            pci_map = {r["Section"]: r for r in pci_results}
            iri_map = {r["Section"]: r for r in iri_results}
            common_sections = sorted(set(pci_map.keys()) & set(iri_map.keys()))

            hybrid_rows = []
            for sec in common_sections:
                p = pci_map[sec]
                r = iri_map[sec]
                combined = hybrid_condition(p["Condition"], r["Condition"])
                hybrid_rows.append({
                    "Section": sec, "PCI": p["PCI"], "PCI Class": p["Condition"],
                    "IRI (m/km)": r["IRI (m/km)"], "IRI Class": r["Condition"],
                    "Combined Condition": combined, "Maintenance Recommendation": hybrid_recommendation(combined)
                })
            df_hybrid = pd.DataFrame(hybrid_rows)

            k1, k2, k3 = st.columns(3)
            k1.metric("Network Avg PCI", f"{df_hybrid['PCI'].mean():.1f}")
            k2.metric("Network Avg IRI", f"{df_hybrid['IRI (m/km)'].mean():.2f} m/km")
            critical = len(df_hybrid[df_hybrid["Combined Condition"].isin(["Poor", "Poor (Rough)"])])
            k3.metric("Critical Sections", f"{critical}/{len(df_hybrid)}")

            st.markdown("---")
            st.markdown("### PCI vs IRI Comparison")
            fig_compare = make_subplots(specs=[[{"secondary_y": True}]])
            fig_compare.add_trace(go.Bar(
                x=[f"S{int(r['Section'])}" for _, r in df_hybrid.iterrows()],
                y=df_hybrid["PCI"], name="PCI", marker_color="#3498db", opacity=0.8), secondary_y=False)
            fig_compare.add_trace(go.Scatter(
                x=[f"S{int(r['Section'])}" for _, r in df_hybrid.iterrows()],
                y=df_hybrid["IRI (m/km)"], name="IRI (m/km)",
                mode="lines+markers", marker=dict(color="#e74c3c", size=10),
                line=dict(color="#e74c3c", width=2)), secondary_y=True)

            # PCI benchmark lines
            fig_compare.add_hline(y=85, line_dash="dot", line_color="#2ecc71",
                                  annotation_text="PCI 85 (Very Good)", secondary_y=False)
            fig_compare.add_hline(y=55, line_dash="dot", line_color="#e74c3c",
                                  annotation_text="PCI 55 (Fair)", secondary_y=False)

            fig_compare.update_yaxes(title_text="PCI Score", secondary_y=False)
            fig_compare.update_yaxes(title_text="IRI (m/km)", secondary_y=True)
            fig_compare.update_layout(height=420, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_compare, use_container_width=True)

            st.markdown("### 📋 Hybrid Summary Table")
            styled_hybrid = df_hybrid.style.map(color_condition, subset=["PCI Class", "IRI Class", "Combined Condition"])
            st.dataframe(styled_hybrid, use_container_width=True, hide_index=True)

            st.markdown("### 🔧 Maintenance Priority Plan")
            priority_map = {"Poor": 1, "Poor (Rough)": 1, "Fair": 2, "Good / Satisfactory": 3,
                            "Good": 3, "Very Good": 4, "Very Good (Smooth)": 4}
            df_hybrid["Priority"] = df_hybrid["Combined Condition"].map(priority_map)
            df_priority = df_hybrid.sort_values("Priority").reset_index(drop=True)
            df_priority["Priority Level"] = df_priority["Priority"].map(
                {1: "🔴 Immediate", 2: "🟡 Short-term", 3: "🟢 Scheduled", 4: "✅ Monitor only"})
            for _, row in df_priority.iterrows():
                with st.expander(f"{row['Priority Level']} — Section {int(row['Section'])} | {row['Combined Condition']}"):
                    c1, c2, c3 = st.columns(3)
                    c1.metric("PCI", f"{row['PCI']:.1f}")
                    c2.metric("IRI", f"{row['IRI (m/km)']:.2f} m/km")
                    c3.metric("Combined", row["Combined Condition"])
                    st.info(f"**Recommended Action:** {row['Maintenance Recommendation']}")

            csv = df_hybrid.drop(columns=["Priority"]).to_csv(index=False)
            st.download_button("⬇️ Download Summary as CSV", data=csv,
                               file_name="TCG633_Pavement_Results.csv", mime="text/csv")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #718096; font-size: 0.85rem;'>
    TCG633 – Digital Pavement Condition Evaluation Tool | UiTM Cawangan Sarawak<br>
    Standards: ASTM D6433 (PCI) | JKR Pavement Maintenance Manual (IRI)
</div>
""", unsafe_allow_html=True)
