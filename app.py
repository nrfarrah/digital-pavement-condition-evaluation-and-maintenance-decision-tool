import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
import io

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

# ── PHOTO HELPERS ──
def extract_section_photos(file_obj, sheet_name="PCI_Input", id_col_letter="A"):
    """
    Extract any images embedded/pasted into an Excel sheet and map each one
    to the nearest Section ID above it in the given ID column.
    Returns: {section_id: photo_bytes}
    """
    photos = {}
    try:
        file_obj.seek(0)
        wb = openpyxl.load_workbook(file_obj)
        if sheet_name not in wb.sheetnames:
            return photos
        ws = wb[sheet_name]

        row_to_section = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.column_letter == id_col_letter and isinstance(cell.value, (int, float)):
                    row_to_section[cell.row] = int(cell.value)

        for img in getattr(ws, "_images", []):
            anchor_row = img.anchor._from.row + 1  # openpyxl anchors are 0-indexed
            candidates = [r for r in row_to_section if r <= anchor_row]
            if not candidates:
                continue
            sec = row_to_section[max(candidates)]
            try:
                photos[sec] = img._data()
            except Exception:
                continue
    except Exception:
        pass
    finally:
        try:
            file_obj.seek(0)
        except Exception:
            pass
    return photos

def show_section_photo(section_id, caption_prefix="Section", width=260):
    """Display a stored photo for a section, if one exists."""
    photo = st.session_state.section_photos.get(section_id)
    if photo:
        st.image(photo, caption=f"{caption_prefix} {section_id} – defect photo", width=width)

# ── MANUAL ENTRY VALIDATION HELPER ──
def validate_section_inputs(road_name_raw, section_raw, defect, custom_defect, severity,
                             area_raw, iri_raw, existing_ids, exclude_id=None):
    """
    Validates and parses one section's manual-entry fields.
    Returns (parsed_dict_or_None, list_of_error_strings).
    PCI info (defect/severity/area) and IRI info are all mandatory.
    """
    errors = []

    road_name = road_name_raw.strip() if road_name_raw and road_name_raw.strip() else "-"

    section_val = None
    if not section_raw or not section_raw.strip():
        errors.append("Section ID is required.")
    else:
        try:
            section_val = int(float(section_raw.strip()))
        except ValueError:
            errors.append("Section ID must be a number.")

    if section_val is not None:
        others = [sid for sid in existing_ids if sid != exclude_id]
        if section_val in others:
            errors.append(f"Section {section_val} already exists — choose a different Section ID.")

    actual_defect = custom_defect.strip() if defect == "Other (Custom)" and custom_defect.strip() else defect
    if defect == "Other (Custom)" and not custom_defect.strip():
        errors.append("Please enter a custom defect name (PCI info required).")

    area_val = None
    if not area_raw or not area_raw.strip():
        errors.append("Area Affected (%) is required (PCI info).")
    else:
        try:
            area_val = float(area_raw.strip())
            if area_val < 0:
                errors.append("Area Affected (%) cannot be negative.")
        except ValueError:
            errors.append("Area Affected (%) must be a number.")

    iri_val = None
    if not iri_raw or not iri_raw.strip():
        errors.append("IRI Value is required (IRI info).")
    else:
        try:
            iri_val = float(iri_raw.strip())
            if not (0 <= iri_val <= 100):
                errors.append("IRI Value must be between 0 and 100.")
        except ValueError:
            errors.append("IRI Value must be a number.")

    if errors:
        return None, errors

    return {
        "Section": section_val,
        "Road Name": road_name,
        "Defect Type": actual_defect,
        "Is Custom": defect == "Other (Custom)",
        "Severity": severity,
        "Area (%)": area_val,
        "IRI (m/km)": iri_val,
    }, []

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

# ── SESSION STATE ──
if "manual_sections" not in st.session_state:
    st.session_state.manual_sections = []
if "section_photos" not in st.session_state:
    st.session_state.section_photos = {}

# ── TABS ──
if mode == "PCI Only":
    tab_labels = ["📖 How to Use", "📥 Data Input", "📊 PCI Analysis"]
elif mode == "IRI Only":
    tab_labels = ["📖 How to Use", "📥 Data Input", "📈 IRI Analysis"]
else:
    tab_labels = ["📖 How to Use", "📥 Data Input", "📊 PCI Analysis", "📈 IRI Analysis", "🧭 Dashboard"]

tabs = st.tabs(tab_labels)

# ══════════════════════════════════════════════
# TAB 0 – HOW TO USE
# ══════════════════════════════════════════════
with tabs[0]:
    st.markdown("## 📖 How to Use This App")
    st.markdown("Welcome to the **TCG633 Digital Pavement Condition Evaluation Tool**. Follow the steps below to evaluate your road pavement condition using PCI and/or IRI data.")
    st.markdown("---")

    st.markdown("""
<div style='background:#eaf4fb; border-left:5px solid #3498db; padding:1rem 1.2rem; border-radius:8px; margin-bottom:1rem;'>
<h4 style='margin:0 0 0.5rem 0; color:#2471a3;'>Step 1 — Choose Your Analysis Mode</h4>
<p style='margin:0; color:#1a252f;'>
On the <b>left sidebar</b>, select one of three modes:<br><br>
🔵 <b>PCI Only</b> — if you only have visual defect survey data<br>
🔵 <b>IRI Only</b> — if you only have roughness measurement data<br>
🔵 <b>Hybrid (PCI + IRI)</b> — recommended; combines both for a more complete picture
</p>
</div>
""", unsafe_allow_html=True)

    st.markdown("""
<div style='background:#eafaf1; border-left:5px solid #2ecc71; padding:1rem 1.2rem; border-radius:8px; margin-bottom:1rem;'>
<h4 style='margin:0 0 0.5rem 0; color:#1e8449;'>Step 2 — Enter Your Data</h4>
<p style='margin:0; color:#1a252f;'>
Go to the <b>📥 Data Input</b> tab. You have two options:<br><br>
📂 <b>Upload Excel File</b> — upload a <code>.xlsx</code> file with sheets named <code>PCI_Input</code> and <code>IRI_Input</code>.<br>
&nbsp;&nbsp;&nbsp;&nbsp;• <code>PCI_Input</code> needs columns: <b>Section ID, Defect Type, Severity, Area Affected (%)</b><br>
&nbsp;&nbsp;&nbsp;&nbsp;• <code>IRI_Input</code> needs columns: <b>Section ID, IRI (m/km)</b><br>
&nbsp;&nbsp;&nbsp;&nbsp;• 📷 If you paste/insert a defect photo into a row of <code>PCI_Input</code>, the app will pick it up and show it against that Section ID.<br><br>
✏️ <b>Enter Data Manually</b> — fill in the form one section at a time and click <b>Add Section</b>.<br>
&nbsp;&nbsp;&nbsp;&nbsp;• Add a Road Name and Section ID, then type in defect, severity, area affected (%, no upper limit) and IRI (0–100)<br>
&nbsp;&nbsp;&nbsp;&nbsp;• All PCI and IRI fields are required before a section can be added<br>
&nbsp;&nbsp;&nbsp;&nbsp;• You can optionally attach a defect photo (jpg/png) per section<br>
&nbsp;&nbsp;&nbsp;&nbsp;• Use the Edit/Delete panel below the table to update or remove any section you've entered
</p>
</div>
""", unsafe_allow_html=True)

    st.markdown("""
<div style='background:#fef9e7; border-left:5px solid #f39c12; padding:1rem 1.2rem; border-radius:8px; margin-bottom:1rem;'>
<h4 style='margin:0 0 0.5rem 0; color:#d68910;'>Step 3 — View the Analysis</h4>
<p style='margin:0; color:#1a252f;'>
Once data is entered, navigate to the analysis tabs:<br><br>
📊 <b>PCI Analysis</b> — shows PCI scores per section, a condition distribution pie chart, and a full computation table<br>
📈 <b>IRI Analysis</b> — shows the IRI roughness profile, condition zones, and results table<br>
🧭 <b>Dashboard</b> — combines both into one overall condition rating per section, with a maintenance priority plan and a section filter
</p>
</div>
""", unsafe_allow_html=True)

    st.markdown("""
<div style='background:#fdf2f8; border-left:5px solid #8e44ad; padding:1rem 1.2rem; border-radius:8px; margin-bottom:1rem;'>
<h4 style='margin:0 0 0.5rem 0; color:#6c3483;'>Step 4 — Read the Results</h4>
<p style='margin:0; color:#1a252f;'>
Each section is rated using colour-coded conditions:<br><br>
🟢 <b>Very Good (PCI 85–100 / IRI &lt; 2.0)</b> — Routine maintenance only<br>
🟢 <b>Good / Satisfactory (PCI 70–84 / IRI 2.0–3.0)</b> — Preventive maintenance (crack sealing, patching)<br>
🟡 <b>Fair (PCI 55–69 / IRI 3.0–4.0)</b> — Surface treatment or localized overlay needed<br>
🔴 <b>Poor (PCI &lt; 55 / IRI &gt; 4.0)</b> — Major rehabilitation or reconstruction required<br><br>
In <b>Hybrid mode</b>, the combined condition is always the <b>worse</b> of PCI and IRI — the conservative approach used in JKR assessments.
</p>
</div>
""", unsafe_allow_html=True)

    st.markdown("""
<div style='background:#f2f3f4; border-left:5px solid #7f8c8d; padding:1rem 1.2rem; border-radius:8px; margin-bottom:1rem;'>
<h4 style='margin:0 0 0.5rem 0; color:#515a5a;'>Step 5 — Export Your Results</h4>
<p style='margin:0; color:#1a252f;'>
In the <b>🧭 Dashboard</b> tab, click <b>⬇️ Download Summary as CSV</b> to save your results for reporting or further analysis.
</p>
</div>
""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### ❓ Frequently Asked Questions")

    with st.expander("What is PCI?"):
        st.markdown("""
**Pavement Condition Index (PCI)** is a numerical rating from **0 to 100** that describes the surface condition of a road section based on visual inspection of defects.
- It is defined by **ASTM D6433** and used worldwide including under **JKR Malaysia** standards.
- Inspectors identify defect types (e.g. cracks, potholes, raveling), their severity (Low/Medium/High), and the area affected.
- The app calculates: `PCI = 100 − (Area % × Defect Weight × Severity Factor)`
""")

    with st.expander("What is IRI?"):
        st.markdown("""
**International Roughness Index (IRI)** measures road surface roughness in **m/km** (metres per kilometre).
- A lower IRI means a smoother road. A higher IRI means a rougher, bumpier surface.
- It is typically measured using a profilometer or smartphone-based tools.
- IRI < 2.0 is considered very smooth (like a new highway); IRI > 4.0 indicates a road needing urgent attention.
""")

    with st.expander("What does Hybrid / Dashboard mode do?"):
        st.markdown("""
The Dashboard combines PCI and IRI into a single **Combined Condition** rating.
- The app compares the condition class of PCI and IRI for each section.
- It always picks the **worse** of the two — this is the conservative approach recommended by JKR.
- Example: If PCI = Very Good but IRI = Fair, the Combined Condition = **Fair**.
""")

    with st.expander("Can I enter my own defect types?"):
        st.markdown("""
Yes! In the manual entry form, the **Defect Type** dropdown includes 22 standard ASTM D6433 defect types.
If your defect is not listed, select **Other (Custom)** and type the name manually.
A default weight of **1.0** will be applied — the PCI result will be an estimate.
""")

    with st.expander("How do defect photos work?"):
        st.markdown("""
There are two ways to attach a defect photo to a section:
- **Excel upload**: paste or insert a picture directly onto a row of the `PCI_Input` sheet, near that section's data. The app finds the nearest Section ID above the picture and links it automatically.
- **Manual entry**: use the photo uploader in the entry form (or the edit panel) to attach a jpg/png for that section.

Photos then appear in the Data Input table, the PCI Analysis table, and the Dashboard's Maintenance Priority Plan.
""")

    with st.expander("Why does my Excel file not load all sections?"):
        st.markdown("""
Make sure your Excel file:
- Has sheets named exactly **`PCI_Input`** and **`IRI_Input`** (case-sensitive)
- Has a header row with columns: **Section ID**, **Defect Type**, **Severity**, **Area Affected (%)**
- Uses **Section ID** as a plain number (e.g. 1, 2, 3 — not "Section 1")
- Uses **Severity** values of exactly: `Low`, `Medium`, or `High`

If a row is skipped, the app will show a warning message explaining why.
""")


# ══════════════════════════════════════════════
# TAB 1 – DATA INPUT
# ══════════════════════════════════════════════
with tabs[1]:
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
        st.caption("Upload your TCG633 Excel file. The app will read the PCI_Input and IRI_Input sheets automatically, and pick up any defect photos pasted into PCI_Input.")

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

                photo_bytes = io.BytesIO(uploaded_file.getvalue())
                extracted_photos = extract_section_photos(photo_bytes, sheet_name="PCI_Input")
                if extracted_photos:
                    st.session_state.section_photos.update(extracted_photos)
                    st.success(f"📷 Found {len(extracted_photos)} defect photo(s) in PCI_Input, linked to sections: {', '.join(str(s) for s in sorted(extracted_photos))}")

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
                            df_pci_input = df_pci_input[df_pci_input["Section ID"].astype(str).str.strip().str.match(r'^\d+(\.\d+)?$')]
                            st.markdown("### 🔍 PCI Data from Excel")
                            st.dataframe(df_pci_input[["Section ID","Defect Type","Severity","Area Affected (%)"]].reset_index(drop=True),
                                         use_container_width=True, hide_index=True)
                            skipped_pci = []
                            for _, row in df_pci_input.iterrows():
                                try:
                                    section = int(float(str(row["Section ID"]).strip()))
                                    defect = str(row["Defect Type"]).strip()
                                    severity = str(row["Severity"]).strip().capitalize()
                                    area = float(str(row["Area Affected (%)"]).strip()) if str(row["Area Affected (%)"]).strip() not in ["nan",""] else 0.0
                                    if severity not in SEVERITY_FACTORS:
                                        skipped_pci.append(f"S{section}: unrecognised severity '{severity}' (use Low/Medium/High)")
                                        continue
                                    pci_data.append({"Section": section, "Road Name": "-", "Defect Type": defect,
                                                     "Severity": severity, "Area (%)": area})
                                except Exception as ex:
                                    skipped_pci.append(f"Row skipped: {ex}")
                                    continue
                            if skipped_pci:
                                st.warning("\u26a0\ufe0f Some PCI rows were skipped:\n" + "\n".join(f"- {s}" for s in skipped_pci))

                            if extracted_photos:
                                st.markdown("#### 📷 Defect Photos")
                                gallery_cols = st.columns(min(4, len(extracted_photos)))
                                for i, (sec, img) in enumerate(sorted(extracted_photos.items())):
                                    with gallery_cols[i % len(gallery_cols)]:
                                        st.image(img, caption=f"Section {sec}", use_container_width=True)
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
                            df_iri_input = df_iri_input[df_iri_input["Section ID"].astype(str).str.strip().str.match(r'^\d+(\.\d+)?$')]
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
    # OPTION B: MANUAL INPUT
    # ══════════════════════════════════════════
    else:
        st.markdown("### ✏️ Add a Section")
        st.caption("Fill in one section at a time, then click **Add Section**. All PCI and IRI fields are required.")

        with st.form("section_form", clear_on_submit=True):
            st.markdown("#### 📋 Section Entry Form")
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**📍 Location**")
                f_road_name = st.text_input("Road Name", placeholder="e.g. Jalan Kolej, Persiaran Ilmu")
                f_section_raw = st.text_input("Section ID *", placeholder="e.g. 1")

                st.markdown("**🔍 PCI – Defect Info (required)**")
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
                f_area_raw = st.text_input("Area Affected (%) *", placeholder="e.g. 5 — type any value, no upper limit",
                                           help="Percentage of the section's surface area affected by the defect. No range limit.")
                f_photo = st.file_uploader("📷 Defect Photo (optional)", type=["jpg", "jpeg", "png"],
                                           help="Attach a site photo of the observed defect for this section")

            with col2:
                st.markdown("**📏 IRI – Roughness Info (required)**")
                f_iri_raw = st.text_input("IRI Value (m/km) *", placeholder="e.g. 2.0 — accepted range 0 to 100",
                                          help="International Roughness Index measured for this section (0–100). Lower = smoother road.")
                if f_photo is not None:
                    st.image(f_photo, caption="Photo preview", use_container_width=True)

                st.markdown("""
<div style='background:#f0f4ff; border-left: 4px solid #3498db;
            padding: 0.8rem 1rem; border-radius: 6px; font-size: 0.85rem; margin-top: 0.5rem;'>
<b>Quick Reference</b><br>
🟢 IRI &lt; 2.0 → Very Good<br>
🟡 IRI 2.0–3.0 → Good<br>
🟠 IRI 3.0–4.0 → Fair<br>
🔴 IRI &gt; 4.0 → Poor<br><br>
<i>* Fields marked with an asterisk are required — all PCI and IRI info must be filled before you can add the section.</i>
</div>
""", unsafe_allow_html=True)

            st.markdown("---")
            submitted = st.form_submit_button("➕ Add Section", use_container_width=True, type="primary")

        if submitted:
            existing_ids = [s["Section"] for s in st.session_state.manual_sections]
            parsed, errors = validate_section_inputs(
                f_road_name, f_section_raw, f_defect, f_custom_defect, f_severity,
                f_area_raw, f_iri_raw, existing_ids
            )
            if errors:
                for e in errors:
                    st.error(f"⚠️ {e}")
            else:
                st.session_state.manual_sections.append(parsed)
                if f_photo is not None:
                    st.session_state.section_photos[parsed["Section"]] = f_photo.getvalue()
                st.success(f"✅ Section {parsed['Section']} added! Defect: **{parsed['Defect Type']}**")

        # ── SECTIONS TABLE ──
        st.markdown("---")
        st.markdown("### 📊 Sections Entered So Far")

        if not st.session_state.manual_sections:
            st.info("No sections added yet. Fill in the form above and click **Add Section** to get started.")
        else:
            df_manual = pd.DataFrame(st.session_state.manual_sections).sort_values("Section").reset_index(drop=True)
            if "Road Name" not in df_manual.columns:
                df_manual["Road Name"] = "-"
            df_manual["Road Name"] = df_manual["Road Name"].fillna("-")

            def quick_pci(row):
                w = DEFECT_WEIGHTS.get(row["Defect Type"], DEFAULT_CUSTOM_WEIGHT)
                sf = SEVERITY_FACTORS.get(row["Severity"], 1.0)
                return round(max(0, 100 - row["Area (%)"] * w * sf), 1)

            df_manual["PCI (preview)"] = df_manual.apply(quick_pci, axis=1)
            df_manual["PCI Rating"] = df_manual["PCI (preview)"].apply(lambda x: classify_pci(x)[0])
            df_manual["IRI Rating"] = df_manual["IRI (m/km)"].apply(lambda x: classify_iri(x)[0])
            df_manual["Photo"] = df_manual["Section"].apply(lambda s: "📷 Yes" if s in st.session_state.section_photos else "—")

            display_cols = ["Section", "Road Name", "Defect Type", "Severity", "Area (%)", "IRI (m/km)",
                            "PCI (preview)", "PCI Rating", "IRI Rating", "Photo"]

            styled_manual = df_manual[display_cols].style.map(
                color_condition, subset=["PCI Rating", "IRI Rating"]
            )
            st.dataframe(styled_manual, use_container_width=True, hide_index=True)

            st.caption(f"**{len(df_manual)} section(s) entered.** PCI (preview) is a quick estimate — full analysis is in the Analysis tabs.")

            sections_with_photo = [s for s in df_manual["Section"] if s in st.session_state.section_photos]
            if sections_with_photo:
                with st.expander(f"📷 View Defect Photos ({len(sections_with_photo)})"):
                    gallery_cols = st.columns(min(4, len(sections_with_photo)))
                    for i, sec in enumerate(sections_with_photo):
                        with gallery_cols[i % len(gallery_cols)]:
                            st.image(st.session_state.section_photos[sec], caption=f"Section {sec}", use_container_width=True)

            # ── EDIT / DELETE A SECTION ──
            st.markdown("---")
            st.markdown("#### ✏️ Edit or Delete a Section")
            section_options = [s["Section"] for s in st.session_state.manual_sections]
            label_map = {s["Section"]: f"Section {s['Section']} – {s.get('Road Name', '-')} ({s['Defect Type']})"
                         for s in st.session_state.manual_sections}
            selected_section = st.selectbox("Select a section to edit or delete", options=section_options,
                                            format_func=lambda s: label_map[s], key="edit_select")
            target = next(s for s in st.session_state.manual_sections if s["Section"] == selected_section)

            with st.form("edit_section_form"):
                st.caption(f"Editing Section {target['Section']}")
                e_col1, e_col2 = st.columns(2)
                with e_col1:
                    e_road_name = st.text_input("Road Name", value=target.get("Road Name", "-"))
                    e_section_raw = st.text_input("Section ID *", value=str(target["Section"]))
                    default_idx = DEFECT_LIST.index(target["Defect Type"]) if target["Defect Type"] in DEFECT_LIST else DEFECT_LIST.index("Other (Custom)")
                    e_defect = st.selectbox("Defect Type", DEFECT_LIST, index=default_idx, key="edit_defect")
                    e_custom_defect = ""
                    if e_defect == "Other (Custom)":
                        e_custom_defect = st.text_input("Custom Defect Name",
                                                        value=target["Defect Type"] if target.get("Is Custom") else "",
                                                        key="edit_custom_defect")
                    e_severity = st.selectbox("Severity Level", SEVERITY_LIST,
                                              index=SEVERITY_LIST.index(target["Severity"]), key="edit_severity")
                    e_area_raw = st.text_input("Area Affected (%) *", value=str(target["Area (%)"]))
                    e_photo = st.file_uploader("Replace Defect Photo (optional)", type=["jpg", "jpeg", "png"], key="edit_photo")
                with e_col2:
                    e_iri_raw = st.text_input("IRI Value (m/km) *", value=str(target["IRI (m/km)"]))
                    if selected_section in st.session_state.section_photos:
                        st.image(st.session_state.section_photos[selected_section], caption="Current photo", use_container_width=True)
                    else:
                        st.caption("No photo attached to this section yet.")

                col_save, col_del = st.columns(2)
                save_clicked = col_save.form_submit_button("💾 Save Changes", use_container_width=True, type="primary")
                delete_clicked = col_del.form_submit_button("🗑️ Delete This Section", use_container_width=True)

            if save_clicked:
                existing_ids = [s["Section"] for s in st.session_state.manual_sections]
                parsed, errors = validate_section_inputs(
                    e_road_name, e_section_raw, e_defect, e_custom_defect, e_severity,
                    e_area_raw, e_iri_raw, existing_ids, exclude_id=selected_section
                )
                if errors:
                    for e in errors:
                        st.error(f"⚠️ {e}")
                else:
                    idx = next(i for i, s in enumerate(st.session_state.manual_sections) if s["Section"] == selected_section)
                    st.session_state.manual_sections[idx] = parsed
                    if parsed["Section"] != selected_section and selected_section in st.session_state.section_photos:
                        st.session_state.section_photos[parsed["Section"]] = st.session_state.section_photos.pop(selected_section)
                    if e_photo is not None:
                        st.session_state.section_photos[parsed["Section"]] = e_photo.getvalue()
                    st.success(f"✅ Section {parsed['Section']} updated!")
                    st.rerun()

            if delete_clicked:
                st.session_state.manual_sections = [s for s in st.session_state.manual_sections if s["Section"] != selected_section]
                st.session_state.section_photos.pop(selected_section, None)
                st.success(f"Deleted Section {selected_section}.")
                st.rerun()

            if st.button("🗑️ Clear All Sections", type="secondary"):
                st.session_state.manual_sections = []
                st.session_state.section_photos = {}
                st.rerun()

        for s in st.session_state.manual_sections:
            pci_data.append({
                "Section": s["Section"],
                "Road Name": s.get("Road Name", "-"),
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
        "Section": row["Section"], "Road Name": row.get("Road Name", "-"), "Defect Type": row["Defect Type"],
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
    pci_tab = tabs[2]
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
            display_pci = df_pci[["Section", "Road Name", "Defect Type", "Severity", "Area (%)", "Weighting",
                                   "Severity Factor", "Deduct Value", "PCI", "Condition", "Recommendation"]].copy()
            styled = display_pci.style.map(color_condition, subset=["Condition"])
            st.dataframe(styled, use_container_width=True, hide_index=True)

            sections_with_photos = [int(s) for s in df_pci["Section"] if int(s) in st.session_state.section_photos]
            if sections_with_photos:
                st.markdown("### 📷 Defect Photos by Section")
                for sec in sections_with_photos:
                    row = df_pci[df_pci["Section"] == sec].iloc[0]
                    with st.expander(f"Section {sec} — {row['Condition']} (PCI {row['PCI']:.1f})"):
                        c1, c2 = st.columns([1, 2])
                        with c1:
                            show_section_photo(sec)
                        with c2:
                            st.write(f"**Road Name:** {row['Road Name']}")
                            st.write(f"**Defect:** {row['Defect Type']}")
                            st.write(f"**Severity:** {row['Severity']}")
                            st.write(f"**Area Affected:** {row['Area (%)']}%")
                            st.write(f"**Recommendation:** {row['Recommendation']}")


# ══════════════════════════════════════════════
# IRI TAB
# ══════════════════════════════════════════════
if mode == "IRI Only":
    iri_tab = tabs[2]
elif mode == "Hybrid (PCI + IRI)":
    iri_tab = tabs[3]
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
            benchmark_table_iri()

            st.markdown("---")
            col1, col2 = st.columns([3, 2])
            with col1:
                st.markdown("### IRI Profile Along Road")
                fig_line = go.Figure()

                fig_line.add_hrect(y0=0, y1=2, fillcolor="#2ecc71", opacity=0.10, line_width=0,
                                   annotation_text="Very Good", annotation_position="left")
                fig_line.add_hrect(y0=2, y1=3, fillcolor="#f1c40f", opacity=0.10, line_width=0,
                                   annotation_text="Good", annotation_position="left")
                fig_line.add_hrect(y0=3, y1=4, fillcolor="#e67e22", opacity=0.10, line_width=0,
                                   annotation_text="Fair", annotation_position="left")
                fig_line.add_hrect(y0=4, y1=max(df_iri["IRI (m/km)"].max() + 1, 6),
                                   fillcolor="#e74c3c", opacity=0.10, line_width=0,
                                   annotation_text="Poor", annotation_position="left")

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
# DASHBOARD TAB (formerly Summary & Hybrid)
# ══════════════════════════════════════════════
if mode == "Hybrid (PCI + IRI)":
    with tabs[4]:
        st.markdown("## 🧭 Dashboard")
        st.caption("Combined PCI + IRI summary, maintenance priority plan, and section filter.")
        if df_pci.empty or df_iri.empty:
            st.warning("⚠️ Both PCI and IRI data are required for the Dashboard.")
        else:
            pci_map = {r["Section"]: r for r in pci_results}
            iri_map = {r["Section"]: r for r in iri_results}
            common_sections = sorted(set(pci_map.keys()) & set(iri_map.keys()))

            hybrid_rows = []
            for sec in common_sections:
                p = pci_map[sec]
                r = iri_map[sec]
                combined = hybrid_condition(p["Condition"], r["Condition"])
                hybrid_rows.append({
                    "Section": sec, "Road Name": p.get("Road Name", "-"), "PCI": p["PCI"], "PCI Class": p["Condition"],
                    "IRI (m/km)": r["IRI (m/km)"], "IRI Class": r["Condition"],
                    "Combined Condition": combined, "Maintenance Recommendation": hybrid_recommendation(combined)
                })
            df_hybrid_full = pd.DataFrame(hybrid_rows)

            st.markdown("### 🔎 Filter Sections")
            all_sections = sorted(df_hybrid_full["Section"].unique())
            selected_sections = st.multiselect(
                "Choose which section(s) to display on the dashboard",
                options=all_sections, default=all_sections,
                format_func=lambda s: f"Section {s}"
            )

            if not selected_sections:
                st.info("Select at least one section above to view the dashboard.")
            else:
                df_hybrid = df_hybrid_full[df_hybrid_full["Section"].isin(selected_sections)].reset_index(drop=True)

                st.markdown("---")
                k1, k2, k3 = st.columns(3)
                k1.metric("Avg PCI (filtered)", f"{df_hybrid['PCI'].mean():.1f}")
                k2.metric("Avg IRI (filtered)", f"{df_hybrid['IRI (m/km)'].mean():.2f} m/km")
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
                    sec = int(row["Section"])
                    with st.expander(f"{row['Priority Level']} — Section {sec} | {row['Road Name']} | {row['Combined Condition']}"):
                        c1, c2, c3 = st.columns(3)
                        c1.metric("PCI", f"{row['PCI']:.1f}")
                        c2.metric("IRI", f"{row['IRI (m/km)']:.2f} m/km")
                        c3.metric("Combined", row["Combined Condition"])
                        st.info(f"**Recommended Action:** {row['Maintenance Recommendation']}")
                        if sec in st.session_state.section_photos:
                            show_section_photo(sec, width=320)

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
