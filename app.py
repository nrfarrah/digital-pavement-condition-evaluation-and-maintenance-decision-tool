import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="TCG633 – Pavement Condition Evaluation Tool",
    page_icon="🛣️",
    layout="wide"
)

# ─────────────────────────────────────────────
# CONSTANTS & LOOKUP TABLES
# ─────────────────────────────────────────────
DEFECT_WEIGHTS = {
    "Longitudinal Crack": 1.0,
    "Alligator (Fatigue) Crack": 2.0,
    "Block Crack": 1.4,
    "Potholes": 2.2,
    "Raveling": 1.2,
    "Patching": 1.5,
    "Bleeding/Flushing": 1.0,
    "Shoving": 1.8,
    "Spalling of Longitudinal Joint": 1.3,
    "Spalling of Transverse Joints": 1.3,
}

SEVERITY_FACTORS = {"Low": 0.6, "Medium": 1.0, "High": 1.4}

DEFECT_LIST = list(DEFECT_WEIGHTS.keys())
SEVERITY_LIST = ["Low", "Medium", "High"]

# PCI classification
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

# IRI classification
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

# Hybrid: take the worse condition
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

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div style='background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460); 
            padding: 2rem; border-radius: 12px; margin-bottom: 1.5rem;'>
    <h1 style='color: white; margin:0; font-size: 2rem;'>🛣️ Digital Pavement Condition Evaluation Tool</h1>
    <p style='color: #a0aec0; margin: 0.5rem 0 0 0; font-size: 1rem;'>
        TCG633 – Bridge & Road Maintenance | UiTM Cawangan Sarawak | JKR Standard
    </p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR – MODE & SETTINGS
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    mode = st.radio("**Analysis Mode**", ["PCI Only", "IRI Only", "Hybrid (PCI + IRI)"], index=2)
    st.markdown("---")
    st.markdown("### 📋 About")
    st.info("This tool evaluates road pavement condition using PCI (visual defect survey) and IRI (roughness measurement) in accordance with JKR and ASTM D6433 standards.")
    st.markdown("---")
    st.markdown("### 📊 Classification Bands")
    st.markdown("""
**PCI**
- 85–100 → Very Good
- 70–85 → Good / Satisfactory  
- 55–70 → Fair
- 0–55 → Poor

**IRI (m/km)**
- 0–2 → Very Good
- 2–3 → Good
- 3–4 → Fair
- >4 → Poor (Rough)
""")

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📥 Data Input", "📊 PCI Analysis", "📈 IRI Analysis", "🔀 Summary & Hybrid"])

# ══════════════════════════════════════════════
# TAB 1 – DATA INPUT
# ══════════════════════════════════════════════
with tab1:
    st.markdown("## 📥 Data Input")
    col_pci, col_iri = st.columns(2)

    # ── PCI INPUT ──
    with col_pci:
        st.markdown("### 🔍 PCI – Defect Survey Data")
        st.caption("Enter one defect per section (100m). Area Affected = % of section area.")

        pci_data = []
        for i in range(1, 11):
            with st.expander(f"Section {i}", expanded=(i <= 3)):
                c1, c2, c3 = st.columns(3)
                defect = c1.selectbox("Defect Type", DEFECT_LIST, key=f"defect_{i}",
                                      index=i - 1 if i <= len(DEFECT_LIST) else 0)
                severity = c2.selectbox("Severity", SEVERITY_LIST, key=f"sev_{i}",
                                        index=["Low", "Medium", "High"].index(
                                            ["Low", "Medium", "High", "Medium", "High",
                                             "Low", "High", "Medium", "Medium", "Medium"][i - 1]))
                area = c3.number_input("Area (%)", min_value=0.0, max_value=100.0,
                                       value=[5.0, 10.0, 25.0, 3.0, 18.0,
                                              10.0, 8.0, 6.0, 6.0, 5.0][i - 1],
                                       step=0.5, key=f"area_{i}")
                pci_data.append({"Section": i, "Defect Type": defect,
                                 "Severity": severity, "Area (%)": area})

    # ── IRI INPUT ──
    with col_iri:
        st.markdown("### 📏 IRI – Roughness Measurements")
        st.caption("Enter average IRI value (m/km) per 100m section.")

        iri_defaults = [1.80, 2.05, 2.30, 2.55, 2.80, 3.05, 3.30, 3.55, 3.80, 4.05]
        iri_data = []
        for i in range(1, 11):
            with st.expander(f"Section {i}", expanded=(i <= 3)):
                iri_val = st.number_input(f"IRI (m/km)", min_value=0.0, max_value=20.0,
                                          value=iri_defaults[i - 1], step=0.05,
                                          key=f"iri_{i}")
                iri_data.append({"Section": i, "IRI (m/km)": iri_val})

    st.success("✅ Data saved! Navigate to the analysis tabs to view results.")

# ══════════════════════════════════════════════
# COMPUTE PCI & IRI
# ══════════════════════════════════════════════
pci_results = []
for row in pci_data:
    w = DEFECT_WEIGHTS.get(row["Defect Type"], 1.0)
    sf = SEVERITY_FACTORS.get(row["Severity"], 1.0)
    deduct = row["Area (%)"] * w * sf
    pci = max(0, 100 - deduct)
    condition, color = classify_pci(pci)
    rec = pci_recommendation(pci)
    pci_results.append({
        "Section": row["Section"],
        "Defect Type": row["Defect Type"],
        "Severity": row["Severity"],
        "Area (%)": row["Area (%)"],
        "Weighting": w,
        "Severity Factor": sf,
        "Deduct Value": round(deduct, 2),
        "PCI": round(pci, 2),
        "Condition": condition,
        "Color": color,
        "Recommendation": rec
    })

iri_results = []
for row in iri_data:
    condition, color = classify_iri(row["IRI (m/km)"])
    rec = iri_recommendation(row["IRI (m/km)"])
    iri_results.append({
        "Section": row["Section"],
        "IRI (m/km)": row["IRI (m/km)"],
        "Condition": condition,
        "Color": color,
        "Recommendation": rec
    })

df_pci = pd.DataFrame(pci_results)
df_iri = pd.DataFrame(iri_results)

# ══════════════════════════════════════════════
# TAB 2 – PCI ANALYSIS
# ══════════════════════════════════════════════
with tab2:
    st.markdown("## 📊 PCI Analysis Results")

    # KPI cards
    avg_pci = df_pci["PCI"].mean()
    worst = df_pci.loc[df_pci["PCI"].idxmin()]
    best = df_pci.loc[df_pci["PCI"].idxmax()]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Average PCI", f"{avg_pci:.1f}", help="Average across all 10 sections")
    k2.metric("Best Section", f"S{int(best['Section'])} ({best['PCI']:.1f})")
    k3.metric("Worst Section", f"S{int(worst['Section'])} ({worst['PCI']:.1f})")
    poor_count = len(df_pci[df_pci["Condition"].isin(["Poor", "Fair"])])
    k4.metric("Sections Needing Attention", f"{poor_count}/10")

    st.markdown("---")
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown("### PCI Score by Section")
        fig_bar = go.Figure()
        for _, row in df_pci.iterrows():
            fig_bar.add_trace(go.Bar(
                x=[f"S{int(row['Section'])}"],
                y=[row["PCI"]],
                marker_color=row["Color"],
                name=row["Condition"],
                showlegend=False,
                text=[f"{row['PCI']:.1f}"],
                textposition="outside"
            ))
        fig_bar.add_hline(y=85, line_dash="dash", line_color="#2ecc71", annotation_text="Very Good (85)")
        fig_bar.add_hline(y=70, line_dash="dash", line_color="#f39c12", annotation_text="Good (70)")
        fig_bar.add_hline(y=55, line_dash="dash", line_color="#e74c3c", annotation_text="Fair (55)")
        fig_bar.update_layout(yaxis_range=[0, 110], yaxis_title="PCI Score",
                               xaxis_title="Section", height=400,
                               plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        st.markdown("### Condition Distribution")
        cond_counts = df_pci["Condition"].value_counts().reset_index()
        cond_counts.columns = ["Condition", "Count"]
        color_map = {"Very Good": "#2ecc71", "Good / Satisfactory": "#27ae60",
                     "Fair": "#f39c12", "Poor": "#e74c3c"}
        fig_pie = px.pie(cond_counts, names="Condition", values="Count",
                         color="Condition", color_discrete_map=color_map, hole=0.4)
        fig_pie.update_layout(height=400, paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("### 📋 PCI Computation Table")
    display_pci = df_pci[["Section", "Defect Type", "Severity", "Area (%)",
                            "Weighting", "Severity Factor", "Deduct Value", "PCI", "Condition", "Recommendation"]].copy()

    def color_condition(val):
        colors = {"Very Good": "background-color: #d5f5e3; color: #1e8449",
                  "Good / Satisfactory": "background-color: #d5f5e3; color: #1e8449",
                  "Fair": "background-color: #fef9e7; color: #d68910",
                  "Poor": "background-color: #fadbd8; color: #922b21"}
        return colors.get(val, "")

    styled = display_pci.style.applymap(color_condition, subset=["Condition"])
    st.dataframe(styled, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
# TAB 3 – IRI ANALYSIS
# ══════════════════════════════════════════════
with tab3:
    st.markdown("## 📈 IRI Analysis Results")

    avg_iri = df_iri["IRI (m/km)"].mean()
    worst_iri = df_iri.loc[df_iri["IRI (m/km)"].idxmax()]
    best_iri = df_iri.loc[df_iri["IRI (m/km)"].idxmin()]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Average IRI", f"{avg_iri:.2f} m/km")
    k2.metric("Smoothest Section", f"S{int(best_iri['Section'])} ({best_iri['IRI (m/km)']:.2f})")
    k3.metric("Roughest Section", f"S{int(worst_iri['Section'])} ({worst_iri['IRI (m/km)']:.2f})")
    rough_count = len(df_iri[df_iri["Condition"].isin(["Poor (Rough)", "Fair"])])
    k4.metric("Sections Needing Attention", f"{rough_count}/10")

    st.markdown("---")
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown("### IRI Profile Along Road")
        fig_line = go.Figure()
        fig_line.add_trace(go.Scatter(
            x=[f"S{int(r['Section'])}" for _, r in df_iri.iterrows()],
            y=df_iri["IRI (m/km)"],
            mode="lines+markers+text",
            text=[f"{v:.2f}" for v in df_iri["IRI (m/km)"]],
            textposition="top center",
            marker=dict(color=[r["Color"] for _, r in df_iri.iterrows()], size=12),
            line=dict(color="#3498db", width=2)
        ))
        fig_line.add_hrect(y0=0, y1=2, fillcolor="#2ecc71", opacity=0.1, annotation_text="Very Good")
        fig_line.add_hrect(y0=2, y1=3, fillcolor="#f1c40f", opacity=0.1, annotation_text="Good")
        fig_line.add_hrect(y0=3, y1=4, fillcolor="#e67e22", opacity=0.1, annotation_text="Fair")
        fig_line.add_hrect(y0=4, y1=10, fillcolor="#e74c3c", opacity=0.1, annotation_text="Poor")
        fig_line.update_layout(yaxis_title="IRI (m/km)", xaxis_title="Section",
                                height=400, plot_bgcolor="rgba(0,0,0,0)",
                                paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_line, use_container_width=True)

    with col2:
        st.markdown("### Condition Distribution")
        iri_cond_counts = df_iri["Condition"].value_counts().reset_index()
        iri_cond_counts.columns = ["Condition", "Count"]
        iri_color_map = {"Very Good (Smooth)": "#2ecc71", "Good": "#27ae60",
                         "Fair": "#f39c12", "Poor (Rough)": "#e74c3c"}
        fig_pie2 = px.pie(iri_cond_counts, names="Condition", values="Count",
                          color="Condition", color_discrete_map=iri_color_map, hole=0.4)
        fig_pie2.update_layout(height=400, paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_pie2, use_container_width=True)

    st.markdown("### 📋 IRI Results Table")
    display_iri = df_iri[["Section", "IRI (m/km)", "Condition", "Recommendation"]].copy()

    def color_iri_condition(val):
        colors = {"Very Good (Smooth)": "background-color: #d5f5e3; color: #1e8449",
                  "Good": "background-color: #d5f5e3; color: #1e8449",
                  "Fair": "background-color: #fef9e7; color: #d68910",
                  "Poor (Rough)": "background-color: #fadbd8; color: #922b21"}
        return colors.get(val, "")

    styled_iri = display_iri.style.applymap(color_iri_condition, subset=["Condition"])
    st.dataframe(styled_iri, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
# TAB 4 – SUMMARY & HYBRID
# ══════════════════════════════════════════════
with tab4:
    st.markdown("## 🔀 Summary & Hybrid Index")

    # Build hybrid table
    hybrid_rows = []
    for i in range(10):
        p = pci_results[i]
        r = iri_results[i]
        combined = hybrid_condition(p["Condition"], r["Condition"])
        rec = hybrid_recommendation(combined)
        hybrid_rows.append({
            "Section": p["Section"],
            "PCI": p["PCI"],
            "PCI Class": p["Condition"],
            "IRI (m/km)": r["IRI (m/km)"],
            "IRI Class": r["Condition"],
            "Combined Condition": combined,
            "Maintenance Recommendation": rec
        })
    df_hybrid = pd.DataFrame(hybrid_rows)

    # Overall network health
    st.markdown("### 🏥 Network Health Overview")
    k1, k2, k3 = st.columns(3)
    avg_pci2 = df_hybrid["PCI"].mean()
    avg_iri2 = df_hybrid["IRI (m/km)"].mean()
    critical = len(df_hybrid[df_hybrid["Combined Condition"].isin(["Poor", "Poor (Rough)"])])
    k1.metric("Network Avg PCI", f"{avg_pci2:.1f}")
    k2.metric("Network Avg IRI", f"{avg_iri2:.2f} m/km")
    k3.metric("Critical Sections", f"{critical}/10", delta=f"-{critical} need rehab" if critical else "All acceptable")

    st.markdown("---")

    # Side by side PCI vs IRI chart
    st.markdown("### PCI vs IRI Comparison")
    fig_compare = make_subplots(specs=[[{"secondary_y": True}]])
    fig_compare.add_trace(go.Bar(
        x=[f"S{int(r['Section'])}" for _, r in df_hybrid.iterrows()],
        y=df_hybrid["PCI"], name="PCI", marker_color="#3498db", opacity=0.8
    ), secondary_y=False)
    fig_compare.add_trace(go.Scatter(
        x=[f"S{int(r['Section'])}" for _, r in df_hybrid.iterrows()],
        y=df_hybrid["IRI (m/km)"], name="IRI (m/km)",
        mode="lines+markers", marker=dict(color="#e74c3c", size=10),
        line=dict(color="#e74c3c", width=2)
    ), secondary_y=True)
    fig_compare.update_yaxes(title_text="PCI Score", secondary_y=False)
    fig_compare.update_yaxes(title_text="IRI (m/km)", secondary_y=True)
    fig_compare.update_layout(height=400, plot_bgcolor="rgba(0,0,0,0)",
                               paper_bgcolor="rgba(0,0,0,0)", legend=dict(x=0.01, y=0.99))
    st.plotly_chart(fig_compare, use_container_width=True)

    st.markdown("### 📋 Hybrid Summary Table")

    def color_combined(val):
        colors = {
            "Very Good": "background-color: #d5f5e3; color: #1e8449",
            "Very Good (Smooth)": "background-color: #d5f5e3; color: #1e8449",
            "Good / Satisfactory": "background-color: #d5f5e3; color: #1e8449",
            "Good": "background-color: #d5f5e3; color: #1e8449",
            "Fair": "background-color: #fef9e7; color: #d68910",
            "Poor": "background-color: #fadbd8; color: #922b21",
            "Poor (Rough)": "background-color: #fadbd8; color: #922b21"
        }
        return colors.get(val, "")

    styled_hybrid = df_hybrid.style.applymap(color_combined, subset=["PCI Class", "IRI Class", "Combined Condition"])
    st.dataframe(styled_hybrid, use_container_width=True, hide_index=True)

    st.markdown("---")
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

    st.markdown("---")
    st.markdown("### 📥 Export Results")
    csv = df_hybrid.drop(columns=["Priority"]).to_csv(index=False)
    st.download_button("⬇️ Download Summary as CSV", data=csv,
                       file_name="TCG633_Pavement_Results.csv", mime="text/csv")

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #718096; font-size: 0.85rem;'>
    TCG633 – Digital Pavement Condition Evaluation Tool | UiTM Cawangan Sarawak<br>
    Standards: ASTM D6433 (PCI) | JKR Pavement Maintenance Manual (IRI)
</div>
""", unsafe_allow_html=True)
