"""
DDR Report Generator — Web App
================================
Upload Inspection PDF + Thermal PDF → Download DDR Report
Run: streamlit run app.py
"""

import os
import sys
import json
import tempfile
import zipfile
import shutil
from pathlib import Path

import streamlit as st

# ── Page config ─────────────────────────────────────────────────
st.set_page_config(
    page_title="DDR Report Generator",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ───────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Background */
.stApp {
    background: #0D1117;
    color: #E6EDF3;
}

/* Hide Streamlit default header */
header[data-testid="stHeader"] { display: none; }
.block-container { padding-top: 2rem !important; }

/* ── Hero title ── */
.hero-title {
    font-family: 'DM Serif Display', serif;
    font-size: 2.8rem;
    color: #E6EDF3;
    line-height: 1.15;
    margin-bottom: 0.3rem;
}
.hero-accent {
    color: #F0883E;
}
.hero-sub {
    color: #8B949E;
    font-size: 1rem;
    font-weight: 300;
    margin-bottom: 2.5rem;
}

/* ── Upload cards ── */
.upload-label {
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #8B949E;
    margin-bottom: 0.4rem;
}
.upload-badge {
    display: inline-block;
    background: #1C2128;
    border: 1px solid #30363D;
    border-radius: 6px;
    padding: 4px 10px;
    font-size: 0.75rem;
    color: #8B949E;
    margin-bottom: 0.5rem;
}
.upload-badge.active {
    background: #0D2016;
    border-color: #2EA043;
    color: #3FB950;
}

/* ── Section divider ── */
.section-divider {
    border: none;
    border-top: 1px solid #21262D;
    margin: 2rem 0;
}

/* ── Status box ── */
.status-box {
    background: #161B22;
    border: 1px solid #30363D;
    border-radius: 10px;
    padding: 1.2rem 1.5rem;
    margin: 1rem 0;
}
.status-row {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 6px 0;
    font-size: 0.9rem;
    color: #8B949E;
}
.status-row.done  { color: #3FB950; }
.status-row.active { color: #F0883E; }
.status-row.wait  { color: #484F58; }

/* ── API key box ── */
.api-hint {
    background: #161B22;
    border-left: 3px solid #F0883E;
    border-radius: 4px;
    padding: 10px 14px;
    font-size: 0.82rem;
    color: #8B949E;
    margin-bottom: 1rem;
}
.api-hint a { color: #F0883E; text-decoration: none; }

/* ── Download button ── */
.stDownloadButton > button {
    background: #238636 !important;
    color: white !important;
    border: 1px solid #2EA043 !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    padding: 0.6rem 1.8rem !important;
    width: 100% !important;
    transition: background 0.2s !important;
}
.stDownloadButton > button:hover {
    background: #2EA043 !important;
}

/* ── Generate button ── */
.stButton > button {
    background: #1F4E79 !important;
    color: white !important;
    border: 1px solid #2563A8 !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    padding: 0.6rem 1.8rem !important;
    width: 100% !important;
}
.stButton > button:hover {
    background: #2563A8 !important;
    border-color: #3B82F6 !important;
}
.stButton > button:disabled {
    background: #21262D !important;
    color: #484F58 !important;
    border-color: #30363D !important;
}

/* ── Info cards ── */
.info-card {
    background: #161B22;
    border: 1px solid #21262D;
    border-radius: 10px;
    padding: 1rem 1.2rem;
}
.info-card h4 {
    color: #E6EDF3;
    margin: 0 0 0.4rem;
    font-size: 0.9rem;
}
.info-card p {
    color: #8B949E;
    font-size: 0.82rem;
    margin: 0;
    line-height: 1.5;
}

/* ── File uploader styling ── */
[data-testid="stFileUploader"] {
    background: #161B22;
    border: 1px dashed #30363D;
    border-radius: 10px;
    padding: 0.5rem;
}
[data-testid="stFileUploader"]:hover {
    border-color: #F0883E;
}

/* ── Text input ── */
[data-testid="stTextInput"] input {
    background: #161B22 !important;
    border: 1px solid #30363D !important;
    color: #E6EDF3 !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', monospace;
}

/* ── Spinner ── */
.stSpinner > div {
    border-top-color: #F0883E !important;
}

/* ── Error / success ── */
.stAlert {
    border-radius: 8px !important;
}

/* ── Progress bar ── */
.stProgress > div > div {
    background: linear-gradient(90deg, #1F4E79, #F0883E) !important;
    border-radius: 4px !important;
}
</style>
""", unsafe_allow_html=True)


# ────────────────────────────────────────────────────────────────
#  HERO SECTION
# ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-title">
    DDR Report <span class="hero-accent">Generator</span>
</div>
<div class="hero-sub">
    Upload your inspection documents → Get a professional diagnostic report in seconds
</div>
""", unsafe_allow_html=True)

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ────────────────────────────────────────────────────────────────
#  API KEY INPUT
# ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="api-hint">
    🔑 Need an API key? Get one free at
    <a href="https://console.anthropic.com" target="_blank">console.anthropic.com</a>
    — takes 2 minutes.
</div>
""", unsafe_allow_html=True)

api_key = st.text_input(
    "Anthropic API Key",
    type="password",
    placeholder="sk-ant-api03-...",
    help="Your key is never stored. It's used only for this session.",
)


st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ────────────────────────────────────────────────────────────────
#  FILE UPLOAD SECTION
# ────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2, gap="medium")

with col1:
    st.markdown('<div class="upload-label">📋 Inspection Report</div>', unsafe_allow_html=True)
    insp_file = st.file_uploader(
        "inspection",
        type=["pdf"],
        label_visibility="collapsed",
        key="inspection_upload",
    )
    if insp_file:
        st.markdown(f'<div class="upload-badge active">✅ {insp_file.name}</div>', unsafe_allow_html=True)
        st.caption(f"Size: {insp_file.size / 1024:.1f} KB")
    else:
        st.markdown('<div class="upload-badge">Drag & drop or browse</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="upload-label">🌡️ Thermal Images Report</div>', unsafe_allow_html=True)
    thermal_file = st.file_uploader(
        "thermal",
        type=["pdf"],
        label_visibility="collapsed",
        key="thermal_upload",
    )
    if thermal_file:
        st.markdown(f'<div class="upload-badge active">✅ {thermal_file.name}</div>', unsafe_allow_html=True)
        st.caption(f"Size: {thermal_file.size / 1024:.1f} KB")
    else:
        st.markdown('<div class="upload-badge">Drag & drop or browse</div>', unsafe_allow_html=True)


st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ────────────────────────────────────────────────────────────────
#  HOW IT WORKS  (collapsed info cards)
# ────────────────────────────────────────────────────────────────
with st.expander("ℹ️  How does it work?"):
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class="info-card">
            <h4>Step 1 — Extract</h4>
            <p>Each page of both PDFs is converted to an image for AI analysis.</p>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="info-card">
            <h4>Step 2 — Analyze</h4>
            <p>Claude AI reads the images, extracts observations and thermal readings.</p>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class="info-card">
            <h4>Step 3 — Generate</h4>
            <p>A structured Word report is built with images, tables and all 7 DDR sections.</p>
        </div>
        """, unsafe_allow_html=True)


# ────────────────────────────────────────────────────────────────
#  GENERATE BUTTON & PIPELINE
# ────────────────────────────────────────────────────────────────
ready = bool(api_key and insp_file and thermal_file)

if not ready:
    missing = []
    if not api_key:    missing.append("API key")
    if not insp_file:  missing.append("Inspection PDF")
    if not thermal_file: missing.append("Thermal PDF")
    if missing:
        st.caption(f"⚠️  Still needed: {', '.join(missing)}")

generate_clicked = st.button(
    "🚀  Generate DDR Report",
    disabled=not ready,
    use_container_width=True,
)

if generate_clicked and ready:

    # ── Setup temp workspace ──────────────────────────────────
    work_dir = tempfile.mkdtemp(prefix="ddr_")
    img_dir  = os.path.join(work_dir, "images")
    os.makedirs(img_dir, exist_ok=True)

    # Save uploaded PDFs to disk
    insp_path    = os.path.join(work_dir, "inspection.pdf")
    thermal_path = os.path.join(work_dir, "thermal.pdf")
    output_docx  = os.path.join(work_dir, "DDR_Report.docx")

    with open(insp_path, "wb") as f:
        f.write(insp_file.read())
    with open(thermal_path, "wb") as f:
        f.write(thermal_file.read())

    # ── Progress UI ───────────────────────────────────────────
    progress_bar = st.progress(0)
    status_area  = st.empty()

    def update_status(msg: str, pct: int):
        progress_bar.progress(pct)
        status_area.markdown(f"""
        <div class="status-box">
            <div class="status-row active">⚙️ &nbsp; {msg}</div>
        </div>
        """, unsafe_allow_html=True)

    try:
        # ── Import pipeline ───────────────────────────────────
        # Add current dir to path so ddr_generator.py is importable
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from ddr_generator import PDFExtractor, AIAnalyzer, DDRDocumentBuilder

        # Step 1: Extract images
        update_status("Extracting pages from PDFs...", 10)
        extractor      = PDFExtractor(img_dir)
        insp_images    = extractor.extract_pages(insp_path,    "insp")
        thermal_images = extractor.extract_pages(thermal_path, "thermal")

        # Step 2a: Analyze inspection report
        update_status("Claude AI is reading the Inspection Report...", 30)
        analyzer        = AIAnalyzer(api_key)
        inspection_data = analyzer.analyze_inspection_report(insp_images, extractor)

        # Step 2b: Analyze thermal report
        update_status("Claude AI is reading the Thermal Report...", 55)
        thermal_data = analyzer.analyze_thermal_report(thermal_images, extractor)

        # Merge thermal readings into areas
        thermal_readings = thermal_data.get("thermal_readings", [])
        areas = inspection_data.get("impacted_areas", [])
        for i, area in enumerate(areas):
            if i < len(thermal_readings):
                tr    = thermal_readings[i]
                hot   = tr.get("hotspot_celsius", "N/A")
                cold  = tr.get("coldspot_celsius", "N/A")
                delta = tr.get("delta_celsius", "N/A")
                area["thermal_reading"] = f"Hotspot: {hot}°C | Coldspot: {cold}°C | Delta: {delta}°C"
                area["thermal_page"]    = tr.get("page_number", i + 1)

        # Step 2c: Generate DDR content
        update_status("Generating DDR report content...", 70)
        ddr_content = analyzer.generate_ddr_content(inspection_data, thermal_data)

        # Merge severity
        severity_map = {
            s.get("area", "").lower(): s.get("severity", "HIGH")
            for s in ddr_content.get("severity_assessment", [])
        }
        ddr_areas = ddr_content.get("area_wise_observations", areas)
        for area in ddr_areas:
            area_name = area.get("area_name", "").lower()
            area["severity"]       = severity_map.get(area_name, "HIGH")
            if "thermal_reading" not in area:
                for a in areas:
                    if a.get("area_number") == area.get("area_number"):
                        area["thermal_reading"] = a.get("thermal_reading", "Not Available")
                        area["thermal_page"]    = a.get("thermal_page", 1)

        # Step 3: Build Word document
        update_status("Building Word document...", 88)
        prop_info = inspection_data.get("property_info", {})
        builder   = DDRDocumentBuilder()
        builder.build_cover(prop_info, thermal_data)
        builder.build_section1(ddr_content.get("property_issue_summary", {}), ddr_areas)
        builder.build_section2(ddr_areas, insp_images, thermal_images)
        builder.build_section3(ddr_content.get("probable_root_causes", []))
        builder.build_section4(ddr_content.get("severity_assessment", []))
        builder.build_section5(ddr_content.get("recommended_actions", []))
        builder.build_section6(ddr_content.get("additional_notes", []))
        builder.build_section7(ddr_content.get("missing_or_unclear_info", []))
        builder.save(output_docx)

        progress_bar.progress(100)
        status_area.empty()

        # ── Success UI ────────────────────────────────────────
        st.success("✅ DDR Report generated successfully!")

        # Show summary stats
        prop = inspection_data.get("property_info", {})
        s1, s2, s3 = st.columns(3)
        with s1:
            st.metric("Areas Inspected", len(ddr_areas))
        with s2:
            st.metric("Thermal Scans",   len(thermal_readings))
        with s3:
            st.metric("Inspection Score", prop.get("score", "N/A"))

        # Download button
        with open(output_docx, "rb") as f:
            st.download_button(
                label="⬇️  Download DDR Report (.docx)",
                data=f.read(),
                file_name="DDR_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

    except Exception as e:
        progress_bar.empty()
        status_area.empty()
        st.error(f"❌ Error: {str(e)}")
        with st.expander("🔍 Full error details"):
            import traceback
            st.code(traceback.format_exc())

    finally:
        # Clean up temp files
        try:
            shutil.rmtree(work_dir, ignore_errors=True)
        except Exception:
            pass


# ────────────────────────────────────────────────────────────────
#  FOOTER
# ────────────────────────────────────────────────────────────────
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center; color:#484F58; font-size:0.78rem;">
    Built with Claude AI (Anthropic) · python-docx · Streamlit
    &nbsp;·&nbsp;
    <a href="https://github.com/ayush22310418-prog/ddr-report-generator"
       style="color:#8B949E; text-decoration:none;">GitHub ↗</a>
</div>
""", unsafe_allow_html=True)
