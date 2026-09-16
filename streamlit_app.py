"""
EOY Nomination Report – Streamlit Web App
Upload Excel, click Generate, download PDF. Simple as that.
"""

import streamlit as st
import os
import tempfile
import shutil
import sys
import runpy
from datetime import date
from pathlib import Path

# Page config
st.set_page_config(
    page_title="EOY Report Builder",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Professional Styling
st.markdown("""
<style>
    * {
        margin: 0;
        padding: 0;
    }

    body {
        background: linear-gradient(135deg, #f5f7fa 0%, #ffffff 100%);
    }

    .main {
        max-width: 700px;
        margin: 0 auto;
        padding: 0;
    }

    /* Header */
    .header-container {
        background: linear-gradient(135deg, #1F2933 0%, #2d3e50 100%);
        padding: 40px 24px;
        border-radius: 0;
        margin-bottom: 40px;
        box-shadow: 0 4px 12px rgba(31, 41, 51, 0.15);
    }

    .header-top {
        display: flex;
        align-items: center;
        gap: 16px;
        margin-bottom: 12px;
    }

    .header-icon {
        font-size: 40px;
    }

    .header-title {
        color: white;
        font-size: 32px;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.5px;
    }

    .header-subtitle {
        color: #cbd2d9;
        font-size: 15px;
        margin-top: 8px;
        font-weight: 400;
    }

    /* Section Headers */
    .section-header {
        color: #1F2933;
        font-size: 18px;
        font-weight: 700;
        margin-top: 28px;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    .step-number {
        background: #F26522;
        color: white;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        font-weight: 700;
    }

    /* Upload Box */
    .upload-box {
        background: #f9fafb;
        border: 2px dashed #E4E7EB;
        border-radius: 8px;
        padding: 24px;
        text-align: center;
        transition: all 0.3s ease;
        margin-bottom: 16px;
    }

    .upload-box:hover {
        border-color: #F26522;
        background: #fff9f6;
    }

    /* Success Message */
    .success-box {
        background: linear-gradient(135deg, #E2F2EA 0%, #d4ebe5 100%);
        border-left: 4px solid #2E7D5B;
        padding: 16px;
        border-radius: 6px;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 12px;
    }

    .success-text {
        color: #2E7D5B;
        font-weight: 500;
        font-size: 14px;
    }

    /* Button Styling */
    .stButton > button {
        background: linear-gradient(135deg, #F26522 0%, #e8571f 100%) !important;
        color: white !important;
        border: none !important;
        padding: 12px 32px !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(242, 101, 34, 0.3) !important;
        width: 100% !important;
        height: 48px !important;
    }

    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(242, 101, 34, 0.4) !important;
    }

    .stButton > button:active {
        transform: translateY(0) !important;
    }

    /* Download Button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1F5F8B 0%, #184b6d 100%) !important;
        color: white !important;
        border: none !important;
        padding: 12px 32px !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(31, 95, 139, 0.3) !important;
        width: 100% !important;
        height: 48px !important;
    }

    .stDownloadButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(31, 95, 139, 0.4) !important;
    }

    /* Info Box */
    .info-box {
        background: #E3EEF6;
        border-left: 4px solid #1F5F8B;
        padding: 16px;
        border-radius: 6px;
        margin-bottom: 20px;
        font-size: 14px;
        color: #184b6d;
        line-height: 1.5;
    }

    .info-box ul {
        margin-left: 20px;
        margin-top: 8px;
    }

    .info-box li {
        margin-bottom: 6px;
    }

    /* Footer */
    .footer {
        text-align: center;
        color: #9AA5B1;
        font-size: 12px;
        margin-top: 40px;
        padding-top: 20px;
        border-top: 1px solid #E4E7EB;
    }

    /* Error Box */
    .error-box {
        background: #FDE9DD;
        border-left: 4px solid #F26522;
        padding: 16px;
        border-radius: 6px;
        margin-bottom: 20px;
        color: #7d3918;
        font-size: 14px;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="header-container">
    <div class="header-top">
        <div class="header-icon">📊</div>
        <h1 class="header-title">EOY Report Builder</h1>
    </div>
    <p class="header-subtitle">Generate your Employee of the Year nomination report in seconds</p>
</div>
""", unsafe_allow_html=True)

# Upload widget
st.markdown('<div class="section-header"><div class="step-number">1</div> Upload Excel File</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Choose the EOY_Report_Data.xlsx file",
    type=["xlsx"],
    help="Must have Export, Settings, and optional Awards/History sheets",
    label_visibility="collapsed"
)

if uploaded_file:
    st.markdown(f'<div class="success-box"><span style="font-size: 18px;">✓</span><span class="success-text">File ready: <strong>{uploaded_file.name}</strong></span></div>', unsafe_allow_html=True)

    # Generate button
    st.markdown('<div class="section-header"><div class="step-number">2</div> Generate Report</div>', unsafe_allow_html=True)
    if st.button("⚡ Generate Report", use_container_width=True):

        with st.spinner("🔄 Building your report... (about 10 seconds)"):
            try:
                # Save uploaded file to temp location
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
                    tmp_input.write(uploaded_file.getbuffer())
                    input_xlsx = tmp_input.name

                # Output PDF path
                output_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name

                # Get build_report.py script location
                script_dir = Path(__file__).parent
                build_script = script_dir / "build_report.py"

                if not build_script.exists():
                    # Fallback for Streamlit Cloud
                    build_script = Path("/app") / "build_report.py"

                if not build_script.exists():
                    raise FileNotFoundError(f"build_report.py not found at {script_dir}")

                # Run the generator
                sys.argv = [str(build_script), input_xlsx, output_pdf]
                runpy.run_path(str(build_script), run_name="__main__")

                # Verify output
                if not os.path.exists(output_pdf):
                    raise RuntimeError("Report generation completed but no PDF was created.")

                # Read the PDF
                with open(output_pdf, "rb") as f:
                    pdf_bytes = f.read()

                # Cleanup
                os.unlink(input_xlsx)
                os.unlink(output_pdf)

                # Success!
                st.markdown(f'<div class="success-box"><span style="font-size: 18px;">✓</span><span class="success-text">Report generated successfully!</span></div>', unsafe_allow_html=True)
                st.markdown('<div class="section-header"><div class="step-number">3</div> Download</div>', unsafe_allow_html=True)
                st.download_button(
                    label="📥 Download PDF Report",
                    data=pdf_bytes,
                    file_name=f"EOY_Nomination_Report_{date.today().isoformat()}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            except PermissionError:
                st.markdown('<div class="error-box"><strong>❌ Excel file is locked</strong><br>Close it in Excel and try again.</div>', unsafe_allow_html=True)
            except FileNotFoundError as e:
                st.markdown(f'<div class="error-box"><strong>❌ Error:</strong><br>{str(e)}<br><br>Make sure build_report.py is in the same folder as streamlit_app.py</div>', unsafe_allow_html=True)
            except RuntimeError as e:
                st.markdown(f'<div class="error-box"><strong>❌ Report generation failed:</strong><br>{str(e)}</div>', unsafe_allow_html=True)
            except Exception as e:
                st.markdown(f'<div class="error-box"><strong>❌ Unexpected error:</strong><br>{str(e)}<br><br>Check that your Excel file has the required sheets and columns.</div>', unsafe_allow_html=True)

else:
    st.markdown("""
    <div class="info-box">
        <strong>👆 Start by uploading your Excel file above</strong><br><br>
        <strong>What you need:</strong>
        <ul>
            <li><strong>EOY_Report_Data.xlsx</strong> with these sheets:</li>
            <li><strong>Export</strong> or <strong>Nominations</strong> sheet (columns: Rec. No., Achievement, E/S, Title, Department, Work Location)</li>
            <li><strong>Settings</strong> (optional: Award Year, Report Date, Prepared For, etc.)</li>
            <li><strong>Awards</strong> (optional: Rec. No., EOY, Nominator, Department, Notes)</li>
            <li><strong>History</strong> (optional: year-by-year trends)</li>
        </ul>
        <strong>Your report will include:</strong> KPIs, charts, department breakdowns, trends, and all nominations.
    </div>
    """, unsafe_allow_html=True)

st.markdown("""
<div class="footer">
    Built with Python, ReportLab, and Streamlit | Updated quarterly for executive review
</div>
""", unsafe_allow_html=True)
