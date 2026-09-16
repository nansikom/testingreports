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
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Styling
st.markdown("""
<style>
    .main {
        max-width: 600px;
        margin: 0 auto;
    }
    .header-section {
        border-left: 4px solid #F26522;
        padding-left: 16px;
        margin-bottom: 24px;
    }
    h1 {
        color: #1F2933;
        margin: 0;
        font-size: 28px;
    }
    .subtext {
        color: #52606D;
        font-size: 14px;
        margin-top: 4px;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="header-section">
    <h1>📊 EOY Report Builder</h1>
    <div class="subtext">Upload your Excel data and generate a clean, modernized report</div>
</div>
""", unsafe_allow_html=True)

# Upload widget
st.markdown("### Step 1: Upload Excel File")
uploaded_file = st.file_uploader(
    "Choose the EOY_Report_Data.xlsx file",
    type=["xlsx"],
    help="Must have Export, Settings, and optional Awards/History sheets"
)

if uploaded_file:
    st.success(f"✓ Loaded: {uploaded_file.name}")

    # Generate button
    st.markdown("### Step 2: Generate Report")
    if st.button("🎯 Generate Report", use_container_width=True, type="primary"):

        with st.spinner("Building your report... This takes about 10 seconds."):
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
                st.success("✅ Report generated successfully!")
                st.markdown("### Step 3: Download")
                st.download_button(
                    label="📥 Download PDF Report",
                    data=pdf_bytes,
                    file_name=f"EOY_Nomination_Report_{date.today().isoformat()}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            except PermissionError:
                st.error("❌ Excel file is locked. Close it in Excel and try again.")
            except FileNotFoundError as e:
                st.error(f"❌ Error: {str(e)}\n\nMake sure build_report.py is in the same folder as streamlit_app.py")
            except RuntimeError as e:
                st.error(f"❌ Report generation failed:\n\n{str(e)}")
            except Exception as e:
                st.error(f"❌ Unexpected error:\n\n{str(e)}")
                st.warning("Check that your Excel file has the required sheets: Export/Nominations (with Rec. No., Achievement, E/S columns)")

else:
    st.info("👆 Start by uploading your Excel file above")
    st.markdown("""
    ### What you need:
    - **EOY_Report_Data.xlsx** with these sheets:
      - **Export** or **Nominations** (columns: Rec. No., Achievement, E/S, Title, Department, Work Location)
      - **Settings** (optional: Award Year, Report Date, Prepared For, etc.)
      - **Awards** (optional: Rec. No., EOY, Nominator, Department, Notes)
      - **History** (optional: year-by-year trends)

    Your report will include KPIs, charts, department breakdowns, and all nominations.
    """)

st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #9AA5B1; font-size: 12px; margin-top: 24px;">
    Built with Python, ReportLab, and Streamlit | Updated quarterly for executive review
</div>
""", unsafe_allow_html=True)
