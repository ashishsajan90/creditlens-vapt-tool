import base64
from datetime import datetime
import io
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import pdfplumber
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="CreditLens VAPT Alignment Engine",
    page_icon="favicon.png",
    layout="wide",
)

# --- CUSTOM CSS FOR COMPACT HEADER & CORRECT BUTTON SELECTORS ---
st.markdown(
    """
    <style>
        /* Tighten up top padding of the app container */
        .block-container {
            padding-top: 0.5rem;
            padding-bottom: 1.5rem;
        }
        /* Reduce extra margins on headers and paragraphs */
        h1 {
            margin-bottom: 0px !important;
            padding-bottom: 0px !important;
        }
        p {
            margin-bottom: 0px !important;
        }
        /* Tighten divider spacing */
        hr {
            margin-top: 0.5rem;
            margin-bottom: 0.8rem;
        }
        /* --- ACCURATE BUTTON & DOWNLOAD LINK STYLING --- */
        div.stButton > button, 
        div[data-testid="stDownloadButton"] a, 
        div[data-testid="stDownloadButton"] button {
            background-color: #1F4E78 !important;
            color: #FFFFFF !important;
            border: 1px solid #326294 !important;
            border-radius: 6px !important;
            font-weight: 500 !important;
            text-decoration: none !important;
            transition: all 0.3s ease !important;
        }
        div.stButton > button:hover, 
        div[data-testid="stDownloadButton"] a:hover, 
        div[data-testid="stDownloadButton"] button:hover {
            background-color: #2E6B9E !important;
            border-color: #4A89C5 !important;
            color: #FFFFFF !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# Load Master Tracker
@st.cache_data
def load_master_tracker():
  return pd.read_excel("CL_Vulnerability_Master_Tracker.xlsx")


# Load data safely
try:
  master_df = load_master_tracker()
except Exception as e:
  st.sidebar.error(f"❌ Error loading Master Tracker: {e}")
  master_df = pd.DataFrame()


# --- SECURE FILE SIGNATURE VALIDATION (MAGIC NUMBERS) ---
def validate_file_signature(uploaded_file):
  """Inspects the binary header bytes of an uploaded file to guarantee

  it is genuinely a PDF or Excel document, preventing extension spoofing.
  """
  try:
    header = uploaded_file.read(8)
    uploaded_file.seek(0)

    # 1. Check for PDF signature (%PDF-)
    if header.startswith(b"%PDF"):
      return True, "pdf"

    # 2. Check for modern Excel .xlsx signature (ZIP container format: PK..)
    if header.startswith(b"PK\x03\x04"):
      return True, "xlsx"

    # 3. Check for legacy Excel .xls signature (OLE Compound File format)
    if header.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
      return True, "xls"

    return False, None
  except Exception:
    return False, None


# --- SIDEBAR TEMPLATE DOWNLOADER ---
with st.sidebar:
  st.subheader("📋 Master Baseline Template")
  try:
    with open("CL_Vulnerability_Master_Tracker.xlsx", "rb") as f:
      master_file_bytes = f.read()

    st.download_button(
        label="📥 Download Master Tracker",
        data=master_file_bytes,
        file_name="CL_Vulnerability_Master_Tracker.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        help=(
            "Download the current baseline tracker used for cross-referencing."
        ),
    )
  except Exception:
    st.warning("Master tracker file not found in root directory.")


# Helper to convert local SVG to base64 for seamless HTML embedding
def get_svg_base64(path):
  try:
    with open(path, "rb") as f:
      return base64.b64encode(f.read()).decode()
  except Exception:
    return ""


logo_b64 = get_svg_base64("Logo.svg")

# --- ULTRA-COMPACT TOP DASHBOARD HEADER ---
header_col1, header_col2 = st.columns([3, 1], vertical_alignment="center")

with header_col1:
  st.markdown(
      f"""
    <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 0px;">
        <div style="height: 38px; display: flex; align-items: center; overflow: visible; flex-shrink: 0;">
            <img src="data:image/svg+xml;base64,{logo_b64}" style="height: 32px; width: auto; display: block; overflow: visible;" />
        </div>
        <h1 style="margin: 0; padding: 0; font-size: 1.7rem; font-weight: 700; line-height: 1.1;">CreditLens VAPT Alignment Engine</h1>
    </div>
    """,
      unsafe_allow_html=True,
  )
  st.markdown(
      "<p style='margin: 2px 0 0 0; color: #95a5a6; font-size: 0.85rem;'>Automated"
      " cross-referencing, semantic vulnerability matching, vendor response"
      " mapping, and executive report generation.</p>",
      unsafe_allow_html=True,
  )

with header_col2:
  if not master_df.empty:
    st.metric(
        label="Master Baseline",
        value=f"{len(master_df)} Issues",
        delta="Synced & Active",
    )

st.divider()

# --- UPLOAD SECTION IN A CLEAN CONTAINER ---
with st.container():
  st.subheader("📁 Import External Findings")

  st.info(
      "ℹ️ **Supported Formats:** Upload your raw VAPT report as an **Excel"
      " (.xlsx / .xls)** or **PDF (.pdf)** file. Ensure the vulnerability"
      " name is in the first column and Description is in the second"
      " column."
  )

  uploaded_file = st.file_uploader(
      "Upload raw VAPT report (Excel or PDF)",
      type=["xlsx", "xls", "pdf"],
      help="Supports Excel and PDF VAPT reports",
  )

bank_df = pd.DataFrame()

if uploaded_file is not None and not master_df.empty:
  # SECURITY: Validate true file signature bytes
  is_valid_sig, detected_type = validate_file_signature(uploaded_file)

  if not is_valid_sig:
    st.error(
        "❌ **Security Warning:** The uploaded file does not match a valid,"
        " untampered PDF or Excel file format."
    )
  else:
    if detected_type in ["xlsx", "xls"]:
      try:
        bank_df = pd.read_excel(uploaded_file)
      except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")

    elif detected_type == "pdf":
      with st.spinner("Extracting tables from PDF report..."):
        extracted_rows = []
        try:
          with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
              tables = page.extract_tables()
              for table in tables:
                for row in table:
                  if (
                      len(row) >= 2
                      and row[0]
                      and row[1]
                      and "vulnerability" not in str(row[0]).lower()
                  ):
                    extracted_rows.append({
                        "Vulnerability Name": str(row[0]).strip(),
                        "Vulnerability Description & Impact": str(
                            row[1]
                        ).strip(),
                    })
          bank_df = pd.DataFrame(extracted_rows)
        except Exception as e:
          st.error(f"❌ Error reading PDF file: {e}")

      if bank_df.empty:
        st.warning(
            "⚠️ Could not automatically extract structured tables from this PDF."
            " Please ensure the PDF has clean formatted tables or use an Excel"
            " sheet."
        )

  if not bank_df.empty:
    st.subheader("📥 Preview of Uploaded Findings")
    st.dataframe(bank_df.head(), use_container_width=True)

    if st.button("🚀 Run Smart Cross-Reference & Generate Response"):
      with st.spinner(
          "Analyzing findings with intelligent keyword matching against Master"
          " Tracker..."
      ):
        enriched_rows = []

        for idx, row in bank_df.iterrows():
          bank_name = str(
              row.get("Vulnerability Name", row.iloc[0] if len(row) > 0 else "")
          ).strip()
          bank_desc = str(
              row.get(
                  "Vulnerability Description & Impact",
                  row.iloc[1] if len(row) > 1 else "",
              )
          ).strip()
          bank_combined = f"{bank_name} {bank_desc}".lower()

          best_match = None
          highest_score = 0.0

          # --- SMART SEMANTIC & KEYWORD MATCHING ALGORITHM ---
          for m_idx, m_row in master_df.iterrows():
            m_name = str(m_row["Vulnerability Name"])
            m_name_lower = m_name.lower()

            # 1. Exact match
            if m_name_lower == bank_name.lower():
              best_match = m_idx
              highest_score = 1.0
              break

            # 2. Substring match
            if (
                m_name_lower in bank_name.lower()
                or bank_name.lower() in m_name_lower
            ):
              if highest_score < 0.9:
                highest_score = 0.9
                best_match = m_idx

            # 3. Keyword / Token overlap check for differently worded issues
            keywords = [
                w
                for w in m_name_lower.split()
                if len(w) > 3
                and w
                not in [
                    "with",
                    "from",
                    "this",
                    "that",
                    "and",
                    "the",
                    "missing",
                    "feature",
                ]
            ]
            if keywords:
              matched_kw = sum(1 for kw in keywords if kw in bank_combined)
              kw_score = matched_kw / len(keywords)
              if kw_score >= 0.5 and kw_score > highest_score:
                highest_score = kw_score
                best_match = m_idx

          # Evaluate match results
          if best_match is not None and highest_score >= 0.5:
            match_idx = best_match
            excel_row_num = match_idx + 2
            match_record = master_df.iloc[match_idx]

            enriched_rows.append({
                **row.to_dict(),
                "Master Tracker Row": f"Row {excel_row_num}",
                "Master Category": match_record["Category"],
                "Master Severity": match_record["Severity"],
                "Matched Master Issue": match_record["Vulnerability Name"],
                "Master Status": match_record["Status"],
                "Vendor Response": match_record["Moody's / Vendor Response"],
                "Suggested Rebuttal": match_record["Client Rebuttal / Notes"],
            })
          else:
            enriched_rows.append({
                **row.to_dict(),
                "Master Tracker Row": "New Finding",
                "Master Category": "N/A",
                "Master Severity": "N/A",
                "Matched Master Issue": "New / Unmatched Finding",
                "Master Status": "Requires Review",
                "Vendor Response": (
                    "No historical record found in Master Tracker."
                ),
                "Suggested Rebuttal": (
                    "Draft custom response or log for vendor review."
                ),
            })

        result_df = pd.DataFrame(enriched_rows)

        st.divider()
        st.subheader("📊 Enriched VAPT Report & Smart Action Matrix")
        st.dataframe(result_df, use_container_width=True)

        # --- BEAUTIFIED EXCEL EXPORT USING OPENPYXL ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          result_df.to_excel(writer, index=False, sheet_name="Enriched Report")

        output.seek(0)
        wb = load_workbook(output)
        ws = wb.active

        # Styling Definitions
        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        green_fill = PatternFill(
            start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
        )
        green_font = Font(name="Calibri", size=10, color="006100", bold=True)

        red_fill = PatternFill(
            start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
        )
        red_font = Font(name="Calibri", size=10, color="9C0006", bold=True)

        regular_font = Font(name="Calibri", size=10)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        # Locate Master Tracker Row column
        master_row_col_idx = None
        for col_num in range(1, ws.max_column + 1):
          if ws.cell(row=1, column=col_num).value == "Master Tracker Row":
            master_row_col_idx = col_num
            break

        # Format Header Row
        for col_num in range(1, ws.max_column + 1):
          cell = ws.cell(row=1, column=col_num)
          cell.fill = header_fill
          cell.font = header_font
          cell.alignment = Alignment(
              horizontal="center", vertical="center", wrap_text=True
          )
          cell.border = thin_border
        ws.row_dimensions[1].height = 28

        # Format Data Rows & Conditional Formatting
        for row_num in range(2, ws.max_row + 1):
          ws.row_dimensions[row_num].height = 20
          for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

            if col_num == master_row_col_idx:
              cell.alignment = Alignment(
                  horizontal="center", vertical="center", wrap_text=True
              )
              val = str(cell.value)
              if "Row" in val:
                cell.fill = green_fill
                cell.font = green_font
              else:
                cell.fill = red_fill
                cell.font = red_font

        # Auto-adjust column widths
        for col in ws.columns:
          max_len = 0
          col_letter = get_column_letter(col[0].column)
          for cell in col:
            if cell.value:
              val_str = str(cell.value)
              if len(val_str) > 50:
                val_str = val_str[:50]
              max_len = max(max_len, len(val_str))
          ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

        # Save workbook
        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        # Filename with Timestamp
        timestamp_str = datetime.now().strftime("%d-%b-%y_%H%M%S")
        output_filename = f"Enriched_VAPT_Report_{timestamp_str}.xlsx"

        st.success("✨ Report successfully generated!")

        st.download_button(
            label="📥 Download The Enriched Report (Excel)",
            data=final_output,
            file_name=output_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )