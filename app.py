from datetime import datetime
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="CreditLens VAPT Cross-Referencer",
    page_icon="favicon.ico",  # <--- Cleaned up without the hidden character
    layout="wide",
)

st.title("🛡️ CreditLens VAPT Report Cross-Referencer & Team Assistant")
st.markdown("""
Upload a bank's raw VAPT Excel sheet below. The app uses **Smart Semantic Matching** to detect 
recurring vulnerabilities even when phrased with different wording, pulls Moody's responses, and generates a beautified report.
""")


# Load Master Tracker
@st.cache_data
def load_master_tracker():
  return pd.read_excel("CL_Vulnerability_Master_Tracker.xlsx")


try:
  master_df = load_master_tracker()
  st.sidebar.success(
      f"✅ Master Tracker Loaded Successfully ({len(master_df)} known issues"
      " indexed)."
  )
except Exception as e:
  st.sidebar.error(f"❌ Error loading Master Tracker: {e}")
  master_df = pd.DataFrame()

# File Uploader for Bank VAPT Report
st.divider()
uploaded_file = st.file_uploader(
    "📂 Upload Bank VAPT Excel Report (.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None and not master_df.empty:
  bank_df = pd.read_excel(uploaded_file)

  st.subheader("📥 Preview of Uploaded Bank Findings")
  st.dataframe(bank_df.head(), use_container_width=True)

  if st.button("🚀 Run Smart Cross-Reference & Generate Response"):
    with st.spinner(
        "Analyzing findings with intelligent keyword matching against Master"
        " Tracker..."
    ):
      enriched_rows = []

      for idx, row in bank_df.iterrows():
        bank_name = str(
            row.get("Vulnerability Name", row.get("Issue", ""))
        ).strip()
        bank_desc = str(
            row.get(
                "Vulnerability Description & Impact",
                row.get("Description", ""),
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

          label_prefix = (
              "Row" if highest_score >= 0.9 else f"Row {excel_row_num} (Similar)"
          )
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
      from openpyxl import load_workbook

      wb = load_workbook(output)
      ws = wb.active

      # Styling Definitions
      header_fill = PatternFill(
          start_color="1F4E78", end_color="1F4E78", fill_type="solid"
      )  # Dark Blue
      header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

      green_fill = PatternFill(
          start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
      )  # Soft Green for Matched
      green_font = Font(name="Calibri", size=10, color="006100", bold=True)

      red_fill = PatternFill(
          start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
      )  # Soft Red for New Findings
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

      # Filename with Timestamp (dd-mmm-yy hh:mm:ss)
      timestamp_str = datetime.now().strftime("%d-%b-%y %H-%M-%S")
      output_filename = f"Enriched_VAPT_Report_{timestamp_str}.xlsx"

      st.success(
          "✨ Smart cross-reference and beautified export completed successfully!"
      )

      st.download_button(
          label="📥 Download Beautified Enriched Report (Excel)",
          data=final_output,
          file_name=output_filename,
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )