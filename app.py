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
    page_icon="🛡️",
    layout="wide",
)

st.title("🛡️ CreditLens VAPT Report Cross-Referencer & Team Assistant")
st.markdown("""
Upload a bank's raw VAPT Excel sheet below. The app will automatically cross-reference 
the findings against your consolidated Master Tracker, pull Moody's responses, and generate a beautified report.
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

  if st.button("🚀 Run Cross-Reference & Generate Response"):
    with st.spinner("Analyzing findings against Master Tracker..."):
      enriched_rows = []

      for idx, row in bank_df.iterrows():
        issue_name = str(
            row.get("Vulnerability Name", row.get("Issue", ""))
        ).strip()

        matched = master_df[
            master_df["Vulnerability Name"].str.contains(
                issue_name, case=False, na=False, regex=False
            )
        ]

        if not matched.empty:
          match_idx = matched.index[0]
          excel_row_num = match_idx + 2
          match_record = matched.iloc[0]

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
      st.subheader("📊 Enriched VAPT Report & Action Matrix")
      st.dataframe(result_df, use_container_width=True)

      # --- BEAUTIFIED EXCEL EXPORT USING OPENPYXL ---
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Enriched Report")

      # Reload workbook via openpyxl to apply formatting
      output.seek(0)
      from openpyxl import load_workbook

      wb = load_workbook(output)
      ws = wb.active

      # Styling Definitions
      header_fill = PatternFill(
          start_color="1F4E78", end_color="1F4E78", fill_type="solid"
      )  # Professional Dark Blue
      header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

      green_fill = PatternFill(
          start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
      )  # Soft Green for Matched Rows
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

      # Find which column index is "Master Tracker Row"
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

      # Format Data Rows
      for row_num in range(2, ws.max_row + 1):
        ws.row_dimensions[row_num].height = 20
        for col_num in range(1, ws.max_column + 1):
          cell = ws.cell(row=row_num, column=col_num)
          cell.font = regular_font
          cell.border = thin_border
          cell.alignment = Alignment(vertical="center", wrap_text=True)

          # Conditional formatting for Master Tracker Row column
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

      # Auto-adjust column widths for neatness
      for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
          if cell.value:
            # Avoid long descriptions forcing excessively wide columns
            val_str = str(cell.value)
            if len(val_str) > 50:
              val_str = val_str[:50]
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

      # Save styled workbook to bytes
      final_output = io.BytesIO()
      wb.save(final_output)
      final_output.seek(0)

      # Generate Timestamp for Filename (dd-mmm-yy hh:mm:ss)
      timestamp_str = datetime.now().strftime("%d-%b-%y %H-%M-%S")
      output_filename = f"Enriched_VAPT_Report_{timestamp_str}.xlsx"

      st.success("✨ Report successfully styled and generated!")

      st.download_button(
          label="📥 Download Beautified Enriched Report (Excel)",
          data=final_output,
          file_name=output_filename,
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )