import pandas as pd
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
the findings against your consolidated Master Tracker, pull Moody's responses, and include the exact **Master Tracker Row number** for quick referencing.
""")


# Load Master Tracker
@st.cache_data
def load_master_tracker():
  return pd.read_excel("Moodys_CreditLens_Vulnerability_Master_Tracker.xlsx")


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

        # NOTE: regex=False is added here so parentheses/special characters are matched literally!
        matched = master_df[
            master_df["Vulnerability Name"].str.contains(
                issue_name, case=False, na=False, regex=False
            )
        ]

        if not matched.empty:
          # Get the matching row index. In Excel, row number is index + 2 (accounting for header row)
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
              "Master Tracker Row": "N/A",
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

      # Export to Excel
      output_filename = "Enriched_VAPT_Report_Output.xlsx"
      result_df.to_excel(output_filename, index=False)

      with open(output_filename, "rb") as f:
        st.download_button(
            label="📥 Download Final Enriched Report (Excel)",
            data=f,
            file_name="Enriched_VAPT_Report_Output.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )