import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Sunstone Society Billing Automation", layout="wide")

st.title("🏢 Sunstone Co-operative Housing Society")
st.subheader("Smart Name-Token Payment Matcher")

# --- SIDEBAR: FILE UPLOADERS ---
st.sidebar.header("📁 Step 1: Upload Source Files")
bank_file = st.sidebar.file_uploader("Upload New Bank Statement (Excel)", type=["xlsx", "xls"])
directory_file = st.sidebar.file_uploader("Upload Old Completed Sheet (Excel/CSV)", type=["xlsx", "xls", "csv"])

# List of common bank junk words to ignore when matching names
BANK_JUNK = {'neft', 'imps', 'upi', 'cr', 'dr', 'ft', 'transfer', 'from', 'to', 'society', 'coop', 'housing', 'co', 'op', 'hsg', 'soc', 'ltd', 'bank', 'bccb', 'sbin', 'hdfc', 'icic', 'utib', 'payment', 'maintenance', 'oct', 'nov', 'dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep'}

def extract_name_tokens(text):
    """Cleans bank descriptions down to just unique lowercase name words."""
    if pd.isna(text):
        return set()
    # Remove numbers, punctuation, and symbols
    text_clean = re.sub(r'[^a-zA-Z\s]', ' ', str(text)).lower()
    # Split into individual words
    words = text_clean.split()
    # Filter out bank words and short initials (length <= 2)
    name_tokens = {w for w in words if w not in BANK_JUNK and len(w) > 2}
    return name_tokens

if bank_file and directory_file:
    # Load Dataframes
    dir_df = pd.read_csv(directory_file) if directory_file.name.endswith('.csv') else pd.read_excel(directory_file)
    raw_bank = pd.read_excel(bank_file)
    
    # Clean up column spaces
    dir_df.columns = dir_df.columns.str.strip()
    
    # Dynamically locate bank statement headers
    header_row_idx = 0
    for idx, row in raw_bank.iterrows():
        row_str = [str(val).strip().lower() for val in row.values]
        if 'description' in row_str or 'credit' in row_str:
            header_row_idx = idx + 1
            break
    bank_df = pd.read_excel(bank_file, skiprows=header_row_idx)
    bank_df.columns = bank_df.columns.str.strip()

    # --- SIDEBAR DROPDOWNS ---
    st.sidebar.divider()
    st.sidebar.header("⚙️ Step 2: Map Columns")
    
    bank_desc_col = st.sidebar.selectbox("New Bank Description Column:", bank_df.columns, index=1 if len(bank_df.columns) > 1 else 0)
    bank_credit_col = st.sidebar.selectbox("New Bank Credit Column:", bank_df.columns, index=2 if len(bank_df.columns) > 2 else 0)
    bank_remark_col = st.sidebar.selectbox("New Bank Remark Column:", bank_df.columns, index=3 if len(bank_df.columns) > 3 else 0)
    
    st.sidebar.divider()
    dir_desc_col = st.sidebar.selectbox("Old Sheet Description Column:", dir_df.columns, index=1 if len(dir_df.columns) > 1 else 0)
    dir_flat_col = st.sidebar.selectbox("Old Sheet FLAT NO Column:", dir_df.columns, index=3 if len(dir_df.columns) > 3 else 0)

    # Force string type to prevent pandas dtype float errors
    bank_df[bank_remark_col] = bank_df[bank_remark_col].astype(object)

    # --- MATCHING ENGINE ---
    if st.sidebar.button("🚀 Run Smart Matcher"):
        for index, row in bank_df.iterrows():
            new_desc = str(row[bank_desc_col])
            credit_val = row[bank_credit_col]
            
            # Skip rows where no money was received
            if pd.isna(credit_val) or str(credit_val).strip() == "" or float(str(credit_val).replace(',', '')) == 0:
                continue
                
            new_tokens = extract_name_tokens(new_desc)
            if not new_tokens:
                continue
                
            best_match_flat = ""
            max_shared_words = 0
            
            # Loop through old database sheet rows
            for _, dir_row in dir_df.iterrows():
                old_desc = str(dir_row[dir_desc_col])
                old_flat = str(dir_row[dir_flat_col]).strip()
                
                # --- NEW FILTER: Skip if empty or if it looks like a date/timestamp ---
                if pd.isna(old_flat) or old_flat.lower() == 'nan' or old_flat == "":
                    continue
                if "00:00:00" in old_flat or re.search(r'\d{4}-\d{2}-\d{2}', old_flat):
                    continue
                # ---------------------------------------------------------------------
                    
                old_tokens = extract_name_tokens(old_desc)
                
                # Find overlapping name words
                shared_tokens = new_tokens.intersection(old_tokens)
                
                # Match based on shared unique name words
                if len(shared_tokens) > max_shared_words and len(shared_tokens) >= 1:
                    max_shared_words = len(shared_tokens)
                    best_match_flat = old_flat.upper()
            
            # Assign the best flat match found, otherwise leave empty
            if max_shared_words >= 1:
                bank_df.at[index, bank_remark_col] = best_match_flat
            else:
                bank_df.at[index, bank_remark_col] = ""

        # --- DISPLAY & EXPORT ---
        st.write("### 📊 Cleaned & Processed Bank Statement")
        st.dataframe(bank_df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            bank_df.to_excel(writer, index=False, sheet_name='Matched Payments')
        excel_bytes = output.getvalue()
        
        st.sidebar.success("Matching Complete!")
        st.sidebar.download_button(
            label="📥 Download Corrected Excel",
            data=excel_bytes,
            file_name="Final_Society_Payments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )