import os
import re
import glob
import pandas as pd
import pdfplumber
import streamlit as st

# ==========================================
# 1. PRICING PIPELINE LOGIC
# ==========================================

def calculate_customer_investment(base_equipment_price, misc_cost, labor_cost, margin_decimal, markup_pct, tax_rate_pct):
    """Calculates final customer investment from base equipment price."""
    if pd.isna(base_equipment_price) or base_equipment_price is None or base_equipment_price <= 0:
        return 0.0, 0.0, 0.0, 0.0, 0.0

    tax_decimal = tax_rate_pct / 100.0

    # Step 1: Equipment + Misc Cost
    equipment_misc_subtotal = base_equipment_price + misc_cost
    
    # Step 2: Sales Tax (applied to Equipment + Misc)
    tax_amount = equipment_misc_subtotal * tax_decimal
    post_tax_subtotal = equipment_misc_subtotal + tax_amount
    
    # Step 3: Add Labor Cost
    subtotal_with_labor = post_tax_subtotal + labor_cost
    
    # Step 4: Divide by Margin Decimal
    effective_margin = max(margin_decimal, 0.01) 
    total_cost_base = subtotal_with_labor / effective_margin
    
    # Step 5: Increase by Markup %
    final_customer_investment = total_cost_base * (1 + markup_pct)
    
    return (
        equipment_misc_subtotal,
        tax_amount,
        subtotal_with_labor,
        total_cost_base,
        final_customer_investment
    )

# ==========================================
# 2. PARSING & EXTRACTION LOGIC
# ==========================================

def extract_number_from_cell(cell_val):
    if pd.isna(cell_val):
        return None
    cell_str = str(cell_val).strip()
    clean_str = re.sub(r'[^\d.]', '', cell_str)
    try:
        val = float(clean_str)
        return val if val > 10 else None
    except ValueError:
        return None


def extract_all_numbers_from_str(text):
    if not text or not isinstance(text, str):
        return []

    strict_dollar_pattern = r'\$\s*([\d,]+(?:\.\d{1,2})?)'
    matches = re.findall(strict_dollar_pattern, text)

    found_vals = []
    for m in matches:
        clean_str = m.replace(',', '').strip()
        if not clean_str:
            continue
        try:
            val = float(clean_str)
            if val > 10:
                found_vals.append(val)
        except ValueError:
            continue

    return found_vals


def process_excel_csv_vendor_file(file_path):
    records = []
    vendor_name = os.path.basename(file_path)

    try:
        if file_path.lower().endswith('.csv'):
            sheets_dict = {'CSV': pd.read_csv(file_path, header=None, dtype=str)}
        else:
            sheets_dict = pd.read_excel(file_path, header=None, sheet_name=None, dtype=str)

        for sheet_name, df in sheets_dict.items():
            if df.empty:
                continue

            for idx, row in df.iterrows():
                row_vals = row.dropna().tolist()
                if not row_vals:
                    continue

                row_str = " ".join([str(v) for v in row_vals])

                ahri_matches = re.findall(r'(?<!\d)\d{9}(?!\d)', row_str)
                if not ahri_matches:
                    continue

                last_cell = str(row_vals[-1]).strip()
                system_price = extract_number_from_cell(last_cell)

                if system_price is None:
                    extracted = extract_all_numbers_from_str(last_cell)
                    if extracted:
                        system_price = extracted[-1]

                if system_price is not None:
                    for ahri in set(ahri_matches):
                        records.append({
                            'AHRI Number': ahri,
                            'Vendor Sheet': vendor_name,
                            'System Price': system_price
                        })

    except Exception as e:
        st.sidebar.error(f"❌ Error reading `{vendor_name}`: {e}")

    return records


def process_pdf_vendor_file(file_path):
    records = []
    vendor_name = os.path.basename(file_path)

    try:
        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                try:
                    text = page.extract_text(layout=False) or page.extract_text(layout=True)
                    if text:
                        for line in text.split('\n'):
                            clean_line = " ".join(line.split())
                            ahri_matches = re.findall(r'(?<!\d)\d{9}(?!\d)', clean_line)
                            prices = extract_all_numbers_from_str(clean_line)

                            if ahri_matches and prices:
                                system_price = prices[-1]
                                for ahri in set(ahri_matches):
                                    records.append({
                                        'AHRI Number': ahri,
                                        'Vendor Sheet': vendor_name,
                                        'System Price': system_price
                                    })
                except Exception:
                    continue
    except Exception as pdf_err:
        st.sidebar.error(f"❌ Could not open PDF `{vendor_name}`: {pdf_err}")

    return records


@st.cache_data
def load_and_build_vendor_db(vendor_folder):
    all_records = []
    processed_files = []
    failed_or_empty_files = []

    if not os.path.exists(vendor_folder):
        return pd.DataFrame(), processed_files, failed_or_empty_files

    search_path = os.path.join(vendor_folder, "**", "*")
    file_paths = glob.glob(search_path, recursive=True)

    for path in file_paths:
        if os.path.isdir(path):
            continue

        file_name = os.path.basename(path)

        # Skip temporary Excel lock files (~$)
        if file_name.startswith("~$"):
            continue

        ext = os.path.splitext(path)[1].lower()
        file_records = []

        if ext == '.pdf':
            file_records = process_pdf_vendor_file(path)
        elif ext in ['.xlsx', '.xls', '.csv']:
            file_records = process_excel_csv_vendor_file(path)

        if file_records:
            all_records.extend(file_records)
            processed_files.append(f"{file_name} ({len(file_records)} records)")
        elif ext in ['.pdf', '.xlsx', '.xls', '.csv']:
            failed_or_empty_files.append(file_name)

    if not all_records:
        return pd.DataFrame(), processed_files, failed_or_empty_files

    df = pd.DataFrame(all_records)
    df = df.sort_values(by='System Price', ascending=False)
    df = df.drop_duplicates(subset=['AHRI Number', 'Vendor Sheet'], keep='first')

    return df, processed_files, failed_or_empty_files


# ==========================================
# 3. STREAMLIT INTERFACE
# ==========================================

APP_PASSWORD = "Pr3$t1g375098!"  # Set your desired password here

def main():
    st.set_page_config(page_title="HVAC Unit Price Estimator", layout="wide")
    st.title("⚡ HVAC Unit Price Estimator")
    st.subheader("Field Tech Look-Up Portal")

    # ------------------------------------------
    # SIDEBAR AUTHENTICATION
    # ------------------------------------------
    st.sidebar.header("🔒 Access Control")
    password_input = st.sidebar.text_input("Enter Password", type="password")

    if password_input != APP_PASSWORD:
        if password_input:
            st.sidebar.error("Incorrect Password")
        else:
            st.sidebar.info("Please enter the password to access controls.")
        st.warning("👈 Please enter the sidebar password to unlock the estimator settings and data.")
        return

    st.sidebar.success("Access Granted")
    st.sidebar.markdown("---")

    # ------------------------------------------
    # SIDEBAR CONTROLS (UNLOCKED)
    # ------------------------------------------
    st.sidebar.header("⚙️ Pricing & Margin Setup")

    misc_cost = st.sidebar.number_input(
        "Misc Cost ($)", 
        min_value=0.0, 
        value=700.0, 
        step=50.0
    )

    tax_rate_pct = st.sidebar.number_input(
        "Sales Tax Rate (%)", 
        min_value=0.0, 
        value=8.25, 
        step=0.25,
        format="%.2f",
        help="Enter 8.25 for an 8.25% tax rate."
    )

    labor_cost = st.sidebar.number_input(
        "Labor Cost ($)", 
        min_value=0.0, 
        value=1200.0, 
        step=50.0
    )

    margin_decimal = st.sidebar.number_input(
        "Margin Factor (Decimal)", 
        min_value=0.01, 
        max_value=0.99, 
        value=0.60, 
        step=0.05
    )

    markup_pct = st.sidebar.number_input(
        "Markup Percentage (%)", 
        min_value=0.0, 
        value=15.0, 
        step=1.0
    ) / 100.0

    st.sidebar.markdown("---")
    st.sidebar.header("📁 File Configuration")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(script_dir, "data")

    excel_files = glob.glob(os.path.join(data_dir, "*.xlsx")) + glob.glob(os.path.join(data_dir, "*.xls"))
    default_master_path = excel_files[0] if excel_files else os.path.join(data_dir, "Master_ahri.xlsx")
    default_vendor_folder = os.path.join(data_dir, "vendor_files")

    if st.sidebar.button("🔄 Force Re-scan Folder"):
        st.cache_data.clear()
        st.rerun()

    uploaded_master = st.sidebar.file_uploader("Upload Master AHRI Sheet (.xlsx)", type=["xlsx"])
    vendor_folder = st.sidebar.text_input("Vendor Sheets Directory Path", value=default_vendor_folder)

    # ------------------------------------------
    # MASTER & VENDOR SHEET PROCESSING
    # ------------------------------------------
    master_source = uploaded_master if uploaded_master is not None else (
        default_master_path if os.path.exists(default_master_path) else None
    )

    if master_source:
        try:
            master_sheets = pd.read_excel(master_source, header=None, sheet_name=None, dtype=str)
            tab_names = list(master_sheets.keys())
        except Exception as e:
            st.error(f"Failed to read Master AHRI file: {e}")
            return

        with st.spinner("Processing vendor sheets..."):
            vendor_db, processed_files, skipped_files = load_and_build_vendor_db(vendor_folder)

        st.sidebar.markdown("---")
        st.sidebar.write("### File Scan Details")
        st.sidebar.success(f"**Successfully extracted ({len(processed_files)} files):**")
        for f in processed_files:
            st.sidebar.write(f"- {f}")

        if skipped_files:
            st.sidebar.warning(f"**0 AHRI/Price matches extracted ({len(skipped_files)} files):**")
            for f in skipped_files:
                st.sidebar.write(f"- {f}")

        if vendor_db.empty:
            st.warning("No valid AHRI pricing records extracted from vendor folder.")
            return

        selected_tonnage = st.tabs(tab_names)

        for idx, tab in enumerate(selected_tonnage):
            ton_label = tab_names[idx]
            with tab:
                st.write(f"### Master Units for {ton_label}")
                master_df = master_sheets[ton_label]

                def find_ahri_in_row(row):
                    row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
                    matches = re.findall(r'(?<!\d)\d{9}(?!\d)', row_str)
                    return matches[0] if matches else None

                master_df['Clean_AHRI'] = master_df.apply(find_ahri_in_row, axis=1)
                master_clean = master_df.dropna(subset=['Clean_AHRI']).copy()

                merged = pd.merge(
                    master_clean,
                    vendor_db,
                    left_on='Clean_AHRI',
                    right_on='AHRI Number',
                    how='inner'
                )

                if merged.empty:
                    st.info(f"No vendor matches found for AHRI units in {ton_label}.")
                else:
                    vendors = merged['Vendor Sheet'].unique()
                    selected_vendor = st.selectbox(
                        f"Filter by Vendor Sheet ({ton_label}):",
                        ["All Vendors"] + list(vendors),
                        key=f"vendor_select_{ton_label}"
                    )

                    display_df = merged if selected_vendor == "All Vendors" else merged[merged['Vendor Sheet'] == selected_vendor]
                    
                    final_table = display_df[['Clean_AHRI', 'Vendor Sheet', 'System Price']].copy()
                    final_table.columns = ['AHRI Number', 'Vendor Sheet', 'System Price']
                    final_table['System Price'] = final_table['System Price'].map("${:,.2f}".format)

                    st.dataframe(final_table, use_container_width=True)

                    # ------------------------------------------
                    # SINGLE UNIT CALCULATION OUTPUT
                    # ------------------------------------------
                    st.markdown("---")
                    st.subheader("💡 Customer Investment Calculator")

                    default_unit_price = float(display_df['System Price'].iloc[0]) if not display_df.empty else 0.0
                    selected_price = st.number_input(
                        f"Select or enter System Price from above table ({ton_label}):",
                        min_value=0.0,
                        value=default_unit_price,
                        step=50.0,
                        key=f"price_input_{ton_label}"
                    )

                    if selected_price > 0:
                        subtotal_misc, tax_amt, subtotal_labor, cost_base, final_investment = calculate_customer_investment(
                            base_equipment_price=selected_price,
                            misc_cost=misc_cost,
                            labor_cost=labor_cost,
                            margin_decimal=margin_decimal,
                            markup_pct=markup_pct,
                            tax_rate_pct=tax_rate_pct
                        )

                        col1, col2, col3, col4 = st.columns(4)
                        col1.metric("Base Price + Misc", f"${subtotal_misc:,.2f}")
                        col2.metric(f"Sales Tax ({tax_rate_pct:.2f}%)", f"${tax_amt:,.2f}")
                        col3.metric("Total Cost Base", f"${cost_base:,.2f}")
                        col4.metric("Final Customer Investment", f"${final_investment:,.2f}")

    else:
        st.error(f"Could not locate Master Excel file inside `{data_dir}`.")

if __name__ == "__main__":
    main()