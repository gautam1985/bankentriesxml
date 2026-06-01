import xml.etree.ElementTree as ET
from datetime import datetime
import io
import pandas as pd
import streamlit as st


def clean_xml_string(val):
    """Escapes XML special characters and strips whitespaces."""
    if pd.isna(val):
        return ""
    return str(val).strip().replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# ==========================================
# CORE ENGINE 1: GST PURCHASE INVOICES (MULTI-ITEM)
# ==========================================
def advanced_excel_to_tally_xml(df):
    columns = df.columns.tolist()
    item_col_header = columns[5]         
    qty_col_header = columns[6]          
    rate_col_header = columns[7]         
    purchase_ledger_name = columns[8]   
    igst_ledger_name = columns[9]       
    cgst_ledger_name = columns[10]      
    sgst_ledger_name = columns[11]      
    round_off_ledger_name = columns[12] 
    
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    request_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(request_desc, "REPORTNAME").text = "Vouchers"
    static_vars = ET.SubElement(request_desc, "STATICVARIABLES")
    ET.SubElement(static_vars, "SVCURRENTCOMPANY").text = "Decor Test"
    
    request_data = ET.SubElement(import_data, "REQUESTDATA")
    df = df.dropna(subset=['Supplier Name', 'Invoice Number'])
    grouped = df.groupby('Invoice Number', sort=False)
    
    for invoice_no, group in grouped:
        first_row = group.iloc[0]
        supplier_name = clean_xml_string(first_row['Supplier Name'])
        invoice_no_str = clean_xml_string(invoice_no)
        ref_no = clean_xml_string(first_row['Reference Number'])
        gstin = clean_xml_string(first_row['GST Number'])
        
        raw_date = first_row['Invoice Date']
        if isinstance(raw_date, datetime):
            date_str = raw_date.strftime("%Y%m%d")
        else:
            try:
                date_str = pd.to_datetime(raw_date).strftime("%Y%m%d")
            except:
                date_str = datetime.today().strftime("%Y%m%d")
                
        has_inventory = any(pd.notna(row[item_col_header]) and str(row[item_col_header]).strip() != "" for _, row in group.iterrows())
        
        total_taxable_amt = 0.0
        total_igst = 0.0
        total_cgst = 0.0
        total_sgst = 0.0
        total_round_off = 0.0
        
        tally_msg = ET.SubElement(request_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        voucher = ET.SubElement(tally_msg, "VOUCHER", {
            "VTYPE": "Purchase", 
            "ACTION": "Create",
            "OBJVIEW": "Invoice Voucher View" if has_inventory else "Accounting Voucher View"
        })
        
        ET.SubElement(voucher, "DATE").text = date_str
        ET.SubElement(voucher, "REFERENCEDATE").text = date_str
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = "Purchase"
        ET.SubElement(voucher, "PARTYLEDGERNAME").text = supplier_name
        ET.SubElement(voucher, "VOUCHERNUMBER").text = invoice_no_str
        ET.SubElement(voucher, "REFERENCE").text = ref_no  
        ET.SubElement(voucher, "PERSISTEDVIEW").text = "Invoice Voucher View" if has_inventory else "Accounting Voucher View"
        ET.SubElement(voucher, "VCHENTRYMODE").text = "Item Invoice" if has_inventory else "Accounting Voucher"
        ET.SubElement(voucher, "ISINVOICE").text = "Yes" if has_inventory else "No"
        ET.SubElement(voucher, "NARRATION").text = f"Imported Multi-Item Ref: {ref_no}. GSTIN: {gstin}"
        
        for _, row in group.iterrows():
            row_item_name = row[item_col_header]
            row_has_inv = pd.notna(row_item_name) and str(row_item_name).strip() != ""
            
            taxable_amt = float(row[purchase_ledger_name]) if pd.notna(row[purchase_ledger_name]) else 0.0
            igst_amt = float(row[igst_ledger_name]) if pd.notna(row[igst_ledger_name]) else 0.0
            cgst_amt = float(row[cgst_ledger_name]) if pd.notna(row[cgst_ledger_name]) else 0.0
            sgst_amt = float(row[sgst_ledger_name]) if pd.notna(row[sgst_ledger_name]) else 0.0
            round_off_amt = float(row[round_off_ledger_name]) if pd.notna(row[round_off_ledger_name]) else 0.0
            
            total_taxable_amt += taxable_amt
            total_igst += igst_amt
            total_cgst += cgst_amt
            total_sgst += sgst_amt
            total_round_off += round_off_amt
            
            if row_has_inv:
                actual_stock_item = clean_xml_string(row_item_name)
                quantity = float(row[qty_col_header]) if pd.notna(row[qty_col_header]) else 0.0
                rate = float(row[rate_col_header]) if pd.notna(row[rate_col_header]) else 0.0
                
                inv_entry = ET.SubElement(voucher, "ALLINVENTORYENTRIES.LIST")
                ET.SubElement(inv_entry, "STOCKITEMNAME").text = actual_stock_item  
                ET.SubElement(inv_entry, "ISDEEMEDPOSITIVE").text = "Yes"
                ET.SubElement(inv_entry, "RATE").text = f"{rate:.2f}/PCS"
                ET.SubElement(inv_entry, "AMOUNT").text = f"-{taxable_amt:.2f}"
                ET.SubElement(inv_entry, "ACTUALQTY").text = f" {int(quantity)} PCS"
                ET.SubElement(inv_entry, "BILLEDQTY").text = f" {int(quantity)} PCS"
                
                accounting_alloc = ET.SubElement(inv_entry, "ACCOUNTINGALLOCATIONS.LIST")
                ET.SubElement(accounting_alloc, "LEDGERNAME").text = clean_xml_string(purchase_ledger_name)
                ET.SubElement(accounting_alloc, "ISDEEMEDPOSITIVE").text = "Yes"
                ET.SubElement(accounting_alloc, "AMOUNT").text = f"-{taxable_amt:.2f}"
            else:
                if taxable_amt != 0:
                    dr_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
                    ET.SubElement(dr_entry, "LEDGERNAME").text = clean_xml_string(purchase_ledger_name)
                    ET.SubElement(dr_entry, "ISDEEMEDPOSITIVE").text = "Yes" 
                    ET.SubElement(dr_entry, "AMOUNT").text = f"-{taxable_amt:.2f}"

        grand_total_amt = total_taxable_amt + total_igst + total_cgst + total_sgst + total_round_off
        
        cr_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(cr_entry, "LEDGERNAME").text = supplier_name
        ET.SubElement(cr_entry, "ISDEEMEDPOSITIVE").text = "No" 
        ET.SubElement(cr_entry, "AMOUNT").text = f"{grand_total_amt:.2f}"
        ET.SubElement(cr_entry, "ISPARTYLEDGER").text = "Yes"
        
        bill_alloc = ET.SubElement(cr_entry, "BILLALLOCATIONS.LIST")
        ET.SubElement(bill_alloc, "NAME").text = invoice_no_str
        ET.SubElement(bill_alloc, "BILLTYPE").text = "New Ref"
        ET.SubElement(bill_alloc, "AMOUNT").text = f"{grand_total_amt:.2f}"
        
        if total_igst != 0:
            igst_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(igst_entry, "LEDGERNAME").text = clean_xml_string(igst_ledger_name)
            ET.SubElement(igst_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(igst_entry, "AMOUNT").text = f"-{total_igst:.2f}"
            
        if total_cgst != 0:
            cgst_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(cgst_entry, "LEDGERNAME").text = clean_xml_string(cgst_ledger_name)
            ET.SubElement(cgst_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(cgst_entry, "AMOUNT").text = f"-{total_cgst:.2f}"
            
        if total_sgst != 0:
            sgst_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(sgst_entry, "LEDGERNAME").text = clean_xml_string(sgst_ledger_name)
            ET.SubElement(sgst_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(sgst_entry, "AMOUNT").text = f"-{total_sgst:.2f}"
            
        if total_round_off != 0:
            ro_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(ro_entry, "LEDGERNAME").text = clean_xml_string(round_off_ledger_name)
            if total_round_off > 0:
                ET.SubElement(ro_entry, "ISDEEMEDPOSITIVE").text = "Yes"
                ET.SubElement(ro_entry, "AMOUNT").text = f"-{total_round_off:.2f}"
            else:
                ET.SubElement(ro_entry, "ISDEEMEDPOSITIVE").text = "No"
                ET.SubElement(ro_entry, "AMOUNT").text = f"{abs(total_round_off):.2f}"

    buffer = io.BytesIO()
    tree = ET.ElementTree(envelope)
    tree.write(buffer, encoding="utf-8", xml_declaration=True)
    return buffer.getvalue()


# ==========================================
# CORE ENGINE 2: BANK TRANSACTIONS (UPDATED)
# ==========================================
def bank_statement_to_tally_xml(df, bank_ledger_name):
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    request_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(request_desc, "REPORTNAME").text = "Vouchers"
    static_vars = ET.SubElement(request_desc, "STATICVARIABLES")
    ET.SubElement(static_vars, "SVCURRENTCOMPANY").text = "Decor Test"
    
    request_data = ET.SubElement(import_data, "REQUESTDATA")
    
    df.columns = [str(c).strip() for c in df.columns]
    
    for index, row in df.iterrows():
        # Check that Ledger Name field exists and contains text
        if pd.isna(row['Ledger Name']) or str(row['Ledger Name']).strip() == "":
            continue
            
        tally_ledger = clean_xml_string(row['Ledger Name'])
        raw_narration = clean_xml_string(row['Particulars']) if pd.notna(row['Particulars']) else ""
        
        # Clean Date Formatting
        raw_date = row['Date']
        if isinstance(raw_date, datetime):
            date_str = raw_date.strftime("%Y%m%d")
        else:
            try:
                date_str = pd.to_datetime(raw_date).strftime("%Y%m%d")
            except:
                date_str = datetime.today().strftime("%Y%m%d")
                
        # Read debit and credit flows
        debit_val = float(row['Debit Amount']) if pd.notna(row['Debit Amount']) else 0.0
        credit_val = float(row['Credit Amount']) if pd.notna(row['Credit Amount']) else 0.0
        
        if debit_val == 0.0 and credit_val == 0.0:
            continue  
            
        # Dynamically set accounting voucher headers and modes
        if debit_val > 0:
            vch_type = "Payment"
            amount_str = f"{debit_val:.2f}"
            is_debit_positive_bank = "No"  # Bank is Credited (Money Left)
            is_debit_positive_part = "Yes" # Target Ledger is Debited
        else:
            vch_type = "Receipt"
            amount_str = f"{credit_val:.2f}"
            is_debit_positive_bank = "Yes" # Bank is Debited (Money Entered)
            is_debit_positive_part = "No"  # Target Ledger is Credited

        tally_msg = ET.SubElement(request_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        voucher = ET.SubElement(tally_msg, "VOUCHER", {
            "VTYPE": vch_type, 
            "ACTION": "Create",
            "OBJVIEW": "Accounting Voucher View"
        })
        
        # Core Voucher Header Properties
        ET.SubElement(voucher, "DATE").text = date_str
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
        ET.SubElement(voucher, "PARTYLEDGERNAME").text = tally_ledger if vch_type == "Receipt" else bank_ledger_name
        ET.SubElement(voucher, "PERSISTEDVIEW").text = "Accounting Voucher View"
        ET.SubElement(voucher, "VCHENTRYMODE").text = "Accounting Voucher"
        ET.SubElement(voucher, "ISINVOICE").text = "No"
        
        # --- NARRATION UPDATED ---
        # Raw statement transaction text injected directly into voucher narrative summary
        ET.SubElement(voucher, "NARRATION").text = raw_narration if raw_narration != "" else "Imported Bank Transaction"
        
        # Line 1: Main Bank Leg Mapping
        bank_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(bank_entry, "LEDGERNAME").text = clean_xml_string(bank_ledger_name)
        ET.SubElement(bank_entry, "ISDEEMEDPOSITIVE").text = is_debit_positive_bank
        ET.SubElement(bank_entry, "AMOUNT").text = f"-{amount_str}" if is_debit_positive_bank == "Yes" else amount_str
        
        # Line 2: Target Double-Entry Mapping (Now reads clean custom Ledger Name)
        party_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(party_entry, "LEDGERNAME").text = tally_ledger
        ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = is_debit_positive_part
        ET.SubElement(party_entry, "AMOUNT").text = f"-{amount_str}" if is_debit_positive_part == "Yes" else amount_str

    buffer = io.BytesIO()
    tree = ET.ElementTree(envelope)
    tree.write(buffer, encoding="utf-8", xml_declaration=True)
    return buffer.getvalue()


# ==========================================
# STREAMLIT USER GATEWAY SYSTEM CONTROL
# ==========================================
st.set_page_config(page_title="Universal Tally Master Dashboard", page_icon="⚙️", layout="centered")

st.sidebar.title("Workspace Control")
app_mode = st.sidebar.radio(
    "Select Import Data Category:",
    ["GST Purchase Invoices", "Bank Statement Ledger Transactions"]
)

if app_mode == "GST Purchase Invoices":
    st.title("Universal Excel ➔ Tally Prime XML Converter")
    st.markdown("### Accounts & Inventory Purchase Matrix Mode")
    st.write("Groups multi-item tracking lines sharing matching receipt identification codes automatically.")
    
    uploaded_file = st.file_uploader("Upload 13-Column GST Invoice Excel Sheet", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("Purchase row matrix mapped successfully!")
            st.dataframe(df.head(5))
            
            if st.button("⚙️ Compile GST Purchase XML Asset", type="primary"):
                xml_data = advanced_excel_to_tally_xml(df)
                st.balloons()
                st.download_button(
                    label="📥 Download Purchase Import File",
                    data=xml_data,
                    file_name="Tally_MultiItem_Purchase.xml",
                    mime="application/xml"
                )
        except Exception as e:
            st.error(f"Processing Error: {str(e)}")

else:
    st.title("Bank Statement ➔ Tally Auto-Posting Module")
    st.markdown("### Automatic Payment & Receipt Generator")
    st.write("Generates pure accounting vouchers dynamically directly from your historical passbook data entries.")
    
    target_bank = st.text_input("Enter Tally Bank Ledger Name (e.g., HDFC Bank, SBI Account):", "HDFC BANK A/C")
    
    uploaded_file = st.file_uploader("Upload Updated 5-Column Bank Statement Sheet", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("Bank Statement structural template detected!")
            st.dataframe(df.head(5))
            
            # Form validation check for newly structured required banking data keys
            required_bank_cols = ['Date', 'Particulars', 'Debit Amount', 'Credit Amount', 'Ledger Name']
            found_cols = [str(c).strip() for c in df.columns]
            missing_cols = [c for c in required_bank_cols if c not in found_cols]
            
            if missing_cols:
                st.warning(f"Column Mismatch! Your sheet must include these exact headers: {missing_cols}")
            else:
                if st.button("⚙️ Compile Bank Transactions XML Asset", type="primary"):
                    xml_data = bank_statement_to_tally_xml(df, target_bank)
                    st.balloons()
                    st.download_button(
                        label="📥 Download Bank Import File",
                        data=xml_data,
                        file_name="Tally_Bank_Transactions.xml",
                        mime="application/xml"
                    )
        except Exception as e:
            st.error(f"Processing Error: {str(e)}")