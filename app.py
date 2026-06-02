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
# ENGINE 1: GST PURCHASE INVOICES
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
# ENGINE 2: GST SALES INVOICES (NEW)
# ==========================================
def sales_excel_to_tally_xml(df, sales_ledger_name, cgst_ledger, sgst_ledger, igst_ledger, delivery_ledger, round_off_ledger):
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
    df = df.dropna(subset=['Buyer Name', 'Invoice Number'])
    
    # Group rows by Invoice Number to compile multi-item bills natively
    grouped = df.groupby('Invoice Number', sort=False)
    
    for invoice_no, group in grouped:
        first_row = group.iloc[0]
        buyer_name = clean_xml_string(first_row['Buyer Name'])
        invoice_no_str = clean_xml_string(invoice_no)
        
        # Clean Date Handling
        raw_date = first_row['Date']
        if isinstance(raw_date, datetime):
            date_str = raw_date.strftime("%Y%m%d")
        else:
            try:
                date_str = pd.to_datetime(raw_date).strftime("%Y%m%d")
            except:
                date_str = datetime.today().strftime("%Y%m%d")
        
        tally_msg = ET.SubElement(request_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        voucher = ET.SubElement(tally_msg, "VOUCHER", {
            "VTYPE": "Sales", 
            "ACTION": "Create",
            "OBJVIEW": "Invoice Voucher View"
        })
        
        # Sales Header Configuration
        ET.SubElement(voucher, "DATE").text = date_str
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = "Sales"
        ET.SubElement(voucher, "PARTYLEDGERNAME").text = buyer_name
        ET.SubElement(voucher, "VOUCHERNUMBER").text = invoice_no_str
        ET.SubElement(voucher, "PERSISTEDVIEW").text = "Invoice Voucher View"
        ET.SubElement(voucher, "VCHENTRYMODE").text = "Item Invoice"
        ET.SubElement(voucher, "ISINVOICE").text = "Yes"
        ET.SubElement(voucher, "NARRATION").text = f"Imported Sales Bill No: {invoice_no_str}"
        
        # Financial Accumulators
        total_item_amount = 0.0
        total_igst = 0.0
        total_cgst = 0.0
        total_sgst = 0.0
        total_round_off = 0.0
        total_delivery = 0.0
        
        # Loop through items assigned to this sales invoice block
        for _, row in group.iterrows():
            item_name = clean_xml_string(row['Item Name'])
            qty = float(row['Quantity']) if pd.notna(row['Quantity']) else 0.0
            rate = float(row['Rate']) if pd.notna(row['Rate']) else 0.0
            item_amt = float(row['Item amount']) if pd.notna(row['Item amount']) else 0.0
            
            total_item_amount += item_amt
            total_igst += float(row['IGST Amount']) if pd.notna(row['IGST Amount']) else 0.0
            total_cgst += float(row['CGST Amount']) if pd.notna(row['CGST Amount']) else 0.0
            total_sgst += float(row['SGST Amount']) if pd.notna(row['SGST Amount']) else 0.0
            total_round_off += float(row['Round Off']) if pd.notna(row['Round Off']) else 0.0
            total_delivery += float(row['Deilvery Charges']) if pd.notna(row['Deilvery Charges']) else 0.0
            
            # Line Item Block
            inv_entry = ET.SubElement(voucher, "ALLINVENTORYENTRIES.LIST")
            ET.SubElement(inv_entry, "STOCKITEMNAME").text = item_name
            ET.SubElement(inv_entry, "ISDEEMEDPOSITIVE").text = "No"  # Sales inventory is reduced (Negative)
            ET.SubElement(inv_entry, "RATE").text = f"{rate:.2f}/PCS"
            ET.SubElement(inv_entry, "AMOUNT").text = f"{item_amt:.2f}"
            ET.SubElement(inv_entry, "ACTUALQTY").text = f" {int(qty)} PCS"
            ET.SubElement(inv_entry, "BILLEDQTY").text = f" {int(qty)} PCS"
            
            # Sales Account mapping linked to the item allocation track
            acc_alloc = ET.SubElement(inv_entry, "ACCOUNTINGALLOCATIONS.LIST")
            ET.SubElement(acc_alloc, "LEDGERNAME").text = clean_xml_string(sales_ledger_name)
            ET.SubElement(acc_alloc, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(acc_alloc, "AMOUNT").text = f"{item_amt:.2f}"

        # Calculate absolute invoice grand total
        grand_total_amt = total_item_amount + total_igst + total_cgst + total_sgst + total_round_off + total_delivery
        
        # Debited Leg: The Buyer Master Ledger Link (Run once per invoice)
        buyer_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(buyer_entry, "LEDGERNAME").text = buyer_name
        ET.SubElement(buyer_entry, "ISDEEMEDPOSITIVE").text = "Yes"  # Debtors are positive/debit entries in sales
        ET.SubElement(buyer_entry, "AMOUNT").text = f"-{grand_total_amt:.2f}"
        ET.SubElement(buyer_entry, "ISPARTYLEDGER").text = "Yes"
        
        bill_alloc = ET.SubElement(buyer_entry, "BILLALLOCATIONS.LIST")
        ET.SubElement(bill_alloc, "NAME").text = invoice_no_str
        ET.SubElement(bill_alloc, "BILLTYPE").text = "New Ref"
        ET.SubElement(bill_alloc, "AMOUNT").text = f"-{grand_total_amt:.2f}"
        
        # Additional Global Ledger Sub-Entries (Taxes, Charges, Adjustments)
        if total_igst > 0:
            e = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(e, "LEDGERNAME").text = clean_xml_string(igst_ledger)
            ET.SubElement(e, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(e, "AMOUNT").text = f"{total_igst:.2f}"
            
        if total_cgst > 0:
            e = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(e, "LEDGERNAME").text = clean_xml_string(cgst_ledger)
            ET.SubElement(e, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(e, "AMOUNT").text = f"{total_cgst:.2f}"
            
        if total_sgst > 0:
            e = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(e, "LEDGERNAME").text = clean_xml_string(sgst_ledger)
            ET.SubElement(e, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(e, "AMOUNT").text = f"{total_sgst:.2f}"
            
        if total_delivery > 0:
            e = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(e, "LEDGERNAME").text = clean_xml_string(delivery_ledger)
            ET.SubElement(e, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(e, "AMOUNT").text = f"{total_delivery:.2f}"
            
        if total_round_off != 0:
            e = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
            ET.SubElement(e, "LEDGERNAME").text = clean_xml_string(round_off_ledger)
            if total_round_off < 0:
                ET.SubElement(e, "ISDEEMEDPOSITIVE").text = "Yes"
                ET.SubElement(e, "AMOUNT").text = f"-{abs(total_round_off):.2f}"
            else:
                ET.SubElement(e, "ISDEEMEDPOSITIVE").text = "No"
                ET.SubElement(e, "AMOUNT").text = f"{total_round_off:.2f}"

    buffer = io.BytesIO()
    tree = ET.ElementTree(envelope)
    tree.write(buffer, encoding="utf-8", xml_declaration=True)
    return buffer.getvalue()


# ==========================================
# ENGINE 3: BANK TRANSACTIONS
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
        if pd.isna(row['Ledger Name']) or str(row['Ledger Name']).strip() == "":
            continue
            
        tally_ledger = clean_xml_string(row['Ledger Name'])
        raw_narration = clean_xml_string(row['Particulars']) if pd.notna(row['Particulars']) else ""
        
        raw_date = row['Date']
        if isinstance(raw_date, datetime):
            date_str = raw_date.strftime("%Y%m%d")
        else:
            try:
                date_str = pd.to_datetime(raw_date).strftime("%Y%m%d")
            except:
                date_str = datetime.today().strftime("%Y%m%d")
                
        debit_val = float(row['Debit Amount']) if pd.notna(row['Debit Amount']) else 0.0
        credit_val = float(row['Credit Amount']) if pd.notna(row['Credit Amount']) else 0.0
        
        if debit_val == 0.0 and credit_val == 0.0:
            continue  
            
        if debit_val > 0:
            vch_type = "Payment"
            amount_str = f"{debit_val:.2f}"
            is_debit_positive_bank = "No"  
            is_debit_positive_part = "Yes" 
        else:
            vch_type = "Receipt"
            amount_str = f"{credit_val:.2f}"
            is_debit_positive_bank = "Yes" 
            is_debit_positive_part = "No"  

        tally_msg = ET.SubElement(request_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
        voucher = ET.SubElement(tally_msg, "VOUCHER", {
            "VTYPE": vch_type, 
            "ACTION": "Create",
            "OBJVIEW": "Accounting Voucher View"
        })
        
        ET.SubElement(voucher, "DATE").text = date_str
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
        ET.SubElement(voucher, "PARTYLEDGERNAME").text = tally_ledger if vch_type == "Receipt" else bank_ledger_name
        ET.SubElement(voucher, "PERSISTEDVIEW").text = "Accounting Voucher View"
        ET.SubElement(voucher, "VCHENTRYMODE").text = "Accounting Voucher"
        ET.SubElement(voucher, "ISINVOICE").text = "No"
        ET.SubElement(voucher, "NARRATION").text = raw_narration if raw_narration != "" else "Imported Bank Transaction"
        
        bank_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(bank_entry, "LEDGERNAME").text = clean_xml_string(bank_ledger_name)
        ET.SubElement(bank_entry, "ISDEEMEDPOSITIVE").text = is_debit_positive_bank
        ET.SubElement(bank_entry, "AMOUNT").text = f"-{amount_str}" if is_debit_positive_bank == "Yes" else amount_str
        
        party_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(party_entry, "LEDGERNAME").text = tally_ledger
        ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = is_debit_positive_part
        ET.SubElement(party_entry, "AMOUNT").text = f"-{amount_str}" if is_debit_positive_part == "Yes" else amount_str

    buffer = io.BytesIO()
    tree = ET.ElementTree(envelope)
    tree.write(buffer, encoding="utf-8", xml_declaration=True)
    return buffer.getvalue()


# ==========================================
# STREAMLIT USER GATEWAY INTERFACE CONTROLLER
# ==========================================
st.set_page_config(page_title="Universal Tally Master Dashboard", page_icon="⚙️", layout="centered")

st.sidebar.title("Workspace Control")
app_mode = st.sidebar.radio(
    "Select Import Data Category:",
    ["GST Purchase Invoices", "GST Sales Invoices", "Bank Statement Ledger Transactions"]
)

# --- WORKSPACE MODE: PURCHASE ---
if app_mode == "GST Purchase Invoices":
    st.title("Excel ➔ Tally Purchase XML Engine")
    uploaded_file = st.file_uploader("Upload GST Purchase Excel Sheet", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("Matrix mapped successfully!")
            st.dataframe(df.head(5))
            if st.button("⚙️ Compile GST Purchase XML Asset", type="primary"):
                xml_data = advanced_excel_to_tally_xml(df)
                st.balloons()
                st.download_button(label="📥 Download Purchase File", data=xml_data, file_name="Tally_Purchase.xml")
        except Exception as e:
            st.error(f"Error: {str(e)}")

# --- WORKSPACE MODE: SALES (NEW) ---
elif app_mode == "GST Sales Invoices":
    st.title("Excel ➔ Tally Sales XML Engine")
    st.markdown("### Retail & Counter Sales Compilation Matrix")
    
    # Configure defaults matching standard setup
    col1, col2 = st.columns(2)
    with col1:
        s_led = st.text_input("Sales Account Ledger:", "Sales Account")
        cgst_led = st.text_input("Output CGST Ledger:", "Output CGST")
        sgst_led = st.text_input("Output SGST Ledger:", "Output SGST")
    with col2:
        igst_led = st.text_input("Output IGST Ledger:", "Output IGST")
        del_led = st.text_input("Delivery Charges Ledger:", "Delivery Charges")
        ro_led = st.text_input("Round Off Ledger:", "Round Off")

    uploaded_file = st.file_uploader("Upload Counter Sales Excel Sheet", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("Sales data template loaded successfully!")
            st.dataframe(df.head(5))
            
            # Map validation for safety matching exact required names
            required_sales_cols = [
                'Date', 'Invoice Number', 'Buyer Name', 'Item Name', 'Quantity', 
                'Rate', 'Item amount', 'IGST Amount', 'CGST Amount', 'SGST Amount', 
                'Round Off', 'Deilvery Charges', 'Total Amount'
            ]
            found_cols = [str(c).strip() for c in df.columns]
            missing_cols = [c for c in required_sales_cols if c not in found_cols]
            
            if missing_cols:
                st.warning(f"Header Name Mismatch! Ensure your columns include: {missing_cols}")
            else:
                if st.button("⚙️ Compile GST Sales XML Asset", type="primary"):
                    xml_data = sales_excel_to_tally_xml(df, s_led, cgst_led, sgst_led, igst_led, del_led, ro_led)
                    st.balloons()
                    st.download_button(
                        label="📥 Download Sales Import File", 
                        data=xml_data, 
                        file_name="Tally_Sales_Import.xml",
                        mime="application/xml"
                    )
        except Exception as e:
            st.error(f"Processing Error: {str(e)}")

# --- WORKSPACE MODE: BANK ---
else:
    st.title("Bank Statement Auto-Posting Module")
    target_bank = st.text_input("Enter Tally Bank Ledger Name:", "HDFC BANK A/C")
    uploaded_file = st.file_uploader("Upload 5-Column Bank Statement Sheet", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("Bank Statement structural template detected!")
            st.dataframe(df.head(5))
            required_bank_cols = ['Date', 'Particulars', 'Debit Amount', 'Credit Amount', 'Ledger Name']
            found_cols = [str(c).strip() for c in df.columns]
            missing_cols = [c for c in required_bank_cols if c not in found_cols]
            
            if missing_cols:
                st.warning(f"Column Mismatch! Required: {missing_cols}")
            else:
                if st.button("⚙️ Compile Bank Transactions XML Asset", type="primary"):
                    xml_data = bank_statement_to_tally_xml(df, target_bank)
                    st.balloons()
                    st.download_button(label="📥 Download Bank Import File", data=xml_data, file_name="Tally_Bank.xml")
        except Exception as e:
            st.error(f"Error: {str(e)}")
