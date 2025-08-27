import os
import warnings
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
from io import BytesIO
import base64
import unicodedata
import openpyxl
import xlrd

warnings.filterwarnings("ignore", category=UserWarning, module="streamlit.runtime.scriptrunner.script_runner")

# Set page config for title and favicon
st.set_page_config(
    page_title="Map It",
    layout="centered"
)

# Custom CSS for purple and gold theme with orchid watermark
st.markdown("""
<style>
:root {
    --primary: #63328a;
    --secondary: #f9c846;
    --accent: #b178d6;
    --light: #f8fafc;
    --dark: #2d124d;
    --success: #0ACF83;
    --text: #2d3748;
    --background: #f7f1fa;
}

.stApp {
    background: var(--background);
    font-family: 'Inter', 'Montserrat', 'Segoe UI', sans-serif;
}

.header {
    text-align: center;
    margin: -1rem -1rem 2rem;
    padding: 3rem 1rem 2rem;
    background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
    color: white;
    position: relative;
    overflow: hidden;
    border-radius: 0 0 20px 20px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

.title {
    font-size: 2.7rem;
    font-weight: 800;
    margin-bottom: 0.5rem;
    letter-spacing: -0.5px;
    font-family: 'Montserrat', 'Inter', sans-serif;
}

.subtitle {
    font-size: 1.12rem;
    opacity: 0.95;
    font-weight: 400;
    letter-spacing: 0.5px;
}

.stButton>button {
    background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
    color: white;
    border-radius: 14px;
    border: none;
    padding: 1rem 2.3rem;
    font-weight: 700;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(99,50,138,0.13);
    font-size: 1.02rem;
    width: 100%;
}

.stButton>button:hover {
    transform: translateY(-2px) scale(1.01);
    box-shadow: 0 6px 20px rgba(99,50,138,0.22);
}

.stDateInput>div>div>input {
    border: 2px solid var(--primary) !important;
    border-radius: 14px !important;
    padding: 13px 18px !important;
    font-size: 1.04rem !important;
    color: var(--dark) !important;
}

.stAlert {
    border-radius: 14px !important;
    border: none !important;
}

.stDataFrame {
    border-radius: 14px !important;
    box-shadow: 0 4px 20px rgba(99,50,138,0.04) !important;
    border: 1px solid rgba(99,50,138,0.07) !important;
}

.stMetric {
    border-radius: 14px !important;
    box-shadow: 0 2px 10px rgba(99,50,138,0.04) !important;
    border: 1px solid rgba(99,50,138,0.07) !important;
    padding: 1.1rem !important;
}

.orchid-text {
    position: fixed;
    bottom: 20px;
    right: 20px;
    font-family: 'Playfair Display', 'Arial', serif;
    font-size: 0.8rem;
    color: #f9c846;
    font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
    z-index: 1000;
    opacity: 0.98;
    text-shadow: 0 2px 7px rgba(99,50,138,0.13);
}

.error-details {
    background: #fee;
    border: 1px solid #fcc;
    border-radius: 8px;
    padding: 1rem;
    margin: 0.5rem 0;
    font-family: monospace;
    font-size: 0.9rem;
}

.success-details {
    background: #efe;
    border: 1px solid #cfc;
    border-radius: 8px;
    padding: 1rem;
    margin: 0.5rem 0;
}
</style>
""", unsafe_allow_html=True)

# Mapping dictionary for descriptions
DESCRIPTION_MAPPING = {
    # Interest related
    'interest': 'Monthly Interest',
    'monthly interest': 'Monthly Interest',
    'interest reversal': 'Interest Reversal',
    'reversal interest': 'Interest Reversal',
    'int reversal': 'Interest Reversal',
    'reversal int': 'Interest Reversal',
    'arrears interest': 'Arrear Interest',
    'penalty interest': 'Penalty Interest',

    # Direct deposits
    'direct deposits': 'Direct Deposits',
    'direct deposit': 'Direct Deposits',

    # Refunds
    'refund': 'Refund',
    'refunds': 'Refund',

    # Fees
    'service fee': 'Service Fee',
    'card fee': 'Card Fee',
    'legal fee': 'Legal Fees',
    'legal fees': 'Legal Fees',
    'funeral fee': 'Funeral Fee',
    'arrangement fee': 'Arrangement Fee',
    'transfer fee': 'Transfer Fee',
    'early settlement fee': 'Early Settlement Fee',

    # Reversals
    'receipt reversal': 'Receipts Reversals',
    'reversal receipt': 'Receipts Reversals',
    'cash drawer receipt reversal': 'Cash Drawer Receipt Reversal',
    'reversal service fee': 'Reversal Service Fee',
    'reversal card fee': 'Reversal Card Fee',
    'reverse legal fees': 'Reverse Legal Fees',
    'reverse funeral fee': 'Reverse Funeral Fee',
    'reverse arrangement fee': 'Reversal Initiation Fee',
    'reverse write off': 'Reverse Write Off',
    'write off reversal': 'Reverse Write Off',
    'writeoff reversal': 'Reverse Write Off',
    'reverse write - off': 'Reverse Write Off',
    'reverse writeoff': 'Reverse Write Off',
    'reverse complete write off': 'Reverse Write Off',
    'complete write off reversal': 'Reverse Write Off',
    'reverse complete write - off': 'Reverse Write Off',
    'reverse complete writeoff': 'Reverse Write Off',

    # Status changes
    'contract status - active': 'Contract Status - Active',
    'status-active': 'Contract Status - Active',
    'contract status - complete': 'Contract Status - Complete',
    'status - complete': 'Contract Status - Complete',
    'status complete': 'Contract Status - Complete',
    'status-complete': 'Contract Status - Complete',
    'status - settled': 'Status - Settled',
    'status - legal': 'Status - Legal',
    'change status to legal': 'Status - Legal',
    'status - cancelled': 'Status - Cancelled',

    # Other transactions
    'receipts': 'Receipts',
    'receipt': 'Receipts',
    'cash disbursement': 'Cash Disbursement',
    'cash receipt': 'Cash Receipt',
    'bank deposit': 'Bank Deposit',
    'bank withdrawal': 'Bank Withdrawal',
    'instalment': 'Instalment',
    'write off': 'Write Off',
    'complete write off': 'Complete Write Off',
    'small balance write off': 'Small Balance Credit W-Off'
}

def clean_cell_value(value):
    """
    Comprehensive cleaning function to handle invisible characters and formatting issues
    """
    if pd.isna(value):
        return ""

    # Convert to string and handle float values properly
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value)

    # Remove BOM (Byte Order Mark) characters
    text = text.replace('\ufeff', '')
    text = text.replace('\ufffe', '')

    # Normalize unicode characters (convert special spaces, etc.)
    text = unicodedata.normalize('NFKD', text)

    # Remove various types of invisible characters commonly found in Excel
    invisible_chars = [
        '\u00a0',  # Non-breaking space
        '\u1680',  # Ogham space mark
        '\u2000', '\u2001', '\u2002', '\u2003', '\u2004', '\u2005',  # En quad, Em quad, etc.
        '\u2006', '\u2007', '\u2008', '\u2009', '\u200a',  # Six-per-em space, etc.
        '\u200b', '\u200c', '\u200d', '\u200e', '\u200f',  # Zero-width spaces and marks
        '\u2028', '\u2029',  # Line separator, paragraph separator
        '\u202a', '\u202b', '\u202c', '\u202d', '\u202e',  # Bidirectional text control
        '\u202f',  # Narrow no-break space
        '\u205f',  # Medium mathematical space
        '\u2060',  # Word joiner
        '\u3000',  # Ideographic space
        '\ufeff',  # Zero-width no-break space (BOM)
        '\u180e',  # Mongolian vowel separator
        '\u061c',  # Arabic letter mark
        '\u2066', '\u2067', '\u2068', '\u2069',  # Directional isolates
    ]

    for char in invisible_chars:
        text = text.replace(char, ' ')

    # Remove control characters (except common ones like \n, \t, \r)
    text = ''.join(char for char in text if unicodedata.category(char)[0] != 'C' or char in '\n\t\r')

    # Handle Excel-specific formatting artifacts
    text = text.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')

    # Remove multiple consecutive spaces and normalize
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()

    # Handle empty strings that might contain only invisible characters
    if not text or text.isspace():
        return ""

    return text


def normalize_column_name(col_name):
    """
    Normalize column names for better matching
    """
    if pd.isna(col_name):
        return ""

    # Clean the column name first
    clean_name = clean_cell_value(col_name)

    # Convert to lowercase and remove extra characters
    normalized = re.sub(r'[^\w\s]', ' ', clean_name.lower())
    normalized = re.sub(r'\s+', ' ', normalized).strip()

    return normalized


def enhanced_column_finder(df, possible_names):
    """
    Enhanced column finder with comprehensive cleaning and matching
    """
    # First, clean all column names and create multiple matching variations
    cleaned_columns = {}
    for col in df.columns:
        original_col = str(col)
        # Split at pipe and take first part if pipe exists
        base_col = original_col.split('|')[0].strip()
        cleaned_col = normalize_column_name(base_col)

        # Create multiple variations for better matching
        variations = [
            cleaned_col,
            cleaned_col.replace(' ', ''),  # No spaces
            cleaned_col.replace(' ', '_'),  # Underscores
            re.sub(r'[^\w]', '', cleaned_col),  # Alphanumeric only
        ]

        for variation in variations:
            if variation:  # Only add non-empty variations
                cleaned_columns[variation] = original_col

    # Try to match with possible names
    for target_name in possible_names:
        target_normalized = normalize_column_name(target_name)
        target_variations = [
            target_normalized,
            target_normalized.replace(' ', ''),
            target_normalized.replace(' ', '_'),
            re.sub(r'[^\w]', '', target_normalized),
        ]

        # Try exact matches first
        for target_var in target_variations:
            if target_var in cleaned_columns:
                return cleaned_columns[target_var]

        # Try partial matches
        for target_var in target_variations:
            for cleaned_col, original_col in cleaned_columns.items():
                if target_var and cleaned_col and (target_var in cleaned_col or cleaned_col in target_var):
                    return original_col

        # Try word-by-word matching for multi-word columns
        target_words = [word for word in target_normalized.split() if len(word) > 2]
        if target_words:
            for cleaned_col, original_col in cleaned_columns.items():
                col_words = cleaned_col.split()
                if any(word in col_words for word in target_words):
                    return original_col

    return None


def extract_amount(comment_part):
    """Extract amount from a comment part that might contain an amount"""
    if not isinstance(comment_part, str):
        return 0.0

    # First look for explicit currency format ($25.97 or $1,000.00)
    currency_match = re.search(r'\$?\s*([\d,]+(?:\.\d{1,2})?)', comment_part)
    if currency_match:
        amount_str = currency_match.group(1).replace(',', '')
        try:
            return float(amount_str)
        except ValueError:
            pass

    # Then look for general numbers that might be amounts (25.97 or 1,000.00)
    number_match = re.search(r'([\d,]+(?:\.\d{1,2})?)', comment_part)
    if number_match:
        amount_str = number_match.group(1).replace(',', '')
        try:
            return float(amount_str)
        except ValueError:
            pass

    return 0.0


def clean_comment(comment_part):
    """Remove amount information from comment part while preserving key phrases"""
    # First remove dollar amounts
    cleaned = re.sub(r'\$[\d,]+\.?\d*', '', comment_part)
    # Then remove standalone numbers
    cleaned = re.sub(r'(^|\s)\d[\d,]*\.?\d*\b', ' ', cleaned)
    # Clean but preserve important hyphenated terms
    cleaned = re.sub(r'[^\w\s-]', ' ', cleaned.lower())
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned

def get_transaction_type(description):
    """Get transaction type from description"""
    if description == "Monthly Interest":
        return "INT"
    elif description == "Reverse Write Off":
        return "REV_WOFF"
    elif description == "Reverse Complete Write Off":
        return "REV_COMP_WOFF"
    elif description == "Interest Reversal":
        return "INT_ADJ_CR"
    elif description == "Receipts":
        return "REC"
    elif description == "Arrear Interest":
        return "INT_ARR"
    elif description == "Penalty Interest":
        return "P_INT"
    elif description == "Direct Deposits":
        return "DIRECT_DEPOSIT"
    elif description == "Refund":
        return "REFUND"
    elif description == "Service Fee":
        return "SERV_FEE"
    elif description == "Card Fee":
        return "CARD_FEE"
    elif description == "Legal Fees":
        return "LEG_FEE"
    elif description == "Funeral Fee":
        return "FUN_FEE"
    elif description == "Arrangement Fee":
        return "INIT_FEE"
    elif description == "Transfer Fee":
        return "TRF_FEE"
    elif description == "Early Settlement Fee":
        return "SET_FEE"
    elif description == "Receipts Reversals":
        return "REV_REC"
    elif description == "Cash Drawer Receipt Reversal":
        return "R_CASH_REC"
    elif description == "Reversal Service Fee":
        return "R_SERV_FEE"
    elif description == "Reversal Card Fee":
        return "R_CARD_FEE"
    elif description == "Reverse Legal Fees":
        return "REV_LEG_FEE"
    elif description == "Reverse Funeral Fee":
        return "REV_FUN_FEE"
    elif description == "Reversal Initiation Fee":
        return "REV_INI_FEE"
    elif description == "Contract Status - Active":
        return "STAT_ACTIVE"
    elif description == "Contract Status - Complete":
        return "STAT_COMPLETE"
    elif description == "Status - Settled":
        return "STAT_SETT"
    elif description == "Status - Legal":
        return "STAT_LEGAL"
    elif description == "Status - Cancelled":
        return "STAT_CANCEL"
    elif description == "Cash Disbursement":
        return "CASH_DISB"
    elif description == "Cash Receipt":
        return "CASH_REC"
    elif description == "Bank Deposit":
        return "BANK_DEP"
    elif description == "Bank Withdrawal":
        return "BANK_WTHDRW"
    elif description == "Instalment":
        return "INS"
    elif description == "Write Off":
        return "WOFF"
    elif description == "Complete Write Off":
        return "COMP_WOFF"
    elif description == "Small Balance Credit W-Off":
        return "SB_C_WOFF"
    return "GEN"


def process_comment(comment):
    """Process a comment string into individual transactions"""
    if not isinstance(comment, str):
        return []

    # Clean the comment first
    comment = clean_cell_value(comment)

    # Normalize the comment - more aggressive cleaning
    comment = re.sub(r'\s*,\s*', ', ', comment)  # Normalize commas
    comment = re.sub(r'(\$)\s*(\d)', r'\1\2', comment)  # Fix $ spacing
    comment = re.sub(r'[^\w\s\-,]', ' ', comment.lower())  # Remove special chars except hyphens
    comment = re.sub(r'\s+', ' ', comment).strip()  # Normalize whitespace

    transactions = []
    parts = [part.strip() for part in re.split(r',(?![^()]*\))', comment)]

    for part in parts:
        if not part:
            continue

        # Extract amount first (before cleaning alters the string)
        amount = extract_amount(part)

        # Clean the part for matching
        cleaned = clean_comment(part).lower()
        cleaned = normalize_column_name(cleaned)

        # Try exact matches first
        matched = False
        for key, value in DESCRIPTION_MAPPING.items():
            if cleaned == normalize_column_name(key):
                transactions.append({
                    'description': value,
                    'amount': amount
                })
                matched = True
                break

        # If no exact match, try partial matches
        if not matched:
            for key, value in DESCRIPTION_MAPPING.items():
                norm_key = normalize_column_name(key)
                if norm_key in cleaned or cleaned in norm_key:
                    transactions.append({
                        'description': value,
                        'amount': amount
                    })
                    matched = True
                    break

        # If still no match, try word-by-word matching
        if not matched:
            key_words = set(cleaned.split())
            for key, value in DESCRIPTION_MAPPING.items():
                mapping_words = set(normalize_column_name(key).split())
                if key_words & mapping_words:  # Any common words
                    transactions.append({
                        'description': value,
                        'amount': amount
                    })
                    break

    return transactions


def validate_and_clean_dataframe(df):
    """
    Validate and clean the entire dataframe
    """
    # Clean all cell values in the dataframe
    for col in df.columns:
        df[col] = df[col].apply(clean_cell_value)

    return df


def detect_merged_comments(df, comment_col, contract_col, payee_col, employee_col):
    """Detect and combine merged comments that span multiple rows"""
    merged_rows = []
    current_merge = None

    for idx, row in df.iterrows():
        contract_val = clean_cell_value(row[contract_col])

        if current_merge and contract_val == current_merge['contract_no']:
            current_merge['comment'] += ' ' + clean_cell_value(row[comment_col])
            current_merge['end_idx'] = idx
        else:
            if current_merge:
                merged_rows.append(current_merge)
            current_merge = {
                'contract_no': contract_val,
                'payee': clean_cell_value(row[payee_col]),
                'employee_no': clean_cell_value(row[employee_col]),
                'comment': clean_cell_value(row[comment_col]),
                'start_idx': idx,
                'end_idx': idx
            }

    if current_merge:
        merged_rows.append(current_merge)

    return merged_rows


def validate_row_data(row, idx, contract_col, payee_col, employee_col, comment_col):
    """Validate a single row with enhanced cleaning"""
    errors = []
    warnings = []

    # Clean and validate contract number
    contract_val = clean_cell_value(row.get(contract_col, ''))
    if not contract_val:
        errors.append(f"Row {idx + 1}: Missing or empty contract number")

    # Clean and validate payee
    payee_val = clean_cell_value(row.get(payee_col, ''))
    if not payee_val:
        errors.append(f"Row {idx + 1}: Missing or empty payee name")

    # Clean and validate employee number
    employee_val = clean_cell_value(row.get(employee_col, ''))
    if not employee_val:
        errors.append(f"Row {idx + 1}: Missing or empty employee/EC number")

    # Clean and validate comment
    comment_val = clean_cell_value(row.get(comment_col, ''))
    if not comment_val:
        errors.append(f"Row {idx + 1}: Missing or empty comment")

    return errors, warnings


def convert_file(uploaded_file, effective_date):
    """Convert the uploaded file with enhanced data cleaning and validation"""
    try:
        # Reset file pointer to beginning
        uploaded_file.seek(0)

        # First verify it's actually an Excel file by checking the file signature
        file_signature = uploaded_file.read(8)
        uploaded_file.seek(0)

        # Excel file signatures
        excel_signatures = [
            b'\x50\x4B\x03\x04',  # Modern Excel (xlsx)
            b'\xD0\xCF\x11\xE0',  # Older Excel (xls)
            b'\x09\x08\x10\x00',  # Older Excel variants
            b'\xFD\xFF\xFF\xFF'  # Some older Excel versions
        ]

        is_excel = any(file_signature.startswith(sig) for sig in excel_signatures)
        if not is_excel:
            return None, "❌ The uploaded file doesn't appear to be a valid Excel file. Please upload a .xlsx or .xls file."

        # Try multiple approaches to read the Excel file
        df = None
        read_errors = []

        # Method 1: Try with openpyxl engine (for .xlsx)
        try:
            df = pd.read_excel(uploaded_file, engine='openpyxl', header=0)
        except Exception as e:
            read_errors.append(f"openpyxl engine: {str(e)}")
            uploaded_file.seek(0)

        # Method 2: Try with xlrd engine (for older .xls files) only if needed
        if df is None and file_signature.startswith(b'\xD0\xCF\x11\xE0'):  # Only for .xls files
            try:
                # Check if xlrd is installed
                try:
                    import xlrd
                    xlrd_installed = True
                except ImportError:
                    xlrd_installed = False
                    read_errors.append("xlrd not installed (required for .xls files)")

                if xlrd_installed:
                    df = pd.read_excel(uploaded_file, engine='xlrd', header=0)
            except Exception as e:
                read_errors.append(f"xlrd engine: {str(e)}")
                uploaded_file.seek(0)

        # Method 3: Try default engine (will use openpyxl for .xlsx)
        if df is None:
            try:
                df = pd.read_excel(uploaded_file, header=0)
            except Exception as e:
                read_errors.append(f"default engine: {str(e)}")

        if df is None:
            error_msg = "❌ **Unable to read Excel file.**\n\n"
            error_msg += "**Possible solutions:**\n"
            error_msg += "- Make sure you're uploading a valid Excel file (.xlsx or .xls)\n"

            if any("xlrd" in err.lower() for err in read_errors):
                error_msg += "\n**For .xls files:**\n"
                error_msg += "This environment doesn't have the 'xlrd' package installed.\n"
                error_msg += "You can either:\n"
                error_msg += "1. Save your file as .xlsx format instead, or\n"
                error_msg += "2. Install xlrd with: `pip install xlrd`\n"

            error_msg += "\n**Technical details:**\n"
            error_msg += "\n".join(f"- {err}" for err in read_errors)
            return None, error_msg

        # Rest of your existing processing code remains exactly the same...
        df = validate_and_clean_dataframe(df)

        # Find required columns with enhanced detection
        contract_col = enhanced_column_finder(df, [
            'contract no', 'contract', 'contract number', 'contractno', 'contract num'
        ])

        payee_col = enhanced_column_finder(df, [
            'name', 'customer name', 'payee', 'fullname', 'employee', 'full name',
            'client name', 'employee name', 'employeename', 'customer', 'clientname'
        ])

        employee_col = enhanced_column_finder(df, [
            'employee number', 'ec number', 'ec num', 'employee num', 'ec number',
            'employeenumber', 'ecnumber', 'emp num', 'emp number', 'id number'
        ])

        comment_col = enhanced_column_finder(df, [
            'comment', 'description', 'transaction description', 'comments',
            'transaction', 'desc', 'transaction desc'
        ])

        # Check for missing columns with detailed feedback
        missing_columns = []
        column_suggestions = {}

        if not contract_col:
            missing_columns.append("Contract Number")
            column_suggestions["Contract Number"] = "Try columns like: Contract No, Contract, Contract Number"

        if not payee_col:
            missing_columns.append("Payee/Customer Name")
            column_suggestions["Payee/Customer Name"] = "Try columns like: Name, Customer Name, Payee, Client Name"

        if not employee_col:
            missing_columns.append("Employee Number")
            column_suggestions["Employee Number"] = "Try columns like: Employee Number, EC Number, ID Number"

        if not comment_col:
            missing_columns.append("Comment/Description")
            column_suggestions[
                "Comment/Description"] = "Try columns like: Comment, Description, Transaction Description"

        if missing_columns:
            error_msg = f"❌ **Missing required columns:** {', '.join(missing_columns)}\n\n"
            error_msg += "**Suggestions:**\n"
            for col, suggestion in column_suggestions.items():
                error_msg += f"- **{col}**: {suggestion}\n"
            error_msg += "\n**Available columns in your file:**\n"
            for col in df.columns:
                error_msg += f"- `{col}`\n"
            return None, error_msg

        # Show successful column mapping
        st.success(f"✅ **Column mapping successful:**")
        st.write(f"- Contract Number: `{contract_col}`")
        st.write(f"- Payee/Customer: `{payee_col}`")
        st.write(f"- Employee Number: `{employee_col}`")
        st.write(f"- Comments: `{comment_col}`")

        # Validate data quality
        validation_errors = []
        validation_warnings = []

        for idx, row in df.iterrows():
            errors, warnings = validate_row_data(row, idx, contract_col, payee_col, employee_col, comment_col)
            validation_errors.extend(errors)
            validation_warnings.extend(warnings)

        if validation_errors:
            error_msg = "❌ **Data validation errors found:**\n\n"
            for error in validation_errors[:10]:  # Show first 10 errors
                error_msg += f"- {error}\n"
            if len(validation_errors) > 10:
                error_msg += f"\n... and {len(validation_errors) - 10} more errors"
            return None, error_msg

        # Process merged comments
        merged_comments = detect_merged_comments(df, comment_col, contract_col, payee_col, employee_col)
        converted_data = []
        processed_indices = set()

        # Process merged rows first
        for merge in merged_comments:
            if merge['end_idx'] > merge['start_idx']:
                for idx in range(merge['start_idx'], merge['end_idx'] + 1):
                    processed_indices.add(idx)

                transactions = process_comment(merge['comment'])
                for transaction in transactions:
                    converted_data.append({
                        'Contract No': merge['contract_no'],
                        'Transaction Type': get_transaction_type(transaction['description']),
                        'Description': transaction['description'],
                        'Amount': transaction['amount'],
                        'Value Date': effective_date.strftime('%d-%b-%y'),
                        'Effective Date': effective_date.strftime('%d-%b-%y'),
                        'Post Date': datetime.now().strftime('%d-%b-%y'),
                        'Employee Number': merge['employee_no'],
                        'Payee': merge['payee']
                    })

        # Process individual rows
        for idx, row in df.iterrows():
            if idx not in processed_indices:
                contract_no = clean_cell_value(row[contract_col])
                payee = clean_cell_value(row[payee_col])
                employee_no = clean_cell_value(row[employee_col])
                comment = clean_cell_value(row[comment_col])

                transactions = process_comment(comment)
                for transaction in transactions:
                    converted_data.append({
                        'Contract No': contract_no,
                        'Transaction Type': get_transaction_type(transaction['description']),
                        'Description': transaction['description'],
                        'Amount': transaction['amount'],
                        'Value Date': effective_date.strftime('%d-%b-%y'),
                        'Effective Date': effective_date.strftime('%d-%b-%y'),
                        'Post Date': datetime.now().strftime('%d-%b-%y'),
                        'Employee Number': employee_no,
                        'Payee': payee
                    })

        if not converted_data:
            return None, "❌ No valid transactions found in the file. Please check your comment/description format."

        result_df = pd.DataFrame(converted_data)
        return result_df, None

    except Exception as e:
        error_msg = f"❌ **Error processing file:** {str(e)}\n\n"
        error_msg += "**Common solutions:**\n"
        error_msg += "- Ensure the file is a valid Excel (.xlsx or .xls) file\n"
        error_msg += "- Check that the file is not corrupted\n"
        error_msg += "- Verify that required columns exist\n"
        error_msg += "- Try saving the Excel file in a different format\n"
        return None, error_msg


def create_download_link(df, uploaded_filename):
    """Create a download link for the CSV file"""
    # Get the uploaded filename without extension and add '_converted.csv'
    base_name = os.path.splitext(uploaded_filename)[0]
    filename = f"{base_name}_converted.csv"

    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}" style="text-decoration: none;">'
    href += '<button style="background: linear-gradient(135deg, #63328a 0%, #f9c846 100%); '
    href += 'color: white; border: none; padding: 12px 24px; border-radius: 8px; '
    href += 'font-weight: bold; cursor: pointer;">📥 Download CSV</button></a>'
    return href


def main():
    # Header
    st.markdown("""
    <div class="header">
        <div class="title">Map It</div>
        <div class="subtitle">Posting Made Easier</div>
    </div>
    """, unsafe_allow_html=True)

    # Main content
    st.markdown("### 📁 Upload Excel File")

    # File uploader
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx or .xls)",
        type=['xlsx', 'xls'],
        help="Upload your Excel file containing contract data with columns for Contract No, Payee, Employee Number, and Comments"
    )

    # Date input
    st.markdown("### 📅 Set Effective Date")
    effective_date = st.date_input(
        "Effective Date",
        value=datetime.now().date(),
        help="This date will be used as the Value Date and Effective Date for all transactions"
    )

    if uploaded_file is not None:
        st.markdown("### 🔄 Processing File...")

        with st.spinner("Analyzing file structure and cleaning data..."):
            result_df, error_message = convert_file(uploaded_file, effective_date)

        if error_message:
            st.error("**File Processing Failed**")
            st.markdown(f"""
            <div class="error-details">
            {error_message}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.success("**File processed successfully!**")

            # Show statistics
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Number of Transactions", len(result_df))
            with col2:
                st.metric("Batch Total", f"${result_df['Amount'].sum():,.2f}")

            # Show preview
            st.markdown("### 👀 Preview of Converted Data")
            st.dataframe(result_df.head(5), use_container_width=True)

            if len(result_df) > 5:
                st.info(f"Showing first 5 rows of {len(result_df)} total transactions")

                # Download options
                st.markdown("### 💾 Download Results")
                download_link = create_download_link(result_df, uploaded_file.name)
                st.markdown(download_link, unsafe_allow_html=True)

            # Show transaction type breakdown
            st.markdown("### 📊 Transaction Breakdown")
            transaction_counts = result_df['Transaction Type'].value_counts()
            st.bar_chart(transaction_counts)

    else:
        st.info("👆 Please upload an Excel file to begin conversion")

        # Show expected format
        st.markdown("### 📋 Expected File Format")
        st.markdown("""
        Your Excel file should contain the following columns (names are flexible):

        - **Contract Number**: Contract No, Contract, Contract Number
        - **Payee/Customer**: Name, Customer Name, Payee, Client Name
        - **Employee Number**: Employee Number, EC Number, ID Number
        - **Comments**: Comment, Description, Transaction Description

        """)

    # Orchid watermark
    st.markdown('<div class="orchid-text">Orchid</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()