# Map It - Posting Made Easier

## Overview
Map It is a Streamlit-based web application designed to simplify the processing and conversion of Excel files containing financial transaction data. **I specifically designed this application to make posting bulk transactions via the Financial Transaction section in CreditEase significantly easier and more efficient.**

**The Problem:** Before this application, posting transactions required manually processing each individual contract and its transaction types one at a time through CreditEase's interface - a tedious process that could take hours for large batches. Each transaction had to be entered separately, with manual selection of transaction types, amounts, and contract numbers.

**The Solution:** Now, users can simply upload an Excel file containing all their transaction data, and my application automatically processes, cleans, and converts it into a properly formatted file that can be directly uploaded to CreditEase. This eliminates the need for manual data entry and reduces processing time from hours to minutes.

The application automatically extracts, cleans, and maps transaction descriptions to standardized formats while handling various data quality issues commonly found in Excel files, making bulk transaction posting in CreditEase effortless and error-free.

## Key Features

- **Bulk Transaction Processing**: Convert hundreds of transactions simultaneously instead of manual one-by-one entry
- **CreditEase Compatibility**: Output format specifically optimized for Financial Transaction section uploads
- **Time Savings**: Reduce processing time from hours to minutes for large transaction batches
- **Excel File Processing**: Supports both .xlsx and .xls file formats from various sources
- **Smart Column Detection**: Automatically identifies required columns using flexible matching
- **Data Cleaning**: Handles invisible characters, formatting issues, and Unicode normalization
- **Transaction Mapping**: Converts various transaction descriptions to standardized CreditEase formats
- **Amount Extraction**: Properly handles both whole numbers and decimal amounts
- **CSV Export**: Generates clean, formatted CSV output files ready for CreditEase upload
- **User-Friendly Interface**: Purple and gold themed UI with intuitive controls

## How It Works

1. **Export Data**: Extract transaction data from your source system into Excel
2. **Upload File**: Drag and drop your Excel file into the application
3. **Automatic Processing**: The app cleans, maps, and formats all transactions
4. **Download Results**: Get a CreditEase-ready CSV file in seconds
5. **Upload to CreditEase**: Use the generated file in the Financial Transaction section

## Time-Saving Benefits

- **Eliminates manual data entry** into CreditEase's Financial Transaction section
- **Reduces processing time** from hours to minutes for large transaction batches
- **Minimizes human error** through automated mapping and validation
- **Handles complex transactions** with multiple types per contract automatically
- **Standardizes formatting** ensuring consistency across all transactions
- **Batch processing** allows hundreds of transactions to be processed simultaneously

## Supported Transaction Types

The application recognizes and maps the following transaction types compatible with CreditEase:

### Interest Related
- Monthly Interest
- Interest Reversal
- Arrear Interest
- Penalty Interest

### Deposits & Refunds
- Direct Deposits
- Refunds

### Fees
- Service Fee
- Card Fee
- Legal Fees
- Funeral Fee
- Arrangement Fee
- Transfer Fee
- Early Settlement Fee

### Reversals
- Receipts Reversals
- Cash Drawer Receipt Reversal
- Various fee reversals
- Write-off reversals

### Status Changes
- Contract Status changes (Active, Complete, Settled, Legal, Cancelled)

### Other Transactions
- Receipts
- Cash Disbursement
- Bank transactions
- Instalments
- Write Offs (including Complete Write Off and Small Balance Credit W-Off)

## Installation Requirements

```bash
pip install streamlit pandas numpy openpyxl xlrd
```

## Usage

1. **Upload Excel File**: Select an Excel file (.xlsx or .xls) containing transaction data from your system
2. **Set Effective Date**: Choose the date that will be used for Value Date and Effective Date fields in CreditEase
3. **Automatic Processing**: The application will:
   - Detect and map required columns
   - Clean and validate data
   - Process transaction comments
   - Extract amounts correctly
   - Map descriptions to standardized formats recognized by CreditEase

4. **Download Results**: Export the processed data as a CSV file ready for upload to CreditEase's Financial Transaction section

## Required Columns

Your Excel file should contain these columns (flexible naming supported):

- **Contract Number**: Can be named: Contract No, Contract, Contract Number, etc.
- **Payee/Customer Name**: Can be named: Name, Customer Name, Payee, Client Name, etc.
- **Employee Number**: Can be named: Employee Number, EC Number, ID Number, etc.
- **Comments/Description**: Can be named: Comment, Description, Transaction Description, etc.

## Data Formatting

The application handles:
- Various date formats
- Currency amounts (with or without $ symbol)
- Commas in numbers
- Decimal amounts (properly preserves cents)
- Merged comments across multiple rows
- Invisible characters and Excel formatting artifacts

## Output Format

The generated CSV file contains these columns optimized for CreditEase upload:
- Contract No
- Transaction Type (standardized codes compatible with CreditEase)
- Description (full descriptions without truncation)
- Amount (properly formatted with decimals)
- Value Date
- Effective Date
- Post Date
- Employee Number
- Payee Name

## Common Issues Handled

- Invisible characters and Unicode issues
- Merged transaction comments
- Varied column naming conventions
- Mixed number formats (with/without decimals)
- Currency symbol variations

## CreditEase Integration

The application outputs files in a format specifically designed for seamless upload to CreditEase's Financial Transaction section, eliminating the need for:
- Manual transaction entry
- Individual contract processing
- Format conversion
- Data validation within CreditEase
- Repetitive clicking through multiple screens

## Support

For issues related to:
- File upload problems
- Column detection failures
- Amount extraction errors
- Transaction mapping questions
- CreditEase compatibility issues

Please ensure your Excel file follows the expected format and contains the required columns with appropriate data.

## Browser Compatibility

Works best with modern browsers that support:
- File API
- Base64 encoding
- Modern CSS features

## License

This application is designed for internal use with financial transaction processing in CreditEase.

---

*Built with Streamlit - Transforming hours of manual CreditEase posting into minutes of automated processing*