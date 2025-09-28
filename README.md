# RnB Snow Ticket Matching System - GUI Version

## Overview
This system matches ServiceNow tickets to customer accounts using SAP data. The new GUI interface provides user-friendly access to all features while preserving extracted data between runs.

## Key Features
- **Data Preservation**: Existing ticket data (account matches, email text) is preserved when reloading CSV files
- **Smart Extraction Tracking**: Prevents infinite retries of tickets with no extractable content
- **GUI Interface**: User-friendly graphical interface for file selection and processing
- **Enhanced Pattern Recognition**: Supports multiple account number formats including XX-XXXXXX

## Quick Start

### 1. Launch the GUI
```bash
python gui_main.py
```

### 2. Select Files
- **SAP Data File**: Choose your RnB OP.csv or equivalent SAP export
- **ServiceNow File**: Choose your sc_req_item.csv or equivalent ticket export

### 3. Process Files
Click "Process Files" to:
- Load SAP data (replaces existing SAP records)
- Load ticket data (preserves existing ticket information)
- Run account matching on all tickets

### 4. View Results
The statistics panel shows:
- Total tickets and match percentage
- Extraction status counts
- SAP record count

### 5. Extract Additional Data (Optional)
Click "Launch Selenium Session" to extract email content from unprocessed tickets.

### 6. Export Results
Click "Export to CSV" to generate a report containing:
- Only tickets from the selected ServiceNow CSV file
- All snow table columns (ticket, description, account data, email text, extraction status)
- Current matching and extraction results

## Account Number Pattern Recognition

The system recognizes these account number formats:

### Customer Account Patterns (8-digit lookups):
- `20572883` - Direct 8-digit number
- `20-572883` - **NEW**: XX-XXXXXX format
- `239-63450` - XXX-XXXXX format
- `0020572883` - 10-digit starting with 00 (drops the 00)

### Invoice Number Patterns (10-digit lookups):
- `1234567890` - 10-digit number (not starting with 00)

### Valid Account Ranges:
- 20300000-24999999
- 20600000-20699999
- 25300000-25399999
- 25900000-25900000

## Data Preservation Logic

### SAP Data
- **Always replaced** with new file data
- Ensures clean, up-to-date customer information

### ServiceNow Data
- **New tickets**: Added to database
- **Existing tickets**: Only description and email domain updated
- **Preserved fields**: account_number, account_name, text, extraction_status

## Extraction Status Tracking

### Status Values:
- `NULL`: Not yet processed by Selenium
- `'extracted'`: Email text successfully extracted
- `'nothing_to_extract'`: No email content available (skipped in future runs)

### Benefits:
- Prevents infinite retries of empty tickets
- Allows incremental processing of new batches
- Preserves valuable extracted data

## Files Structure

### Core Files:
- `gui_main.py` - Main GUI application
- `create_database.py` - Database operations and matching logic
- `selenium_debug_session.py` - Web automation for email extraction
- `ticket_matching.db` - SQLite database

### Data Files:
- `RnB OP.csv` - SAP customer/financial data
- `sc_req_item.csv` - ServiceNow ticket data

## Usage Workflow

### Initial Setup:
1. Launch GUI: `python gui_main.py`
2. Select both CSV files
3. Click "Process Files"
4. Review statistics

### Adding New Tickets:
1. Select updated ServiceNow CSV file
2. Keep same SAP file (or update if needed)
3. Click "Process Files"
4. Only new tickets are added, existing data preserved

### Extract Email Content:
1. Click "Launch Selenium Session"
2. Complete manual login in browser
3. Let automation extract email content
4. Refresh statistics to see results

### Export Results:
1. Select ServiceNow CSV file
2. Click "Export to CSV"
3. Choose save location
4. CSV file contains only tickets from CSV with all data columns

## Troubleshooting

### GUI Won't Start:
- Ensure Python has tkinter support
- Run: `python -c "import tkinter"`

### No New Matches:
- Check if account numbers are in valid ranges
- Verify SAP data contains customer records
- Try Selenium extraction for additional account numbers

### Database Issues:
- Database automatically handles schema updates
- Backup `ticket_matching.db` before major changes

### CSV Export Issues:
- **Permission error**: Ensure CSV file is not open in another application
- **Encoding issues**: Files are saved with UTF-8 encoding for international characters
- **Large files**: CSV export handles large datasets efficiently without memory issues

## Technical Details

### Database Schema:
```sql
CREATE TABLE snow (
    ticket TEXT PRIMARY KEY,
    short_description TEXT,
    eml_domain TEXT,
    account_number TEXT,
    account_name TEXT,
    text TEXT,
    extraction_status TEXT
);

CREATE TABLE sap (
    document_number TEXT,
    reference TEXT,
    company_code_currency_value REAL,
    company_code_currency_key TEXT,
    name TEXT,
    customer TEXT,
    PRIMARY KEY (customer, document_number)
);
```

### Dependencies:
- selenium (web automation)
- sqlite3 (database - included with Python)
- csv (data export - included with Python)
- tkinter (GUI - included with Python)

## Migration from Old System

The new system is fully backward compatible. Simply:
1. Keep your existing `ticket_matching.db`
2. Run `python gui_main.py`
3. Your data will be automatically preserved and enhanced

---

For technical support or feature requests, refer to the code documentation or contact the development team.