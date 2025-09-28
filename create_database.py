import sqlite3
import pandas as pd
import re
import os

def create_database():
    """Create SQLite database with snow and sap tables"""
    conn = sqlite3.connect('ticket_matching.db')
    cursor = conn.cursor()

    # Create snow table with unique constraint on ticket
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS snow (
            ticket TEXT PRIMARY KEY,
            short_description TEXT,
            eml_domain TEXT,
            account_number TEXT,
            account_name TEXT,
            text TEXT,
            extraction_status TEXT
        )
    ''')

    # Add extraction_status column if it doesn't exist (for existing databases)
    try:
        cursor.execute('ALTER TABLE snow ADD COLUMN extraction_status TEXT')
        conn.commit()
    except sqlite3.OperationalError:
        # Column already exists
        pass

    # Create sap table with unique constraint on customer
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sap (
            document_number TEXT,
            reference TEXT,
            company_code_currency_value REAL,
            company_code_currency_key TEXT,
            name TEXT,
            customer TEXT,
            PRIMARY KEY (customer, document_number)
        )
    ''')

    conn.commit()
    return conn

def load_sap_data(conn, csv_file='RnB OP.csv'):
    """Load SAP data from CSV file into sap table (updates existing records)"""
    cursor = conn.cursor()

    try:
        # Read CSV file in chunks to handle large files
        chunk_size = 1000
        total_loaded = 0

        for chunk in pd.read_csv(csv_file, chunksize=chunk_size):
            # Clean column names and rename to match our schema
            chunk.columns = chunk.columns.str.strip().str.replace('"', '').str.replace('ï»¿', '')
            chunk = chunk.rename(columns={
                'Document Number': 'document_number',
                'Reference': 'reference',
                'Company Code Currency Value': 'company_code_currency_value',
                'Company Code Currency Key': 'company_code_currency_key',
                'Name': 'name',
                'Customer': 'customer'
            })

            # Remove empty rows
            chunk = chunk.dropna(subset=['customer']).copy()
            chunk = chunk[chunk['customer'] != '']

            # Insert records one by one to avoid SQL variable limit
            for _, row in chunk.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO sap
                    (document_number, reference, company_code_currency_value,
                     company_code_currency_key, name, customer)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (row.get('document_number', ''), row.get('reference', ''),
                      row.get('company_code_currency_value', 0),
                      row.get('company_code_currency_key', ''),
                      row.get('name', ''), row.get('customer', '')))

            total_loaded += len(chunk)
            if total_loaded % 5000 == 0:
                print(f"Loaded {total_loaded} records...")

        conn.commit()
        print(f"Loaded {total_loaded} records into sap table")

    except Exception as e:
        print(f"Error loading SAP data: {e}")

def load_snow_data(conn, csv_file=None, tickets_data=None):
    """Load ServiceNow data from CSV file or list into snow table (preserves existing data)"""
    cursor = conn.cursor()

    if csv_file:
        try:
            df = pd.read_csv(csv_file)
            new_tickets = 0
            updated_tickets = 0

            # Handle sc_req_item.csv format: number, state, assigned_to, sys_created_on, sys_updated_on, short_description, u_sender_address, sys_updated_by, assignment_group
            if 'number' in df.columns and 'short_description' in df.columns:
                for _, row in df.iterrows():
                    ticket_number = row['number']

                    # Check if ticket already exists
                    cursor.execute('SELECT ticket, account_number, account_name, text, extraction_status FROM snow WHERE ticket = ?', (ticket_number,))
                    existing = cursor.fetchone()

                    # Extract email domain from u_sender_address
                    email_domain = None
                    if 'u_sender_address' in row and pd.notna(row['u_sender_address']):
                        email = str(row['u_sender_address'])
                        if '@' in email:
                            email_domain = email.split('@')[1]

                    if existing:
                        # Ticket exists - only update description and email domain, preserve other data
                        cursor.execute('''
                            UPDATE snow
                            SET short_description = ?, eml_domain = ?
                            WHERE ticket = ?
                        ''', (row['short_description'], email_domain, ticket_number))
                        updated_tickets += 1
                    else:
                        # New ticket - insert with NULL values for preserved fields
                        cursor.execute('''
                            INSERT INTO snow (ticket, short_description, eml_domain, account_number, account_name, text, extraction_status)
                            VALUES (?, ?, ?, NULL, NULL, NULL, NULL)
                        ''', (ticket_number, row['short_description'], email_domain))
                        new_tickets += 1
            else:
                # Generic format: TICKET, short description, eml_domain, account number, Account Name
                for _, row in df.iterrows():
                    ticket_number = row[0]

                    # Check if ticket already exists
                    cursor.execute('SELECT ticket FROM snow WHERE ticket = ?', (ticket_number,))
                    existing = cursor.fetchone()

                    if existing:
                        # Ticket exists - only update basic info, preserve extracted data
                        cursor.execute('''
                            UPDATE snow
                            SET short_description = ?, eml_domain = ?
                            WHERE ticket = ?
                        ''', (row[1], row[2] if len(row) > 2 else None, ticket_number))
                        updated_tickets += 1
                    else:
                        # New ticket - insert fresh
                        cursor.execute('''
                            INSERT INTO snow (ticket, short_description, eml_domain, account_number, account_name, text, extraction_status)
                            VALUES (?, ?, ?, ?, ?, NULL, NULL)
                        ''', (ticket_number, row[1], row[2] if len(row) > 2 else None,
                             row[3] if len(row) > 3 else None, row[4] if len(row) > 4 else None))
                        new_tickets += 1

            conn.commit()
            print(f"Loaded from {csv_file}: {new_tickets} new tickets, {updated_tickets} existing tickets updated")
        except Exception as e:
            print(f"Error loading ServiceNow data: {e}")

    elif tickets_data:
        new_tickets = 0
        for ticket_data in tickets_data:
            ticket_number = ticket_data[0]

            # Check if ticket exists
            cursor.execute('SELECT ticket FROM snow WHERE ticket = ?', (ticket_number,))
            existing = cursor.fetchone()

            if not existing:
                cursor.execute('''
                    INSERT INTO snow (ticket, short_description, eml_domain, account_number, account_name, text, extraction_status)
                    VALUES (?, ?, ?, ?, ?, ?, NULL)
                ''', ticket_data)
                new_tickets += 1

        conn.commit()
        print(f"Loaded {new_tickets} new tickets from provided data")

def is_valid_account_range(account_number):
    """
    Check if account number is in valid ranges:
    20600000-20699999, 25300000-25399999, 20300000-24999999, 25900000-25999999
    """
    try:
        num = int(account_number)
        return (
            (20600000 <= num <= 20699999) or
            (25300000 <= num <= 25399999) or
            (20300000 <= num <= 24999999) or
            (25900000 <= num <= 25999999)
        )
    except (ValueError, TypeError):
        return False

def find_account_matches(short_description, conn):
    """
    Find account matches based on the description using the specified logic:
    - 10 digits starting with 00: drop 00, look up 8 digits in customer
    - 8 digits: look up directly in customer
    - 10 digits NOT starting with 00: look up in document_number or reference
    - XXX-XXXXX format: combine digits and try as 8-digit customer lookup
    - XX-XXXXXX format: combine digits and try as 8-digit customer lookup
    - Valid range check: if account is in valid ranges, accept even if not in SAP
    """
    cursor = conn.cursor()
    matches = []

    # Extract regular numbers from description (8-10 digits)
    numbers = re.findall(r'\b\d{8,10}\b', short_description)

    # Extract dash-separated formats like "239-63450" and "20-572883"
    dash_numbers_3_5 = re.findall(r'\b(\d{3})-(\d{5})\b', short_description)
    dash_numbers_2_6 = re.findall(r'\b(\d{2})-(\d{6})\b', short_description)

    # Process dash-separated numbers first
    # Handle XXX-XXXXX format (3-5 digits)
    for part1, part2 in dash_numbers_3_5:
        combined = part1 + part2  # "239" + "63450" = "23963450"
        if len(combined) == 8:
            # First try SAP lookup
            cursor.execute("SELECT * FROM sap WHERE customer = ?", (combined,))
            results = cursor.fetchall()
            if results:
                matches.extend([(combined, 'customer_dash', result) for result in results])
            elif is_valid_account_range(combined):
                # Valid range but not in SAP - create fake record
                fake_record = ('', '', 0, '', 'VALID ACCOUNT (NOT IN SAP)', combined)
                matches.extend([(combined, 'customer_dash_valid', fake_record)])

    # Handle XX-XXXXXX format (2-6 digits)
    for part1, part2 in dash_numbers_2_6:
        combined = part1 + part2  # "20" + "572883" = "20572883"
        if len(combined) == 8:
            # First try SAP lookup
            cursor.execute("SELECT * FROM sap WHERE customer = ?", (combined,))
            results = cursor.fetchall()
            if results:
                matches.extend([(combined, 'customer_dash', result) for result in results])
            elif is_valid_account_range(combined):
                # Valid range but not in SAP - create fake record
                fake_record = ('', '', 0, '', 'VALID ACCOUNT (NOT IN SAP)', combined)
                matches.extend([(combined, 'customer_dash_valid', fake_record)])

    # Process regular numbers
    for number in numbers:
        if len(number) == 10 and number.startswith('00'):
            # Drop 00 and look up 8 digit number in customer
            account_num = number[2:]
            cursor.execute("SELECT * FROM sap WHERE customer = ?", (account_num,))
            results = cursor.fetchall()
            if results:
                matches.extend([(account_num, 'customer', result) for result in results])
            elif is_valid_account_range(account_num):
                # Valid range but not in SAP
                fake_record = ('', '', 0, '', 'VALID ACCOUNT (NOT IN SAP)', account_num)
                matches.extend([(account_num, 'customer_valid', fake_record)])

        elif len(number) == 8:
            # Look up directly in customer
            cursor.execute("SELECT * FROM sap WHERE customer = ?", (number,))
            results = cursor.fetchall()
            if results:
                matches.extend([(number, 'customer', result) for result in results])
            elif is_valid_account_range(number):
                # Valid range but not in SAP
                fake_record = ('', '', 0, '', 'VALID ACCOUNT (NOT IN SAP)', number)
                matches.extend([(number, 'customer_valid', fake_record)])

        elif len(number) == 10 and not number.startswith('00'):
            # Look up in document_number or reference (invoice number)
            cursor.execute("SELECT * FROM sap WHERE document_number = ? OR reference = ?", (number, number))
            results = cursor.fetchall()
            if results:
                matches.extend([(number, 'invoice', result) for result in results])

    return matches

def process_all_tickets(conn):
    """Process all tickets and find account matches"""
    cursor = conn.cursor()

    # Get all tickets without account assignments or with outdated assignments
    cursor.execute("SELECT ticket, short_description FROM snow")
    tickets = cursor.fetchall()

    matched_count = 0
    for ticket_num, description in tickets:
        matches = find_account_matches(description, conn)

        if matches:
            # Take the first match for account assignment
            number, match_type, sap_record = matches[0]
            account_number = sap_record[5] if sap_record[5] else number  # customer field
            account_name = sap_record[4] if sap_record[4] else ''  # name field

            # Update ticket with account info
            cursor.execute('''
                UPDATE snow
                SET account_number = ?, account_name = ?
                WHERE ticket = ?
            ''', (account_number, account_name, ticket_num))

            matched_count += 1
            print(f"Ticket {ticket_num}: Matched to account {account_number} via {match_type}")
        else:
            # Clear account info if no matches found
            cursor.execute('''
                UPDATE snow
                SET account_number = NULL, account_name = NULL
                WHERE ticket = ?
            ''', (ticket_num,))

    conn.commit()
    return matched_count

def show_results(conn):
    """Display results of ticket matching"""
    cursor = conn.cursor()

    print("\n=== TICKET MATCHING RESULTS ===")
    cursor.execute('''
        SELECT ticket, short_description, account_number, account_name
        FROM snow
        ORDER BY ticket
    ''')

    results = cursor.fetchall()
    for ticket, desc, acc_num, acc_name in results:
        status = "MATCHED" if acc_num else "NO MATCH"
        print(f"{ticket} ({status}): {desc[:60]}...")
        if acc_num:
            print(f"  -> Account: {acc_num} ({acc_name})")
        print()

    # Statistics
    cursor.execute("SELECT COUNT(*) FROM snow WHERE account_number IS NOT NULL")
    matched = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM snow")
    total = cursor.fetchone()[0]

    if total > 0:
        print(f"SUMMARY: {matched}/{total} tickets matched to accounts ({matched/total*100:.1f}%)")
    else:
        print("SUMMARY: No tickets in database yet")

def clear_sap_data(conn):
    """Clear all SAP data to prepare for fresh load"""
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sap")
    conn.commit()
    print("Cleared existing SAP data")

def update_database(sap_file=None, snow_file=None, snow_data=None, progress_callback=None):
    """Update database with new data files"""
    conn = create_database()

    if progress_callback:
        progress_callback("Creating database structure...")

    if sap_file and os.path.exists(sap_file):
        if progress_callback:
            progress_callback(f"Clearing old SAP data...")
        clear_sap_data(conn)

        if progress_callback:
            progress_callback(f"Loading SAP data from {sap_file}...")
        load_sap_data(conn, sap_file)

    if snow_file and os.path.exists(snow_file):
        if progress_callback:
            progress_callback(f"Loading ServiceNow data from {snow_file}...")
        load_snow_data(conn, csv_file=snow_file)

    if snow_data:
        if progress_callback:
            progress_callback("Loading ServiceNow data from provided list...")
        load_snow_data(conn, tickets_data=snow_data)

    if progress_callback:
        progress_callback("Processing ticket matches...")
    matched = process_all_tickets(conn)

    if progress_callback:
        progress_callback(f"Complete! {matched} new matches found")

    conn.close()
    return matched

def get_database_stats():
    """Get current database statistics"""
    conn = sqlite3.connect('ticket_matching.db')
    cursor = conn.cursor()

    # Total tickets
    cursor.execute("SELECT COUNT(*) FROM snow")
    total_tickets = cursor.fetchone()[0]

    # Matched tickets
    cursor.execute("SELECT COUNT(*) FROM snow WHERE account_number IS NOT NULL AND account_number != ''")
    matched_tickets = cursor.fetchone()[0]

    # Extraction status counts
    cursor.execute("SELECT COUNT(*) FROM snow WHERE extraction_status = 'extracted'")
    extracted_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM snow WHERE extraction_status = 'nothing_to_extract'")
    nothing_to_extract_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM snow WHERE extraction_status IS NULL OR extraction_status = ''")
    pending_extraction_count = cursor.fetchone()[0]

    # SAP records
    cursor.execute("SELECT COUNT(*) FROM sap")
    sap_records = cursor.fetchone()[0]

    conn.close()

    return {
        'total_tickets': total_tickets,
        'matched_tickets': matched_tickets,
        'match_percentage': (matched_tickets / total_tickets * 100) if total_tickets > 0 else 0,
        'extracted_count': extracted_count,
        'nothing_to_extract_count': nothing_to_extract_count,
        'pending_extraction_count': pending_extraction_count,
        'sap_records': sap_records
    }

if __name__ == "__main__":
    print("=== DATABASE SETUP ===")
    print("This module is now designed to be used with the GUI.")
    print("Run 'python gui_main.py' to use the graphical interface.")
    print("Or import this module to use the functions programmatically.")