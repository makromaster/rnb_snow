import os
# Suppress TensorFlow warnings and logs before any other imports
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'  # Suppress TensorFlow logging
os.environ['PYTHONWARNINGS'] = 'ignore::DeprecationWarning'  # Suppress deprecation warnings

import sqlite3
import time
import pdb
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import re
import tempfile
import shutil
import urllib.request

# OCR imports
try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    print("WARNING: OCR libraries not available. PDF text extraction will be limited.")

def setup_driver():
    """Setup Chrome driver with options to prevent logout"""
    chrome_options = Options()

    # Keep session alive options
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--disable-extensions")

    # User agent to appear more like regular browser
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    driver = webdriver.Chrome(options=chrome_options)

    # Execute script to hide automation indicators
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    return driver

def get_unmatched_tickets():
    """Get tickets that haven't been matched to accounts yet and don't have extraction status"""
    conn = sqlite3.connect('ticket_matching.db')
    cursor = conn.cursor()

    cursor.execute('''
        SELECT ticket, short_description, eml_domain
        FROM snow
        WHERE (account_number IS NULL OR account_number = '')
        AND (extraction_status IS NULL OR extraction_status = '')
        ORDER BY ticket
    ''')

    tickets = cursor.fetchall()
    conn.close()
    return tickets

def update_ticket_account(ticket_number, account_number, account_name):
    """Update ticket with found account information"""
    conn = sqlite3.connect('ticket_matching.db')
    cursor = conn.cursor()

    cursor.execute('''
        UPDATE snow
        SET account_number = ?, account_name = ?
        WHERE ticket = ?
    ''', (account_number, account_name, ticket_number))

    conn.commit()
    conn.close()
    print(f"Updated {ticket_number} with account {account_number} ({account_name})")

def update_ticket_text(ticket_number, text):
    """Update ticket with extracted email text"""
    conn = sqlite3.connect('ticket_matching.db')
    cursor = conn.cursor()

    status = 'extracted' if text and text.strip() else 'nothing_to_extract'

    cursor.execute('''
        UPDATE snow
        SET text = ?, extraction_status = ?
        WHERE ticket = ?
    ''', (text, status, ticket_number))

    conn.commit()
    conn.close()
    print(f"Updated {ticket_number} with email text ({len(text) if text else 0} characters) - Status: {status}")

def scroll_to_bottom(driver):
    element = driver.find_element(By.ID, "sc_req_item.form_scroll")
    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", element)

def click_first_created_link(driver):
    """
    Find the 'Created' column in the email table and click the first link
    """
    try:
        # Find the email table
        table = driver.find_element(By.ID, "sc_req_item.u_email_client.u_item_table")

        # Find the header row to locate the 'Created' column index
        header_row = table.find_element(By.TAG_NAME, "thead").find_element(By.TAG_NAME, "tr")
        headers = header_row.find_elements(By.TAG_NAME, "th")

        created_column_index = None
        for i, header in enumerate(headers):
            header_text = header.get_attribute("textContent").strip().lower()
            if "created" in header_text:
                created_column_index = i
                print(f"Found 'Created' column at index {i}")
                break

        if created_column_index is None:
            print("Could not find 'Created' column in table")
            return False

        # Find the table body and get the first data row
        tbody = table.find_element(By.TAG_NAME, "tbody")
        data_rows = tbody.find_elements(By.TAG_NAME, "tr")

        if not data_rows:
            print("No data rows found in table")
            return False

        # Get the first row and find the cell in the 'Created' column
        first_row = data_rows[0]
        cells = first_row.find_elements(By.TAG_NAME, "td")

        if len(cells) <= created_column_index:
            print(f"Not enough cells in first row (found {len(cells)}, need {created_column_index + 1})")
            return False

        created_cell = cells[created_column_index]

        # Look for a link in the 'Created' cell
        links = created_cell.find_elements(By.TAG_NAME, "a")
        if links:
            first_link = links[0]
            link_text = first_link.get_attribute("textContent").strip()
            href = first_link.get_attribute("href")

            if href:
                print(f"Navigating to 'Created' link: {link_text}")
                print(f"URL: {href}")
                driver.get(href)
                return True
            else:
                print("No href found in 'Created' link")
                return False
        else:
            print("No links found in 'Created' column cell")
            return False

    except Exception as e:
        print(f"Error clicking 'Created' link: {e}")
        return False

def extract_email_text(driver):
    """Extract email message text from the email page"""
    try:
        element = driver.find_element(By.ID, "sys_original.u_email_client.u_message")
        value = element.get_attribute("value")
        return value if value else ""
    except Exception as e:
        print(f"Error extracting email text: {e}")
        return ""

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

def find_account_in_text(text):
    """
    Find account matches in email text using the same logic as create_database.py:
    - 10 digits starting with 00: drop 00, look up 8 digits in customer
    - 8 digits: look up directly in customer
    - 10 digits NOT starting with 00: look up in document_number or reference
    - XXX-XXXXX format: combine digits and try as 8-digit customer lookup
    - XX-XXXXXX format: combine digits and try as 8-digit customer lookup
    - Valid range check: if account is in valid ranges, accept even if not in SAP
    """
    conn = sqlite3.connect('ticket_matching.db')
    cursor = conn.cursor()
    matches = []

    # Extract regular numbers from text (8-10 digits)
    numbers = re.findall(r'\b\d{8,10}\b', text)

    # Extract dash-separated formats like "239-63450" and "20-572883"
    dash_numbers_3_5 = re.findall(r'\b(\d{3})-(\d{5})\b', text)
    dash_numbers_2_6 = re.findall(r'\b(\d{2})-(\d{6})\b', text)

    # Process dash-separated numbers first
    # Handle XXX-XXXXX format (3-5 digits)
    for part1, part2 in dash_numbers_3_5:
        combined = part1 + part2  # "239" + "63450" = "23963450"
        if len(combined) == 8:
            # First try SAP lookup
            cursor.execute("SELECT * FROM sap WHERE customer = ?", (combined,))
            results = cursor.fetchall()
            if results:
                matches.extend([(combined, 'customer_dash', results[0])])
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
                matches.extend([(combined, 'customer_dash', results[0])])
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
                matches.extend([(account_num, 'customer', results[0])])
            elif is_valid_account_range(account_num):
                # Valid range but not in SAP
                fake_record = ('', '', 0, '', 'VALID ACCOUNT (NOT IN SAP)', account_num)
                matches.extend([(account_num, 'customer_valid', fake_record)])

        elif len(number) == 8:
            # Look up directly in customer
            cursor.execute("SELECT * FROM sap WHERE customer = ?", (number,))
            results = cursor.fetchall()
            if results:
                matches.extend([(number, 'customer', results[0])])
            elif is_valid_account_range(number):
                # Valid range but not in SAP
                fake_record = ('', '', 0, '', 'VALID ACCOUNT (NOT IN SAP)', number)
                matches.extend([(number, 'customer_valid', fake_record)])

        elif len(number) == 10 and not number.startswith('00'):
            # Look up in document_number or reference (invoice number)
            cursor.execute("SELECT * FROM sap WHERE document_number = ? OR reference = ?", (number, number))
            results = cursor.fetchall()
            if results:
                matches.extend([(number, 'invoice', results[0])])

    conn.close()
    return matches

def extract_text_from_pdf_ocr(pdf_path):
    """
    Extract text from PDF using OCR (for scanned PDFs or images)
    Returns extracted text or empty string if OCR fails
    """
    if not OCR_AVAILABLE:
        print("OCR libraries not available - cannot extract text from PDF")
        return ""

    try:
        print(f"Running OCR on PDF: {pdf_path}")

        # Convert PDF to images
        images = convert_from_path(pdf_path, dpi=300)
        print(f"Converted PDF to {len(images)} image(s)")

        # Extract text from each page
        extracted_text = []
        for i, image in enumerate(images):
            print(f"Processing page {i+1}/{len(images)} with OCR...")
            text = pytesseract.image_to_string(image)
            if text.strip():
                extracted_text.append(text)

        full_text = '\n\n'.join(extracted_text)
        print(f"OCR extracted {len(full_text)} characters from PDF")
        return full_text

    except Exception as e:
        print(f"Error during OCR extraction: {e}")
        return ""

def find_and_download_pdf_attachments(driver):
    """
    Find PDF attachments on the current ServiceNow email page and download them
    Returns list of downloaded PDF file paths
    """
    downloaded_pdfs = []

    try:
        # Look for attachment links - ServiceNow typically has attachments in a specific section
        # Try multiple patterns to find PDFs
        attachment_links = []

        # Pattern 1: Direct PDF links
        attachment_links.extend(driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]"))

        # Pattern 2: Attachment table links
        attachment_links.extend(driver.find_elements(By.XPATH, "//a[contains(text(), '.pdf')]"))

        # Pattern 3: ServiceNow attachment links (common pattern)
        attachment_links.extend(driver.find_elements(By.XPATH, "//a[contains(@class, 'attachment')]"))

        if not attachment_links:
            print("No PDF attachments found on page")
            return downloaded_pdfs

        # Remove duplicates
        unique_links = []
        seen_hrefs = set()
        for link in attachment_links:
            href = link.get_attribute('href')
            if href and href not in seen_hrefs and '.pdf' in href.lower():
                unique_links.append(link)
                seen_hrefs.add(href)

        if not unique_links:
            print("No unique PDF attachments found")
            return downloaded_pdfs

        print(f"Found {len(unique_links)} unique PDF attachment(s)")

        # Create temp directory for downloads
        temp_dir = tempfile.mkdtemp()
        print(f"Created temp directory: {temp_dir}")

        # Get cookies from selenium session for authenticated download
        cookies = driver.get_cookies()
        cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}

        for i, link in enumerate(unique_links[:3]):  # Limit to first 3 PDFs
            try:
                href = link.get_attribute('href')
                link_text = link.get_attribute('textContent').strip()
                print(f"Processing PDF: {link_text} ({href})")

                pdf_filename = os.path.join(temp_dir, f"attachment_{i}.pdf")

                # Download PDF using urllib with cookies
                opener = urllib.request.build_opener()
                opener.addheaders = [
                    ('Cookie', '; '.join([f"{k}={v}" for k, v in cookie_dict.items()])),
                    ('User-Agent', 'Mozilla/5.0')
                ]
                urllib.request.install_opener(opener)

                print(f"Downloading PDF to: {pdf_filename}")
                urllib.request.urlretrieve(href, pdf_filename)

                if os.path.exists(pdf_filename) and os.path.getsize(pdf_filename) > 0:
                    print(f"✓ Successfully downloaded PDF ({os.path.getsize(pdf_filename)} bytes)")
                    downloaded_pdfs.append(pdf_filename)
                else:
                    print(f"✗ Failed to download PDF or file is empty")

            except Exception as e:
                print(f"Error downloading PDF: {e}")
                continue

    except Exception as e:
        print(f"Error finding PDF attachments: {e}")

    return downloaded_pdfs

def manual_debug_session():
    """
    Main function that sets up Selenium and pauses for manual interaction
    """

    # Get unmatched tickets
    tickets = get_unmatched_tickets()
    driver = setup_driver()
    print(f"\nFound {len(tickets)} unmatched tickets to process")

    # Navigate to ServiceNow login page first
    print("\n=== LOGGING INTO SERVICENOW ===")
    print("Navigating to https://emeops03.service-now.com/")
    driver.get("https://emeops03.service-now.com/")

    # Wait for Microsoft login redirect and manual login
    print("Waiting for login to complete...")
    print("Please complete the Microsoft login process in the browser")
    print("Script will continue when you're back on ServiceNow (emeops03.service-now.com)")
    time.sleep(30)
    # Wait until we're back on ServiceNow domain (not just a parameter in Microsoft login)
    while True:
        current_url = driver.current_url
        if current_url.startswith("https://emeops03.service-now.com") and "microsoftonline.com" not in current_url:
            print(f"✓ Login successful! Current URL: {current_url}")
            break
        else:
            print(f"Still on login page: {current_url[:80]}...")
            time.sleep(2)

    print("Starting ticket processing...")
    time.sleep(2)

    # Now iterate over tickets
    for ticket in tickets:
        ticket_number = ticket[0]
        print(f"\nProcessing ticket: {ticket_number}")
        ticket_url = f"https://emeops03.service-now.com/text_search_exact_match.do?sysparm_search={ticket_number}"
        driver.get(ticket_url)

        try:
            # Wait for page to load
            time.sleep(2)

            # Scroll to see the email table
            scroll_to_bottom(driver)
            time.sleep(1)

            # Click the first link in the 'Created' column
            if click_first_created_link(driver):
                print("Successfully clicked 'Created' link")
                time.sleep(2)  # Wait for new page to load

                # Extract email message text
                email_text = extract_email_text(driver)
                # Always update the ticket text and extraction status
                update_ticket_text(ticket_number, email_text)

                if email_text and email_text.strip():
                    print(f"Extracted email text ({len(email_text)} characters)")

                    # Find account numbers in the email text
                    matches = find_account_in_text(email_text)
                    if matches:
                        # Take the first match for account assignment
                        number, match_type, sap_record = matches[0]
                        account_number = sap_record[5] if sap_record[5] else number  # customer field
                        account_name = sap_record[4] if sap_record[4] else ''  # name field

                        update_ticket_account(ticket_number, account_number, account_name)
                        print(f"Found account {account_number} via {match_type} in email text")
                    else:
                        print("No account numbers found in email text")

                        # Try OCR on PDF attachments if available
                        if OCR_AVAILABLE:
                            print("Attempting to find and process PDF attachments with OCR...")
                            pdf_files = find_and_download_pdf_attachments(driver)

                            if pdf_files:
                                print(f"Downloaded {len(pdf_files)} PDF(s) - processing with OCR...")

                                # Process each PDF with OCR
                                all_ocr_text = []
                                for pdf_path in pdf_files:
                                    ocr_text = extract_text_from_pdf_ocr(pdf_path)
                                    if ocr_text:
                                        all_ocr_text.append(ocr_text)

                                # Combine all OCR text
                                combined_ocr_text = '\n\n'.join(all_ocr_text)

                                if combined_ocr_text:
                                    print(f"Total OCR text extracted: {len(combined_ocr_text)} characters")

                                    # Search for account numbers in OCR text
                                    ocr_matches = find_account_in_text(combined_ocr_text)
                                    if ocr_matches:
                                        # Take the first match for account assignment
                                        number, match_type, sap_record = ocr_matches[0]
                                        account_number = sap_record[5] if sap_record[5] else number
                                        account_name = sap_record[4] if sap_record[4] else ''

                                        update_ticket_account(ticket_number, account_number, account_name)
                                        print(f"✓ Found account {account_number} via {match_type} in PDF (OCR)")

                                        # Also update text field with OCR results
                                        combined_text = email_text + "\n\n[OCR from PDF]\n" + combined_ocr_text
                                        update_ticket_text(ticket_number, combined_text)
                                    else:
                                        print("No account numbers found in OCR text")
                                else:
                                    print("No text extracted from PDFs")

                                # Clean up temp files
                                for pdf_path in pdf_files:
                                    try:
                                        temp_dir = os.path.dirname(pdf_path)
                                        if os.path.exists(temp_dir):
                                            shutil.rmtree(temp_dir)
                                            print(f"Cleaned up temp directory: {temp_dir}")
                                    except Exception as e:
                                        print(f"Error cleaning up temp files: {e}")
                            else:
                                print("No PDF attachments found")
                        else:
                            print("OCR not available - install pytesseract and pdf2image for PDF support")
                else:
                    print("No email text found - marked as 'nothing_to_extract'")

            else:
                print("Failed to click 'Created' link - marked as 'nothing_to_extract'")
                update_ticket_text(ticket_number, "")

        except Exception as e:
            print(f"Error processing ticket {ticket_number}: {e}")

        # Add a small pause between tickets
        time.sleep(1)



if __name__ == "__main__":
    manual_debug_session()