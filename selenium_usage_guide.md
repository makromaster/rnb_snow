# Selenium Debug Session Usage Guide

## Overview
This script solves the ServiceNow logout problem by using Python's debugger (`pdb`) to pause execution, allowing you to manually log in and keep the session alive throughout the automated processing.

## Files
- `selenium_debug_session.py` - Main script with debugger breakpoints
- `ticket_matching.db` - Database with matched/unmatched tickets

## How to Use

### 1. Quick Test (Recommended First)
```bash
python selenium_debug_session.py test
```
This opens a browser and pauses for 5 minutes to test your ServiceNow login.

### 2. Full Processing Session
```bash
python selenium_debug_session.py
```

## Step-by-Step Process

### When Script Starts:
1. Chrome browser opens automatically
2. Script pauses with debugger prompt: `(Pdb)`
3. You see: "DEBUGGER ACTIVATED - Manual Login Time!"

### Your Actions:
1. **In the browser**: Navigate to your ServiceNow instance
2. **Log in manually** with your credentials
3. **Test navigation** - go to any ticket to verify you're logged in
4. **In terminal**: Type `c` and press Enter to continue

### Automated Processing:
- Script processes unmatched tickets one by one
- Extracts account numbers from each ticket page
- Updates database with found matches
- Browser stays logged in throughout

### Debugger Commands:
- `c` - Continue execution
- `n` - Next line/step
- `l` - List current code
- `p variable_name` - Print variable value
- `exit()` - Quit and close browser

## Configuration

### ServiceNow URL Configuration:
The script is already configured for your instance:
```python
ticket_url = f"https://emeops03.service-now.com/text_search_exact_match.do?sysparm_search={ticket_number}"
```

If you need to change the instance, update this line in `selenium_debug_session.py`.

### Adjust Account Extraction Patterns:
The script looks for these patterns in ticket pages:
- `konto: 12345678`
- `account: 12345678`
- `kunde: 12345678`
- `customer: 12345678`
- `kdkto: 12345678`
- `debitor: 12345678`
- Any 8-digit number

### Processing Limits:
Currently processes 50 unmatched tickets at a time. Change in:
```python
LIMIT 50  # Increase this number
```

## Troubleshooting

### If Browser Closes Unexpectedly:
- Restart script
- Browser settings are configured to prevent automation detection
- User agent is set to mimic regular Chrome browser

### If Login Fails:
- Try the test mode first: `python selenium_debug_session.py test`
- Manually verify you can log into ServiceNow in regular browser
- Check if your organization requires VPN

### If No Account Numbers Found:
- Use debugger to inspect page source
- Add additional regex patterns for your specific format
- Check if account numbers are in different fields/sections

## Example Session Output

```
=== SELENIUM DEBUG SESSION FOR SERVICENOW ===
Found 1313 unmatched tickets to process

DEBUGGER ACTIVATED - Manual Login Time!
(Pdb) c

Processing ticket 1/50: RITM17699925
Description: Some issue with customer account...
Found account: 23589010
Updated RITM17699925 with account 23589010 (CUSTOMER NAME)

Processing ticket 2/50: RITM17699924
Description: Another issue...
No account number found on page

...
```

## Benefits of This Approach

1. **Session Persistence**: Manual login keeps session alive
2. **Full Control**: Can inspect any ticket manually during processing
3. **Flexible**: Can pause/resume at any point
4. **Safe**: No automated login attempts that might trigger security
5. **Debuggable**: Can inspect variables and page content in real-time

## Next Steps After Processing

1. Check results: `python -c "import sqlite3; conn=sqlite3.connect('ticket_matching.db'); cursor=conn.cursor(); cursor.execute('SELECT COUNT(*) FROM snow WHERE account_number IS NOT NULL'); print(f'Matched: {cursor.fetchone()[0]}')"`

2. Re-run pattern matching: `python create_database.py`

3. Process remaining unmatched tickets with updated patterns

## Security Notes

- Browser configured to avoid automation detection
- No credentials stored in code
- Manual login reduces security risks
- Session stays in browser for manual verification