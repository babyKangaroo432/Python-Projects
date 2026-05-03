"""
Google Sheets to HTML Email Automator
Author: Joseph Norman

This script automates the process of fetching scout payment data from Google Sheets,
converting specific spreadsheet ranges into styled HTML tables, and sending 
personalized email notifications via Gmail SMTP. 

Key Features:
- Google Sheets API integration for real-time data retrieval.
- Dynamic HTML generation with CSS styling preserved from the spreadsheet.
- Secure credential management using environment variables.
"""


import os
import socket
import smtplib
import gspread
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
from google.oauth2 import service_account

# --- CONFIGURATION (Use Environment Variables for Security) ---
MY_EMAIL = os.environ.get("EMAIL_USER", "your-email@gmail.com")
MY_PASSWORD = os.environ.get("EMAIL_PASS", "")  # Use Google App Password
SPREADSHEET_NAME = "" 
WORKSHEET_NAME = "Data"  # Ensure this matches your tab name
SPREADSHEET_ID = ""      # Add your ID here
RANGE_NAME = "A1:E10"    # Add your range here

def get_cell_html(creds):
    """Retrieves Google Sheet data and converts it to a styled HTML table."""
    try:
        service = build('sheets', 'v4', credentials=creds)
        result = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID, 
            ranges=[RANGE_NAME], 
            includeGridData=True
        ).execute()

        table_html = '<table style="border-collapse: collapse; font-family: sans-serif;">'
        
        def get_rgb(color_obj, default_color="rgb(0,0,0)"):
            if not color_obj or not any(key in color_obj for key in ['red', 'green', 'blue']):
                return default_color
            r = int(color_obj.get('red', 0) * 255)
            g = int(color_obj.get('green', 0) * 255)
            b = int(color_obj.get('blue', 0) * 255)
            return f"rgb({r},{g},{b})"

        rows = result['sheets'][0]['data'][0].get('rowData', [])
        max_cols = max([len(row.get('values', [])) for row in rows]) if rows else 0

        for row in rows:
            table_html += '<tr>'
            cells = row.get('values', [])
            for i in range(max_cols):
                cell = cells[i] if i < len(cells) else {}
                text = cell.get('formattedValue', '')
                eff_format = cell.get('effectiveFormat', {})
                text_format = eff_format.get('textFormat', {})
                
                f_size = text_format.get('fontSize', 10)
                bg = get_rgb(eff_format.get('backgroundColor', {}), "rgb(255,255,255)")
                fg = get_rgb(text_format.get('foregroundColor', {}), "rgb(0,0,0)")
                bold = "bold" if text_format.get('bold') else "normal"
                italic = "italic" if text_format.get('italic') else "normal"
                
                # Border Logic
                borders = eff_format.get('borders', {})
                b_styles = ""
                for side in ['top', 'bottom', 'left', 'right']:
                    if side in borders:
                        b_color = get_rgb(borders[side].get('color', {}))
                        b_styles += f"border-{side}: 2px solid {b_color};"
                    else:
                        b_styles += f"border-{side}: 1px solid #e0e0e0;"

                cell_style = (f"background-color:{bg}; color:{fg}; font-weight:{bold}; "
                              f"font-style:{italic}; font-size:{f_size}pt; padding:8px; {b_styles}")
                
                table_html += f'<td style="{cell_style}">{text if text else "&nbsp;"}</td>'
            table_html += '</tr>'
        
        return table_html + '</table>'
    except Exception as e:
        return f"Error generating HTML: {e}"

def safe_split(text, index, delimiter=None):
    try:
        return text.split(delimiter)[index]
    except (IndexError, AttributeError):
        return ""

def pull_scout_information(all_data):
    """Parses raw sheet data into a structured dictionary."""
    if not all_data: return {}
    headers = all_data[0]
    
    def get_idx(name):
        return headers.index(name) if name in headers else -1

    indices = {
        "name": get_idx('Full Name'), "email": get_idx('Email'),
        "sub": get_idx('Subject'), "body": get_idx('Body'),
        "send": get_idx('Send Email?'), "parent": get_idx('Parent Name'),
        "payment": get_idx('Payment Link'), "extra": get_idx('Additional Email Recipients')
    }

    scout_list = {}
    for row in all_data[1:]:
        # Basic validation
        if len(row) <= max(indices.values()):
            continue
            
        full_name = str(row[indices["name"]]).strip()
        send_flag = str(row[indices["send"]]).strip()

        if send_flag == "Yes":
            scout_list[full_name] = {
                "email": f"{row[indices['email']]}, {row[indices['extra']]}",
                "first_name": safe_split(full_name, 1),
                "last_name": safe_split(full_name, 0, ", "),
                "parent_first": safe_split(str(row[indices["parent"]]), 1),
                "subject": str(row[indices["sub"]]),
                "body": str(row[indices["body"]]),
                "payment": str(row[indices["payment"]])
            }
    return scout_list

def send_emails(scout_data, creds, worksheet):
    """Connects to SMTP and sends formatted HTML emails."""
    try:
        host = socket.gethostbyname("smtp.gmail.com")
    except:
        host = "smtp.gmail.com"

    with smtplib.SMTP_SSL(host, 465) as server:
        server.login(MY_EMAIL, MY_PASSWORD)
        
        for name, data in scout_data.items():
            formatted_table = get_cell_html(creds)
            
            # Dynamic String Formatting
            final_body = data["body"].format(
                scout_name=data["first_name"], 
                parent_name=data["parent_first"],
                payment_link=data["payment"]
            )

            msg = MIMEMultipart()
            msg["Subject"] = data["subject"]
            msg["From"] = MY_EMAIL
            msg["To"] = data["email"]
            
            html_content = f"<html><body><p>{final_body}</p>{formatted_table}</body></html>"
            msg.attach(MIMEText(html_content, "html"))
            
            server.send_message(msg)
            worksheet.update_acell('C2', name) # Logging progress
            print(f"Sent to {name}")

if __name__ == "__main__":
    # Standard authentication setup
    creds_path = 'creds.json'
    gc = gspread.service_account(filename=creds_path)
    auth_creds = service_account.Credentials.from_service_account_file(creds_path)
    
    sh = gc.open(SPREADSHEET_NAME)
    ws = sh.worksheet(WORKSHEET_NAME)
    
    raw_data = ws.get_all_values()
    processed_scouts = pull_scout_information(raw_data)
    
    send_emails(processed_scouts, auth_creds, ws)
