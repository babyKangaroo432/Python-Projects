pip install google-api-python-client

from googleapiclient.discovery import build
from google.oauth2 import service_account

def get_cell_html():
    try:
        creds = service_account.Credentials.from_service_account_file('creds.json')
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

        # 1. FIND THE MAX NUMBER OF COLUMNS
        max_cols = 0
        for row in rows:
            max_cols = max(max_cols, len(row.get('values', [])))

        # 2. BUILD THE TABLE
        for row in rows:
            table_html += '<tr>'
            cells = row.get('values', [])

            for i in range(max_cols):
                cell = cells[i] if i < len(cells) else {}

                text = cell.get('formattedValue', '')
                eff_format = cell.get('effectiveFormat', {})
                text_format = eff_format.get('textFormat', {})

                # --- FONT SIZE LOGIC ---
                # Default to 10pt if not specified
                f_size = text_format.get('fontSize', 10)

                bg = get_rgb(eff_format.get('backgroundColor', {}), "rgb(255,255,255)")
                fg = get_rgb(text_format.get('foregroundColor', {}), "rgb(0,0,0)")
                bold = "bold" if text_format.get('bold') else "normal"
                italic = "italic" if text_format.get('italic') else "normal"

                borders = eff_format.get('borders', {})
                b_styles = ""
                for side in ['top', 'bottom', 'left', 'right']:
                    if side in borders:
                        b_info = borders[side]
                        b_color = get_rgb(b_info.get('color', {}), "rgb(0,0,0)")
                        b_styles += f"border-{side}: 2px solid {b_color};"
                    else:
                        b_styles += f"border-{side}: 1px solid #e0e0e0;"

                cell_style = (
                    f"background-color:{bg}; color:{fg}; font-weight:{bold}; "
                    f"font-style:{italic}; font-size:{f_size}pt; " # Added font-size
                    f"padding:8px 12px; min-width:80px; {b_styles}"
                )

                display_text = text if text else "&nbsp;"
                table_html += f'<td style="{cell_style}">{display_text}</td>'

            table_html += '</tr>'

        table_html += '</table>'
        return table_html

    except Exception as e:
        print(f"DEBUG ERROR: {e}")
        return "Error retrieving data."


def read_whole_sheet():
  gc = gspread.service_account(filename='creds.json')
  sh = gc.open(SPREADSHEET_NAME)
  worksheet = sh.worksheet(WORKSHEET_NAME)

  # Get all values from the worksheet
  all_data = worksheet.get_all_values()

  return all_data


def get_header_location(headers_row, header_name):
  if not headers_row:
      return -1 # Return -1 if headers row is empty

  # Find the column indices for 'header_name'
  for i, header_val in enumerate(headers_row):
      if header_val == header_name:
          return i

  print(f"Warning: '{header_name}' column not found in headers. Please check sheet headers.")
  return -1 # Or handle this case as appropriate


def pull_scout_information(all_data):
    if not all_data:
        return {}

    headers = all_data[0]
    data_rows = all_data[1:]

    full_name_col_idx = get_header_location(headers, 'Full Name')
    email_col_idx = get_header_location(headers, 'Email')
    subject_col_idx = get_header_location(headers, 'Subject')
    body_col_idx = get_header_location(headers, 'Body')
    send_email_col_idx = get_header_location(headers, 'Send Email?')
    parent_name_col_idx = get_header_location(headers, 'Parent Name')
    aditional_email_recipients_col_idx = get_header_location(headers, 'Additional Email Recipients')
    payment_link_col_idx = get_header_location(headers, 'Payment Link')

    if full_name_col_idx == -1 or email_col_idx == -1:
        return {} # Return empty dict if essential headers are missing

    scout_list = {}

    for row in data_rows:
        full_name = ''
        email = ''
        subject = ''
        body = ''
        send_email = ''
        parent_name = ''
        additional_email_recipients = ''
        payment_link = ''

        # Safely access name and email using their indices
        if full_name_col_idx < len(row):
            full_name = str(row[full_name_col_idx]).strip()
        if email_col_idx < len(row):
            email = str(row[email_col_idx]).strip()
        if subject_col_idx < len(row):
            subject = str(row[subject_col_idx]).strip()
        if body_col_idx < len(row):
            body = str(row[body_col_idx]).strip()
        if send_email_col_idx < len(row):
            send_email = str(row[send_email_col_idx]).strip()
        if parent_name_col_idx < len(row):
            parent_name = str(row[parent_name_col_idx]).strip()
        if aditional_email_recipients_col_idx < len(row):
            additional_email_recipients = str(row[aditional_email_recipients_col_idx]).strip()
        if payment_link_col_idx < len(row):
            payment_link = str(row[payment_link_col_idx]).strip()

        # Contingency: Check if either is empty
        unfilled_catagories = []
        if (not full_name or not email or not subject or not body or not parent_name or ", " not in full_name or ", " not in parent_name or not payment_link) and send_email == "Yes":
            if not full_name or ", " not in full_name: unfilled_catagories.append("Full Name")
            if not email: unfilled_catagories.append("Email")
            if not subject: unfilled_catagories.append("Subject")
            if not body: unfilled_catagories.append("Body")
            if not parent_name or ", " not in parent_name: unfilled_catagories.append("Parent Name")
            if not payment_link: unfilled_catagories.append("Payment Link")

            print(f"Skipping email recipiant '{full_name}' to '{email}'\nERROR:\n    MISSING COLUMN INFORMATION: {unfilled_catagories}\n\n")
            #print(f"Skipping email recipiant '{full_name}' to '{email}':\nSend Email?='{send_email}', Full Name='{full_name}', Email='{email}', Subject='{subject}', Body='{body}', Parent Name={parent_name}, Additional Email Recipients={additional_email_recipients}\n\n")
            continue

        scout_last_name = safe_split(full_name, 0, ", ")
        scout_first_name = safe_split(full_name, 1)  # Default delimiter is space
        parent_first_name = safe_split(parent_name, 1)
        scout_list[full_name] = [send_email, email+", "+additional_email_recipients, scout_first_name, scout_last_name, parent_first_name, subject, body, payment_link]

    return scout_list


def safe_split(text, index, delimiter=None):
    try:
        return text.split(delimiter)[index]
    except (IndexError, AttributeError):
        return ""


def send_emails(scout_data, MY_EMAIL, MY_PASSWORD, worksheet):
  # --- 1. GET IP ADDRESS ---
  try:
     ip_address = socket.gethostbyname("smtp.gmail.com")
     print(f"Server IP: {ip_address}")
  except Exception as e:
      ip_address = "smtp.gmail.com"

  # Loop for every scout
  for full_name in scout_data:
    # set each var
    scout = scout_data[full_name]
    send_email = scout[0]
    email = scout[1]
    first_name = scout[2]
    last_name = scout[3]
    parent_name = scout[4]
    subject = scout[5]
    body = scout[6]
    payment_link = scout[7]

    # Replace {name} placeholder with the scout's first name
    body = body.format(scout_name=first_name, email=email, parent_name=parent_name, last_name=last_name, payment_link=payment_link)
    subject = subject.format(scout_name=first_name, email=email, parent_name=parent_name, last_name=last_name, payment_link=payment_link)

    worksheet.update_acell('C2', full_name)
    if send_email == "Yes":
      formatted_data = get_cell_html()

      msg = MIMEMultipart()
      msg["Subject"] = subject
      msg["From"] = MY_EMAIL
      msg["To"] = email

      html_body = f"""
      <html>
        <body>
          <p>{body}</p>
          <div style="background-color: #f9f9f9; padding: 15px; border: 1px solid #ddd; font-family: sans-serif;">
            {formatted_data}
          </div>
        </body>
      </html>
      """

      msg.attach(MIMEText(html_body, "html"))

      try:
          print("Connecting and sending...")
          with smtplib.SMTP_SSL(ip_address, 465) as server:
              server.login(MY_EMAIL, MY_PASSWORD)
              server.send_message(msg)
          print("Email sent successfully with formatting!")
      except Exception as e:
          print(f"Failed to send email. Error: {e}")


# Main
# Add your creds.json from 2 factor auth

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
import smtplib
import socket

MY_EMAIL = "" # myemail.gmail.com
MY_PASSWORD = "" # Google App Password
WORKSHEET_NAME = ""

# Define SPREADSHEET_ID and RANGE_NAME globally for get_cell_html
SPREADSHEET_ID = "" # After the "https://docs.google.com/spreadsheets/d/"
SPREADSHEET_NAME = "2026 Scout Payments" # Name of whole spreadsheet
RANGE_NAME = "" # Invoice range ie. "Invoices!B2:D"

# Initialize gspread objects once
gc = gspread.service_account(filename="creds.json")
sh = gc.open(SPREADSHEET_NAME)
worksheet = sh.worksheet(WORKSHEET_NAME)

all_data = read_whole_sheet()
scout_data = pull_scout_information(all_data)
send_emails(scout_data, MY_EMAIL, MY_PASSWORD, worksheet)
