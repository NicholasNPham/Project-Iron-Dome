# STANDARD LIBRARY IMPORTS
from collections import Counter
import re
from datetime import datetime

# THIRD PARTY IMPORTS
import win32com.client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# CONSTANTS
SUBJECT = "[External] Message meets Alert condition"
OUTPUT_FILE = "PARSED_INTRUSION_EMAILS.xlsx"

# INPUTS
DATE_START = datetime.strptime(input("Start date (YYYY-MM-DD): "), "%Y-%m-%d")
DATE_END   = datetime.strptime(input("End date (YYYY-MM-DD): "), "%Y-%m-%d")

# FUNCTION
def get_emails(subject: str, start: datetime, end: datetime) -> list[dict]:
    """
    Connect to the Alerts shared mailbox in Outlook and retrieve emails
    filtered by exact subject and date range. Parses email body fields
    using regex.

    Args:
        subject (str): Exact subject line to match against.
        start (datetime): Start of the date range (inclusive).
        end (datetime): End of the date range (inclusive).

    Returns:
        list[dict]: Each dict contains 'date', 'srcip', 'srccountry',
                    'proto', 'service', and 'dstip'.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    inbox = None

    for store in namespace.Stores:
        if store.DisplayName == "Alerts":
            inbox = store.GetDefaultFolder(6)
            break

    if not inbox:
        raise ValueError("Alerts mailbox not found.")

    filter_str = (f"[Subject] = '{subject}' AND "
                  f"[ReceivedTime] >= '{start.strftime('%m/%d/%Y')}' AND "
                  f"[ReceivedTime] <= '{end.strftime('%m/%d/%Y')}'")

    all_email_items = inbox.Items
    all_email_items.Sort("[ReceivedTime]", True)
    filtered_email_items = all_email_items.Restrict(filter_str)

    results_list = []

    for email in filtered_email_items:
        body = email.Body

        # regex
        ip = re.search(r'srcip=([\d.]+)', body)
        country = re.search(r'srccountry="([^"]+)"', body)
        proto = re.search(r'proto=(\d+)', body)
        service = re.search(r'service="([^"]+)"', body)
        destip = re.search(r'dstip=([\d.]+)', body)

        # errors
        if ip:
            srcip = ip.group(1)
        else:
            srcip = "NOT FOUND"

        if country:
            srccountry = country.group(1)
        else:
            srccountry = "NOT FOUND"

        if proto:
            srcproto = proto.group(1)
        else:
            srcproto = "NOT FOUND"

        if service:
            srcservice = service.group(1)
        else:
            srcservice = "NOT FOUND"

        if destip:
            dstip = destip.group(1)
        else:
            dstip = "NOT FOUND"

        # dict creation
        results_list.append({
            "date"      : email.ReceivedTime.strftime("%Y-%m-%d %H:%M"),
            "srcip"     : srcip,
            "srccountry": srccountry,
            "proto" : srcproto,
            "service" : srcservice,
            "dstip" : dstip,})

    return results_list

def write_excel(result_list: list[dict], output_file: str):

# Create a workbook and sheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = "PARSED EMAILS"
# Write styled headers
headers = ["DATE", "SRCIP", "SRCCOUNTRY", "PROTO", "SERVICE", "DSTIP"]
header_font = Font(name="Times New Roman", size = 14, bold=True, italic=False, color="FFFFFF")
header_fill = PatternFill(fill_type="solid", fgColor="2F4F7F")
header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True,)

for column, header in enumerate(headers, 1):
    cell = worksheet.cell(row=1, column=column, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment
# Sort data by IP frequency, NOT FOUND at the bottom

# Write each row
# Set column widths
# Save the file
