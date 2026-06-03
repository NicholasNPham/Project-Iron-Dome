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
    pass