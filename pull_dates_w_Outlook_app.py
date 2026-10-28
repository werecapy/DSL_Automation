"""
Pull raw Outlook calendar events into Excel. No cleaning.
 
Setup:  pip install pywin32 pandas openpyxl
"""
 
from datetime import datetime
 
import pandas as pd
import win32com.client
 
START = datetime(2026, 1, 1)
END = datetime(2026, 12, 31)
OUTPUT_FILE = "calendar_raw.xlsx"
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
items = outlook.GetDefaultFolder(9).Items  # 9 = Calendar
items.IncludeRecurrences = True  # set before Sort/Restrict
items.Sort("[Start]")
 
fmt = "%m/%d/%Y %I:%M %p"
items = items.Restrict(f"[Start] >= '{START.strftime(fmt)}' AND [End] <= '{END.strftime(fmt)}'")
 
rows = []
for item in items:
    try:
        rows.append({
            "Subject": item.Subject,
            "Start": item.Start.replace(tzinfo=None),
            "End": item.End.replace(tzinfo=None),
            "Location": item.Location,
            "Organizer": item.Organizer,
            "Categories": item.Categories,
            "AllDay": item.AllDayEvent,
            "Recurring": item.IsRecurring,
            "Body": item.Body,
        })
    except Exception as e:
        print(f"Skipped an item: {e}")
 
pd.DataFrame(rows).to_excel(OUTPUT_FILE, index=False)
print(f"Saved {len(rows)} events to {OUTPUT_FILE}")
