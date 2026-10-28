
"""
Pull Outlook calendar events (desktop Outlook, Windows) and clean them into Excel.
 
Setup:
    pip install pywin32 pandas openpyxl
 
Output workbook has two sheets:
    Raw   - exactly what Outlook returned (never modified)
    Clean - parsed columns + a Needs_Review flag for incomplete rows
"""
 
import re
from datetime import datetime
 
import pandas as pd
import win32com.client
 
# ----------------------------- SETTINGS ------------------------------------
START = datetime(2026, 1, 1)
END = datetime(2026, 12, 31)
OUTPUT_FILE = "calendar_export.xlsx"
 
# Known clients/projects, used as a fallback when the subject has no clear format.
KNOWN_CLIENTS = ["Smith", "Acme", "Jones"]
 
# Variants -> standard label (add rows as you find new spellings).
TYPE_MAP = {
    "mtg": "Meeting",
    "meeting": "Meeting",
    "site visit": "Site Visit",
    "call": "Call",
}
 
REQUIRED_FIELDS = ["Client", "Type"]  # rows missing these get flagged
# ---------------------------------------------------------------------------
 
 
def pull_events(start, end):
    """Return a DataFrame of raw calendar events between start and end."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)  # 9 = olFolderCalendar
 
    items = calendar.Items
    items.IncludeRecurrences = True  # must be set BEFORE Sort/Restrict
    items.Sort("[Start]")
 
    fmt = "%m/%d/%Y %I:%M %p"
    restriction = f"[Start] >= '{start.strftime(fmt)}' AND [End] <= '{end.strftime(fmt)}'"
    items = items.Restrict(restriction)
 
    rows = []
    for item in items:
        try:
            rows.append(
                {
                    "Subject": item.Subject,
                    "Start": item.Start.replace(tzinfo=None),
                    "End": item.End.replace(tzinfo=None),
                    "Location": item.Location,
                    "Organizer": item.Organizer,
                    "Categories": item.Categories,
                    "AllDay": item.AllDayEvent,
                    "Recurring": item.IsRecurring,
                    "Body": (item.Body or "")[:2000],  # truncate long bodies
                }
            )
        except Exception as e:  # some items (e.g. private/corrupt) can fail
            print(f"Skipped an item: {e}")
    return pd.DataFrame(rows)
 
 
def extract_client(row):
    """Try several sources in order; return the first hit, else None."""
    subject = str(row["Subject"] or "")
    body = str(row["Body"] or "")
 
    # 1. "Client - Task" in the subject
    if " - " in subject:
        return subject.split(" - ", 1)[0].strip()
    # 2. Known client name anywhere in the subject
    for name in KNOWN_CLIENTS:
        if re.search(rf"\b{re.escape(name)}\b", subject, flags=re.I):
            return name
    # 3. "Client: X" line in the body
    m = re.search(r"Client:\s*(.+)", body, flags=re.I)
    if m:
        return m.group(1).strip()
    return None
 
 
def extract_type(subject):
    """Map keywords found in the subject to a standard type."""
    s = str(subject or "").lower()
    for key, standard in TYPE_MAP.items():
        if key in s:
            return standard
    return None
 
 
def clean(raw):
    df = raw.copy()
 
    # Normalize text
    for col in ["Subject", "Location", "Organizer", "Categories"]:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["Subject"] = df["Subject"].str.replace(r"\s+", " ", regex=True)
 
    # Parsed fields (fallback rules above)
    df["Client"] = df.apply(extract_client, axis=1)
    df["Type"] = df["Subject"].apply(extract_type)
 
    # Dates and duration
    df["Date"] = df["Start"].dt.date
    df["Duration_Hours"] = (df["End"] - df["Start"]).dt.total_seconds() / 3600
    df.loc[df["AllDay"], "Duration_Hours"] = None  # don't let all-day skew totals
 
    # Missing location -> explicit placeholder
    df["Location"] = df["Location"].replace("", "Unspecified")
 
    # Remove duplicates (same subject + start)
    df = df.drop_duplicates(subset=["Subject", "Start"])
 
    # Flag rows that need a human look
    missing = df[REQUIRED_FIELDS].isna().any(axis=1)
    df["Needs_Review"] = missing.map({True: "Review", False: "OK"})
 
    cols = [
        "Date", "Start", "End", "Duration_Hours", "Client", "Type",
        "Location", "Organizer", "Categories", "AllDay", "Recurring",
        "Needs_Review", "Subject",
    ]
    return df[cols].sort_values("Start")
 
 
if __name__ == "__main__":
    raw = pull_events(START, END)
    print(f"Pulled {len(raw)} events")
 
    cleaned = clean(raw)
    print(f"{(cleaned['Needs_Review'] == 'Review').sum()} rows flagged for review")
 
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="Raw", index=False)
        cleaned.to_excel(writer, sheet_name="Clean", index=False)
    print(f"Saved to {OUTPUT_FILE}")
