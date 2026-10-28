"""
Parse raw calendar entries like:
    Nova Avon IDES 333 "Revit for Designers" | REAL TIME 12:40pm-2:00pm
into structured columns. Reads calendar_raw.xlsx (from the pull scripts)
and writes calendar_clean.xlsx with Raw + Clean sheets.
 
Setup:  pip install pandas openpyxl
"""
 
import re
from datetime import datetime
 
import pandas as pd
 
INPUT_FILE = "calendar_raw.xlsx"
OUTPUT_FILE = "calendar_clean.xlsx"
 
# Rows missing any of these get flagged for review
REQUIRED = ["Course", "Class_Start"]
 
# Department code -> College/Dept name. Fill in the ones you use; unknown codes stay blank.
DEPT_LOOKUP = {
    "IDES": "Interior Design",  # example only: replace with your real college/dept names
}
 
# Each piece is parsed independently so partly-formatted entries still yield what they can.
COURSE_RE = re.compile(r"\b([A-Z]{2,5})\s?(\d{3}[A-Z]?)\b")
TITLE_RE = re.compile(r'["\u201c\u201d](.+?)["\u201c\u201d]')
MODE_RE = re.compile(r"\|\s*(.+?)\s*(?=\d{1,2}(?::\d{2})?\s*(?:[ap]m)?\s*[-\u2013]|$)", re.I)
TIME_RANGE_RE = re.compile(
    r"(\d{1,2}(?::\d{2})?)\s*([ap]m)?\s*[-\u2013]\s*(\d{1,2}(?::\d{2})?)\s*([ap]m)", re.I
)
 
def to_time(text, meridiem):
    """'12:40' + 'pm' -> datetime.time"""
    fmt = "%I:%M%p" if ":" in text else "%I%p"
    return datetime.strptime(f"{text}{meridiem}".lower(), fmt).time()
 
 
def parse_subject(subject):
    s = str(subject or "")
    out = {"Instructor": None, "Dept": None, "Course_Number": None, "Course": None,
           "Course_Title": None, "Session_Mode": None, "Class_Start": None, "Class_End": None}
 
    m = COURSE_RE.search(s)
    if m:
        out["Dept"], out["Course_Number"] = m.group(1), m.group(2)
        out["Course"] = f"{m.group(1)} {m.group(2)}"
        before = s[: m.start()].strip()
        out["Instructor"] = before or None  # text before the course code
 
    m = TITLE_RE.search(s)
    if m:
        out["Course_Title"] = m.group(1).strip()
 
    m = MODE_RE.search(s)
    if m and m.group(1):
        out["Session_Mode"] = m.group(1).strip()
 
    m = TIME_RANGE_RE.search(s)
    if m:
        t1, mer1, t2, mer2 = m.group(1), m.group(2) or m.group(4), m.group(3), m.group(4)
        try:
            start, end = to_time(t1, mer1), to_time(t2, mer2)
            if start > end and not m.group(2):  # e.g. "11-1pm" read as 11pm-1pm
                start = to_time(t1, "am" if mer1.lower() == "pm" else "pm")
            out["Class_Start"], out["Class_End"] = start, end
        except ValueError:
            pass
    return out
 
 
def minutes_between(date_part, t_later, t_earlier):
    if t_later is None or t_earlier is None:
        return None
    a = datetime.combine(date_part, t_later)
    b = datetime.combine(date_part, t_earlier)
    return (a - b).total_seconds() / 60
 
 
raw = pd.read_excel(INPUT_FILE)
for optional in ["Location", "Organizer", "Categories", "Body"]:
    if optional not in raw.columns:  # some exports omit these
        raw[optional] = ""
raw["Start"] = pd.to_datetime(raw["Start"])
raw["End"] = pd.to_datetime(raw["End"])
 
parsed = pd.DataFrame([parse_subject(s) for s in raw["Subject"]], index=raw.index)
df = pd.concat([raw, parsed], axis=1)
 
df["Date"] = df["Start"].dt.date
df["Weekday"] = df["Start"].dt.strftime("%a")
df["Scheduled_Start"] = df["Start"].dt.time
df["Scheduled_End"] = df["End"].dt.time
df["Notes"] = df["Body"].fillna("").astype(str).str.strip()
 
# Room booked vs. actual class time (e.g. 12:00 booked, class starts 12:40 = 40 min setup)
df["Setup_Minutes"] = [
    minutes_between(d, cs, ss) for d, cs, ss in zip(df["Date"], df["Class_Start"], df["Scheduled_Start"])
]
df["Class_Minutes"] = [
    minutes_between(d, ce, cs) for d, ce, cs in zip(df["Date"], df["Class_End"], df["Class_Start"])
]
 
df = df.drop_duplicates(subset=["Subject", "Start"])
df["Needs_Review"] = df[REQUIRED].isna().any(axis=1).map({True: "Review", False: "OK"})
 
cols = ["Date", "Weekday", "Scheduled_Start", "Scheduled_End", "Class_Start", "Class_End",
        "Setup_Minutes", "Class_Minutes", "Instructor", "Dept", "Course_Number", "Course",
        "Course_Title", "Session_Mode", "Notes", "Location",
        "Organizer", "Categories", "Needs_Review", "Subject"]
clean = df[cols].sort_values(["Date", "Scheduled_Start"])
 
# Sheet laid out with the same headings as your tracking spreadsheet
event_name = (clean["Course"].fillna("") + " " + clean["Course_Title"].fillna("")).str.strip()
dept_code = clean["Dept"]
tracker = pd.DataFrame({
    "Date": clean["Date"],
    "Requestor Email/Name": clean["Instructor"].fillna(df.loc[clean.index, "Organizer"]),
    "Class/Event Name": event_name.where(event_name != "", clean["Subject"]),
    "College/Dept": dept_code.map(DEPT_LOOKUP),
    "College/Department Codes": dept_code,
    "Notes": clean["Notes"],
})
 
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    tracker.to_excel(writer, sheet_name="Tracker", index=False)
    clean.to_excel(writer, sheet_name="Clean", index=False)
    raw.to_excel(writer, sheet_name="Raw", index=False)
 
print(f"{len(clean)} rows, {(clean['Needs_Review'] == 'Review').sum()} flagged for review")
print(f"Saved to {OUTPUT_FILE}")
