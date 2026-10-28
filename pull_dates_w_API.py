"""
Pull raw calendar events from Outlook on the web (Microsoft 365) via Microsoft Graph.
 
Setup:  pip install msal requests pandas openpyxl
 
Two ways to authenticate (pick one):
  A) ACCESS_TOKEN: paste a token from Graph Explorer (https://developer.microsoft.com/graph/graph-explorer),
     sign in, consent to Calendars.Read, then copy it from the "Access token" tab.
     Quickest, no app registration, but it expires in about an hour.
  B) CLIENT_ID / TENANT_ID: from your own Azure app registration (see notes).
     Leave ACCESS_TOKEN empty to use this.
"""
 
import msal
import pandas as pd
import requests
 
# ----------------------------- SETTINGS ------------------------------------
START = "2026-01-01T00:00:00"
END = "2026-12-31T23:59:59"
TIMEZONE = "Eastern Standard Time"  # Windows time zone name
OUTPUT_FILE = "calendar_raw.xlsx"
 
ACCESS_TOKEN = ""  # option A
CLIENT_ID = ""     # option B: Application (client) ID
TENANT_ID = ""     # option B: Directory (tenant) ID, or "organizations"
# ---------------------------------------------------------------------------
 
 
def get_token():
    if ACCESS_TOKEN:
        return ACCESS_TOKEN
    app = msal.PublicClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )
    flow = app.initiate_device_flow(scopes=["Calendars.Read"])
    print(flow["message"])  # follow the on-screen instructions to sign in
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", result))
    return result["access_token"]
 
 
def pull_events(token):
    headers = {
        "Authorization": f"Bearer {token}",
        "Prefer": f'outlook.timezone="{TIMEZONE}", outlook.body-content-type="text"',
    }
    url = "https://graph.microsoft.com/v1.0/me/calendarView"
    params = {
        "startDateTime": START,
        "endDateTime": END,
        "$top": 100,
        "$orderby": "start/dateTime",
        "$select": "subject,start,end,location,organizer,categories,isAllDay,type,body",
    }
 
    rows = []
    while url:
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        for e in data["value"]:
            rows.append({
                "Subject": e.get("subject"),
                "Start": e["start"]["dateTime"],
                "End": e["end"]["dateTime"],
                "Location": (e.get("location") or {}).get("displayName"),
                "Organizer": ((e.get("organizer") or {}).get("emailAddress") or {}).get("name"),
                "Categories": ", ".join(e.get("categories", [])),
                "AllDay": e.get("isAllDay"),
                "Type": e.get("type"),  # singleInstance, occurrence, exception, seriesMaster
                "Body": (e.get("body") or {}).get("content"),
            })
        url = data.get("@odata.nextLink")  # follow paging
        params = None  # nextLink already contains the query
    return pd.DataFrame(rows)
 
 
if __name__ == "__main__":
    df = pull_events(get_token())
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(df)} events to {OUTPUT_FILE}")
