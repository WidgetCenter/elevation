import os
import smtplib
from datetime import datetime, timedelta, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import requests

BASE = "https://cwms-data.usace.army.mil/cwms-data"
HEADERS = {"Accept": "application/json"}
OFFICE = "SWT"
OUTPUT_DIR = "output"

# Keep this False while testing locally
SEND_EMAIL = False

# Reservoir/lake projects - report pool Elevation
ELEVATION_PROJECTS = [
    "ALTU", "ARBU", "ARCA", "BIGH", "BIRC", "BROK", "CANT", "CHEN",
    "COPA", "COUN", "DENI", "ELDR", "ELKC", "EUFA", "FALL",
    "FCOB", "FGIB", "FOSS", "FSUP", "GSAL", "HEYB", "HUDS", "HUGO",
    "HULA", "JOHN", "KAWL", "KEMP", "KEYS", "MARI", "MCGE", "MERE",
    "OOLO", "PATM", "PENS", "PINE", "SARD", "SKIA", "TENK", "THUN",
    "TOMS", "TORO", "TRUS", "WAUR", "WIST"
]

# Levee projects - report river Stage instead of pool Elevation
STAGE_PROJECTS = [
    "CUMB",
]


def get_latest_value(ts_id):
    """Pull the most recent reading for a single timeseries ID."""
    end = datetime.now(timezone.utc)
    begin = end - timedelta(hours=6)

    params = {
        "office": OFFICE,
        "name": ts_id,
        "begin": begin.isoformat(),
        "end": end.isoformat(),
        "page-size": 1
    }

    try:
        r = requests.get(
            f"{BASE}/timeseries",
            params=params,
            headers=HEADERS,
            timeout=30
        )
        r.raise_for_status()

        values = r.json().get("values", [])
        if not values:
            return {
                "timestamp": None,
                "value": None,
                "status": r.status_code,
                "error": "No data returned"
            }

        timestamp_ms, value = values[-1][0], values[-1][1]
        timestamp_utc = datetime.fromtimestamp(timestamp_ms / 1000, tz=timezone.utc)
        timestamp_naive = timestamp_utc.replace(tzinfo=None)  # drop timezone for Excel

        return {
            "timestamp": timestamp_naive,
            "value": value,
            "status": r.status_code,
            "error": None
        }

    except requests.RequestException as e:
        status_code = None
        if getattr(e, "response", None) is not None:
            status_code = e.response.status_code

        return {
            "timestamp": None,
            "value": None,
            "status": status_code,
            "error": str(e)
        }


def build_report():
    """Build the SWT elevation/stage Excel report."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    rows = []

    for code in ELEVATION_PROJECTS:
        ts_id = f"{code}.Elev.Inst.1Hour.0.Ccp-Rev"
        result = get_latest_value(ts_id)

        rows.append({
            "Project": code,
            "Parameter": "Elevation (ft)",
            "Value": result["value"],
            "ReadingTime_UTC": result["timestamp"],
            "TS_ID": ts_id,
            "HTTP_Status": result["status"],
            "Error": result["error"]
        })

    for code in STAGE_PROJECTS:
        ts_id = f"{code}.Stage.Inst.1Hour.0.Ccp-Rev"
        result = get_latest_value(ts_id)

        rows.append({
            "Project": code,
            "Parameter": "Stage (ft)",
            "Value": result["value"],
            "ReadingTime_UTC": result["timestamp"],
            "TS_ID": ts_id,
            "HTTP_Status": result["status"],
            "Error": result["error"]
        })

    df = pd.DataFrame(rows).sort_values(["Project", "Parameter"])

    filename = os.path.join(
        OUTPUT_DIR,
        f"swt_elevations_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )

    with pd.ExcelWriter(
        filename,
        engine="openpyxl",
        datetime_format="yyyy-mm-dd hh:mm:ss"
    ) as writer:
        df.to_excel(writer, index=False, sheet_name="SWT Report")

        worksheet = writer.sheets["SWT Report"]
        worksheet.column_dimensions["A"].width = 12
        worksheet.column_dimensions["B"].width = 18
        worksheet.column_dimensions["C"].width = 14
        worksheet.column_dimensions["D"].width = 24
        worksheet.column_dimensions["E"].width = 35
        worksheet.column_dimensions["F"].width = 12
        worksheet.column_dimensions["G"].width = 40

    return filename


def send_email(filename):
    """Send the report by email using SMTP environment variables."""
    msg = MIMEMultipart()
    msg["Subject"] = f"SWT Lakes Daily Elevation Report - {datetime.now().strftime('%Y-%m-%d')}"
    msg["From"] = os.environ["SMTP_USERNAME"]
    msg["To"] = os.environ["RECIPIENT_EMAIL"]

    body = "Attached: current elevation/stage readings for Tulsa District (SWT) projects."
    msg.attach(MIMEText(body, "plain"))

    with open(filename, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f'attachment; filename="{os.path.basename(filename)}"'
    )
    msg.attach(part)

    with smtplib.SMTP(os.environ["SMTP_SERVER"], int(os.environ["SMTP_PORT"])) as server:
        server.starttls()
        server.login(os.environ["SMTP_USERNAME"], os.environ["SMTP_PASSWORD"])
        server.send_message(msg)


def check_projects(codes, parameter):
    """Optional helper to test a smaller list of project codes."""
    found, missing = [], []

    for code in codes:
        ts_id = f"{code}.{parameter}.Inst.1Hour.0.Ccp-Rev"
        result = get_latest_value(ts_id)

        if result["value"] is not None:
            print(f"{code:8s} ({parameter:5s}) {result['value']:>10.2f} @ {result['timestamp']}")
            found.append(code)
        else:
            print(f"{code:8s} ({parameter:5s}) NO DATA (status {result['status']})")
            missing.append(code)

    return found, missing


if __name__ == "__main__":
    filename = build_report()
    print(f"Report built: {filename}")

    if SEND_EMAIL:
        #send_email(filename)
        print("Email sent.")
    else:
        print("Email step skipped for testing.")
