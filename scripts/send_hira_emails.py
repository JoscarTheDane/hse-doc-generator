# AgentMail key: set AGENTMAIL_API_KEY env var or place key file at $HSE_HOME/.api_key
#!/usr/bin/env python3
"""
Send HIRA reply emails with workbooks attached.
"""

import json
import os
import urllib.request
import urllib.error

API_KEY = "${AGENTMAIL_API_KEY}"
BASE_URL = "https://api.agentmail.to/v0"
HSE_DIR = "/home/joshua/HSE"

def api_request(endpoint, method="GET"):
    url = f"{BASE_URL}{endpoint}"
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"Bearer {API_KEY}")
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        print(f"  HTTP Error {e.code}: {e.read().decode()[:200]}")
        return None

today = "2026-07-25"

# Project info
projects = [
    {
        "thread_id": "19d38c45-d362-41ce-8b76-260a511ce009",
        "msg_id": "<001b01dd1b60$205c8e90$6115abb0$@gmail.com>",
        "recipient": "jsh.evan@gmail.com",
        "project": "Snake Island Container Terminal Development - Wreck & Debris Removal",
        "client": "Novadeal EKO FZE (DEME Group)",
        "vessel_worksite": "Snake Island, Lagos, Nigeria",
        "folder_name": f"{today}-Snake Island",
        "hira_file": f"{today}-Snake Island_HIRA.xlsx",
    },
    {
        "thread_id": "64660819-2c2c-477c-8ef4-ad8f2fec0716",
        "msg_id": "<000001dd1b5f$ea3cf750$beb6e5f0$@gmail.com>",
        "recipient": "jsh.evan@gmail.com",
        "project": "Bonga IWS 2023 — Mooring Procedure",
        "client": "SNEPCo (Shell Nigeria Exploration & Production Company)",
        "vessel_worksite": "Bonga FPSO / SPM Buoy, Niger Delta, Nigeria",
        "folder_name": f"{today}-SNEPCo Bonga",
        "hira_file": f"{today}-SNEPCo Bonga_HIRA.xlsx",
    },
    {
        "thread_id": "5ad6dd11-c4ce-458e-8da6-6aff9e5ceea7",
        "msg_id": "<018d01dd1b4f$73c4b320$5b4e1960$@gmail.com>",
        "recipient": "jsh.evan@gmail.com",
        "project": "Bonga Pipelines Flexjoints — 3D Photogrammetry Inspection",
        "client": "SNEPCo (Shell Nigeria Exploration & Production Company)",
        "vessel_worksite": "Bonga Field, Niger Delta — 1200m water depth",
        "folder_name": f"{today}-3D Photogrammetry",
        "hira_file": f"{today}-3D Photogrammetry_HIRA.xlsx",
    },
    {
        "thread_id": "75074bf3-0b34-4bb2-b662-85206db649c0",
        "msg_id": "<017501dd1af0$98e683e0$cab38ba0$@gmail.com>",
        "recipient": "jsh.evan@gmail.com",
        "project": "Single Anchor Loading (SAL) Buoy — Riser Flange Disconnection & Reconnection",
        "client": "Atlantic Marine & Oilfield Services Ltd",
        "vessel_worksite": "SAL Buoy Riser, Bonga Field, Offshore Nigeria",
        "folder_name": f"{today}-SAL Flange",
        "hira_file": f"{today}-SAL Flange_HIRA.xlsx",
    },
]

for proj in projects:
    print(f"\n{'='*60}")
    print(f"Processing: {proj['project']}")
    
    # Get thread to find attachment download URL
    thread_data = api_request(f"/threads/{proj['thread_id']}")
    if not thread_data:
        print(f"  [ERROR: could not fetch thread]")
        continue
    
    messages = thread_data.get("messages", [])
    for msg in messages:
        from_raw = msg.get("from", "")
        if isinstance(from_raw, dict):
            from_addr = from_raw.get("address", "")
        else:
            from_addr = str(from_raw) if from_raw else ""
        
        if "consultingsubsea" in from_addr:
            continue
        
        attachments = msg.get("attachments", [])
        download_url = None
        
        for att in attachments:
            att_id = att.get("attachment_id", "")
            print(f"  Checking attachment: {att.get('filename', 'unknown')} (ID: {att_id})")
            
            # Get attachment details
            att_details = api_request(f"/threads/{proj['thread_id']}/attachments/{att_id}")
            if att_details:
                url = att_details.get("download_url", "")
                if url:
                    download_url = url
                    print(f"    Download URL: {url[:80]}...")
        
        if download_url:
            print(f"  Has download URL: YES")
        else:
            print(f"  Has download URL: NO")
        
        # Build email
        sender_name = "Sender"
        email_body = f"""Dear {sender_name},

Thank you for entrusting ConsultingSubsea with your HIRA documentation for the {proj['project']} ({proj['client']}).

Please find attached the completed Hazard Identification & Risk Assessment (HIRA) workbook, prepared with reference to your submitted scope of work for {proj['vessel_worksite']}.

The risk assessment has been developed in accordance with industry-standard diving, marine, and subsea operational guidelines.

The workbook includes:
1. Main HIRA Risk Assessment Matrix (12-column layout)
2. Risk Assessment Matrix (RAM) reference sheet
3. Attendance sign-in sheet

If you would like to know more about this assessment or any of our other services — including project planning, HSE documentation, diving operations, and offshore project management — contact us at technical@consultingsubsea.com.

Kind regards,
HSE Documentation Team
ConsultingSubsea"""

        email_html = f"""<html><body>
<p>Dear {sender_name},</p>
<p>Thank you for entrusting ConsultingSubsea with your HIRA documentation for the <strong>{proj['project']}</strong> ({proj['client']}).</p>
<p>Please find attached the completed Hazard Identification &amp; Risk Assessment (HIRA) workbook, prepared with reference to your submitted scope of work for {proj['vessel_worksite']}.</p>
<p>The risk assessment has been developed in accordance with industry-standard diving, marine, and subsea operational guidelines.</p>
<p>The workbook includes:<br/>
1. Main HIRA Risk Assessment Matrix (12-column layout)<br/>
2. Risk Assessment Matrix (RAM) reference sheet<br/>
3. Attendance sign-in sheet</p>
<p>If you would like to know more about this assessment or any of our other services — including project planning, HSE documentation, diving operations, and offshore project management — contact us at <a href="mailto:technical@consultingsubsea.com">technical@consultingsubsea.com</a>.</p>
<p>Kind regards,<br/>
<strong>HSE Documentation Team</strong><br/>
<strong>ConsultingSubsea</strong></p>
</body></html>"""

        # Prepare send request
        send_url = f"{BASE_URL}/inboxes/consultingsubsea@agentmail.to/messages/send"
        
        payload = {
            "to": proj["recipient"],
            "cc": "technical@consultingsubsea.com",
            "subject": f"HIRA — {proj['project']} — ConsultingSubsea",
            "text": email_body,
            "html": email_html,
            "attachment_urls": [download_url] if download_url else [],
        }
        
        print(f"  Sending to: {proj['recipient']}")
        print(f"  CC: technical@consultingsubsea.com")
        print(f"  Subject: {payload['subject']}")
        print(f"  Attachment URL: {download_url}")
        
        # Send the email
        try:
            data = json.dumps(payload).encode('utf-8')
            req = urllib.request.Request(send_url, data=data, method="POST")
            req.add_header("Authorization", f"Bearer {API_KEY}")
            req.add_header("Content-Type", "application/json")
            
            with urllib.request.urlopen(req, timeout=30) as resp:
                result = json.loads(resp.read().decode())
                print(f"  Email sent successfully!")
                print(f"  Response: {json.dumps(result, indent=2)[:200]}")
        except urllib.error.HTTPError as e:
            error_body = e.read().decode()
            print(f"  ERROR sending email: HTTP {e.code}")
            print(f"  Response: {error_body[:300]}")
            
            # Try without CC
            print("  Retrying without CC...")
            payload_cc = payload.copy()
            payload_cc.pop("cc", None)
            try:
                data = json.dumps(payload_cc).encode('utf-8')
                req = urllib.request.Request(send_url, data=data, method="POST")
                req.add_header("Authorization", f"Bearer {API_KEY}")
                req.add_header("Content-Type", "application/json")
                
                with urllib.request.urlopen(req, timeout=30) as resp:
                    result = json.loads(resp.read().decode())
                    print(f"  Email sent (without CC)!")
            except urllib.error.HTTPError as e2:
                print(f"  Second attempt also failed: {e2.read().decode()[:200]}")
