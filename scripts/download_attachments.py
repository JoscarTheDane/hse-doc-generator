# AgentMail key: set AGENTMAIL_API_KEY env var or place key file at $HSE_HOME/.api_key
#!/usr/bin/env python3
"""Download attachments from HIRA threads and extract their content."""

import json
import os
import urllib.request
import urllib.error

API_KEY = "${AGENTMAIL_API_KEY}"
BASE_URL = "https://api.agentmail.to/v0"
SAVE_DIR = "/home/joshua/.hermes/skills/productivity/hse-risk-assessment-generator/scripts/temp_attachments"

os.makedirs(SAVE_DIR, exist_ok=True)

def api_request(endpoint, method="GET"):
    url = f"{BASE_URL}{endpoint}"
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"Bearer {API_KEY}")
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        print(f"  HTTP Error {e.code}: {e.read().decode()[:200]}")
        return None

threads_info = [
    {
        "thread_id": "19d38c45-d362-41ce-8b76-260a511ce009",
        "project": "Snake_Island_Wrecks_Removal",
    },
    {
        "thread_id": "64660819-2c2c-477c-8ef4-ad8f2fec0716",
        "project": "SNEPCo_Bonga_Mooring",
    },
    {
        "thread_id": "5ad6dd11-c4ce-458e-8da6-6aff9e5ceea7",
        "project": "3D_Photogrammetry_Inspection",
    },
    {
        "thread_id": "75074bf3-0b34-4bb2-b662-85206db649c0",
        "project": "Flange_Disconnection_Reconnection",
    },
]

for tinfo in threads_info:
    thread_id = tinfo["thread_id"]
    project = tinfo["project"]
    
    print(f"\n{'='*60}")
    print(f"Processing: {project}")
    
    thread_data = api_request(f"/threads/{thread_id}")
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
        
        for att in attachments:
            filename = att.get("filename", "unknown")
            att_id = att.get("attachment_id", "")
            size = att.get("size", 0)
            content_type = att.get("content_type", "")
            
            print(f"    Attachment: {filename} ({size} bytes, {content_type})")
            
            # Get attachment details for download URL
            att_details = api_request(f"/threads/{thread_id}/attachments/{att_id}")
            if att_details:
                download_url = att_details.get("download_url", "")
                extracted_text = att_details.get("extracted_text", "") or att_details.get("text", "")
                
                if download_url:
                    out_path = os.path.join(SAVE_DIR, f"{project}_{os.path.basename(filename)}")
                    try:
                        with urllib.request.urlopen(download_url, timeout=60) as resp:
                            data = resp.read()
                            with open(out_path, "wb") as f:
                                f.write(data)
                        print(f"    Downloaded to: {out_path} ({len(data)} bytes)")
                    except Exception as e:
                        print(f"    Download failed: {e}")
                
                if extracted_text:
                    text_path = os.path.join(SAVE_DIR, f"{project}_{os.path.basename(filename)}.txt")
                    with open(text_path, "w") as f:
                        f.write(extracted_text)
                    print(f"    Extracted text: {text_path} ({len(extracted_text)} chars)")
            
            print()
