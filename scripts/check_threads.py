# AgentMail key: set AGENTMAIL_API_KEY env var or place key file at $HSE_HOME/.api_key
#!/usr/bin/env python3
import json
import urllib.request
import urllib.error

API_KEY = "${AGENTMAIL_API_KEY}"
BASE_URL = "https://api.agentmail.to/v0"

def api_request(endpoint, method="GET"):
    url = f"{BASE_URL}{endpoint}"
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"Bearer {API_KEY}")
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        print(f"HTTP Error {e.code}: {e.read().decode()}")
        return None

# List threads
result = api_request("/inboxes/consultingsubsea@agentmail.to/threads?limit=20")
if result:
    threads = result.get("threads", [])
    print(f"Found {len(threads)} threads")
    for t in threads:
        thread_id = t.get("thread_id", "N/A")
        subject = t.get("subject", "N/A")
        from_addr = t.get("from", {}).get("address", "N/A")
        created = t.get("created_at", "N/A")
        has_create_hira = "CREATE HIRA" in subject.upper()
        print(f"  Thread: {thread_id[:20]}... | Subject: {subject} | From: {from_addr} | Has CREATE HIRA: {has_create_hira} | Created: {created}")
else:
    print("Failed to fetch threads")
