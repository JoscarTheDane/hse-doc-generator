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
        print(f"  HTTP Error {e.code}: {e.read().decode()[:200]}")
        return None

# Read tracking file
tracked_ids = set()
try:
    with open("$HSE_HOME/hse-tracking.txt", "r") as f:
        for line in f:
            stripped = line.strip().strip("<>").strip()
            if stripped:
                tracked_ids.add(stripped)
except FileNotFoundError:
    pass

print(f"Tracked message IDs: {len(tracked_ids)}")

# List threads
result = api_request("/inboxes/consultingsubsea@agentmail.to/threads?limit=20")
if not result:
    print("Failed to fetch threads")
    exit(1)

threads = result.get("threads", [])

for t in threads:
    thread_id = t.get("thread_id", "")
    subject = t.get("subject", "")
    if "CREATE HIRA" not in subject.upper():
        continue
    
    print(f"\n{'='*60}")
    print(f"Thread: {thread_id}")
    print(f"Subject: {subject}")
    
    thread_data = api_request(f"/threads/{thread_id}")
    if not thread_data:
        continue
    
    messages = thread_data.get("messages", [])
    
    # Find the original request (first non-reply from jsh.evan or similar)
    # and find our replies
    our_replies = []
    original_request = None
    request_msg_id = None
    
    for msg in messages:
        msg_id = msg.get("message_id", "")
        from_raw = msg.get("from", "")
        if isinstance(from_raw, dict):
            from_addr = from_raw.get("address", "")
        else:
            from_addr = str(from_raw) if from_raw else ""
        to_raw = msg.get("to", "")
        if isinstance(to_raw, list):
            to_addrs = [a.get("address", "") if isinstance(a, dict) else str(a) for a in to_raw]
        elif isinstance(to_raw, dict):
            to_addrs = [to_raw.get("address", "")]
        else:
            to_addrs = [str(to_raw)] if to_raw else []
        
        is_ours = "consultingsubsea@agentmail.to" in from_addr
        
        print(f"  msg_id: {msg_id[:80]}")
        print(f"    From: {from_addr} | To: {', '.join(to_addrs[:3])}")
        print(f"    Subject: {msg.get('subject', '')}")
        print(f"    Is ours: {is_ours} | Tracked: {msg_id in tracked_ids}")
        print(f"    Attachments: {len(msg.get('attachments', []))}")
        
        if is_ours:
            our_replies.append(msg_id)
        
        if not is_ours and not request_msg_id:
            request_msg_id = msg_id
            original_request = msg
    
    # Check if we've already replied
    has_our_reply = len(our_replies) > 0
    print(f"\n  SUMMARY: {len(messages)} messages, {len(our_replies)} replies from us")
    print(f"  Original request msg_id: {request_msg_id}")
    print(f"  Request in tracked: {request_msg_id in tracked_ids if request_msg_id else 'N/A'}")
    print(f"  Has our reply: {has_our_reply}")
    if has_our_reply:
        print(f"  STATUS: ALREADY PROCESSED (we replied)")
    else:
        print(f"  STATUS: NEEDS PROCESSING")
