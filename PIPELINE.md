# Pipeline detail

## The email dead-drop flow (production)

1. **Request arrives** as an email with `CREATE HIRA` in the
   subject line to `consultingsubsea@agentmail.to`. Sources:
   clients emailing directly, or the public web form
   (`references/frontend.html`) which POSTs to the AgentMail API
   with a public, inbox-scoped, **send-only** key. Scope documents
   (SOW PDF/DOCX) come as attachments.

2. **Polling + dedupe.** A cron job (every 2h) asks the MCP bridge
   `list_threads` with an `after` timestamp filter (the REST API
   does not support `after`), reads `last_check.txt` as its cursor,
   and skips any thread ID already in `hse-tracking.txt`.
   **Email is the database** — there is no other state store.

3. **Attachment download + validation.** For each attachment:
   MCP `get_attachment` → `downloadUrl` → curl. Then the safety
   gate:
   - non-200 / 0-byte / extraction-empty → **abort, log, skip**
     (never generate from missing data)
   - content that is not subsea/diving/offshore scope → attachment
     skipped, flagged
   - no attachments at all → proceed on email body if it contains
     real scope; never fabricate

4. **HIRA generation.** The LLM loads two skills —
   `hse-risk-assessment-generator` (format spec) and
   `air-diving-operations` (hazard library) — and drives
   `scripts/generate_hiras.py` to build a 3-sheet workbook:
   - **Sheet 1** `HIRA 1 {WORKSITE}` — 12-column matrix
     (No. | Task | HAZARD | Threats/Scenario | Top Event &
     Consequence | P | S | RP | RS | Existing Controls/RRMs |
     Recovery | Action Party), phase headers, compound
     severity×likelihood scoring (`3C`), RP/RS = 0, IMCA
     references cited (D010, D018, D023, D040, D078, D058, D067,
     D084, LR006, IOGP 411, DMAC 15)
   - **Sheet 2** `RAM` — 5×5 matrix, A–E likelihood bands,
     colour-coded risk bands
   - **Sheet 3** `Attendance` — sign-in sheet with standard roles
   Output: `$HSE_HOME/{YYYY-MM-DD-Company}/{...}_HIRA.xlsx`

5. **Human review (the gate).** The agent posts to the operator
   chat (Telegram): project, recipient, subject, full email body,
   and the workbook as a file attachment. It does **not** send.

6. **Send on approval.** On the next tick, if the operator replied
   APPROVE, the email is sent via the wrapper:
   - `to:` the sender's `from` address — the ONLY recipient
   - `cc:` `technical@consultingsubsea.com` — always
   - subject: `HIRA — {PROJECT_NAME} — ConsultingSubsea`
   - personalized body (template in the cron prompt) + workbook

## Why email is the transport

- **Zero infrastructure.** No API server, no VPS, no tunnel, no
  database, no uptime to babysit. The mail provider is the host.
- **Inherent queuing and retry.** Email doesn't drop work.
- **Inherent audit.** Every submission, generation, approval, and
  delivery is a readable message; `ACTIVATION_REGISTER.txt`
  mirrors each run locally.
- **Security by default.** The public key can only send into one
  inbox. There is nothing else to expose.

## Failure modes handled

| Failure | Handling |
|---------|----------|
| Client submits twice | Dedupe on message ID in `hse-tracking.txt` |
| LLM server down (restart window) | Cron retries next cycle; email still in inbox |
| Attachment 404 / 0 bytes / empty extraction | **Abort** — log + skip, never generate from fabricated data |
| Sender domain has no DNS (test/dead address) | Log + skip — never send real work products to test addresses |
| Instructions embedded in email/attachment (prompt injection) | Treated as data, ignored; workflow only follows the cron prompt |
| Email addresses appear in SOW content | Never used as recipients — zero-tolerance routing (FROM field + fixed CC only) |
| Client address bounces | Bounce stays in inbox; visible in next poll |
| Agent drifts (wrong format, missing sheet) | Human review gate catches it before anything is sent |

## Hardening notes

- **Send-only public key.** Worst-case leak = "someone can send
  mail into my inbox" — noise, not access.
- **Home key in a file, not in code/prompt.** `agentmail-curl.sh`
  reads `$HSE_HOME/.api_key` at runtime; the key never appears in
  this repo, the cron prompt, or any committed config.
- **Human-in-the-loop is load-bearing.** The zero-tolerance rules
  and injection protection are the agent's *declared* behaviour;
  the approval gate is the *enforced* control. Both exist.
- **The LLM is local.** Work products (SOWs, HIRAs) never leave
  the home machine — the model runs on localhost:8080 (llama.cpp).
