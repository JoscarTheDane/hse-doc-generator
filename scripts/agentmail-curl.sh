#!/bin/bash
# Wrapper for AgentMail API calls - reads key from the correct source
# Usage: agentmail-curl.sh <method> <path> [data]
set -e

KEY=$(cat /home/joshua/HSE/.api_key 2>/dev/null | tr -d '\n' | tr -d '\r' | head -c 255)

if [ -z "$KEY" ]; then
    # Fall back to .env
    source /home/joshua/.hermes/.env 2>/dev/null
    KEY="$AGENTMAIL_API_KEY"
fi

METHOD="${1:-GET}"
PATH_ARG="${2:-/inboxes}"
DATA="${3:-}"

BASE="https://api.agentmail.to"
MCP_BASE="https://mcp.agentmail.to"

# Use the key from the wrapper variable
AUTH_KEY="$KEY"

case "$PATH_ARG" in
mcp:*)
    # MCP JSON-RPC call - strip "mcp:" prefix
    BODY="${PATH_ARG#mcp:}"
    curl -s -X POST "$MCP_BASE/mcp" \
      -H "Authorization: Bearer $AUTH_KEY" \
      -H "Content-Type: application/json" \
      -H "Accept: application/json, text/event-stream" \
      -d "$BODY"
    ;;
*)
    # REST API call
    if [ -n "$DATA" ]; then
        curl -s -X "$METHOD" "$BASE$PATH_ARG" \
          -H "Authorization: Bearer $AUTH_KEY" \
          -H "Content-Type: application/json" \
          -d "$DATA"
    else
        curl -s -X "$METHOD" "$BASE$PATH_ARG" \
          -H "Authorization: Bearer $AUTH_KEY" \
          -H "Accept: application/json"
    fi
    ;;
esac
