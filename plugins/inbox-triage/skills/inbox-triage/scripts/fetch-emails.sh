#!/bin/bash
# Fetch unread emails for inbox triage
# Usage: ./fetch-emails.sh [timeframe] [limit]
#
# Examples:
#   ./fetch-emails.sh             # Default: last 24h, limit 30
#   ./fetch-emails.sh 7d 50       # Last 7 days, limit 50
#   ./fetch-emails.sh 1d 10       # Last 24h, limit 10

TIMEFRAME="${1:-1d}"
LIMIT="${2:-30}"

# Validate environment
if [ -z "$GOG_ACCOUNT" ]; then
    echo "Error: GOG_ACCOUNT not set"
    echo "Run: source ./scripts/setup-env.sh"
    exit 1
fi

echo "Fetching unread emails (last ${TIMEFRAME}, limit ${LIMIT})..."
echo "Account: $GOG_ACCOUNT"
echo ""

gog gmail search "is:unread newer_than:${TIMEFRAME}" --no-input --limit "$LIMIT"
