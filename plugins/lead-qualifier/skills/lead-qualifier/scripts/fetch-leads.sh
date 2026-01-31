#!/bin/bash
# Fetch inbound leads for qualification
# Usage: ./fetch-leads.sh [timeframe] [limit]
#
# Examples:
#   ./fetch-leads.sh             # Default: last 7 days, limit 20
#   ./fetch-leads.sh 14d 50      # Last 14 days, limit 50

TIMEFRAME="${1:-7d}"
LIMIT="${2:-20}"

# Validate environment
if [ -z "$GOG_ACCOUNT" ]; then
    echo "Error: GOG_ACCOUNT not set"
    echo "Run: export GOG_ACCOUNT='user@company.com'"
    exit 1
fi

echo "Fetching leads (last ${TIMEFRAME}, limit ${LIMIT})..."
echo "Account: $GOG_ACCOUNT"
echo ""

# Common lead sources - adjust labels/queries for your setup
# Option 1: Labeled leads
# gog gmail search "label:leads newer_than:${TIMEFRAME}" --no-input --limit "$LIMIT"

# Option 2: Form submissions
# gog gmail search "(from:typeform.com OR from:hubspot.com) newer_than:${TIMEFRAME}" --no-input --limit "$LIMIT"

# Option 3: Direct inquiries (adjust for your sales email)
gog gmail search "(to:sales OR to:info OR subject:inquiry OR subject:contact) newer_than:${TIMEFRAME}" --no-input --limit "$LIMIT"
