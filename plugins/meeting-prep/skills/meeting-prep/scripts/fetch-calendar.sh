#!/bin/bash
# Fetch calendar events for meeting prep
# Usage: ./fetch-calendar.sh [from] [to]
#
# Examples:
#   ./fetch-calendar.sh                        # Default: tomorrow
#   ./fetch-calendar.sh today today+1          # Today's meetings
#   ./fetch-calendar.sh "2024-02-15" "2024-02-16"  # Specific date

FROM="${1:-tomorrow}"
TO="${2:-tomorrow+1}"

# Validate environment
if [ -z "$GOG_ACCOUNT" ]; then
    echo "Error: GOG_ACCOUNT not set"
    echo "Run: export GOG_ACCOUNT='user@company.com'"
    exit 1
fi

echo "Fetching calendar events..."
echo "Account: $GOG_ACCOUNT"
echo "From: $FROM"
echo "To: $TO"
echo ""

gog calendar list --from "$FROM" --to "$TO" --no-input
