#!/bin/bash
# Setup environment for inbox-triage skill
# Usage: source ./setup-env.sh

# Prompt for account if not set
if [ -z "$GOG_ACCOUNT" ]; then
    echo "Enter your Gmail account (e.g., user@company.com):"
    read -r GOG_ACCOUNT
    export GOG_ACCOUNT
fi

# Set keyring password from config file if it exists
if [ -f "$HOME/.config/gog/keyring_pass" ]; then
    export GOG_KEYRING_PASSWORD="$(cat "$HOME/.config/gog/keyring_pass")"
    echo "✓ Keyring password loaded from config"
else
    echo "⚠ No keyring password found at ~/.config/gog/keyring_pass"
    echo "  You may need to run: gog auth login --account $GOG_ACCOUNT"
fi

echo "✓ GOG_ACCOUNT set to: $GOG_ACCOUNT"
