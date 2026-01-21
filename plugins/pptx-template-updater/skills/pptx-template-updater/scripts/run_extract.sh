#!/bin/bash
# Wrapper script to run extract_template_structure.py in isolated environment

set -e  # Exit on error

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_DIR="$(dirname "$SCRIPT_DIR")"
VENV_DIR="$SKILL_DIR/.venv"
PYTHON_SCRIPT="$SCRIPT_DIR/extract_template_structure.py"

# Check if virtual environment exists
if [ ! -d "$VENV_DIR" ]; then
    echo "❌ Error: Virtual environment not found"
    echo "Please run setup.sh first:"
    echo "  cd $SKILL_DIR && ./setup.sh"
    exit 1
fi

# Check if the Python script exists
if [ ! -f "$PYTHON_SCRIPT" ]; then
    echo "❌ Error: Python script not found: $PYTHON_SCRIPT"
    exit 1
fi

# Execute the Python script in the virtual environment
"$VENV_DIR/bin/python" "$PYTHON_SCRIPT" "$@"
