#!/bin/bash

# Absolute path to the python executable
PYTHON_EXEC="/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"

# Absolute path to the script
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPT_PATH="$SCRIPT_DIR/Transformer.py"

# Run the script with all arguments passed to this wrapper
"$PYTHON_EXEC" "$SCRIPT_PATH" "$@"
