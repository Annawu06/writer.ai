#!/bin/bash
# /opt/libreoffice26.2/program/soffice --writer --accept="socket,host=localhost,port=2002;urp;"

# Extension name
EXTENSION_NAME="writer.ai"

# Remove old package if it exists
if [ -f "${EXTENSION_NAME}.oxt" ]; then
    echo "Removing old package..."
    rm "${EXTENSION_NAME}.oxt"
fi

# Create the new package
echo "Creating package ${EXTENSION_NAME}.oxt..."
zip -r "${EXTENSION_NAME}.oxt" description.xml Addons.xcu main.py LICENSE META-INF assets description

if [ $? -eq 0 ]; then
    if unzip -Z1 "${EXTENSION_NAME}.oxt" | rg -qi '(^|/)(AGENTS?\.md|task\.md|mao[^/]*)$|(^|/)Select$'; then
        echo "Error: release package contains a private file."
        exit 1
    fi
    echo "Package created successfully: ${EXTENSION_NAME}.oxt"
    echo ""
    echo "To install:"
    echo "  1. Open LibreOffice"
    echo "  2. Tools > Extension Manager"
    echo "  3. Add > Select ${EXTENSION_NAME}.oxt"
    echo ""
else
    echo "Error creating package"
    exit 1
fi
