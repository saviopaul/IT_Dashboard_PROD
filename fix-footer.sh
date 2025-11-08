#!/bin/bash
# Auto-update footer text after rebuilds
TARGET_FILE=$(grep -ril "Made with Dyad" /var/www/it-dashboard/assets 2>/dev/null)

if [ -n "$TARGET_FILE" ]; then
  sed -i 's/Made with Dyad/Made by Savio Paul/g' "$TARGET_FILE"
  echo "✅ Footer text updated in $TARGET_FILE"
else
  echo "⚠️ No Dyad footer text found. Skipping."
fi

sudo systemctl reload nginx
