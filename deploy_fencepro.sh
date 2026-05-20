#!/bin/bash
# FencePro Deploy Script
# Usage: ./deploy_fencepro.sh
# Grabs the most recently downloaded HTML file and deploys it to the FencePro directory

DOWNLOADS=~/Downloads
DEST=~/Desktop/fencepro/index.html

# Find the most recently modified .html file in Downloads
LATEST=$(ls -t "$DOWNLOADS"/*.html 2>/dev/null | head -1)

if [ -z "$LATEST" ]; then
  echo "❌ No HTML files found in $DOWNLOADS"
  exit 1
fi

# Show what we found
echo "📄 Found: $(basename "$LATEST")"
echo "   Modified: $(date -r "$LATEST" '+%b %d %Y %I:%M %p')"

# Backup existing if it exists
if [ -f "$DEST" ]; then
  BACKUP="${DEST%.html}_backup_$(date +%Y%m%d_%H%M%S).html"
  cp "$DEST" "$BACKUP"
  echo "💾 Backed up existing to: $(basename "$BACKUP")"
fi

# Deploy
cp "$LATEST" "$DEST"

if [ $? -eq 0 ]; then
  echo "✅ Deployed to $DEST"
  echo "🗑  Removing downloaded file..."
  rm "$LATEST"
  echo "Done."
else
  echo "❌ Deploy failed"
  exit 1
fi
