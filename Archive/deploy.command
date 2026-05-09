#!/bin/bash

# FencePro Deploy Script
# Double-click this file to deploy your latest changes to the web

cd ~/Desktop/fencepro

echo "🚀 Deploying FencePro..."
echo ""

# Stage all changes
git add .

# Commit with timestamp
git commit -m "Update $(date '+%Y-%m-%d %I:%M %p')"

# Push to GitHub (Netlify auto-deploys from there)
git push

echo ""
echo "✅ Done! Your changes will be live at:"
echo "   https://fencepro-blackwolf.netlify.app"
echo ""
echo "Press any key to close..."
read -n 1
