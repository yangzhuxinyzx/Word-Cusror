#!/usr/bin/env bash
# Word-Cursor Website Deploy Script
# Author: Claude
# Usage: ./scripts/deploy-website.sh

set -e

# Target server configuration from WEBSITE-DEV-DOC.md
SERVER_IP="8.141.124.194"
SERVER_USER="root"
SSH_KEY="~/.ssh/id_ed25519"
TARGET_DIR="/www/wwwroot/yangyzx.com/"

echo "==========================================="
echo "🚀 Starting Word-Cursor Website Deployment..."
echo "==========================================="

# Check if SSH key exists locally
if [ ! -f ~/.ssh/id_ed25519 ]; then
    echo "⚠️ Warning: SSH key ~/.ssh/id_ed25519 not found."
    echo "Please make sure your SSH key is located at ~/.ssh/id_ed25519 or edit this script."
fi

# Step 1: Upload index.html and favicon.svg
echo "📤 Uploading core assets to website root..."
scp -i "$SSH_KEY" website/index.html "$SERVER_USER@$SERVER_IP:$TARGET_DIR"
scp -i "$SSH_KEY" website/favicon.svg "$SERVER_USER@$SERVER_IP:$TARGET_DIR"

# Step 2: Upload assets directory
echo "📤 Uploading screenshots & gallery assets..."
scp -i "$SSH_KEY" -r website/assets/ "$SERVER_USER@$SERVER_IP:$TARGET_DIR"

# Step 3: Fix Nginx permissions and reload
echo "🔄 Restoring owner permissions and reloading Nginx..."
ssh -i "$SSH_KEY" "$SERVER_USER@$SERVER_IP" "chown -R www-data:www-data $TARGET_DIR && nginx -s reload"

echo "==========================================="
echo "🎉 Deployment successful!"
echo "🌐 URL: http://yangyzx.com/"
echo "==========================================="
