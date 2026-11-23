#!/usr/bin/env bash
# Focus Institute - One-Click Portal & Online Tunnel Starter

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "=========================================================="
echo " FOCUS INSTITUTE - STARTING TELECALLING PORTAL & TUNNEL"
echo "=========================================================="

# 1. Check if Flask is already running
if ! curl -s http://127.0.0.1:5500/api/me >/dev/null 2>&1; then
    echo "Starting Flask Server on port 5500..."
    python3 app.py > /tmp/focus_flask.log 2>&1 &
    sleep 2
fi

echo "✅ Portal Server is RUNNING at http://127.0.0.1:5500"

# 2. Start Cloudflare Tunnel
CLOUDFLARED="/home/harsh/.gemini/antigravity/scratch/cloudflared"
if [ ! -f "$CLOUDFLARED" ]; then
    echo "Downloading cloudflared..."
    curl -sL https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64 -o "$CLOUDFLARED"
    chmod +x "$CLOUDFLARED"
fi

echo "Connecting to Cloudflare edge network for public access..."
TUNNEL_LOG="/tmp/focus_tunnel.log"
pkill -f "cloudflared tunnel --url" 2>/dev/null || true
"$CLOUDFLARED" tunnel --url http://127.0.0.1:5500 > "$TUNNEL_LOG" 2>&1 &
TUNNEL_PID=$!

# Wait for tunnel URL
echo -n "Generating public internet link"
for i in {1..15}; do
    echo -n "."
    URL=$(grep -o 'https://[-a-zA-Z0-9.]*\.trycloudflare\.com' "$TUNNEL_LOG" 2>/dev/null | head -n 1)
    if [ -n "$URL" ]; then
        echo ""
        echo "=========================================================="
        echo " 🌐 ONLINE PUBLIC LINK (Works on any room, Wi-Fi & Mobile):"
        echo " 👉 $URL"
        echo "=========================================================="
        echo " Share this link with Madam Savita & your calling staff!"
        echo " Portal will stay active as long as this terminal is open."
        echo " Press Ctrl+C to stop."
        wait $TUNNEL_PID
        exit 0
    fi
    sleep 1
done

echo ""
echo "Tunnel started. Check log: $TUNNEL_LOG"
