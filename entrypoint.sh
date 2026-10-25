#!/usr/bin/env bash
set -e

# 1) Virtual X display
Xvfb :99 -screen 0 1440x900x24 -nolisten tcp &
sleep 1

# 2) Lightweight window manager (needed for dialogs / focus)
fluxbox >/dev/null 2>&1 &
sleep 1

# 3) VNC server exporting :99 (port 5900, no password – local use)
x11vnc -display :99 -forever -shared -nopw -rfbport 5900 >/dev/null 2>&1 &

# 4) noVNC – exposes the VNC session over HTTP/WebSocket on 6080
websockify --web=/usr/share/novnc 6080 localhost:5900 >/dev/null 2>&1 &

# 5) Run the app
if [ -f /app/app.py ]; then
    exec python3 /app/app.py
else
    exec python3 /opt/app/app.py
fi