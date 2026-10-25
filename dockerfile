FROM python:3.11

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    DISPLAY=:99 \
    PYTHONPATH=/opt/pysite

# GUI stack + Persian fonts + helpers
RUN apt-get update && apt-get install -y --no-install-recommends \
        xvfb \
        x11vnc \
        fluxbox \
        novnc \
        websockify \
        xdg-utils \
        fonts-noto-core \
        fonts-farsiweb \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
RUN mkdir -p /app /opt/app /opt/pysite

# Python deps
COPY requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt

# Linux replacement for os.startfile() (auto-imported via sitecustomize)
RUN cat > /opt/pysite/sitecustomize.py <<'PY'
import os, subprocess
if not hasattr(os, "startfile"):
    def _startfile(path):
        try:
            subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            pass
    os.startfile = _startfile
PY

# Fallback copy of the app (used if the host dir is not mounted)
COPY app.py /opt/app/app.py

COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

EXPOSE 5900 6080
ENTRYPOINT ["/entrypoint.sh"]