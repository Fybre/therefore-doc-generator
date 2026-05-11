FROM python:3.11-bookworm

# Node.js via NodeSource (more reliable than apt's outdated version)
RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs

# Chromium + all shared libs it needs
RUN apt-get install -y --no-install-recommends \
        chromium \
        libatk1.0-0 \
        libatk-bridge2.0-0 \
        libcups2 \
        libdrm2 \
        libgbm1 \
        libgtk-3-0 \
        libnspr4 \
        libnss3 \
        libx11-xcb1 \
        libxss1 \
        libxtst6 \
        fonts-liberation \
        fonts-noto \
    && rm -rf /var/lib/apt/lists/*

# Confirm chromium is where we expect it
RUN chromium --version

# Tell Puppeteer to skip downloading its own Chromium and use the system one
ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true
ENV PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium

# Install Mermaid CLI
RUN npm install -g @mermaid-js/mermaid-cli

# Confirm mmdc works
RUN mmdc --version

# Python dependencies
COPY requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt

WORKDIR /app
COPY . .

# Write puppeteer config pointing at system Chromium with sandbox disabled
RUN echo '{"executablePath":"/usr/bin/chromium","args":["--no-sandbox","--disable-setuid-sandbox","--disable-dev-shm-usage"]}' \
    > /app/puppeteer.json

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=10s --start-period=15s \
    CMD curl -f http://localhost:8000/ || exit 1

CMD ["uvicorn", "web.app:app", "--host", "0.0.0.0", "--port", "8000"]
