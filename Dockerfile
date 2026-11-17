# Paper Desk in a container (Linux, LibreOffice instead of Word).  See DEPLOY.md.
FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer libreoffice-java-common nodejs fontconfig \
        fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Optional: the exact fonts the certificates use (Calibri, Algerian, Times New Roman ...), copied into ./fonts
# before building.  Without them LibreOffice swaps in look-alikes and the certificates look a little different.
COPY fonts/ /usr/share/fonts/truetype/paperdesk/
RUN fc-cache -f

COPY . .

ENV PAPERDESK_HOST=0.0.0.0 \
    PAPERDESK_ENGINE=libreoffice \
    PAPERDESK_FONTS=/usr/share/fonts/truetype/paperdesk \
    PAPERDESK_DATA=/data
VOLUME /data
EXPOSE 5000
CMD ["python", "webapp/app.py"]
