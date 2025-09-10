#!/usr/bin/env bash
# End-to-end Poste.io mail server setup (internal + external)
# Using existing SSL cert (no Let's Encrypt)
# Akshay's Version — minimal changes needed

set -euo pipefail

# ======== CONFIG ========
DOMAIN="${DOMAIN:-example.com}"   # Your domain
HOSTNAME="mail.${DOMAIN}"
ADMIN_EMAIL="${ADMIN_EMAIL:-admin@${DOMAIN}}"
TZ="Asia/Kolkata"

# SSL cert + key (already existing)
SSL_CERT_PATH="/etc/ssl/mail/fullchain.pem"
SSL_KEY_PATH="/etc/ssl/mail/privkey.pem"

PROJECT_DIR="/opt/poste-mail"

# ======== CHECK SSL FILES ========
if [ ! -f "$SSL_CERT_PATH" ] || [ ! -f "$SSL_KEY_PATH" ]; then
    echo "[ERROR] SSL cert or key not found."
    echo "  Cert path: $SSL_CERT_PATH"
    echo "  Key path:  $SSL_KEY_PATH"
    exit 1
fi

# ======== INSTALL DOCKER ========
if ! command -v docker &>/dev/null; then
    echo "[INFO] Installing Docker..."
    curl -fsSL https://get.docker.com | sh
fi

# ======== INSTALL DOCKER COMPOSE ========
if ! docker compose version &>/dev/null; then
    echo "[INFO] Installing Docker Compose plugin..."
    DOCKER_PLUGIN_DIR="/usr/local/lib/docker/cli-plugins"
    sudo mkdir -p "$DOCKER_PLUGIN_DIR"
    curl -fsSL "https://github.com/docker/compose/releases/latest/download/docker-compose-linux-x86_64" \
        -o "$DOCKER_PLUGIN_DIR/docker-compose"
    chmod +x "$DOCKER_PLUGIN_DIR/docker-compose"
fi

# ======== CREATE PROJECT ========
mkdir -p "$PROJECT_DIR"
cd "$PROJECT_DIR"

# ======== CREATE docker-compose.yml ========
cat > docker-compose.yml <<EOF
version: "3.8"
services:
  poste:
    image: analogic/poste.io:latest
    container_name: poste
    restart: unless-stopped
    environment:
      - TZ=${TZ}
      - HTTPS=ON
    ports:
      - "25:25"
      - "80:80"
      - "443:443"
      - "465:465"
      - "587:587"
      - "993:993"
    volumes:
      - poste_data:/data
      - ${SSL_CERT_PATH}:/etc/ssl/certs/ssl-cert.crt:ro
      - ${SSL_KEY_PATH}:/etc/ssl/private/ssl-cert.key:ro
    networks:
      - mailnet

volumes:
  poste_data:

networks:
  mailnet:
    driver: bridge
EOF

# ======== OPEN FIREWALL ========
if command -v ufw &>/dev/null; then
    echo "[INFO] Opening firewall ports..."
    ufw allow 25/tcp
    ufw allow 465/tcp
    ufw allow 587/tcp
    ufw allow 993/tcp
    ufw allow 80/tcp
    ufw allow 443/tcp
else
    echo "[WARN] UFW not installed — open ports manually."
fi

# ======== START POSTE.IO ========
docker compose pull
docker compose up -d

# ======== DONE ========
echo
echo "======================================================="
echo " Poste.io is deployed!"
echo " Access admin panel at: https://${HOSTNAME} (or IP)"
echo " First time: set up admin account from browser."
echo
echo " Mailbox management, DKIM config, spam filter, attachments — all from GUI."
echo
echo " To check logs:"
echo "   docker logs -f poste"
echo "======================================================="
