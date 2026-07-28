#!/usr/bin/env bash
# Healthcheck periódico do Chatwoot: verifica se o painel responde e,
# se não responder, envia um e-mail de alerta.
#
# Requer: `mail` (pacote mailutils/bsd-mailx) configurado no host para
# conseguir enviar e-mail (ou adapte para usar `msmtp` com as credenciais
# SMTP do .env).
#
# Exemplo de crontab (a cada 5 minutos):
#   */5 * * * * /caminho/para/infra/chatwoot-helpdesk/scripts/healthcheck.sh >> /var/log/chatwoot-healthcheck.log 2>&1

set -euo pipefail

COMPOSE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$COMPOSE_DIR"

# shellcheck disable=SC1091
source .env

STATUS_CODE="$(curl -s -o /dev/null -w '%{http_code}' --max-time 10 "${FRONTEND_URL}" || echo "000")"

if [ "${STATUS_CODE}" != "200" ]; then
  MESSAGE="[$(date)] Chatwoot indisponível em ${FRONTEND_URL} (HTTP ${STATUS_CODE})."
  echo "${MESSAGE}"
  echo "${MESSAGE}" | mail -s "ALERTA: Helpdesk fora do ar" "${HEALTHCHECK_ALERT_EMAIL}"
  exit 1
fi

echo "[$(date)] Chatwoot OK (HTTP ${STATUS_CODE})."
