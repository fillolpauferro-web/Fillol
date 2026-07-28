#!/usr/bin/env bash
# Backup diário do Postgres do Chatwoot: dump + upload para object storage
# (DigitalOcean Spaces ou Backblaze B2, ambos compatíveis com a API S3).
#
# Requer: aws-cli instalado no host (não no container).
# Uso: rodar a partir do diretório infra/chatwoot-helpdesk/ (ou ajustar COMPOSE_DIR).
#
# Exemplo de crontab (todo dia às 3h da manhã):
#   0 3 * * * /caminho/para/infra/chatwoot-helpdesk/scripts/backup.sh >> /var/log/chatwoot-backup.log 2>&1

set -euo pipefail

COMPOSE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$COMPOSE_DIR"

# shellcheck disable=SC1091
source .env

RETENTION_DAYS=14
TIMESTAMP="$(date +%Y-%m-%d_%H-%M-%S)"
DUMP_FILE="chatwoot_${TIMESTAMP}.sql.gz"
TMP_PATH="/tmp/${DUMP_FILE}"

echo "[$(date)] Gerando dump do Postgres..."
docker compose exec -T postgres pg_dump -U "${POSTGRES_USER}" "${POSTGRES_DATABASE}" | gzip > "${TMP_PATH}"

echo "[$(date)] Enviando ${DUMP_FILE} para ${BACKUP_S3_BUCKET}..."
AWS_ACCESS_KEY_ID="${BACKUP_S3_ACCESS_KEY}" \
AWS_SECRET_ACCESS_KEY="${BACKUP_S3_SECRET_KEY}" \
aws s3 cp "${TMP_PATH}" "s3://${BACKUP_S3_BUCKET}/${DUMP_FILE}" \
  --endpoint-url "${BACKUP_S3_ENDPOINT}"

rm -f "${TMP_PATH}"

echo "[$(date)] Removendo backups com mais de ${RETENTION_DAYS} dias..."
CUTOFF_EPOCH=$(date -d "-${RETENTION_DAYS} days" +%s)

AWS_ACCESS_KEY_ID="${BACKUP_S3_ACCESS_KEY}" \
AWS_SECRET_ACCESS_KEY="${BACKUP_S3_SECRET_KEY}" \
aws s3api list-objects-v2 \
  --bucket "${BACKUP_S3_BUCKET}" \
  --endpoint-url "${BACKUP_S3_ENDPOINT}" \
  --query "Contents[?starts_with(Key, 'chatwoot_')].[Key,LastModified]" \
  --output text |
while read -r KEY LAST_MODIFIED; do
  [ -z "${KEY:-}" ] && continue
  FILE_EPOCH=$(date -d "${LAST_MODIFIED}" +%s)
  if [ "${FILE_EPOCH}" -lt "${CUTOFF_EPOCH}" ]; then
    echo "  removendo ${KEY} (${LAST_MODIFIED})"
    AWS_ACCESS_KEY_ID="${BACKUP_S3_ACCESS_KEY}" \
    AWS_SECRET_ACCESS_KEY="${BACKUP_S3_SECRET_KEY}" \
    aws s3 rm "s3://${BACKUP_S3_BUCKET}/${KEY}" --endpoint-url "${BACKUP_S3_ENDPOINT}"
  fi
done

echo "[$(date)] Backup concluído com sucesso."
