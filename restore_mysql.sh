#!/bin/bash

set -euo pipefail

CONTAINER_NAME="${MYSQL_CONTAINER:-assets}"
MYSQL_USER="${MYSQL_USER:-root}"
MYSQL_PASSWORD="${MYSQL_PASSWORD:-sa123}"
MYSQL_DATABASE="${MYSQL_DATABASE:-asset}"
BACKUP_FILE="${1:-}"

if [ -z "$BACKUP_FILE" ]; then
    echo "Usage: $0 <backup-file.sql>"
    echo ""
    echo "Environment overrides (optional):"
    echo "  MYSQL_CONTAINER   Docker container name (default: assets)"
    echo "  MYSQL_USER        MySQL username (default: root)"
    echo "  MYSQL_PASSWORD    MySQL password (default: sa123)"
    echo "  MYSQL_DATABASE    Target database (default: asset)"
    echo ""
    echo "Example:"
    echo "  $0 backup/asset_2026-02-27.sql"
    exit 1
fi

if [ ! -f "$BACKUP_FILE" ]; then
    echo "Error: Backup file not found: $BACKUP_FILE"
    exit 1
fi

if ! docker ps --format '{{.Names}}' | grep -q "^${CONTAINER_NAME}$"; then
    echo "Error: Docker container '${CONTAINER_NAME}' is not running."
    echo "Start MySQL first (example: ./startmysql.sh)."
    exit 1
fi

echo "Ensuring database '${MYSQL_DATABASE}' exists..."
docker exec -i "$CONTAINER_NAME" mysql -u"$MYSQL_USER" -p"$MYSQL_PASSWORD" -e "CREATE DATABASE IF NOT EXISTS ${MYSQL_DATABASE};"

echo "Restoring '${BACKUP_FILE}' into '${MYSQL_DATABASE}' on container '${CONTAINER_NAME}'..."
docker exec -i "$CONTAINER_NAME" mysql -u"$MYSQL_USER" -p"$MYSQL_PASSWORD" "$MYSQL_DATABASE" < "$BACKUP_FILE"

echo "Restore completed successfully."
