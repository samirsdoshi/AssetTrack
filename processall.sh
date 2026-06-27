#!/bin/bash

set -euo pipefail

CONTAINER_NAME="assets"
MYSQL_ROOT_PASSWORD="sa123"
MYSQL_DATA_DIR="/Users/Sdoshi/development/samir/docs/Investments/US/data"

ensure_docker_ready() {
    if ! command -v docker >/dev/null 2>&1; then
        echo "Error: docker command not found in PATH."
        exit 1
    fi

    if docker info >/dev/null 2>&1; then
        return
    fi

    if command -v colima >/dev/null 2>&1; then
        echo "Docker daemon not running. Starting Colima..."
        colima start >/dev/null
    fi

    if ! docker info >/dev/null 2>&1; then
        echo "Error: Docker daemon is not running."
        echo "Start Docker Desktop or Colima and retry."
        exit 1
    fi
}

ensure_assets_container_running() {
    if docker ps --format '{{.Names}}' | grep -q "^${CONTAINER_NAME}$"; then
        echo "Docker container '${CONTAINER_NAME}' is already running."
        return
    fi

    if docker ps -a --format '{{.Names}}' | grep -q "^${CONTAINER_NAME}$"; then
        echo "Starting existing Docker container '${CONTAINER_NAME}'..."
        docker start "${CONTAINER_NAME}" >/dev/null
        echo "Docker container '${CONTAINER_NAME}' started."
        return
    fi

    echo "Creating and starting Docker container '${CONTAINER_NAME}'..."
    docker run -d \
        --name "${CONTAINER_NAME}" \
        -p 3306:3306 \
        -e "MYSQL_ROOT_PASSWORD=${MYSQL_ROOT_PASSWORD}" \
        --volume "${MYSQL_DATA_DIR}:/var/lib/mysql" \
        mysql:latest >/dev/null
    echo "Docker container '${CONTAINER_NAME}' created and started."
}

wait_for_mysql_ready() {
    local max_attempts=30
    local attempt=1

    echo "Waiting for MySQL in '${CONTAINER_NAME}' to become ready..."
    while [ "$attempt" -le "$max_attempts" ]; do
        if docker exec "${CONTAINER_NAME}" \
            mysqladmin ping -h 127.0.0.1 -u root -p"${MYSQL_ROOT_PASSWORD}" --silent >/dev/null 2>&1; then
            echo "MySQL is ready."
            return
        fi

        sleep 2
        attempt=$((attempt + 1))
    done

    echo "Error: MySQL did not become ready in time in container '${CONTAINER_NAME}'."
    exit 1
}

# Check parameter
action=${1:-main}
threshold=${2:-0.5}

source ~/development/python/.venv/bin/activate
currdate="2026-06-19"
prevdate="2026-06-12"

ensure_docker_ready
ensure_assets_container_running
wait_for_mysql_ready

if [ "$action" = "main" ]; then
    python process_assets.py --delete-only --date "$currdate"
    python process_assets.py --normalize --date "$currdate" --datetocompare "$prevdate"
    python process_assets.py --process --date "$currdate" --datetocompare "$prevdate"
    python process_assets.py --refresh-dataconn --currdate "$currdate" --datetocompare "$prevdate"
    python process_assets.py --compare-dates --currdate "$currdate" --datetocompare "$prevdate" --threshold "$threshold" --show-all
elif [ "$action" = "compare" ]; then
    python process_assets.py --refresh-dataconn --currdate "$currdate" --datetocompare "$prevdate"
    python process_assets.py --compare-dates --currdate "$currdate" --datetocompare "$prevdate" --threshold "$threshold" --show-all > out.txt
elif [ "$action" = "backup" ]; then
    echo "Backing up database for ${currdate}..."
    docker exec -i assets mysqldump -u root -psa123 --single-transaction --set-gtid-purged=OFF --databases asset > "backup/asset_${currdate}.sql"
    echo "Database backup complete: backup/asset_${currdate}.sql"
    
    echo "Backing up CSV source files for ${currdate}..."
    [ -f "Fidelity.csv" ] && cp "Fidelity.csv" "backup/Fidelity_${currdate}.csv" && echo "  Backed up: Fidelity_${currdate}.csv"
    [ -f "trow.csv" ] && cp "trow.csv" "backup/trow_${currdate}.csv" && echo "  Backed up: trow_${currdate}.csv"
    [ -f "stocks.csv" ] && cp "stocks.csv" "backup/stocks_${currdate}.csv" && echo "  Backed up: stocks_${currdate}.csv"
    [ -f "allaccounts.csv" ] && cp "allaccounts.csv" "backup/allaccounts_${currdate}.csv" && echo "  Backed up: allaccounts_${currdate}.csv"
    
    echo "Uploading backups to Google Drive..."
    python upload_to_gdrive.py
elif [ "$action" = "restore" ]; then
    backup_file="${2:-}"
    if [ -z "$backup_file" ]; then
        echo "Error: restore requires a backup file path"
        echo "Example: $0 restore backup/asset_2026-02-27.sql"
        exit 1
    fi
    ./restore_mysql.sh "$backup_file"
elif [ "$action" = "show-dates" ]; then
    if [ -n "${2:-}" ]; then
        python process_assets.py --show-dates --after-date "$2"
    else
        python process_assets.py --show-dates
    fi
else
    echo "Usage: $0 [main|compare|backup|restore|show-dates] [threshold|backup-file|after-date]"
    echo "  main       - Run main processing (default)"
    echo "  compare    - Compare two dates and write report to out.txt"
    echo "  backup     - Run backup and upload to Google Drive"
    echo "  restore    - Restore database from a backup SQL file"
    echo "  show-dates - Show unique dates with data in database"
    echo "  threshold  - Percentage change threshold for comparison report (default: 0.5)"
    echo "  backup-file- Path to SQL backup file for restore action"
    echo "  after-date - Filter dates to show only those after this date (YYYY-MM-DD)"
    echo ""
    echo "Examples:"
    echo "  $0 main"
    echo "  $0 main 0.5"
    echo "  $0 compare 0.5"
    echo "  $0 restore backup/asset_2026-02-27.sql"
    echo "  $0 show-dates"
    echo "  $0 show-dates 2026-01-01"
    exit 1
fi