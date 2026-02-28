#!/bin/bash

# Check parameter
action=${1:-main}
threshold=${2:-0.5}  # Default 5% threshold, can be overridden

source ~/development/python/.venv/bin/activate
prevdate="2026-02-20"
currdate="2026-02-27"


if [ "$action" = "main" ]; then
    python process_assets.py --delete-only --date $currdate
    python process_assets.py --normalize --date $currdate --datetocompare $prevdate
    python process_assets.py --process --date $currdate --datetocompare $prevdate
    python process_assets.py --refresh-dataconn  --currdate $currdate --datetocompare $prevdate
    python process_assets.py --compare-dates --currdate $currdate --datetocompare $prevdate --threshold $threshold --show-all
elif [ "$action" = "compare" ]; then
    python process_assets.py --refresh-dataconn  --currdate $currdate --datetocompare $prevdate
    python process_assets.py --compare-dates --currdate $currdate --datetocompare $prevdate --threshold $threshold --show-all > out.txt
elif [ "$action" = "backup" ]; then
    docker exec -i assets mysqldump -u root -psa123 --single-transaction --set-gtid-purged=OFF --databases asset  > backup/asset_$currdate.sql
    python upload_to_gdrive.py
elif [ "$action" = "restore" ]; then
    backup_file="$2"
    if [ -z "$backup_file" ]; then
        echo "Error: restore requires a backup file path"
        echo "Example: $0 restore backup/asset_2026-02-27.sql"
        exit 1
    fi
    ./restore_mysql.sh "$backup_file"
elif [ "$action" = "show-dates" ]; then
    # Show available dates - optionally filter after a date
    if [ -n "$2" ]; then
        python process_assets.py --show-dates --after-date $2
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
    echo "  $0 main              # Run with default 0.5% threshold"
    echo "  $0 main 0.5          # Run with 0.5% threshold"
    echo "  $0 main 10           # Run with 10% threshold"
    echo "  $0 compare 0.5       # Compare with 0.5% threshold"
    echo "  $0 restore backup/asset_2026-02-27.sql"
    echo "  $0 show-dates        # Show all available dates"
    echo "  $0 show-dates 2026-01-01  # Show dates after 2026-01-01"
    exit 1
fi