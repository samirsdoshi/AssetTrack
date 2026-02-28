#!/bin/bash

# Check parameter
action=${1:-main}
threshold=${2:-0.5}  # Default 5% threshold, can be overridden

source ~/development/python/.venv/bin/activate
currdate="2026-02-20"
prevdate="2026-02-13"


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
elif [ "$action" = "show-dates" ]; then
    # Show available dates - optionally filter after a date
    if [ -n "$2" ]; then
        python process_assets.py --show-dates --after-date $2
    else
        python process_assets.py --show-dates
    fi
else
    echo "Usage: $0 [main|backup|show-dates] [threshold|after-date]"
    echo "  main       - Run main processing (default)"
    echo "  backup     - Run backup and upload to Google Drive"
    echo "  show-dates - Show unique dates with data in database"
    echo "  threshold  - Percentage change threshold for comparison report (default: 0.5)"
    echo "  after-date - Filter dates to show only those after this date (YYYY-MM-DD)"
    echo ""
    echo "Examples:"
    echo "  $0 main              # Run with default 0.5% threshold"
    echo "  $0 main 0.5          # Run with 0.5% threshold"
    echo "  $0 main 10           # Run with 10% threshold"
    echo "  $0 show-dates        # Show all available dates"
    echo "  $0 show-dates 2026-01-01  # Show dates after 2026-01-01"
    exit 1
fi