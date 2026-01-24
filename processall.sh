#!/bin/bash

# Check parameter
action=${1:-main}
threshold=${2:-0.5}  # Default 5% threshold, can be overridden

source ~/development/python/.venv/bin/activate
currdate="2026-01-23"
prevdate="2026-01-16"

if [ "$action" = "main" ]; then
    python process_assets.py --delete-only --date $currdate
    python process_assets.py --normalize --date $currdate --datetocompare $prevdate
    python process_assets.py --process --date $currdate --datetocompare $prevdate
    python process_assets.py --refresh-dataconn  --currdate $currdate --datetocompare $prevdate
    python process_assets.py --compare-dates --currdate $currdate --datetocompare $prevdate --threshold $threshold --show-all
elif [ "$action" = "backup" ]; then
    docker exec -i assets mysqldump -u root -psa123 --single-transaction --set-gtid-purged=OFF --databases asset  > backup/asset_$currdate.sql
    python upload_to_gdrive.py
else
    echo "Usage: $0 [main|backup] [threshold]"
    echo "  main      - Run main processing (default)"
    echo "  backup    - Run backup and upload to Google Drive"
    echo "  threshold - Percentage change threshold for comparison report (default: 5.0)"
    echo ""
    echo "Examples:"
    echo "  $0 main         # Run with default 5% threshold"
    echo "  $0 main 0.5     # Run with 0.5% threshold"
    echo "  $0 main 10      # Run with 10% threshold"
    exit 1
fi