#!/bin/bash

# Check parameter
action=${1:-main}

startenv
currdate="2026-01-09"
prevdate="2025-12-31"

if [ "$action" = "main" ]; then
    python process_assets.py --delete-only --date $currdate
    python process_assets.py --normalize --date $currdate
    python process_assets.py --process --date $currdate
    python process_assets.py --refresh-dataconn  --currdate $currdate --datetocompare $prevdae
elif [ "$action" = "backup" ]; then
    docker exec -i assets mysqldump -u root -psa123 --databases asset  > backup/asset_$currdate.sql
    python upload_to_gdrive.py
else
    echo "Usage: $0 [main|backup]"
    echo "  main   - Run main processing (default)"
    echo "  backup - Run backup and upload to Google Drive"
    exit 1
fi