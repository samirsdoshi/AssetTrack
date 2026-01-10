startenv
currdate="2026-01-09"
prevdate="2025-12-31"
python process_assets.py --delete-only --date $currdate
python process_assets.py --normalize --date $currdate
python process_assets.py --process --date $currdate
python process_assets.py --refresh-dataconn  --currdate $currdate --datetocompare $prevdate
#docker exec -i assets mysqldump -u root -psa123 --databases asset  > backup/asset_$currdate.sql
#python upload_to_gdrive.py     


#python process_assets.py --refresh-dataconn  --currdate 2025-12-31 --datetocompare 2024-12-27