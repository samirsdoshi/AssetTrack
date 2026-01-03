currdate="2025-01-02"
prevdate="2025-12-31"
python process_assets.py --delete-only --date $currdate
python process_assets.py --normalize --date $currdate
python process_assets.py --updateassetref --date $currdate
python process_assets.py --process --date $currdate
python process_assets.py --refresh-dataconn  --currdate $currdate --datetocompare $prevdate
docker exec -i assets mysqldump -u root -psa123 --databases asset  > backup/asset_$currdate.sql