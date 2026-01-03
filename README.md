# Asset Processing - Python Conversion

This project converts VBA-based asset processing to Python, connecting to MySQL database to manage investment portfolio allocations and performance calculations.

## Converted Files

- **module1.vba** → **asset_processor.py** - Main asset processing logic
- **module2.vba** → **utils.py** - Utility functions
- **New:** **process_assets.py** - Main script to process Asset.xls

## Features

### AssetDatabase
- MySQL connection management (port 3306, schema 'asset')
- Transaction support with commit/rollback
- Query execution with parameterized statements

### AssetAllocator
- `allocate()` - Allocate assets based on templates
- `reallocate()` - Delete and reallocate existing assets
- `allocate_asset_ref()` - Process asset reference sheet data
- `delete_asset_info()` - Clean up old asset data

### GainCalculator
- Calculate performance gains from Yahoo Finance
- Calculate gains from Morningstar (when available)
- Support for multiple time periods (1 week, 2 weeks, 1 month, 3 months, 6 months, 1 year)
- Market trading day adjustments

### TemplateManager
- Manage allocation templates
- Add allocation type details
- Add sector/industry details
- Add interest rate details

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Create `.env` file with Excel password (if your Excel file is password-protected):
```bash
echo "password=YourExcelPassword" > .env
```

3. (Optional) Set up Google Drive API for backup uploads:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing one
   - Enable Google Drive API
   - Create OAuth 2.0 credentials (Desktop app)
   - Download `credentials.json` to this directory

4. Ensure MySQL is running on port 3306 (use your existing startmysql.sh):
```bash
./startmysql.sh
```

4. Verify database connection:
```bash
mysql -h localhost -P 3306 -u root -psa123 -e "USE asset; SHOW TABLES;"
```

## Usage

### Process Asset.xls (Full Workflow)

Process the Asset.xls file with current date:
```bash
python process_assets.py --file Asset.xls
```

Process with specific date:
```bash
python process_assets.py --file Asset.xls --date 2025-12-31
```

Process specific sheet:
```bash
python process_assets.py --file Asset.xls --sheet assetref --date 2025-12-31
```

### Options

- `--file, -f` - Path to Excel file (default: Asset.xls)
- `--sheet, -s` - Sheet name to process (default: assetref)
- `--date, -d` - As-of date in YYYY-MM-DD format (default: today)
- `--process` - Run main asset allocation workflow (default if no other mode specified)
- `--no-delete` - Skip deleting existing data
- `--no-gains` - Skip gain calculations
- `--gains-only` - Only calculate gains, skip allocation
- `--delete-only` - Only delete data for the date
- `--normalize` - Normalize and aggregate fullview data by account
- `--normalize-sheet` - Sheet name for normalize (default: fullview)
- `--updateassetref` - Update assetref sheet with allocation data from database
- `--refresh-dataconn` - Refresh DataConn sheet with database queries
- `--output, -o` - Output Excel file for normalize results

### Examples

Run main asset allocation workflow (explicit):
```bash
python process_assets.py --process --date 2025-12-31
```

Delete existing data and reprocess (default):
```bash
python process_assets.py --date 2025-12-31
```

Process without calculating gains:
```bash
python process_assets.py --date 2025-12-31 --no-gains
```

Only calculate gains for existing data:
```bash
python process_assets.py --date 2025-12-31 --gains-only
```

Delete data only:
```bash
python process_assets.py --date 2025-12-31 --delete-only
```

### Normalize Full View

Normalize and aggregate account data from fullview sheet:
```bash
python process_assets.py --normalize --file Asset_clean.xls
```

Save normalized results to Excel:
```bash
python process_assets.py --normalize --output normalized_results.xlsx
```

Process custom sheet:
```bash
python process_assets.py --normalize --normalize-sheet mysheet --output results.xlsx
```

This operation:
- Reads the 'fullview' sheet (or custom sheet)
- Categorizes accounts (CollegeAdv, TRPInv, FidelityIRA, Vanguard, etc.)
- Aggregates fund holdings by account
- Writes aggregated data to columns K, L, N, O in the fullview sheet
- Preserves all existing formulas
- Outputs detailed fund list and summary totals

### Update Asset Reference

Update assetref sheet with allocation data from database for a specific date:
```bash
python process_assets.py --updateassetref --date 2025-12-31
```

Specify custom file:
```bash
python process_assets.py --updateassetref --file Asset_clean.xls --date 2025-12-31
```

This operation:
- Queries the database for allocation data on the specified date
- Updates the assetref sheet with alloctype percentages
- Writes data to columns M onwards in the assetref sheet
- Preserves existing data in other columns

### Refresh DataConn Sheet

Refresh DataConn sheet with data from database queries:
```bash
python process_assets.py --refresh-dataconn
```

Specify custom file:
```bash
python process_assets.py --refresh-dataconn --file Asset_clean.xls
```

This operation:
- Executes 4 predefined database queries
- Populates the DataConn sheet with results from each query
- Used to sync Excel with current database state

### Refresh Connection Options

The `--refresh-dataconn` option supports multiple database connection configurations and date parameters:

**Default Local Connection:**
```bash
python process_assets.py --refresh-dataconn
```
Uses default values:
- Host: localhost
- Port: 3306
- User: root
- Password: sa123
- Database: asset

**With Current Date and Previous Date:**
```bash
python process_assets.py --refresh-dataconn --currdate 2025-12-31 --datetocompare 2025-12-24
```

This updates:
- `wkdates.currdate` with the current date (2025-12-31)
- `wkdates.datetocompare` with the comparison date (2025-12-24)
- Used for comparing performance between two dates in Excel reports

**Custom Host:**
```bash
python process_assets.py --refresh-dataconn --db-host 192.168.1.10
```

**Custom Port:**
```bash
python process_assets.py --refresh-dataconn --db-port 3307
```

**Custom User Credentials:**
```bash
python process_assets.py --refresh-dataconn --db-user myuser --db-password mypass
```

**Alternative Database:**
```bash
python process_assets.py --refresh-dataconn --db-name asset_backup
```

**Combined Options with Dates:**
```bash
python process_assets.py --refresh-dataconn --file Asset_clean.xls --currdate 2025-12-31 --datetocompare 2025-12-24
```

**Environment Variable Configuration:**
Set database connection via environment variables (.env file):
```
DB_HOST=localhost
DB_PORT=3306
DB_USER=root
DB_PASSWORD=sa123
DB_NAME=asset
```

Then run:
```bash
python process_assets.py --refresh-dataconn --currdate 2025-12-31 --datetocompare 2025-12-24
```

#### Date Parameters Details

- `--currdate` - Current date for wkdates table and assetref cell N4 (YYYY-MM-DD format)
  - Updates current/latest portfolio value
  - Reflects today's or most recent data
  
- `--datetocompare` - Comparison date for wkdates table and assetref cell N3 (YYYY-MM-DD format)
  - Previous date to compare against
  - Used to calculate period-over-period changes
  - Must have data in database for the specified date

## Excel File Format

The script expects an Excel file (Asset.xls or .xlsx) with an 'assetref' sheet containing:

| Column | Description |
|--------|-------------|
| Ticker/Symbol | Asset ticker symbol |
| Amount/Value | Investment amount |
| HeldAt | Location where asset is held (e.g., FidelityInv, TRPRoth, Vanguard) |

The script will process rows until it encounters "ENDOFPORTFOLIO" in the Ticker column.

**Password Protection**: If your Excel file is password-protected, the script will automatically decrypt it using the password from the `.env` file. Create a `.env` file with:
```
password=YourExcelPassword
```

The script attempts to read the file directly first, and only uses decryption if needed.

## Database Schema

The script works with the following tables in the 'asset' schema:

- `asset` - Asset definitions with templates
- `assetinv` - Asset investments by date
- `assetinvalloc` - Allocation type breakdowns
- `assetinvsecind` - Sector/Industry breakdowns
- `assetinvinter` - Interest rate breakdowns
- `assetgain` - Performance gains
- `templatedetails` - Allocation templates
- `alloctype` - Allocation types
- `sector` - Sectors
- `industry` - Industries
- `inter` - Interest rate types
- `holiday` - Market holidays

## Python API Usage

### Using as a Library

```python
from datetime import datetime
from asset_processor import AssetDatabase, AssetAllocator, GainCalculator

# Initialize database connection
db = AssetDatabase(
    host='localhost',
    port=3306,
    user='root',
    password='sa123',
    database='asset'
)

# Allocate a single asset
allocator = AssetAllocator(db)
allocator.allocate(
    asset_id=123,
    as_of_date=datetime(2025, 12, 31),
    amount=50000.0
)

# Calculate gains
gain_calc = GainCalculator(db)
gain_calc.calculate_gains(datetime(2025, 12, 31))
```

### Custom Processing

```python
from process_assets import AssetProcessor

processor = AssetProcessor('MyAssets.xlsx')

# Delete existing data
processor.delete_existing_data(datetime(2025, 12, 31))

# Process allocation
df = processor.read_asset_reference_sheet('Sheet1')
processor.process_asset_allocation(df, datetime(2025, 12, 31))

# Calculate gains
processor.calculate_gains(datetime(2025, 12, 31))
```

## Key Differences from VBA

1. **Database Connection**: Uses mysql-connector-python instead of ADODB
2. **Excel Reading**: Uses pandas instead of COM automation
3. **Web Scraping**: Uses requests and BeautifulSoup instead of XMLHTTP
4. **Error Handling**: Python exceptions instead of VBA On Error Resume Next
5. **Date Handling**: Python datetime instead of VBA Date functions

## Held At Location Codes

The following location codes are recognized:
- `CollegeAdv` - CollegeAdvantage 529 Plan
- `FidelityInv` - Fidelity Individual Account
- `FidelityIRA` - Fidelity Rollover IRA
- `FidelityRoth` - Fidelity Roth IRA
- `Fidelity401k` - Fidelity 401(k)
- `TRPInv` - T. Rowe Price Individual
- `TRPRoth` - T. Rowe Price Roth IRA
- `TRPRollover` - T. Rowe Price Rollover IRA
- `Vanguard` - Vanguard Brokerage
- `WellsFargo401k` - Wells Fargo 401(k)
- `Etrade`, `Robinhood`, `Ameritrade`, `TradeStation`

## Troubleshooting

### Connection Issues
```bash
# Test MySQL connection
mysql -h localhost -P 3306 -u root -psa123

# Check if MySQL is running
ps aux | grep mysql
```

### Module Import Errors
```bash
# Ensure all dependencies are installed
pip install -r requirements.txt

# Run from the correct directory
cd /Users/Sdoshi/development/samir/docs/Investments/US
python process_assets.py
```

### Excel File Not Found
Ensure Asset.xls is in the same directory or provide full path:
```bash
python process_assets.py --file /path/to/Asset.xls
```

## Backup to Google Drive

Upload SQL backups and Excel file to Google Drive:
```bash
python upload_to_gdrive.py
```

This will:
- Upload all `asset*.sql` files in the current directory
- Upload `Asset.xlsx` file
- Save to Google Drive folder: https://drive.google.com/drive/u/1/folders/1143-kZ1KCLy8yQsL8Dkms1mowIndLfRu

**First-time setup:**
1. Install Google Drive API dependencies: `pip install -r requirements.txt`
2. Download `credentials.json` from Google Cloud Console (see Installation step 3)
3. Run the script - it will open a browser for OAuth authentication
4. Grant permissions - a `token.pickle` file will be saved for future runs

**Subsequent runs:**
Just run `python upload_to_gdrive.py` - authentication is cached in `token.pickle`

## Notes

- The script automatically handles market holidays and weekends when calculating gains
- Transactions are used to ensure data consistency
- Failed operations are rolled back automatically
- Yahoo Finance data is fetched via CSV API for reliability
- Morningstar is used as fallback when Yahoo data is unavailable
