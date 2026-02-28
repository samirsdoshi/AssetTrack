"""
Main Asset Processing Script
Reads Asset.xls and processes investment data
"""

import pandas as pd
from datetime import datetime
from asset_processor import AssetDatabase, AssetAllocator, GainCalculator, TemplateManager
from utils import filter_ticker, get_held_at, empty_to_default, clean_up
import sys
import os
from dotenv import load_dotenv
import msoffcrypto
import io


class AssetProcessor:
    """Main class to process Asset.xls file"""
    
    def __init__(self, excel_file: str = 'Asset.xls'):
        self.excel_file = excel_file
        # Load environment variables
        load_dotenv()
        self.excel_password = os.getenv('password', '')
        
        # Auto-convert .xls to .xlsx if it's an old format
        if excel_file.endswith('.xls') and not excel_file.endswith('.xlsx'):
            xlsx_file = excel_file + 'x'  # Asset.xls -> Asset.xlsx
            if not os.path.exists(xlsx_file):
                print(f"Converting {excel_file} to modern .xlsx format...")
                self.convert_xls_to_xlsx(excel_file, xlsx_file)
                self.excel_file = xlsx_file
                print(f"Using converted file: {xlsx_file}")
            else:
                print(f"Using existing .xlsx file: {xlsx_file}")
                self.excel_file = xlsx_file
        
        self.db = AssetDatabase(
            host='localhost',
            port=3306,
            user='root',
            password='sa123',
            database='asset'
        )
        self.allocator = AssetAllocator(self.db)
        self.gain_calculator = GainCalculator(self.db)
        self.template_manager = TemplateManager(self.db)
    
    def convert_xls_to_xlsx(self, xls_file: str, xlsx_file: str):
        """
        Convert .xls file to .xlsx format
        
        Args:
            xls_file: Source .xls file path
            xlsx_file: Destination .xlsx file path
        """
        try:
            from openpyxl import Workbook
            
            # Decrypt if needed
            if self.excel_password:
                try:
                    decrypted = io.BytesIO()
                    with open(xls_file, 'rb') as f:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password=self.excel_password)
                        office_file.decrypt(decrypted)
                    decrypted.seek(0)
                    
                    # Read all sheets from .xls
                    xls_data = pd.read_excel(decrypted, sheet_name=None, header=None)
                except:
                    xls_data = pd.read_excel(xls_file, sheet_name=None, header=None)
            else:
                xls_data = pd.read_excel(xls_file, sheet_name=None, header=None)
            
            # Write to .xlsx
            with pd.ExcelWriter(xlsx_file, engine='openpyxl') as writer:
                for sheet_name, df in xls_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            
            print(f"Successfully converted {xls_file} to {xlsx_file}")
            
        except Exception as e:
            print(f"Error converting file: {e}")
            raise
    
    def update_assetalloc_dates(self, prevdate: str = None, currdate: str = None):
        """Update Assetalloc sheet B1 (prevdate) and C1 (currdate) with direct values"""
        if not prevdate and not currdate:
            return
        
        print(f"\nUpdating Assetalloc sheet dates...")
        try:
            from openpyxl import load_workbook
            import msoffcrypto
            import io
            
            # Load workbook
            wb = load_workbook(self.excel_file, data_only=False)
            
            # Check available sheet names (case-insensitive search)
            sheet_name = None
            for name in wb.sheetnames:
                if name.lower() == 'assetalloc':
                    sheet_name = name
                    break
            
            if not sheet_name:
                print(f"  Warning: Assetalloc sheet not found in {self.excel_file}")
                print(f"  Available sheets: {', '.join(wb.sheetnames)}")
                return
            
            ws = wb[sheet_name]
            # Set direct values (this will replace any formulas)
            if prevdate:
                ws['B1'].value = prevdate
                print(f"  Set B1 (prevdate) = {prevdate}")
            if currdate:
                ws['C1'].value = currdate
                print(f"  Set C1 (currdate) = {currdate}")
            
            wb.save(self.excel_file)
            print(f"  Assetalloc sheet updated successfully")
        except Exception as e:
            print(f"  Error: Could not update Assetalloc dates: {e}")
            import traceback
            traceback.print_exc()
    
    def normalize_full_view(self, sheet_name: str = 'fidfullview', output_file: str = None) -> pd.DataFrame:
        """
        Normalize and aggregate fidfullview data by account
        Converted from VBA normalizefullview() function
        
        Args:
            sheet_name: Name of the sheet to process (default: 'fidfullview')
            output_file: Optional output Excel file to save results
            
        Returns:
            DataFrame with normalized data
        """
        print(f"Normalizing full view from sheet '{sheet_name}'...")
        
        # Read the fidfullview sheet (has header row)
        df = self.decrypt_and_read_excel(sheet_name, header=0)
        
        # Account mapping based on Column B values
        account_mapping = {
            'Individual - TOD': 'FidelityInv',
            'Rollover IRA': 'FidelityIRA',
            'Samir S Doshi - Brokerage Account - 10498558': 'Vanguard',
            'Samir S Doshi - Rollover IRA': 'Vanguard IRA'
        }
        
        account_data = {}  # Dictionary to store aggregated amounts by account
        ticker_data = {}  # Dictionary to aggregate by account_ticker combination
        
        # Process each row
        for idx, row in df.iterrows():
            # Get values from named columns (B='Account Name', C='Symbol', H='Current Value')
            account_name = row.get('Account Name') or row.iloc[1] if len(row) > 1 else None
            symbol = row.get('Symbol') or row.iloc[2] if len(row) > 2 else None
            value = row.get('Current Value') or row.iloc[7] if len(row) > 7 else None
            
            # Skip if any required field is missing or empty
            if pd.isna(account_name) or pd.isna(symbol) or pd.isna(value):
                continue
            
            account_name = str(account_name).strip()
            symbol = str(symbol).strip()
            
            if account_name == '' or symbol == '':
                continue
            
            # Strip trailing ** (e.g. FCASH** -> FCASH, FDRXX** -> FDRXX)
            symbol = symbol.rstrip('*')
            
            # Map specific symbols to Cash
            if symbol in ['FDRXX', 'FCASH', 'Pending activity', 'VMRXX', 'VMFXX']:
                symbol = 'Cash'
            
            # Clean up value if it's a string
            if isinstance(value, str):
                value = value.replace('$', '').replace(',', '')
                try:
                    value = float(value)
                except:
                    continue
            
            # Skip zero values
            if value == 0:
                continue
            
            # Map account name to prefix
            account_prefix = account_mapping.get(account_name, None)
            
            # Skip if account not in mapping (will be handled by Trow sheet)
            if not account_prefix:
                continue
            
            # Create account_ticker key
            account_ticker = f"{account_prefix}_{symbol}"
            
            # Aggregate by account_ticker (sum duplicates)
            if account_ticker not in ticker_data:
                ticker_data[account_ticker] = 0
            ticker_data[account_ticker] += value
            
            # Aggregate by account
            if account_prefix not in account_data:
                account_data[account_prefix] = 0
            account_data[account_prefix] += value
        
        # Convert ticker_data dictionary to list of results
        results = []
        for account_ticker, amount in ticker_data.items():
            results.append({
                'account_ticker': account_ticker,
                'amount': amount
            })
        
        # Create results DataFrame
        results_df = pd.DataFrame(results)
        
        # Remove entries with 0 value (shouldn't happen but safety check)
        if len(results_df) > 0:
            results_df = results_df[results_df['amount'] != 0]
        
        # Read Trow sheet and manually compute columns N and O from source columns
        # Column N formula: =CONCAT(M,"_",C) - concatenates column M (account) and C (ticker)
        # Column O formula: =I - market value
        # Also read pre-formatted entries directly from N and O
        try:
            print("Reading Trow sheet data...")
            trow_df = self.decrypt_and_read_excel('Trow', header=None)
            
            # First pass: Extract and compute values from columns C (ticker), I (amount), M (account)
            # Skip header row (index 0)
            added_count = 0
            for idx, row in trow_df.iterrows():
                if idx == 0:  # Skip header row
                    continue
                    
                account = row.get(12)  # Column M (index 12)
                ticker = row.get(2)    # Column C (index 2)
                amount = row.get(8)    # Column I (index 8)
                
                # Skip if any required field is empty
                if pd.isna(account) or pd.isna(ticker) or pd.isna(amount):
                    continue
                
                # Skip if account or ticker is empty string
                if str(account).strip() == '' or str(ticker).strip() == '':
                    continue
                
                # Clean up amount if it's a string
                if isinstance(amount, str):
                    amount = amount.replace('$', '').replace(',', '')
                    try:
                        amount = float(amount)
                    except:
                        continue
                
                # Skip if amount is 0
                if amount == 0:
                    continue
                
                # Construct account_ticker (equivalent to column N formula)
                account_ticker = f"{account}_{ticker}"
                
                # Add to results
                results_df = pd.concat([results_df, pd.DataFrame([{
                    'account_ticker': account_ticker,
                    'amount': amount
                }])], ignore_index=True)
                added_count += 1
            
            # Second pass: Read pre-formatted entries from columns N and O
            # This handles rows where N already has "Account_Ticker" format
            # Need to evaluate formulas in column O by reading from data_only=False and calculating
            from openpyxl import load_workbook
            try:
                wb_formulas = load_workbook(self.excel_file, data_only=False)
                wb_data = load_workbook(self.excel_file, data_only=True)
                ws_f = wb_formulas['Trow']
                ws_d = wb_data['Trow']
                
                for row_idx in range(1, 100):
                    n_val = ws_d[f'N{row_idx}'].value
                    o_formula = ws_f[f'O{row_idx}'].value
                    
                    # Skip if N is empty or doesn't contain "_"
                    if not n_val or '_' not in str(n_val):
                        continue
                    
                    # Try to evaluate formula in O
                    amount = None
                    if o_formula and str(o_formula).startswith('='):
                        # Formula exists but not evaluated - try to manually evaluate simple cases
                        formula_str = str(o_formula)[1:]  # Remove '='
                        
                        # Handle simple cell references like "A45"
                        if formula_str.startswith('A') and formula_str[1:].isdigit():
                            ref_row = int(formula_str[1:])
                            amount = ws_d[f'A{ref_row}'].value
                        # Handle addition like "A37+A41"
                        elif '+' in formula_str:
                            parts = formula_str.split('+')
                            total = 0
                            for part in parts:
                                part = part.strip()
                                if part.startswith('A') and part[1:].isdigit():
                                    ref_row = int(part[1:])
                                    val = ws_d[f'A{ref_row}'].value
                                    if val:
                                        total += float(val)
                            amount = total if total > 0 else None
                    elif o_formula:
                        # Direct value
                        amount = o_formula
                    
                    if amount and amount != 0:
                        # Clean up amount if needed
                        if isinstance(amount, str):
                            amount = amount.replace('$', '').replace(',', '')
                            try:
                                amount = float(amount)
                            except:
                                continue
                        
                        results_df = pd.concat([results_df, pd.DataFrame([{
                            'account_ticker': str(n_val),
                            'amount': amount
                        }])], ignore_index=True)
                        added_count += 1
            except Exception as e:
                print(f"Could not read pre-formatted N/O columns: {e}")
            
            print(f"Added {added_count} entries from Trow sheet")
        except Exception as e:
            print(f"Warning: Could not read Trow sheet: {e}")
        
        # Sort final results
        results_df = results_df.sort_values(by='account_ticker', ascending=True)
        
        # Create summary DataFrame
        summary_df = pd.DataFrame([
            {'account': account, 'total': total}
            for account, total in account_data.items()
        ])
        
        print(f"\nNormalized {len(results_df)} fund entries across {len(summary_df)} accounts")
        print("\nAccount Totals:")
        for _, row in summary_df.iterrows():
            print(f"  {row['account']}: ${row['total']:,.2f}")
        
        # Write results back to the fullview sheet (columns K, L, N, O)
        # Note: We read from fldfullview sheet but write results to fullview sheet for compatibility
        try:
            from openpyxl import load_workbook
            from zipfile import ZipFile
            
            print("\nWriting results back to fullview sheet...")
            
            # Handle .xlsx format with openpyxl
            # Decrypt and load workbook if needed
            if self.excel_password:
                try:
                    decrypted = io.BytesIO()
                    with open(self.excel_file, 'rb') as f:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password=self.excel_password)
                        office_file.decrypt(decrypted)
                    decrypted.seek(0)
                    
                    # Remove external references from the zip
                    cleaned = io.BytesIO()
                    with ZipFile(decrypted, 'r') as zin:
                        with ZipFile(cleaned, 'w') as zout:
                            for item in zin.infolist():
                                if 'externalLink' not in item.filename and 'externalReferences' not in item.filename:
                                    data = zin.read(item.filename)
                                    zout.writestr(item, data)
                    cleaned.seek(0)
                    wb = load_workbook(cleaned, data_only=False)
                except:
                    wb = load_workbook(self.excel_file, data_only=False)
            else:
                wb = load_workbook(self.excel_file, data_only=False)
            
            # Always write results to 'fullview' sheet (not the source sheet)
            ws = wb['fullview']
            
            # IMPORTANT: Read stock account data from N, O, P BEFORE clearing columns
            # Need to evaluate formulas in column P, so load both formula and data versions
            wb_data = load_workbook(self.excel_file, data_only=True)
            ws_data = wb_data['fullview']
            
            stock_account_entries = []
            print("\nReading stock account entries from columns N, O, P...")
            for row_idx in range(1, 1000):
                account = ws[f'N{row_idx}'].value
                ticker = ws[f'O{row_idx}'].value
                amount_formula = ws[f'P{row_idx}'].value
                amount_value = ws_data[f'P{row_idx}'].value
                
                # Skip if account or ticker is empty or if ticker is "Total"
                if not account or not ticker or ticker == 'Total':
                    continue
                
                # Get amount - prefer calculated value, fallback to formula evaluation
                amount = amount_value
                
                # If amount is still None/empty but there's a formula, try to evaluate it
                if not amount and amount_formula and str(amount_formula).startswith('='):
                    formula_str = str(amount_formula)[1:]  # Remove '='
                    # Handle simple cases like "P3-P2"
                    if '-' in formula_str:
                        parts = formula_str.split('-')
                        if len(parts) == 2:
                            val1 = val2 = 0
                            p1 = parts[0].strip()
                            p2 = parts[1].strip()
                            if p1.startswith('P') and p1[1:].isdigit():
                                val1 = ws_data[p1].value or 0
                            if p2.startswith('P') and p2[1:].isdigit():
                                val2 = ws_data[p2].value or 0
                            amount = val1 - val2
                
                # Skip if amount is still empty or 0
                if not amount or amount == 0:
                    continue
                
                # Store the entry
                stock_account_entries.append({
                    'account': account,
                    'ticker': ticker,
                    'amount': amount
                })
            
            print(f"Found {len(stock_account_entries)} stock account entries")
            
            # Clear ONLY columns K and L (rows 1-2000) - DO NOT touch N, O, P
            for row in range(1, 2001):
                ws[f'K{row}'].value = None
                ws[f'L{row}'].value = None
            
            # Write fund results to columns K and L
            j = 1
            for _, row in results_df.iterrows():
                ws[f'K{j}'].value = row['account_ticker']
                ws[f'L{j}'].value = row['amount']
                j += 1
            
            # Append stock account entries (read earlier from N, O, P) to K and L
            print(f"\nAppending {len(stock_account_entries)} stock account entries to columns K and L...")
            stock_appended = 0
            
            for entry in stock_account_entries:
                account = entry['account']
                ticker = entry['ticker']
                amount = entry['amount']
                
                # Create account_ticker format: Account_Ticker
                account_ticker = f"{account}_{ticker}"
                
                # Append to columns K and L
                ws[f'K{j}'].value = account_ticker
                ws[f'L{j}'].value = amount
                j += 1
                stock_appended += 1
            
            print(f"Appended {stock_appended} stock account entries")
            
            # Save the workbook
            wb.save(self.excel_file)
            print(f"Results written to {self.excel_file} in columns K and L")
            
        except Exception as e:
            print(f"Warning: Could not write back to Excel file: {e}")
            print("Results are still available in the returned DataFrames")
        
        # Save to separate output file if specified
        if output_file:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                results_df.to_excel(writer, sheet_name='Details', index=False)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            print(f"\nResults also saved to {output_file}")
        
        return results_df, summary_df

    
    def decrypt_and_read_excel(self, sheet_name: str, header=0) -> pd.DataFrame:
        """
        Decrypt password-protected Excel file and read sheet
        
        Args:
            sheet_name: Name of the sheet to read
            header: Row to use as column names (default: 0)
            
        Returns:
            DataFrame with data
        """
        try:
            # Try to read directly first (unprotected file)
            # Use openpyxl with data_only=True to read calculated formula values
            from openpyxl import load_workbook as openpyxl_load
            wb = openpyxl_load(self.excel_file, data_only=True)
            
            # Check if sheet exists
            if sheet_name not in wb.sheetnames:
                available_sheets = ', '.join(wb.sheetnames)
                raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {available_sheets}")
            
            ws = wb[sheet_name]
            
            # Convert to DataFrame manually
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(row)
            
            if header is not None and len(data) > header:
                df = pd.DataFrame(data[header+1:], columns=data[header])
            else:
                df = pd.DataFrame(data)
            
            return df
        except Exception as e:
            # If failed, try with password decryption
            if self.excel_password:
                try:
                    print(f"Attempting to decrypt {self.excel_file}...")
                    decrypted = io.BytesIO()
                    with open(self.excel_file, 'rb') as f:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password=self.excel_password)
                        office_file.decrypt(decrypted)
                    
                    decrypted.seek(0)
                    
                    # Remove external references from the zip to avoid openpyxl errors
                    from zipfile import ZipFile
                    cleaned = io.BytesIO()
                    with ZipFile(decrypted, 'r') as zin:
                        with ZipFile(cleaned, 'w') as zout:
                            for item in zin.infolist():
                                # Skip external links and related files
                                if 'externalLink' not in item.filename and 'externalReferences' not in item.filename:
                                    data = zin.read(item.filename)
                                    zout.writestr(item, data)
                    
                    cleaned.seek(0)
                    
                    # Use openpyxl to load cleaned file
                    from openpyxl import load_workbook as openpyxl_load
                    wb = openpyxl_load(cleaned, data_only=True)
                    ws = wb[sheet_name]
                    
                    # Convert to DataFrame manually
                    data = []
                    for row in ws.iter_rows(values_only=True):
                        data.append(row)
                    
                    if header is not None and len(data) > header:
                        df = pd.DataFrame(data[header+1:], columns=data[header])
                    else:
                        df = pd.DataFrame(data)
                    
                    print(f"Successfully decrypted and read sheet '{sheet_name}'")
                    return df
                except Exception as decrypt_error:
                    print(f"Error decrypting Excel file: {decrypt_error}")
                    raise
            else:
                print(f"Error reading Excel file: {e}")
                print("No password found in .env file")
                raise
    
    def read_asset_reference_sheet(self, sheet_name: str = 'fullview') -> pd.DataFrame:
        """
        Read asset data directly from fullview sheet columns K and L
        Column K contains account_ticker (e.g., "FidelityInv_FXAIX", "Etrade_Cash")
        Column L contains amount
        
        Args:
            sheet_name: Name of the sheet to read (default: fullview)
            
        Returns:
            DataFrame with columns: Ticker (account_ticker), Amount, HeldAt
        """
        try:
            # Read fullview sheet without header
            df = self.decrypt_and_read_excel(sheet_name, header=None)
            
            # Extract data from columns K (index 10) and L (index 11)
            all_data = []
            
            for idx, row in df.iterrows():
                # Get values from columns K and L
                account_ticker_k = row.get(10)  # Column K
                amount_l = row.get(11)  # Column L
                
                if pd.notna(account_ticker_k) and pd.notna(amount_l) and amount_l != 0:
                    # Clean up amount
                    if isinstance(amount_l, str):
                        amount_l = amount_l.replace('$', '').replace(',', '')
                        try:
                            amount_l = float(amount_l)
                        except:
                            continue
                    
                    # Split account_ticker to get account and ticker
                    account_ticker_str = str(account_ticker_k)
                    if '_' in account_ticker_str:
                        held_at, ticker = account_ticker_str.split('_', 1)
                        all_data.append({
                            'Ticker': ticker,
                            'Amount': amount_l,
                            'HeldAt': held_at
                        })
            
            result_df = pd.DataFrame(all_data)
            
            print(f"Successfully read {len(result_df)} entries from '{sheet_name}' columns K and L")
            return result_df
        except FileNotFoundError:
            print(f"Error: File '{self.excel_file}' not found")
            sys.exit(1)
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            sys.exit(1)
    
    def process_asset_allocation(self, df: pd.DataFrame, as_of_date: datetime, held_at_column: str = 'HeldAt'):
        """
        Process asset allocation from DataFrame
        
        Args:
            df: DataFrame with asset data
            as_of_date: Date for the asset allocation
            held_at_column: Column name containing held at information
        """
        self.db.open_db()
        
        processed_count = 0
        error_count = 0
        error_details = []  # Track detailed error information
        
        # Assuming columns: Ticker (A), Symbol (B), Amount (E), HeldAt (J)
        # Adjust column names based on actual Excel structure
        
        # Track current stock account being processed
        current_stock_account = None
        stock_account_cash = {}  # Track cash amounts for calculating stock value
        stock_accounts = ['Etrade', 'Ameritrade', 'TradeStation', 'Robinhood']
        
        for index, row in df.iterrows():
            try:
                # Get ticker
                ticker = row.get('Ticker', row.get('Symbol', ''))
                if pd.isna(ticker):
                    continue
                
                ticker = str(ticker).strip()
                
                # Check for end marker
                if ticker == "ENDOFPORTFOLIO":
                    break
                
                # Check if this is a stock account header
                if ticker in stock_accounts:
                    current_stock_account = ticker
                    continue
                
                # For stock accounts, process both Cash and Stock rows
                if current_stock_account:
                    if ticker == "Cash":
                        # Process Cash row
                        amount = row.get('Amount', row.get('Value', 0))
                        if pd.isna(amount):
                            amount = 0
                        else:
                            # Clean up amount
                            if isinstance(amount, str):
                                amount = clean_up(amount)
                                amount = float(amount) if amount else 0
                            else:
                                amount = float(amount)
                        
                        # Store cash amount for this account
                        stock_account_cash[current_stock_account] = amount
                        
                        if amount > 0:
                            # Process as Cash
                            held_at = current_stock_account
                            
                            # Get asset ID for Cash
                            query = """
                                SELECT assetid FROM asset 
                                WHERE ticker = %s OR assetname = %s
                            """
                            results = self.db.execute_query(query, ("Cash", "Cash"))
                            
                            if not results:
                                print(f"Warning: Asset not found for ticker Cash")
                                error_count += 1
                            else:
                                asset_id = results[0]['assetid']
                                print(f"Processing: Cash - ${amount:,.2f} at {held_at}")
                                self.allocator.allocate_asset_ref(asset_id, as_of_date, amount, held_at)
                                processed_count += 1
                        continue
                        
                    elif ticker == "Stock":
                        # Skip Stock row (it's a formula, we'll calculate from Total - Cash)
                        continue
                        
                    elif ticker == "Total":
                        # Process Total row - calculate Stock value
                        total_amount = row.get('Amount', row.get('Value', 0))
                        if pd.isna(total_amount):
                            total_amount = 0
                        else:
                            # Clean up amount
                            if isinstance(total_amount, str):
                                total_amount = clean_up(total_amount)
                                total_amount = float(total_amount) if total_amount else 0
                            else:
                                total_amount = float(total_amount)
                        
                        # Calculate Stock = Total - Cash
                        cash_amount = stock_account_cash.get(current_stock_account, 0)
                        stock_amount = total_amount - cash_amount
                        
                        if stock_amount > 0:
                            # Process as Stock
                            held_at = current_stock_account
                            
                            # Get asset ID for Stock
                            query = """
                                SELECT assetid FROM asset 
                                WHERE ticker = %s OR assetname = %s
                            """
                            results = self.db.execute_query(query, ("Stock", "Stock"))
                            
                            if not results:
                                print(f"Warning: Asset not found for ticker Stock")
                                error_count += 1
                                error_details.append(f"Asset not found: Stock (for {held_at})")
                            else:
                                asset_id = results[0]['assetid']
                                print(f"Processing: Stock - ${stock_amount:,.2f} at {held_at}")
                                self.allocator.allocate_asset_ref(asset_id, as_of_date, stock_amount, held_at)
                                processed_count += 1
                        
                        # Reset after Total row
                        current_stock_account = None
                        continue
                    else:
                        # Skip other rows
                        continue
                
                # Skip total/summary rows for regular accounts
                if 'Total' in ticker or 'total' in ticker:
                    continue
                
                # Filter invalid tickers
                ticker = filter_ticker(ticker)
                if not ticker:
                    continue
                
                # Apply fund mapping (e.g., VMRXX -> VMMXX)
                fund_mapping = {
                    'VMRXX': 'VMMXX'
                }
                if ticker in fund_mapping:
                    original_ticker = ticker
                    ticker = fund_mapping[ticker]
                    print(f"Mapped {original_ticker} -> {ticker}")
                
                # Get held at location
                held_at = row.get(held_at_column, row.get('HeldAt', ''))
                if pd.isna(held_at):
                    held_at = ''
                else:
                    held_at = str(held_at).strip()
                
                if not held_at:
                    print(f"Warning: No 'HeldAt' location for ticker {ticker}")
                    continue
                
                # Get amount
                amount = row.get('Amount', row.get('Value', 0))
                if pd.isna(amount):
                    amount = 0
                else:
                    # Clean up amount (remove $ and ,)
                    if isinstance(amount, str):
                        amount = clean_up(amount)
                        amount = float(amount) if amount else 0
                    else:
                        amount = float(amount)
                
                if amount == 0:
                    continue
                
                # Get asset ID from database
                query = """
                    SELECT assetid FROM asset 
                    WHERE ticker = %s OR assetname = %s
                """
                results = self.db.execute_query(query, (ticker, ticker))
                
                if not results:
                    print(f"Warning: Asset not found for ticker {ticker}")
                    error_count += 1
                    error_details.append(f"Asset not found: {ticker}")
                    continue
                
                asset_id = results[0]['assetid']
                
                # Allocate the asset
                print(f"Processing: {ticker} - ${amount:,.2f} at {held_at}")
                self.allocator.allocate_asset_ref(asset_id, as_of_date, amount, held_at)
                processed_count += 1
                
            except Exception as e:
                print(f"Error processing row {index}: {e}")
                error_count += 1
                error_details.append(f"Row {index}: {e}")
                continue
        
        self.db.close_db()
        
        print(f"\n=== Processing Complete ===")
        print(f"Processed: {processed_count} assets")
        print(f"Errors: {error_count}")
        
        if error_count > 0 and error_details:
            print(f"\nError details:")
            for error in error_details:
                print(f"  - {error}")
    
    def compare_dates_report(self, currdate: datetime, datetocompare: datetime, threshold_percent: float = 5.0, show_all: bool = False):
        """
        Compare asset values between two dates and report significant changes
        
        Args:
            currdate: Current date to compare
            datetocompare: Previous date to compare against
            threshold_percent: Minimum percentage change to report (default: 5%)
            show_all: If True, show all changes regardless of threshold (default: False)
        """
        print(f"\n{'='*60}")
        print(f"Significant Changes Report")
        print(f"{'='*60}")
        print(f"Comparing: {datetocompare.strftime('%Y-%m-%d')} vs {currdate.strftime('%Y-%m-%d')}")
        if show_all:
            print(f"Showing: ALL changes\n")
        else:
            print(f"Threshold: {threshold_percent}% change or more\n")
        
        self.db.open_db()
        
        # Query to get asset values by account and ticker for both dates
        query = """
            SELECT 
                ai.asofdate,
                ai.heldat as account,
                a.ticker,
                SUM(aia.amount) as total_amount
            FROM assetinv ai
            JOIN assetinvalloc aia ON ai.assetinvid = aia.assetinvid
            JOIN asset a ON ai.assetid = a.assetid
            WHERE ai.asofdate IN (%s, %s)
            GROUP BY ai.asofdate, ai.heldat, a.ticker
            ORDER BY ai.heldat, a.ticker, ai.asofdate
        """
        
        results = self.db.execute_query(query, 
            (datetocompare.strftime('%Y-%m-%d'), currdate.strftime('%Y-%m-%d')))
        
        # Organize data by account_ticker and also track account totals
        data = {}
        account_totals = {}
        
        for row in results:
            date = row['asofdate']
            account = row['account']
            ticker = row['ticker']
            amount = float(row['total_amount'])
            
            key = f"{account}_{ticker}"
            if key not in data:
                data[key] = {}
            data[key][date] = amount
            
            # Track account totals
            if account not in account_totals:
                account_totals[account] = {}
            if date not in account_totals[account]:
                account_totals[account][date] = 0
            account_totals[account][date] += amount
        
        # Calculate changes
        changes = []
        # Convert to date objects for comparison
        prev_date = datetocompare.date() if hasattr(datetocompare, 'date') else datetocompare
        curr_date = currdate.date() if hasattr(currdate, 'date') else currdate
        
        for key, dates in data.items():
            # dates dict has date/datetime keys from DB - try both formats
            prev_amount = 0
            curr_amount = 0
            for date_key, amount in dates.items():
                # Convert date_key to date for comparison
                dk = date_key.date() if hasattr(date_key, 'date') else date_key
                if dk == prev_date:
                    prev_amount = amount
                if dk == curr_date:
                    curr_amount = amount
            
            account_name = key.split('_')[0]  # Extract account name from key
            
            if prev_amount == 0 and curr_amount > 0:
                # New position
                changes.append({
                    'key': key,
                    'account': account_name,
                    'prev': prev_amount,
                    'curr': curr_amount,
                    'change': curr_amount,
                    'pct_change': 100.0,
                    'type': 'NEW'
                })
            elif prev_amount != 0 and curr_amount == 0:
                # Closed position
                changes.append({
                    'key': key,
                    'account': account_name,
                    'prev': prev_amount,
                    'curr': curr_amount,
                    'change': -prev_amount,
                    'pct_change': -100.0,
                    'type': 'CLOSED'
                })
            elif prev_amount != 0:
                # Changed position (prev may be negative, e.g. margin balance)
                change_amount = curr_amount - prev_amount
                pct_change = (change_amount / abs(prev_amount)) * 100
                
                # Add to list if show_all or if above threshold
                if show_all or abs(pct_change) >= threshold_percent:
                    changes.append({
                        'key': key,
                        'account': account_name,
                        'prev': prev_amount,
                        'curr': curr_amount,
                        'change': change_amount,
                        'pct_change': pct_change,
                        'type': 'CHANGE'
                    })
        
        # Sort by account first, then by absolute percentage change (descending)
        changes.sort(key=lambda x: (x['account'], -abs(x['pct_change'])))
        
        # Calculate account totals and changes
        account_summary = {}
        for account, dates in account_totals.items():
            prev_total = 0
            curr_total = 0
            for date_key, amount in dates.items():
                dk = date_key.date() if hasattr(date_key, 'date') else date_key
                if dk == prev_date:
                    prev_total = amount
                if dk == curr_date:
                    curr_total = amount
            change_total = curr_total - prev_total
            pct_change_total = (change_total / prev_total * 100) if prev_total > 0 else 0
            
            account_summary[account] = {
                'prev': prev_total,
                'curr': curr_total,
                'change': change_total,
                'pct_change': pct_change_total
            }
        
        # Print account summary first
        if account_summary:
            print(f"{'='*60}")
            print("ACCOUNT TOTALS")
            print(f"{'='*60}")
            print(f"{'Account':<30} {'Previous':>12} {'Current':>12} {'Change':>12} {'% Change':>10}")
            print(f"{'-'*30} {'-'*12} {'-'*12} {'-'*12} {'-'*10}")
            
            for account in sorted(account_summary.keys()):
                summary = account_summary[account]
                print(f"{account:<30} ${summary['prev']:>10,.2f} ${summary['curr']:>10,.2f} "
                      f"${summary['change']:>10,.2f} {summary['pct_change']:>9.2f}%")
            print()
        
        # Print individual changes
        if changes:
            print(f"{'='*60}")
            print("INDIVIDUAL HOLDINGS")
            print(f"{'='*60}")
            
            current_account = None
            for i, item in enumerate(changes):
                # Print account header when it changes
                if item['account'] != current_account:
                    # Print summary line for the previous account before switching
                    if current_account is not None:
                        summary = account_summary.get(current_account, {})
                        if summary:
                            print(f"  {'-'*38} {'-'*12} {'-'*12} {'-'*12} {'-'*10}")
                            sprev = summary['prev']
                            scurr = summary['curr']
                            schange = summary['change']
                            spct = summary['pct_change']
                            print(f"  {'TOTAL':<38} ${sprev:>10,.2f} ${scurr:>10,.2f} ${schange:>10,.2f} {spct:>9.2f}%")
                        print()  # Blank line between accounts
                    current_account = item['account']
                    print(f"\n{current_account}:")
                    print(f"{'  Ticker':<38} {'Previous':>12} {'Current':>12} {'Change':>12} {'% Change':>10}")
                    print(f"  {'-'*38} {'-'*12} {'-'*12} {'-'*12} {'-'*10}")
                
                # Extract ticker from key (remove account prefix)
                ticker = '_'.join(item['key'].split('_')[1:])
                prev = item['prev']
                curr = item['curr']
                change = item['change']
                pct = item['pct_change']
                
                if item['type'] == 'NEW':
                    print(f"  {ticker:<38} {'NEW':>12} ${curr:>10,.2f} ${change:>10,.2f} {'NEW':>10}")
                elif item['type'] == 'CLOSED':
                    print(f"  {ticker:<38} ${prev:>10,.2f} {'CLOSED':>12} ${change:>10,.2f} {'CLOSED':>10}")
                else:
                    print(f"  {ticker:<38} ${prev:>10,.2f} ${curr:>10,.2f} ${change:>10,.2f} {pct:>9.2f}%")
                
                # Print summary line after the last item in the last account
                if i == len(changes) - 1 and current_account is not None:
                    summary = account_summary.get(current_account, {})
                    if summary:
                        print(f"  {'-'*38} {'-'*12} {'-'*12} {'-'*12} {'-'*10}")
                        sprev = summary['prev']
                        scurr = summary['curr']
                        schange = summary['change']
                        spct = summary['pct_change']
                        print(f"  {'TOTAL':<38} ${sprev:>10,.2f} ${scurr:>10,.2f} ${schange:>10,.2f} {spct:>9.2f}%")
            
            print(f"\nTotal holdings shown: {len(changes)}")
        else:
            print("No changes found.")
        
        print(f"{'='*60}\n")
        
        self.db.close_db()
    
    def delete_existing_data(self, as_of_date: datetime):
        """
        Delete existing asset data for a given date
        
        Args:
            as_of_date: Date to delete data for
        """
        print(f"Deleting existing data for {as_of_date.strftime('%Y-%m-%d')}...")
        self.allocator.delete_asset_info(as_of_date)
        print("Deletion complete")
    
    def calculate_gains(self, as_of_date: datetime):
        """
        Calculate gains for all assets
        
        Args:
            as_of_date: Date to calculate gains for
        """
        print(f"Calculating gains for {as_of_date.strftime('%Y-%m-%d')}...")
        self.gain_calculator.calculate_gains(as_of_date)
        print("Gain calculation complete")
    
    def refresh_dataconn(self, currdate: datetime = None, datetocompare: datetime = None):
        """
        Refresh dataconn sheet with data from database queries
        
        Args:
            currdate: Current date to update in wkdates table
            datetocompare: Date to compare to update in wkdates table
        """
        print("=" * 60)
        print("Refreshing Dataconn Sheet")
        print("=" * 60)
        
        # Update Assetalloc sheet dates if provided
        if currdate and datetocompare:
            self.update_assetalloc_dates(
                prevdate=datetocompare.strftime('%Y-%m-%d'),
                currdate=currdate.strftime('%Y-%m-%d')
            )
        
        from openpyxl import load_workbook
        
        # Open database connection
        self.db.open_db()
        
        try:
            # Update wkdates table if dates provided
            if currdate or datetocompare:
                print("\nUpdating wkdates table...")
                if currdate:
                    update_query = "UPDATE wkdates SET currdate = %s"
                    self.db.cursor.execute(update_query, (currdate,))
                    print(f"  Updated currdate to {currdate.strftime('%Y-%m-%d')}")
                
                if datetocompare:
                    update_query = "UPDATE wkdates SET datetocompare = %s"
                    self.db.cursor.execute(update_query, (datetocompare,))
                    print(f"  Updated datetocompare to {datetocompare.strftime('%Y-%m-%d')}")
                
                self.db.connection.commit()
            
            # Validate that data exists for both dates
            if currdate and datetocompare:
                print("\nValidating data exists for both dates...")
                
                # Check what dates have data in the views
                date_check_query = """
                    SELECT DISTINCT asofdate FROM totalbyalloctypedate
                    ORDER BY asofdate
                """
                available_dates = self.db.execute_query(date_check_query)
                available_date_strs = set([d['asofdate'].strftime('%Y-%m-%d') if isinstance(d['asofdate'], datetime) else str(d['asofdate']) for d in available_dates])
                
                currdate_str = currdate.strftime('%Y-%m-%d')
                datetocompare_str = datetocompare.strftime('%Y-%m-%d')
                
                print(f"  Available dates in database: {sorted(available_date_strs)}")
                print(f"  Requested currdate: {currdate_str}")
                print(f"  Requested datetocompare: {datetocompare_str}")
                
                missing_dates = []
                if currdate_str not in available_date_strs:
                    missing_dates.append(f"currdate ({currdate_str})")
                if datetocompare_str not in available_date_strs:
                    missing_dates.append(f"datetocompare ({datetocompare_str})")
                
                if missing_dates:
                    error_msg = f"\nERROR: No data found for {' and '.join(missing_dates)}"
                    print(error_msg)
                    print("Please ensure data exists for both dates before running refresh-dataconn.")
                    return
                
                print("  ✓ Data exists for both dates")
            
            # Execute queries
            print("\nExecuting database queries...")
            
            # Query 1: totalbyalloctypedate → A1:C17
            query1 = """
                SELECT totalbyalloctypedate_0.allocdesc, totalbyalloctypedate_0.asofdate, 
                       totalbyalloctypedate_0.`sum(assetinvalloc.amount)`
                FROM totalbyalloctypedate totalbyalloctypedate_0
            """
            results1 = self.db.execute_query(query1)
            print(f"totalbyalloctypedate: {len(results1)} rows")
            
            # Query 2: heldatbydate → F1:H27
            query2 = """
                SELECT heldatbydate_0.heldat, heldatbydate_0.asofdate, 
                       heldatbydate_0.`sum(assetinvalloc.amount)`
                FROM heldatbydate heldatbydate_0
            """
            results2 = self.db.execute_query(query2)
            print(f"heldatbydate: {len(results2)} rows")
            
            # Query 3: cashheldatbydate → J1:L27
            query3 = """
                SELECT cashheldatbydate_0.heldat, cashheldatbydate_0.asofdate, 
                       cashheldatbydate_0.`sum(assetinvalloc.amount)`
                FROM cashheldatbydate cashheldatbydate_0
            """
            results3 = self.db.execute_query(query3)
            print(f"cashheldatbydate: {len(results3)} rows")
            
            # Query 4: assetbydate → O1:U69
            query4 = """
                SELECT * FROM `asset`.`assetbydate` 
                WHERE assetname NOT IN ('Cash', 'Stock')
            """
            results4 = self.db.execute_query(query4)
            print(f"assetbydate: {len(results4)} rows")
            
            # Load Excel workbook
            print(f"\nLoading Excel file: {self.excel_file}")
            wb = load_workbook(self.excel_file)
            
            # Get DataConn sheet (note: capital D and C)
            if 'DataConn' not in wb.sheetnames:
                ws = wb.create_sheet('DataConn')
            else:
                ws = wb['DataConn']
            
            # Clear existing data in all the columns where we write
            print("Clearing existing data in DataConn sheet...")
            # Clear columns A-C (rows 1-100), F-H (rows 1-100), J-L (rows 1-100), O-U (rows 1-100)
            for row in range(1, 101):
                # Clear A1:C100
                ws.cell(row, 1).value = None
                ws.cell(row, 2).value = None
                ws.cell(row, 3).value = None
                # Clear F1:H100
                ws.cell(row, 6).value = None
                ws.cell(row, 7).value = None
                ws.cell(row, 8).value = None
                # Clear J1:L100
                ws.cell(row, 10).value = None
                ws.cell(row, 11).value = None
                ws.cell(row, 12).value = None
                # Clear O1:U100
                for col in range(15, 22):  # O=15 to U=21
                    ws.cell(row, col).value = None
            
            print("Writing data to DataConn sheet...")
            
            # Write Query 1 results to A1:C17
            if results1:
                for row_idx, row_data in enumerate(results1, start=1):
                    ws.cell(row_idx, 1, row_data.get('allocdesc'))
                    ws.cell(row_idx, 2, row_data.get('asofdate'))
                    ws.cell(row_idx, 3, row_data.get('sum(assetinvalloc.amount)'))
            
            # Write Query 2 results to F1:H27
            if results2:
                for row_idx, row_data in enumerate(results2, start=1):
                    ws.cell(row_idx, 6, row_data.get('heldat'))
                    ws.cell(row_idx, 7, row_data.get('asofdate'))
                    ws.cell(row_idx, 8, row_data.get('sum(assetinvalloc.amount)'))
            
            # Write Query 3 results to J1:L27
            if results3:
                for row_idx, row_data in enumerate(results3, start=1):
                    ws.cell(row_idx, 10, row_data.get('heldat'))
                    ws.cell(row_idx, 11, row_data.get('asofdate'))
                    ws.cell(row_idx, 12, row_data.get('sum(assetinvalloc.amount)'))
            
            # Write Query 4 results to O1:U69
            if results4:
                for row_idx, row_data in enumerate(results4, start=1):
                    # Get all column names dynamically
                    col_offset = 15  # Column O = 15
                    for col_idx, (key, value) in enumerate(row_data.items()):
                        ws.cell(row_idx, col_offset + col_idx, value)
            
            # Save workbook
            wb.save(self.excel_file)
            print(f"\nDataConn sheet updated successfully in {self.excel_file}")
            
        finally:
            self.db.close_db()
        
        print("=" * 60)
        print("Refresh Complete!")
        print("=" * 60)
    
    def fix_external_references(self):
        """
        Fix external workbook references in formulas
        Note: This only fixes formula references. Excel may still show a warning
        about external links, but the formulas will work correctly.
        """
        print("=" * 60)
        print("Fixing External Workbook References")
        print("=" * 60)
        
        from openpyxl import load_workbook
        import re
        import os
        import shutil
        
        # Create backup first
        backup_file = self.excel_file + '.backup'
        if os.path.exists(self.excel_file):
            shutil.copy2(self.excel_file, backup_file)
            print(f"Created backup: {backup_file}")
        
        # Load Excel workbook with data_only=False to preserve formulas
        print(f"\nLoading Excel file: {self.excel_file}")
        wb = load_workbook(self.excel_file, data_only=False, keep_vba=True)
        
        # Pattern to match external workbook references
        pattern_numbered = r'\[\d+\]'  # [1], [2], etc.
        pattern_fin = r"'\[fin\.xlsx\][^']*'!"
        pattern_fin_simple = r'\[fin\.xlsx\][^!]*!'
        
        fixed_count = 0
        fin_count = 0
        
        # Check all sheets
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_fixes = 0
            sheet_fin_fixes = 0
            
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        old_value = cell.value
                        new_value = old_value
                        
                        # Remove numbered external workbook references like [1], [2], etc.
                        if re.search(r'\[\d+\]', old_value):
                            new_value = re.sub(pattern_numbered, '', new_value)
                            sheet_fixes += 1
                            fixed_count += 1
                        
                        # Remove fin.xlsx references
                        if 'fin.xlsx' in old_value.lower():
                            new_value = re.sub(pattern_fin, '#REF!', new_value, flags=re.IGNORECASE)
                            new_value = re.sub(pattern_fin_simple, '#REF!', new_value, flags=re.IGNORECASE)
                            sheet_fin_fixes += 1
                            fin_count += 1
                        
                        if new_value != old_value:
                            cell.value = new_value
            
            if sheet_fixes > 0:
                print(f"  {sheet_name}: Fixed {sheet_fixes} numbered references")
            if sheet_fin_fixes > 0:
                print(f"  {sheet_name}: Removed {sheet_fin_fixes} fin.xlsx references")
        
        # Save workbook
        print("\nSaving workbook...")
        wb.save(self.excel_file)
        wb.close()
        
        print(f"\nFixed {fixed_count} numbered references and {fin_count} fin.xlsx references")
        print("\nNote: Excel may still show an external links warning on open.")
        print("You can safely update or break those links in Excel's Data > Edit Links menu.")
        print("=" * 60)
        print("Fix Complete!")
        print("=" * 60)
    
    def run_full_process(self, as_of_date: datetime, sheet_name: str = 'fullview', 
                        delete_existing: bool = True, calculate_gains: bool = False):
        """
        Run the full asset processing workflow
        
        Args:
            as_of_date: Date for the asset allocation
            sheet_name: Name of the sheet to read
            delete_existing: Whether to delete existing data first
            calculate_gains: Whether to calculate gains after allocation (default: False)
        """
        print("=" * 60)
        print("Asset Processing Workflow")
        print("=" * 60)
        print(f"Excel File: {self.excel_file}")
        print(f"Sheet: {sheet_name}")
        print(f"As of Date: {as_of_date.strftime('%Y-%m-%d')}")
        print("=" * 60)
        
        # Step 1: Delete existing data if requested
        if delete_existing:
            self.delete_existing_data(as_of_date)
        
        # Step 2: Read Excel data
        print("\nReading Excel file...")
        df = self.read_asset_reference_sheet(sheet_name)
        print(f"Read {len(df)} rows from Excel")
        
        # Step 3: Process asset allocation
        print("\nProcessing asset allocation...")
        self.process_asset_allocation(df, as_of_date)
        
        # Step 4: Calculate gains if requested
        if calculate_gains:
            self.calculate_gains(as_of_date)
        
        print("\n" + "=" * 60)
        print("Processing Complete!")
        print("=" * 60)

    def show_unique_dates(self, after_date: datetime = None):
        """
        Show unique dates for which there is data, optionally filtered after a given date
        
        Args:
            after_date: Optional date to filter results - only show dates after this date
        """
        print(f"\n{'='*60}")
        print(f"Available Data Dates")
        print(f"{'='*60}")
        if after_date:
            print(f"Showing dates after: {after_date.strftime('%Y-%m-%d')}\n")
        else:
            print(f"Showing all available dates\n")
        
        self.db.open_db()
        
        # Query to get unique dates from assetinv table
        if after_date:
            query = """
                SELECT DISTINCT asofdate 
                FROM assetinv 
                WHERE asofdate > %s
                ORDER BY asofdate DESC
            """
            results = self.db.execute_query(query, (after_date.strftime('%Y-%m-%d'),))
        else:
            query = """
                SELECT DISTINCT asofdate 
                FROM assetinv 
                ORDER BY asofdate DESC
            """
            results = self.db.execute_query(query)
        
        self.db.close_db()
        
        if not results:
            if after_date:
                print(f"No data found after {after_date.strftime('%Y-%m-%d')}")
            else:
                print("No data found in database")
            return
        
        # Display the dates
        print(f"Found {len(results)} unique date(s):\n")
        for i, row in enumerate(results, 1):
            date_val = row['asofdate']
            # Handle both datetime and date objects
            if isinstance(date_val, datetime):
                date_str = date_val.strftime('%Y-%m-%d')
            else:
                date_str = str(date_val)
            print(f"  {i:2d}. {date_str}")
        
        print(f"\n{'='*60}")


def main():
    """Main entry point"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Process Asset.xls file and update database')
    parser.add_argument('--file', '-f', default='Asset.xls', 
                       help='Path to Excel file (default: Asset.xls)')
    parser.add_argument('--sheet', '-s', default='fullview',
                       help='Sheet name to process (default: fullview)')
    parser.add_argument('--date', '-d', 
                       help='As-of date in YYYY-MM-DD format (default: today)')
    parser.add_argument('--no-delete', action='store_true',
                       help='Do not delete existing data before processing')
    parser.add_argument('--with-gains', action='store_true',
                       help='Calculate gains after allocation (not default)')
    parser.add_argument('--gains-only', action='store_true',
                       help='Only calculate gains, skip allocation')
    parser.add_argument('--delete-only', action='store_true',
                       help='Only delete data, skip allocation and gains')
    parser.add_argument('--process', action='store_true',
                       help='Run main asset allocation workflow (default if no other mode specified)')
    parser.add_argument('--normalize', action='store_true',
                       help='Normalize full view data and aggregate by account')
    parser.add_argument('--normalize-sheet', default='fidfullview',
                       help='Sheet name for normalize operation (default: fidfullview)')
    parser.add_argument('--output', '-o',
                       help='Output Excel file for normalize results')
    parser.add_argument('--refresh-dataconn', action='store_true',
                       help='Refresh dataconn sheet with database query results')
    parser.add_argument('--compare-dates', action='store_true',
                       help='Compare asset values between two dates and report significant changes')
    parser.add_argument('--show-all', action='store_true',
                       help='Show all changes (ignore threshold) when using --compare-dates')
    parser.add_argument('--threshold', type=float, default=5.0,
                       help='Percentage threshold for significant changes (default: 5.0)')
    parser.add_argument('--currdate',
                       help='Current date for wkdates table and assetref N4 (YYYY-MM-DD)')
    parser.add_argument('--datetocompare',
                       help='Date to compare for wkdates table and assetref N3 (YYYY-MM-DD)')
    parser.add_argument('--fix-references', action='store_true',
                       help='Fix external workbook references in formulas')
    parser.add_argument('--show-dates', action='store_true',
                       help='Show unique dates for which there is data in the database')
    parser.add_argument('--after-date',
                       help='Filter dates to show only those after this date (YYYY-MM-DD). Use with --show-dates')
    
    args = parser.parse_args()
    
    # Parse date
    if args.date:
        try:
            as_of_date = datetime.strptime(args.date, '%Y-%m-%d')
        except ValueError:
            print("Error: Invalid date format. Use YYYY-MM-DD")
            sys.exit(1)
    else:
        as_of_date = datetime.now()
    
    # Parse currdate and datetocompare
    currdate = None
    datetocompare = None
    if args.currdate:
        try:
            currdate = datetime.strptime(args.currdate, '%Y-%m-%d')
        except ValueError:
            print("Error: Invalid currdate format. Use YYYY-MM-DD")
            sys.exit(1)
    
    if args.datetocompare:
        try:
            datetocompare = datetime.strptime(args.datetocompare, '%Y-%m-%d')
        except ValueError:
            print("Error: Invalid datetocompare format. Use YYYY-MM-DD")
            sys.exit(1)
    
    # Parse after_date if provided
    after_date = None
    if args.after_date:
        try:
            after_date = datetime.strptime(args.after_date, '%Y-%m-%d')
        except ValueError:
            print("Error: Invalid after-date format. Use YYYY-MM-DD")
            sys.exit(1)
    
    # Create processor
    processor = AssetProcessor(args.file)
    
    # Execute based on flags
    if args.show_dates:
        # Show unique dates with optional filtering
        processor.show_unique_dates(after_date=after_date)
    elif args.normalize:
        # Update Assetalloc dates
        processor.update_assetalloc_dates(
            prevdate=datetocompare.strftime('%Y-%m-%d') if datetocompare else None,
            currdate=as_of_date.strftime('%Y-%m-%d') if as_of_date else None
        )
        # Run normalize full view
        processor.normalize_full_view(
            sheet_name=args.normalize_sheet,
            output_file=args.output
        )
    elif args.refresh_dataconn:
        # Refresh dataconn sheet
        processor.refresh_dataconn(currdate=currdate, datetocompare=datetocompare)
    elif args.compare_dates:
        # Compare dates and report significant changes
        if not currdate or not datetocompare:
            print("Error: --compare-dates requires --currdate and --datetocompare")
            sys.exit(1)
        processor.compare_dates_report(currdate=currdate, datetocompare=datetocompare, 
                                      threshold_percent=args.threshold, show_all=args.show_all)
    elif args.fix_references:
        # Fix external workbook references
        processor.fix_external_references()
    elif args.delete_only:
        processor.delete_existing_data(as_of_date)
    elif args.gains_only:
        processor.calculate_gains(as_of_date)
    elif args.process or not any([args.normalize, args.refresh_dataconn, 
                                   args.compare_dates, args.fix_references, args.delete_only, 
                                   args.gains_only, args.show_dates]):
        # Update Assetalloc dates
        processor.update_assetalloc_dates(
            prevdate=datetocompare.strftime('%Y-%m-%d') if datetocompare else None,
            currdate=as_of_date.strftime('%Y-%m-%d') if as_of_date else None
        )
        # Run main allocation workflow (default)
        processor.run_full_process(
            as_of_date=as_of_date,
            sheet_name=args.sheet,
            delete_existing=not args.no_delete,
            calculate_gains=args.with_gains
        )


if __name__ == '__main__':
    main()
