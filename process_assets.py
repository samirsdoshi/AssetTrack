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
    
    def normalize_full_view(self, sheet_name: str = 'fullview', output_file: str = None) -> pd.DataFrame:
        """
        Normalize and aggregate fullview data by account
        Converted from VBA normalizefullview() function
        
        Args:
            sheet_name: Name of the sheet to process (default: 'fullview')
            output_file: Optional output Excel file to save results
            
        Returns:
            DataFrame with normalized data
        """
        print(f"Normalizing full view from sheet '{sheet_name}'...")
        
        # Read the fullview sheet (no header for fullview)
        df = self.decrypt_and_read_excel(sheet_name, header=None)
        
        account_data = {}  # Dictionary to store aggregated amounts by account
        results = []  # List to store individual ticker results
        
        curr_account = ""
        old_account = ""
        
        i = 0
        while i < len(df):
            # Check if we've reached the end
            if pd.isna(df.iloc[i, 0]) or df.iloc[i, 0] == "":
                break
            
            account = str(df.iloc[i, 0])
            
            # Get fund name from next row if available
            fund = ""
            if i + 1 < len(df):
                fund = str(df.iloc[i + 1, 0]).lower() if not pd.isna(df.iloc[i + 1, 0]) else ""
            
            curr_account = ""
            
            # Categorize accounts based on patterns
            if "CollegeAdv" in account:
                curr_account = "CollegeAdv"
            elif "- Individual" in account and "rowe" in fund and "price" in fund:
                curr_account = "TRPInv"
            elif "Rollover IRA" in account and "rowe" in fund and "price" in fund:
                curr_account = "TRPRollover"
            elif "Roth IRA" in account and "rowe" in fund and "price" in fund:
                curr_account = "TRPRoth"
            elif "Traditional IRA" in account and "rowe" in fund and "price" in fund:
                curr_account = "TRPRollover"  # Sangeeta's account
            elif "GAPSHARE 401(K) PLAN" in account:
                curr_account = "TRPRps"
            elif "Individual - TOD" in account:
                curr_account = "FidelityInv"
            elif "Samir S Doshi - Rollover IRA" in account:
                curr_account = "Vanguard IRA"
            elif not curr_account and "Rollover IRA" in account:
                curr_account = "FidelityIRA"
            elif "Brokerage Account - 10498558" in account:
                curr_account = "Vanguard"
            elif "Vanguard Investments - Samir S Doshi" in account:
                curr_account = "Vanguard"
            
            # If no current account but we have an old account, this is a fund/ticker row
            if not curr_account and old_account:
                fund = account  # This row is actually the fund/ticker
                ticker = df.iloc[i, 1] if not pd.isna(df.iloc[i, 1]) else ""
                amount = df.iloc[i, 5] if not pd.isna(df.iloc[i, 5]) else 0  # Column F (index 5)
                
                # Clean up amount if it's a string
                if isinstance(amount, str):
                    amount = amount.replace('$', '').replace(',', '')
                    try:
                        amount = float(amount)
                    except:
                        amount = 0
                
                # Store result
                results.append({
                    'account_ticker': f"{old_account}_{ticker}",
                    'amount': amount
                })
                
                # Aggregate by account
                if old_account not in account_data:
                    account_data[old_account] = 0
                account_data[old_account] += amount
            else:
                old_account = curr_account
            
            i += 1
        
        # Create results DataFrame and sort by account_ticker
        results_df = pd.DataFrame(results)
        
        # Aggregate _Stock entries: sum all <Account>_Stock_* into <Account>_Stock
        aggregated_results = []
        stock_aggregates = {}
        
        for _, row in results_df.iterrows():
            ticker = row['account_ticker']
            amount = row['amount']
            
            # Check if this is a stock entry with additional ticker info
            if '_Stock' in ticker:
                # Extract the account_Stock prefix (everything up to and including _Stock)
                parts = ticker.split('_Stock')
                if len(parts) >= 2:
                    base_ticker = parts[0] + '_Stock'
                    # Aggregate under the base ticker
                    if base_ticker not in stock_aggregates:
                        stock_aggregates[base_ticker] = 0
                    stock_aggregates[base_ticker] += amount
                else:
                    # Already in base form
                    if ticker not in stock_aggregates:
                        stock_aggregates[ticker] = 0
                    stock_aggregates[ticker] += amount
            else:
                # Non-stock entry, keep as is
                aggregated_results.append({'account_ticker': ticker, 'amount': amount})
        
        # Add aggregated stock entries
        for ticker, total_amount in stock_aggregates.items():
            aggregated_results.append({'account_ticker': ticker, 'amount': total_amount})
        
        # Create new DataFrame with aggregated results and sort
        results_df = pd.DataFrame(aggregated_results)
        
        # Remove entries with 0 value
        results_df = results_df[results_df['amount'] != 0]
        
        # Remove TRP accounts (will be replaced with Trow sheet data)
        trp_accounts = ['TRPInv', 'TRPRollover', 'TRPRoth', 'TRPRps']
        results_df = results_df[~results_df['account_ticker'].str.startswith(tuple(f"{acc}_" for acc in trp_accounts))]
        
        # Read Trow sheet columns N and O, and add to results
        try:
            print("Reading Trow sheet data...")
            trow_df = self.decrypt_and_read_excel('Trow', header=None)
            
            # Extract columns N (index 13) and O (index 14)
            for idx, row in trow_df.iterrows():
                ticker = row.get(13)  # Column N
                amount = row.get(14)  # Column O
                
                # Skip if either is empty or amount is 0
                if pd.isna(ticker) or pd.isna(amount) or ticker == '' or amount == 0:
                    continue
                
                # Clean up amount if it's a string
                if isinstance(amount, str):
                    amount = amount.replace('$', '').replace(',', '')
                    try:
                        amount = float(amount)
                    except:
                        continue
                
                # Add to results
                results_df = pd.concat([results_df, pd.DataFrame([{
                    'account_ticker': str(ticker),
                    'amount': amount
                }])], ignore_index=True)
            
            print(f"Added {len(trow_df[trow_df[13].notna() & trow_df[14].notna()])} entries from Trow sheet")
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
            
            ws = wb[sheet_name]
            
            # Clear columns K:O (rows 1-2000)
            for row in range(1, 2001):
                ws[f'K{row}'].value = None
                ws[f'L{row}'].value = None
                ws[f'N{row}'].value = None
                ws[f'O{row}'].value = None
            
            # Write results to columns K and L
            j = 1
            for _, row in results_df.iterrows():
                ws[f'K{j}'].value = row['account_ticker']
                ws[f'L{j}'].value = row['amount']
                j += 1
            
            # Write summary to columns N and O (starting at row 8)
            k = 8
            for _, row in summary_df.iterrows():
                ws[f'N{k}'].value = row['account']
                ws[f'O{k}'].value = row['total']
                k += 1
            
            # Save the workbook
            wb.save(self.excel_file)
            print(f"Results written to {self.excel_file} in columns K, L, N, O")
            
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
    
    def update_asset_ref_from_fullview(self):
        """
        Update assetref sheet from fullview normalized data
        Matches Column K (Account_Fund) with Column J (Account) and Column A (Fund)
        Updates Column E with values from Column L in fullview
        """
        print("Updating assetref from fullview...")
        
        try:
            # Read fullview sheet normalized data (columns K and L)
            fullview_df = self.decrypt_and_read_excel('fullview', header=None)
            
            # Read assetref sheet
            assetref_df = self.decrypt_and_read_excel('assetref', header=0)
            
            from openpyxl import load_workbook
            
            # Load workbook for writing
            if self.excel_password and not self.excel_file.endswith('.xlsx'):
                try:
                    decrypted = io.BytesIO()
                    with open(self.excel_file, 'rb') as f:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password=self.excel_password)
                        office_file.decrypt(decrypted)
                    decrypted.seek(0)
                    wb = load_workbook(decrypted)
                except:
                    wb = load_workbook(self.excel_file)
            else:
                wb = load_workbook(self.excel_file)
            
            ws = wb['assetref']
            
            # Extract fullview data from columns K (10) and L (11)
            fullview_data = {}
            
            # Fund mapping dictionary - map alternative fund names
            fund_mapping = {
                'VMRXX': 'VMMXX',  # Vanguard fund mapping
            }
            
            for idx, row in fullview_df.iterrows():
                account_fund = row.get(10)  # Column K
                amount = row.get(11)  # Column L
                
                if pd.isna(account_fund) or account_fund == '':
                    continue
                
                # Parse account_fund into account and fund
                account_fund_str = str(account_fund)
                if '_' in account_fund_str:
                    parts = account_fund_str.split('_', 1)  # Split on first underscore
                    account = parts[0]
                    fund = parts[1]
                    
                    # Store in dictionary with (account, fund) as key
                    fullview_data[(account, fund)] = amount
                    
                    # Also store mapped version if applicable
                    if fund in fund_mapping:
                        mapped_fund = fund_mapping[fund]
                        fullview_data[(account, mapped_fund)] = amount
            
            print(f"Found {len(fullview_data)} entries in fullview columns K and L")
            
            # Track matches and mismatches
            matched_count = 0
            updated_count = 0
            unmatched = []
            account_totals = {}  # Track totals per account
            
            # Iterate through assetref rows (starting from row 2 in Excel, index 1 in 0-based)
            for idx in range(1, ws.max_row + 1):
                # Read Column A (Ticker/Fund) and Column J (HeldAt/Account)
                fund_cell = ws[f'A{idx}'].value
                account_cell = ws[f'J{idx}'].value
                
                if not fund_cell or fund_cell == '':
                    continue
                
                fund = str(fund_cell).strip()
                
                # Stop processing if we encounter Etrade or ENDOFPORTFOLIO
                if fund == 'ENDOFPORTFOLIO' or fund == 'Etrade':
                    break
                
                account = str(account_cell).strip() if account_cell else ''
                
                if not account:
                    continue
                
                # Check if this is a Total row
                if fund.endswith(' Total') or 'Total' in fund:
                    # This is a total row - we'll update it later
                    continue
                
                # Look up in fullview_data
                if (account, fund) in fullview_data:
                    amount = fullview_data[(account, fund)]
                    # Update Column E
                    ws[f'E{idx}'].value = amount
                    matched_count += 1
                    updated_count += 1
                    
                    # Add to account total
                    if account not in account_totals:
                        account_totals[account] = 0
                    account_totals[account] += float(amount) if amount else 0
                else:
                    # Check current value in column E
                    current_value = ws[f'E{idx}'].value
                    if current_value and current_value != 0:
                        unmatched.append({
                            'row': idx,
                            'account': account,
                            'fund': fund,
                            'current_value': current_value
                        })
                        # Add to account total
                        if account not in account_totals:
                            account_totals[account] = 0
                        try:
                            account_totals[account] += float(current_value)
                        except:
                            pass
            
            # Update existing total rows and track which accounts have totals
            print("\nUpdating account totals in column B...")
            accounts_with_totals = set()
            last_row_per_account = {}  # Track last fund row for each account
            
            # First pass: update existing totals and track last rows
            for idx in range(1, ws.max_row + 1):
                fund_cell = ws[f'A{idx}'].value
                account_cell = ws[f'J{idx}'].value
                
                if not fund_cell:
                    continue
                
                fund = str(fund_cell).strip()
                
                if fund == 'ENDOFPORTFOLIO' or fund == 'Etrade':
                    break
                
                account = str(account_cell).strip() if account_cell else ''
                
                # Check if this is a Total row
                if ' Total' in fund:
                    # Extract account name from "<Account> Total"
                    total_account = fund.replace(' Total', '').strip()
                    # Also check column J for account
                    if account_cell:
                        total_account = str(account_cell).strip()
                    
                    accounts_with_totals.add(total_account)
                    
                    if total_account in account_totals:
                        total = account_totals[total_account]
                        ws[f'B{idx}'].value = total
                        print(f"  {total_account} Total: ${total:,.2f} (Row {idx})")
                    else:
                        print(f"  Warning: No total calculated for {total_account}")
                elif account and not fund.endswith(' Total'):
                    # Track last row for each account
                    last_row_per_account[account] = idx
            
            # Second pass: add missing total rows
            accounts_needing_totals = set(account_totals.keys()) - accounts_with_totals
            if accounts_needing_totals:
                print("\nAdding missing total rows...")
                for account in sorted(accounts_needing_totals):
                    if account in last_row_per_account:
                        # Insert after the last fund row for this account
                        insert_row = last_row_per_account[account] + 1
                        ws.insert_rows(insert_row)
                        
                        # Set the total row data
                        ws[f'A{insert_row}'].value = f"{account} Total"
                        ws[f'B{insert_row}'].value = account_totals[account]
                        ws[f'J{insert_row}'].value = account
                        
                        print(f"  Added {account} Total: ${account_totals[account]:,.2f} (Row {insert_row})")
                        
                        # Update last_row_per_account for subsequent accounts
                        for acc, row in last_row_per_account.items():
                            if row >= insert_row:
                                last_row_per_account[acc] += 1
            
            # Save workbook
            wb.save(self.excel_file)
            
            print(f"\n=== Update Summary ===")
            print(f"Matched and updated: {updated_count} entries")
            print(f"Unmatched in assetref: {len(unmatched)} entries")
            
            if unmatched:
                print("\nUnmatched entries (in assetref but not in fullview):")
                for entry in unmatched:
                    print(f"  Row {entry['row']}: {entry['account']}_{entry['fund']} = ${entry['current_value']}")
            
            print(f"\nAssetref sheet updated successfully!")
            
        except Exception as e:
            print(f"Error updating assetref: {e}")
            import traceback
            traceback.print_exc()
    
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
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=header, engine='openpyxl')
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
                    wb = openpyxl_load(cleaned, data_only=False)
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
    
    def read_asset_reference_sheet(self, sheet_name: str = 'assetref') -> pd.DataFrame:
        """
        Read the asset reference sheet from Excel
        
        Args:
            sheet_name: Name of the sheet to read
            
        Returns:
            DataFrame with asset data
        """
        try:
            # Read without header since assetref doesn't have column names
            df = self.decrypt_and_read_excel(sheet_name, header=None)
            
            # Assign column names based on position
            # Column A=0: Ticker, C=2: Quantity, E=4: Amount, J=9: HeldAt
            df.columns = [f'Col{i}' for i in range(len(df.columns))]
            df = df.rename(columns={
                'Col0': 'Ticker',
                'Col2': 'Quantity', 
                'Col4': 'Amount',
                'Col9': 'HeldAt'
            })
            
            print(f"Successfully read sheet '{sheet_name}' from {self.excel_file}")
            return df
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
                    continue
                
                asset_id = results[0]['assetid']
                
                # Allocate the asset
                print(f"Processing: {ticker} - ${amount:,.2f} at {held_at}")
                self.allocator.allocate_asset_ref(asset_id, as_of_date, amount, held_at)
                processed_count += 1
                
            except Exception as e:
                print(f"Error processing row {index}: {e}")
                error_count += 1
                continue
        
        self.db.close_db()
        
        print(f"\n=== Processing Complete ===")
        print(f"Processed: {processed_count} assets")
        print(f"Errors: {error_count}")
    
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
            currdate: Current date to update in wkdates table and assetref N4
            datetocompare: Date to compare to update in wkdates table and assetref N3
        """
        print("=" * 60)
        print("Refreshing Dataconn Sheet")
        print("=" * 60)
        
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
                # Clear only the data cells, not formulas
                # We'll overwrite cells, so no need to delete
            
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
            
            # Update assetref sheet with dates if provided
            if currdate or datetocompare:
                print("\nUpdating assetref sheet with dates...")
                if 'assetref' in wb.sheetnames:
                    ws_assetref = wb['assetref']
                    
                    if datetocompare:
                        ws_assetref['N3'] = datetocompare
                        print(f"  Updated N3 (datetocompare) to {datetocompare.strftime('%Y-%m-%d')}")
                    
                    if currdate:
                        ws_assetref['N4'] = currdate
                        print(f"  Updated N4 (currdate) to {currdate.strftime('%Y-%m-%d')}")
                    
                    wb.save(self.excel_file)
                    print(f"  assetref sheet updated in {self.excel_file}")
                else:
                    print("  Warning: assetref sheet not found in workbook")
            
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
    
    def run_full_process(self, as_of_date: datetime, sheet_name: str = 'assetref', 
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


def main():
    """Main entry point"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Process Asset.xls file and update database')
    parser.add_argument('--file', '-f', default='Asset.xls', 
                       help='Path to Excel file (default: Asset.xls)')
    parser.add_argument('--sheet', '-s', default='assetref',
                       help='Sheet name to process (default: assetref)')
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
    parser.add_argument('--normalize-sheet', default='fullview',
                       help='Sheet name for normalize operation (default: fullview)')
    parser.add_argument('--output', '-o',
                       help='Output Excel file for normalize results')
    parser.add_argument('--updateassetref', action='store_true',
                       help='Update assetref sheet from fullview normalized data')
    parser.add_argument('--refresh-dataconn', action='store_true',
                       help='Refresh dataconn sheet with database query results')
    parser.add_argument('--currdate',
                       help='Current date for wkdates table and assetref N4 (YYYY-MM-DD)')
    parser.add_argument('--datetocompare',
                       help='Date to compare for wkdates table and assetref N3 (YYYY-MM-DD)')
    parser.add_argument('--fix-references', action='store_true',
                       help='Fix external workbook references in formulas')
    
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
    
    # Create processor
    processor = AssetProcessor(args.file)
    
    # Execute based on flags
    if args.normalize:
        # Run normalize full view
        processor.normalize_full_view(
            sheet_name=args.normalize_sheet,
            output_file=args.output
        )
    elif args.updateassetref:
        # Update assetref from fullview
        processor.update_asset_ref_from_fullview()
    elif args.refresh_dataconn:
        # Refresh dataconn sheet
        processor.refresh_dataconn(currdate=currdate, datetocompare=datetocompare)
    elif args.fix_references:
        # Fix external workbook references
        processor.fix_external_references()
    elif args.delete_only:
        processor.delete_existing_data(as_of_date)
    elif args.gains_only:
        processor.calculate_gains(as_of_date)
    elif args.process or not any([args.normalize, args.updateassetref, args.refresh_dataconn, 
                                   args.fix_references, args.delete_only, args.gains_only]):
        # Run main allocation workflow (default)
        processor.run_full_process(
            as_of_date=as_of_date,
            sheet_name=args.sheet,
            delete_existing=not args.no_delete,
            calculate_gains=args.with_gains
        )


if __name__ == '__main__':
    main()
