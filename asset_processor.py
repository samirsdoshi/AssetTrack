"""
Asset Allocation and Investment Processing Module
Converted from VBA to Python
Processes Asset.xls and manages MySQL database operations
"""

import mysql.connector
from mysql.connector import Error
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Tuple
import requests
from bs4 import BeautifulSoup
import csv
from io import StringIO


class AssetDatabase:
    """Handles all database operations for asset management"""
    
    def __init__(self, host='localhost', port=3306, user='root', password='sa123', database='asset'):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.database = database
        self.connection = None
        self.cursor = None
    
    def open_db(self):
        """Open database connection"""
        try:
            if self.connection is None or not self.connection.is_connected():
                self.connection = mysql.connector.connect(
                    host=self.host,
                    port=self.port,
                    user=self.user,
                    password=self.password,
                    database=self.database,
                    autocommit=False
                )
                self.cursor = self.connection.cursor(dictionary=True)
                print(f"Connected to MySQL database: {self.database}")
        except Error as e:
            print(f"Error connecting to MySQL: {e}")
            raise
    
    def close_db(self):
        """Close database connection"""
        try:
            if self.cursor:
                self.cursor.close()
            if self.connection and self.connection.is_connected():
                self.connection.close()
                print("MySQL connection closed")
            self.connection = None
            self.cursor = None
        except Error as e:
            print(f"Error closing MySQL connection: {e}")
    
    def execute_query(self, query: str, params: tuple = None) -> List[Dict]:
        """Execute a SELECT query and return results"""
        try:
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
        except Error as e:
            print(f"Error executing query: {e}")
            print(f"Query: {query}")
            raise
    
    def execute_update(self, query: str, params: tuple = None):
        """Execute an INSERT/UPDATE/DELETE query"""
        try:
            self.cursor.execute(query, params)
        except Error as e:
            print(f"Error executing update: {e}")
            print(f"Query: {query}")
            raise
    
    def begin_transaction(self):
        """Begin a database transaction"""
        self.connection.start_transaction()
    
    def commit(self):
        """Commit the current transaction"""
        try:
            self.connection.commit()
        except Error as e:
            print(f"Error committing transaction: {e}")
            raise
    
    def rollback(self):
        """Rollback the current transaction"""
        try:
            self.connection.rollback()
        except Error as e:
            print(f"Error rolling back transaction: {e}")


class AssetAllocator:
    """Handles asset allocation operations"""
    
    def __init__(self, db: AssetDatabase):
        self.db = db
    
    @staticmethod
    def mysql_date(dt) -> str:
        """Convert date to MySQL format YYYY-MM-DD"""
        if isinstance(dt, str):
            dt = datetime.strptime(dt, '%m/%d/%Y')
        return dt.strftime('%Y-%m-%d')
    
    @staticmethod
    def nullif(value, default_value):
        """Return default_value if value is None, otherwise return value"""
        return default_value if value is None else value
    
    def reallocate(self, asset_id: int, as_of_date, amount: float):
        """Delete existing allocation and reallocate"""
        self.db.open_db()
        
        try:
            # Get existing asset investment IDs
            date_str = self.mysql_date(as_of_date)
            query = """
                SELECT assetinvid FROM assetinv 
                WHERE assetid=%s AND asofdate=%s AND amount=%s 
                ORDER BY assetinvid DESC
            """
            results = self.db.execute_query(query, (asset_id, date_str, amount))
            
            assetinv_ids = [str(row['assetinvid']) for row in results]
            if not assetinv_ids:
                assetinv_ids = ['0']
            
            assetinv_ids_str = ','.join(assetinv_ids)
            
            # Begin transaction to delete old allocations
            self.db.begin_transaction()
            try:
                self.db.execute_update(f"DELETE FROM assetinvalloc WHERE assetinvid IN ({assetinv_ids_str})")
                self.db.execute_update(f"DELETE FROM assetinvsecind WHERE assetinvid IN ({assetinv_ids_str})")
                self.db.execute_update(f"DELETE FROM assetinvinter WHERE assetinvid IN ({assetinv_ids_str})")
                self.db.execute_update(f"DELETE FROM assetinv WHERE assetinvid IN ({assetinv_ids_str})")
                self.db.commit()
            except Exception as e:
                self.db.rollback()
                print(f"Error in reallocate delete: {e}")
                raise
            
            # Now allocate
            self.allocate(asset_id, as_of_date, amount)
            
        finally:
            self.db.close_db()
    
    def allocate(self, asset_id: int, as_of_date, amount: float):
        """Allocate an asset based on its template"""
        self.db.open_db()
        
        try:
            self.db.begin_transaction()
            
            # Insert into assetinv
            date_str = self.mysql_date(as_of_date)
            insert_query = """
                INSERT INTO assetinv(assetid, asofdate, amount) 
                VALUES (%s, %s, %s)
            """
            self.db.execute_update(insert_query, (asset_id, date_str, amount))
            
            # Get the newly inserted assetinvid
            self.db.cursor.execute("SELECT LAST_INSERT_ID() as id")
            assetinv_id = self.db.cursor.fetchone()['id']
            
            # Get template details
            template_query = """
                SELECT tcode, tval1, tval2, prct 
                FROM templatedetails td 
                INNER JOIN asset a ON td.templateid = a.templateid 
                WHERE a.assetid = %s
            """
            template_details = self.db.execute_query(template_query, (asset_id,))
            
            # Process each template detail
            for detail in template_details:
                tcode = detail['tcode'].lower()
                tval1 = self.nullif(detail['tval1'], '0')
                tval2 = self.nullif(detail['tval2'], '0')
                prct = float(detail['prct'])
                
                allocated_amount = round(amount * (prct / 100), 2)
                
                if tcode == 'alloc':
                    insert_alloc = """
                        INSERT INTO assetinvalloc(assetinvid, alloccode, amount) 
                        VALUES (%s, %s, %s)
                    """
                    self.db.execute_update(insert_alloc, (assetinv_id, int(tval1), allocated_amount))
                
                elif tcode == 'secind':
                    insert_secind = """
                        INSERT INTO assetinvsecind(assetinvid, sec_id, ind_id, amount) 
                        VALUES (%s, %s, %s, %s)
                    """
                    self.db.execute_update(insert_secind, (assetinv_id, int(tval1), int(tval2), allocated_amount))
                
                elif tcode == 'inter':
                    insert_inter = """
                        INSERT INTO assetinvinter(assetinvid, intercode, amount) 
                        VALUES (%s, %s, %s)
                    """
                    self.db.execute_update(insert_inter, (assetinv_id, int(tval1), allocated_amount))
            
            self.db.commit()
            print(f"Successfully allocated asset {asset_id} with amount {amount}")
            
        except Exception as e:
            self.db.rollback()
            print(f"Error in allocate: {e}")
            raise
        finally:
            self.db.close_db()
    
    def allocate_asset_ref(self, asset_id: int, as_of_date, amount: float, held_at: str):
        """Allocate an asset from the asset reference sheet
        Note: DB connection must be open before calling this method"""
        
        if amount == 0:
            return
        
        try:
            self.db.begin_transaction()
            
            date_str = self.mysql_date(as_of_date)
            
            # Check if asset investment already exists
            check_query = """
                SELECT assetinvid, amount FROM assetinv 
                WHERE assetid=%s AND asofdate=%s AND heldat=%s
            """
            results = self.db.execute_query(check_query, (asset_id, date_str, held_at))
            
            dup_asset = False
            if results:
                assetinv_id = results[0]['assetinvid']
                org_amount = results[0]['amount']
                dup_asset = True
                
                if org_amount != amount:
                    update_query = "UPDATE assetinv SET amount = amount + %s WHERE assetinvid = %s"
                    self.db.execute_update(update_query, (amount, assetinv_id))
            else:
                # Get max assetinvid
                self.db.cursor.execute("SELECT MAX(assetinvid) as max_id FROM assetinv")
                result = self.db.cursor.fetchone()
                assetinv_id = (result['max_id'] or 0) + 1
                
                # Insert new assetinv
                insert_query = """
                    INSERT INTO assetinv(assetinvid, assetid, asofdate, amount, heldat) 
                    VALUES (%s, %s, %s, %s, %s)
                """
                self.db.execute_update(insert_query, (assetinv_id, asset_id, date_str, amount, held_at))
            
            # Get template details
            template_query = """
                SELECT tcode, tval1, tval2, prct 
                FROM templatedetails td 
                INNER JOIN asset a ON td.templateid = a.templateid 
                WHERE a.assetid = %s
            """
            template_details = self.db.execute_query(template_query, (asset_id,))
            
            if not template_details:
                raise Exception(f"No template details found for asset {asset_id}")
            
            # Process each template detail
            for detail in template_details:
                tcode = detail['tcode'].lower()
                tval1 = self.nullif(detail['tval1'], '0')
                tval2 = self.nullif(detail['tval2'], '0')
                prct = float(detail['prct'])
                
                allocated_amount = round(amount * (prct / 100), 2)
                
                if tcode == 'alloc':
                    if dup_asset:
                        update_alloc = """
                            UPDATE assetinvalloc SET amount = amount + %s 
                            WHERE assetinvid = %s AND alloccode = %s
                        """
                        self.db.execute_update(update_alloc, (allocated_amount, assetinv_id, tval1))
                    else:
                        insert_alloc = """
                            INSERT INTO assetinvalloc(assetinvid, alloccode, amount) 
                            VALUES (%s, %s, %s)
                        """
                        self.db.execute_update(insert_alloc, (assetinv_id, int(tval1), allocated_amount))
                
                elif tcode == 'secind':
                    if dup_asset:
                        update_secind = """
                            UPDATE assetinvsecind SET amount = amount + %s 
                            WHERE assetinvid = %s AND sec_id = %s AND ind_id = %s
                        """
                        self.db.execute_update(update_secind, (allocated_amount, assetinv_id, int(tval1), int(tval2)))
                    else:
                        insert_secind = """
                            INSERT INTO assetinvsecind(assetinvid, sec_id, ind_id, amount) 
                            VALUES (%s, %s, %s, %s)
                        """
                        self.db.execute_update(insert_secind, (assetinv_id, int(tval1), int(tval2), allocated_amount))
                
                elif tcode == 'inter':
                    if dup_asset:
                        update_inter = """
                            UPDATE assetinvinter SET amount = amount + %s 
                            WHERE assetinvid = %s AND intercode = %s
                        """
                        self.db.execute_update(update_inter, (allocated_amount, assetinv_id, int(tval1)))
                    else:
                        insert_inter = """
                            INSERT INTO assetinvinter(assetinvid, intercode, amount) 
                            VALUES (%s, %s, %s)
                        """
                        self.db.execute_update(insert_inter, (assetinv_id, int(tval1), allocated_amount))
            
            self.db.commit()
            print(f"Successfully allocated asset reference {asset_id} at {held_at}")
            
        except Exception as e:
            self.db.rollback()
            print(f"Error in allocate_asset_ref: {e}")
            raise
    
    def delete_asset_info(self, as_of_date):
        """Delete asset information for a given date"""
        self.db.open_db()
        
        try:
            date_str = self.mysql_date(as_of_date)
            
            # Delete gains for the date
            self.db.execute_update("DELETE FROM assetgain WHERE assetdate=%s", (date_str,))
            
            # Delete old gains (older than 24 months)
            old_date = as_of_date - timedelta(days=730)
            old_date_str = self.mysql_date(old_date)
            self.db.execute_update("DELETE FROM assetgain WHERE assetdate<%s", (old_date_str,))
            
            # Get all assetinvid for the date and delete related records
            query = "SELECT assetinvid FROM assetinv WHERE asofdate=%s"
            results = self.db.execute_query(query, (date_str,))
            
            for row in results:
                assetinv_id = row['assetinvid']
                self.db.execute_update("DELETE FROM assetinvalloc WHERE assetinvid=%s", (assetinv_id,))
                self.db.execute_update("DELETE FROM assetinvsecind WHERE assetinvid=%s", (assetinv_id,))
                self.db.execute_update("DELETE FROM assetinvinter WHERE assetinvid=%s", (assetinv_id,))
                self.db.execute_update("DELETE FROM assetinv WHERE assetinvid=%s", (assetinv_id,))
            
            self.db.commit()
            print(f"Deleted asset info for date {date_str}")
            
        except Exception as e:
            print(f"Error deleting asset info: {e}")
            raise
        finally:
            self.db.close_db()


class GainCalculator:
    """Handles gain/performance calculations from Yahoo Finance and Morningstar"""
    
    def __init__(self, db: AssetDatabase):
        self.db = db
    
    @staticmethod
    def is_market_open(dt: datetime, db: AssetDatabase) -> bool:
        """Check if market is open on a given date"""
        # Check if weekend
        if dt.weekday() in [5, 6]:  # Saturday or Sunday
            return False
        
        # Check if holiday
        date_str = AssetAllocator.mysql_date(dt)
        query = "SELECT * FROM holiday WHERE YEAR(holiday_date)=%s AND holiday_date=%s"
        results = db.execute_query(query, (dt.year, date_str))
        
        return len(results) == 0
    
    def trading_date_add(self, dtype: str, num_periods: int, dt: datetime) -> datetime:
        """Add periods to date, adjusting for market closures"""
        if dtype == 'ww':  # weeks
            target_date = dt - timedelta(weeks=num_periods)
        elif dtype == 'd':  # days
            target_date = dt - timedelta(days=num_periods)
        else:
            target_date = dt
        
        # Adjust for non-trading days
        if not self.is_market_open(target_date, self.db):
            for i in range(3):
                target_date = target_date - timedelta(days=1)
                if self.is_market_open(target_date, self.db):
                    break
        
        return target_date
    
    def calc_gain_from_yahoo(self, ticker: str, dates: List[datetime]) -> bool:
        """Calculate gains from Yahoo Finance"""
        print(f"Calculating gains from Yahoo for {ticker}")
        
        if ticker == 'FCASH':
            return False
        
        try:
            # Build Yahoo Finance CSV URL
            start_date = dates[-1]
            end_date = dates[0]
            
            url = (f"https://query1.finance.yahoo.com/v7/finance/download/{ticker}"
                   f"?period1={int(start_date.timestamp())}"
                   f"&period2={int(end_date.timestamp())}"
                   f"&interval=1d&events=history")
            
            response = requests.get(url)
            if response.status_code != 200:
                print(f"Failed to fetch data for {ticker}")
                return False
            
            # Parse CSV data
            csv_data = csv.DictReader(StringIO(response.text))
            price_data = {}
            
            for row in csv_data:
                date_str = row['Date']
                close_price = float(row['Close'])
                price_data[date_str] = close_price
            
            # Calculate gains
            gains = [0] * (len(dates))
            curr_price = None
            
            for i, date in enumerate(dates):
                date_str = date.strftime('%Y-%m-%d')
                if date_str in price_data:
                    if i == 0:
                        curr_price = price_data[date_str]
                    else:
                        prev_price = price_data[date_str]
                        if curr_price and prev_price:
                            gains[i] = round(((curr_price - prev_price) / prev_price) * 100, 2)
            
            # Insert into database
            self.db.open_db()
            date_str = AssetAllocator.mysql_date(dates[0])
            
            insert_query = """
                INSERT INTO assetgain 
                (ticker, assetdate, oneweekgain, twoweekgain, onemonthgain, 
                 threemonthgain, sixmonthgain, oneyeargain) 
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """
            self.db.execute_update(insert_query, (
                ticker, date_str, 
                gains[1] if len(gains) > 1 else 0,
                gains[2] if len(gains) > 2 else 0,
                gains[3] if len(gains) > 3 else 0,
                gains[4] if len(gains) > 4 else 0,
                gains[5] if len(gains) > 5 else 0,
                gains[6] if len(gains) > 6 else 0
            ))
            self.db.commit()
            self.db.close_db()
            
            print(f"Successfully calculated gains for {ticker}")
            return True
            
        except Exception as e:
            print(f"Error calculating gains from Yahoo for {ticker}: {e}")
            return False
    
    def calc_gain_from_morningstar(self, ticker: str, dates: List[datetime]) -> bool:
        """Calculate gains from Morningstar"""
        print(f"Attempting to calculate gains from Morningstar for {ticker}")
        
        try:
            url = f"https://performance.morningstar.com/Performance/fund/trailing-total-returns.action?t={ticker}&ops=clear"
            response = requests.get(url, timeout=10)
            
            if response.status_code != 200:
                return False
            
            # Parse HTML for performance data
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # This is a simplified version - actual parsing would need to match the HTML structure
            # For now, return False to fall back to Yahoo
            return False
            
        except Exception as e:
            print(f"Error fetching from Morningstar: {e}")
            return False
    
    def calculate_gains(self, as_of_date: datetime):
        """Calculate gains for all assets"""
        self.db.open_db()
        
        try:
            # Adjust date if weekend
            dt_today = as_of_date
            if dt_today.weekday() == 6:  # Sunday
                dt_today = dt_today - timedelta(days=2)
            elif dt_today.weekday() == 5:  # Saturday
                dt_today = dt_today - timedelta(days=1)
            
            # Prepare date array
            dates = [
                dt_today,
                self.trading_date_add('ww', 1, dt_today),
                self.trading_date_add('ww', 2, dt_today),
                self.trading_date_add('ww', 4, dt_today),
                self.trading_date_add('ww', 12, dt_today),
                self.trading_date_add('ww', 24, dt_today),
                self.trading_date_add('ww', 52, dt_today)
            ]
            
            # Get distinct tickers
            query = "SELECT DISTINCT ticker FROM asset WHERE benchmark != '' AND benchmark IS NOT NULL"
            results = self.db.execute_query(query)
            
            for row in results:
                ticker = row['ticker']
                # Try Morningstar first, fall back to Yahoo
                if not self.calc_gain_from_morningstar(ticker, dates):
                    self.calc_gain_from_yahoo(ticker, dates)
            
            print("Gain calculation completed")
            
        finally:
            self.db.close_db()


class TemplateManager:
    """Manages asset templates"""
    
    def __init__(self, db: AssetDatabase):
        self.db = db
    
    def add_template_detail_alloc(self, template_id: int, allocations: List[Tuple[str, float]]):
        """Add allocation template details"""
        self.db.open_db()
        
        try:
            for alloc_name, alloc_prct in allocations:
                # Get alloccode
                query = "SELECT alloccode FROM alloctype WHERE allocdesc=%s"
                results = self.db.execute_query(query, (alloc_name,))
                
                if results:
                    alloccode = results[0]['alloccode']
                    
                    insert_query = """
                        INSERT INTO templatedetails(templateid, tcode, tval1, tval2, prct) 
                        VALUES (%s, 'alloc', %s, '', %s)
                    """
                    self.db.execute_update(insert_query, (template_id, alloccode, alloc_prct))
            
            self.db.commit()
            print(f"Added allocation template details for template {template_id}")
            
        finally:
            self.db.close_db()
    
    def add_template_detail_inter(self, template_id: int, interests: List[Tuple[str, float]]):
        """Add interest template details"""
        self.db.open_db()
        
        try:
            for inter_name, alloc_prct in interests:
                # Get intercode
                query = "SELECT intercode FROM inter WHERE inter_name=%s"
                results = self.db.execute_query(query, (inter_name,))
                
                if results:
                    intercode = results[0]['intercode']
                    
                    insert_query = """
                        INSERT INTO templatedetails(templateid, tcode, tval1, tval2, prct) 
                        VALUES (%s, 'inter', %s, '', %s)
                    """
                    self.db.execute_update(insert_query, (template_id, intercode, alloc_prct))
            
            self.db.commit()
            print(f"Added interest template details for template {template_id}")
            
        finally:
            self.db.close_db()
    
    def add_template_detail_secind(self, template_id: int, sectors: List[Tuple[str, str, float]]):
        """Add sector/industry template details"""
        self.db.open_db()
        
        try:
            for sector_name, ind_name, alloc_prct in sectors:
                # Get sec_id
                query = "SELECT sec_id FROM sector WHERE sec_name=%s"
                results = self.db.execute_query(query, (sector_name,))
                
                if results:
                    sec_id = results[0]['sec_id']
                    
                    # Get ind_id if provided
                    ind_id = 0
                    if ind_name and ind_name != '0':
                        ind_query = "SELECT ind_id FROM industry WHERE ind_name=%s"
                        ind_results = self.db.execute_query(ind_query, (ind_name,))
                        if ind_results:
                            ind_id = ind_results[0]['ind_id']
                    
                    insert_query = """
                        INSERT INTO templatedetails(templateid, tcode, tval1, tval2, prct) 
                        VALUES (%s, 'secind', %s, %s, %s)
                    """
                    self.db.execute_update(insert_query, (template_id, sec_id, ind_id, alloc_prct))
            
            self.db.commit()
            print(f"Added sector/industry template details for template {template_id}")
            
        finally:
            self.db.close_db()
    
    def delete_template_details(self, template_id: int):
        """Delete all template details for a template"""
        self.db.open_db()
        
        try:
            self.db.execute_update("DELETE FROM templatedetails WHERE templateid=%s", (template_id,))
            self.db.commit()
            print(f"Deleted template details for template {template_id}")
        finally:
            self.db.close_db()
