"""
Utility Functions Module
Converted from VBA module2.vba
Provides helper functions for HTML parsing, string manipulation, and data processing
"""

import re
from urllib.parse import quote_plus
from typing import Optional


def get_cell(str_ret_val: str, cell_number: int) -> str:
    """
    Extract cell content from HTML table by cell number
    
    Args:
        str_ret_val: HTML string with numbered cells
        cell_number: Cell number to extract
        
    Returns:
        Cell content as string
    """
    cell_marker = f"{cell_number}|"
    i = str_ret_val.upper().find(cell_marker)
    
    if i == -1:
        return ""
    
    # Find the > before the cell marker
    r = str_ret_val.upper().rfind(">", 0, i)
    if r == -1:
        return ""
    
    start_cell_text = r + 1
    
    # Find the end of this cell's text
    # Look for either <TABLE or </TD> - whichever comes first
    table_pos = str_ret_val.upper().find("<TABLE", r)
    td_end_pos = str_ret_val.upper().find("</TD>", r)
    
    if table_pos > 0 and table_pos < td_end_pos:
        this_cell_text = str_ret_val[start_cell_text:table_pos]
    else:
        this_cell_text = str_ret_val[start_cell_text + len(cell_marker):td_end_pos]
    
    return this_cell_text


def number_cell(str_result: str) -> tuple[str, int]:
    """
    Number all cells in an HTML table
    
    Args:
        str_result: HTML string containing table
        
    Returns:
        Tuple of (numbered HTML string, max cell count)
    """
    q = 1
    i = 0
    
    while True:
        # Find next <TD
        i = str_result.upper().find("<TD", i)
        if i == -1:
            break
        
        # Find the end of the <TD tag
        r = str_result.find(">", i)
        if r == -1:
            break
        
        # Insert cell number after the >
        str_result = str_result[:r + 1] + f"{q}|" + str_result[r + 1:]
        
        # Move past this tag
        i = r + 1
        q += 1
    
    max_cell = q
    return str_result, max_cell


def clean_up(str_val: str) -> str:
    """
    Clean up HTML and special characters from string
    
    Args:
        str_val: String to clean
        
    Returns:
        Cleaned string
    """
    replacements = {
        "<b>": "",
        "</b>": "",
        "&nbsp;": "",
        "$": "",
        "</strong>": "",
        "<strong>": "",
        '<font color="red">': "",
        "</font>": "",
        "&mdash;": "",
        ",": ""  # Remove commas from numbers
    }
    
    for old, new in replacements.items():
        str_val = str_val.replace(old, new)
    
    return str_val.strip()


def check_num(str_val: str) -> str:
    """
    Check if string is numeric, return '0' if not
    
    Args:
        str_val: String to check
        
    Returns:
        Original string if numeric, '0' otherwise
    """
    try:
        float(str_val)
        return str_val
    except (ValueError, TypeError):
        return "0"


def remove_whitespace(str_text: str) -> str:
    """
    Remove extra whitespace from string
    
    Args:
        str_text: String to process
        
    Returns:
        String with normalized whitespace
    """
    # Replace multiple whitespace with single space
    return re.sub(r'\s+', ' ', str_text).strip()


def min_val(n1: float, n2: float) -> float:
    """Return minimum of two numbers"""
    return min(n1, n2)


def max_val(n1: float, n2: float) -> float:
    """Return maximum of two numbers"""
    return max(n1, n2)


def url_encode(raw_url: str) -> str:
    """
    URL encode a string
    
    Args:
        raw_url: URL string to encode
        
    Returns:
        URL-encoded string
    """
    if not raw_url:
        return ""
    
    try:
        # Use urllib's quote_plus which handles most cases
        return quote_plus(raw_url)
    except Exception:
        return ""


def nullif(value, default_value):
    """
    Return default_value if value is None, otherwise return value
    
    Args:
        value: Value to check
        default_value: Default value if None
        
    Returns:
        value or default_value
    """
    return default_value if value is None else value


def empty_to_default(value, default_value):
    """
    Return default_value if value is empty/None, otherwise return value
    
    Args:
        value: Value to check
        default_value: Default value if empty
        
    Returns:
        value or default_value
    """
    if value is None:
        return default_value
    
    if isinstance(value, str) and len(value) == 0:
        return default_value
    
    return value


def lpad(value: str, pad_char: str, padded_length: int) -> str:
    """
    Left-pad a string to a specified length
    
    Args:
        value: String to pad
        pad_char: Character to use for padding
        padded_length: Target length
        
    Returns:
        Padded string
    """
    value_str = str(value)
    if len(value_str) < padded_length:
        return pad_char * (padded_length - len(value_str)) + value_str
    return value_str


def rpad(value: str, pad_char: str, padded_length: int) -> str:
    """
    Right-pad a string to a specified length
    
    Args:
        value: String to pad
        pad_char: Character to use for padding
        padded_length: Target length
        
    Returns:
        Padded string
    """
    value_str = str(value)
    if len(value_str) < padded_length:
        return value_str + pad_char * (padded_length - len(value_str))
    return value_str


def filter_ticker(ticker: str) -> str:
    """
    Filter out invalid ticker symbols
    
    Args:
        ticker: Ticker string to filter
        
    Returns:
        Filtered ticker or empty string if invalid
    """
    ticker = ticker.strip()
    
    # Filter out invalid patterns
    if "Go to Site |" in ticker:
        return ""
    
    if ticker.startswith("Symbol"):
        return ""
    
    if ticker.startswith("Total"):
        return ""
    
    if ticker.lower().startswith("samir"):
        return ""
    
    return ticker


def get_held_at(ticker: str) -> str:
    """
    Determine where an asset is held based on ticker/account name
    
    Args:
        ticker: Ticker or account name
        
    Returns:
        Holding location code
    """
    # Replace non-breaking spaces with regular spaces
    ticker = ticker.replace(chr(160), " ")
    
    # Direct mappings
    direct_mappings = {
        "CollegeAdv": "CollegeAdv",
        "FidelityInv": "FidelityInv",
        "FidelityIRA": "FidelityIRA",
        "FidelityRoth": "FidelityRoth",
        "TRPInv": "TRPInv",
        "TRPRoth": "TRPRoth",
        "Vanguard": "Vanguard",
        "Fidelity401k": "Fidelity401k",
        "WellsFargo401k": "WellsFargo401k",
        "Etrade": "Etrade",
        "Robinhood": "Robinhood",
        "Ameritrade": "Ameritrade",
        "TradeStation": "TradeStation"
    }
    
    if ticker in direct_mappings:
        return direct_mappings[ticker]
    
    # Pattern-based mappings
    if "CollegeAdvantage 529 Savings Plan" in ticker:
        return "CollegeAdv"
    
    if "Fidelity Investments" in ticker and "INDIVIDUAL - TOD" in ticker:
        return "FidelityInv"
    
    if "Fidelity Investments" in ticker and "ROLLOVER IRA" in ticker:
        return "FidelityIRA"
    
    if "Fidelity Investments" in ticker and "ROTH IRA" in ticker:
        return "FidelityRoth"
    
    if "T. Rowe Price - Investments - Individual" in ticker:
        return "TRPInv"
    
    if "T. Rowe Price - Investments - Roth IRA" in ticker:
        return "TRPRoth"
    
    if "Vanguard Investments" in ticker:
        return "Vanguard"
    
    if "Fidelity NetBenefits" in ticker and "401(K) PLAN" in ticker:
        return "Fidelity401k"
    
    if "Wells Fargo Retirement Services" in ticker and "401k" in ticker:
        return "WellsFargo401k"
    
    return ""


class FileWriter:
    """Helper class for writing to files"""
    
    def __init__(self):
        self.file_handle = None
    
    def create_file(self, file_path: str):
        """Create/open a file for writing"""
        self.file_handle = open(file_path, 'w', encoding='utf-8')
    
    def write_to_file(self, msg: str):
        """Write a line to the file"""
        if self.file_handle:
            self.file_handle.write(msg + '\n')
    
    def close_file(self):
        """Close the file"""
        if self.file_handle:
            self.file_handle.flush()
            self.file_handle.close()
            self.file_handle = None
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close_file()
