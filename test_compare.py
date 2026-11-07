#!/usr/bin/env python3

from process_assets import AssetProcessor
from datetime import datetime

ap = AssetProcessor()
prevdate = datetime(2026, 3, 20)
currdate = datetime(2026, 3, 27)

# Check if Fidelity data is present for 2026-03-27
ap.db.open_db()
query = """
SELECT COUNT(*) as count FROM assetinv 
WHERE heldat LIKE 'Fidelity%' 
AND DATE(asofdate) = '2026-03-27'
"""
result = ap.db.execute_query(query)
fidelity_count = result[0]['count'] if result else 0
print(f"Fidelity entries on 2026-03-27: {fidelity_count}\n")
ap.db.close_db()

ap.compare_dates_report(currdate, prevdate, show_all=True)
