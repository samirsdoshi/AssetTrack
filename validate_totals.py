from openpyxl import load_workbook

wb = load_workbook('Asset.xlsx', data_only=True)

# Read fullview totals from columns N and O
# But skip stock account entries (N, O, P structure) and only get summary totals
ws_fullview = wb['fullview']
fullview_totals = {}
print('=== Fullview Account Totals (from K, L aggregated) ===')
# Aggregate from K and L columns
for row in range(1, 1000):
    account_ticker = ws_fullview[f'K{row}'].value
    amount = ws_fullview[f'L{row}'].value
    if account_ticker and amount and '_' in str(account_ticker):
        account = str(account_ticker).split('_')[0]
        if account not in fullview_totals:
            fullview_totals[account] = 0
        fullview_totals[account] += amount

for account in sorted(fullview_totals.keys()):
    print(f'{account}: ${fullview_totals[account]:,.2f}')

# Read assetalloc column J (and H for account names)
ws_assetalloc = wb['assetAlloc']
assetalloc_totals = {}
print('\n=== AssetAlloc Account Totals (Column H, J) ===')
for row in range(2, 20):
    account = ws_assetalloc[f'H{row}'].value
    total_j = ws_assetalloc[f'J{row}'].value
    if account and total_j:
        assetalloc_totals[account] = total_j
        print(f'{account}: ${total_j:,.2f}')

# Compare
print('\n=== Comparison ===')
all_accounts = set(list(fullview_totals.keys()) + list(assetalloc_totals.keys()))
mismatches = []
for account in sorted(all_accounts):
    fv_total = fullview_totals.get(account, 0)
    aa_total = assetalloc_totals.get(account, 0)
    match = '✓' if abs(fv_total - aa_total) < 1 else '✗'
    if match == '✗':
        mismatches.append(account)
    print(f'{match} {account:20s} Fullview: ${fv_total:,.2f}  AssetAlloc: ${aa_total:,.2f}')

if mismatches:
    print(f'\n{len(mismatches)} mismatches found!')
else:
    print('\nAll accounts match!')
