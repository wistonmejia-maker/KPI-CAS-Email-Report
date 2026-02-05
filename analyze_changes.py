import pandas as pd
from pathlib import Path

# Cargar ambos archivos
raw_dir = Path('data/raw')
current = pd.read_csv(raw_dir / '20260204_Detailed Opportunity Records.csv')
previous = pd.read_csv(raw_dir / '20260203_Detailed Opportunity Records.csv')

current['USD'] = pd.to_numeric(current['USD'], errors='coerce').fillna(0)
previous['USD'] = pd.to_numeric(previous['USD'], errors='coerce').fillna(0)

current_ids = set(current['Id'].unique())
previous_ids = set(previous['Id'].unique())

new_ids = current_ids - previous_ids
removed_ids = previous_ids - current_ids

print('='*70)
print('NUEVAS OPORTUNIDADES (' + str(len(new_ids)) + ')')
print('='*70)
new_df = current[current['Id'].isin(new_ids)][['Id', 'Responsible', 'Market', 'KPI', 'USD']]
for _, row in new_df.iterrows():
    print(f"  {row['Id'][:18]:18} | {str(row['Responsible'])[:18]:18} | {str(row['Market']):10} | {str(row['KPI']):12}")

print()
print('='*70)
print('ELIMINADAS (' + str(len(removed_ids)) + ') - Top 20')
print('='*70)
removed_df = previous[previous['Id'].isin(removed_ids)][['Id', 'Responsible', 'Market', 'KPI', 'USD']].head(20)
for _, row in removed_df.iterrows():
    print(f"  {row['Id'][:18]:18} | {str(row['Responsible'])[:18]:18} | {str(row['Market']):10} | {str(row['KPI']):12}")

print()
print('='*70)
print('CAMBIOS DETECTADOS - Stage o Responsable')
print('='*70)
common_ids = current_ids & previous_ids
changes_found = 0
for oid in common_ids:
    curr = current[current['Id'] == oid].iloc[0]
    prev = previous[previous['Id'] == oid].iloc[0]
    
    stage_changed = curr['Stage'] != prev['Stage']
    resp_changed = curr['Responsible'] != prev['Responsible']
    
    if stage_changed or resp_changed:
        changes_found += 1
        if changes_found <= 20:
            if stage_changed:
                print(f"  {oid[:15]:15} | STAGE: {str(prev['Stage'])[:18]} -> {str(curr['Stage'])[:18]}")
            if resp_changed:
                print(f"  {oid[:15]:15} | RESP: {str(prev['Responsible'])[:15]} -> {str(curr['Responsible'])[:15]}")

print(f"\nTotal cambios: {changes_found}")
