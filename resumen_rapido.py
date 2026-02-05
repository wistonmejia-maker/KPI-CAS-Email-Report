import pandas as pd
import numpy as np

# Cargar datos
df = pd.read_csv('20260203_Detailed Opportunity Records.csv', encoding='utf-8')
df['USD'] = pd.to_numeric(df['USD'], errors='coerce').fillna(0)
df['Responsible'] = df['Responsible'].fillna('Sin Asignar')

print('='*80)
print('ANALISIS PROFUNDO DE OPORTUNIDADES DE SALESFORCE')
print('='*80)

print(f'\n>>> RESUMEN GENERAL')
print(f'   Total de Oportunidades: {len(df):,}')
print(f'   Valor Total USD: ${df["USD"].sum():,.2f}')
print(f'   Promedio USD: ${df["USD"].mean():,.2f}')

print(f'\n>>> DISTRIBUCION POR KPI:')
kpi_counts = df.groupby('KPI').agg({'Id':'count', 'USD':'sum'}).sort_values('Id', ascending=False)
for kpi, row in kpi_counts.iterrows():
    pct = row['Id']/len(df)*100
    print(f'   - {kpi}: {int(row["Id"]):,} opps ({pct:.1f}%) - ${row["USD"]:,.2f}')

print(f'\n>>> DISTRIBUCION POR PAIS/MERCADO:')
market_counts = df.groupby('Market').agg({'Id':'count', 'USD':'sum', 'Responsible':'nunique'}).sort_values('Id', ascending=False)
for market, row in market_counts.iterrows():
    pct = row['Id']/len(df)*100
    print(f'   - {market}: {int(row["Id"]):,} opps ({pct:.1f}%) - ${row["USD"]:,.2f} - {int(row["Responsible"])} responsables')

print(f'\n>>> TOP 15 RESPONSABLES POR VOLUMEN:')
resp_vol = df.groupby('Responsible').agg({'Id':'count', 'USD':'sum', 'Market':lambda x: ", ".join(x.unique())}).sort_values('Id', ascending=False).head(15)
for i, (resp, row) in enumerate(resp_vol.iterrows(), 1):
    print(f'   {i:2d}. {resp}: {int(row["Id"]):,} opps (${row["USD"]:,.2f}) - {row["Market"]}')

print(f'\n>>> TOP 10 RESPONSABLES POR VALOR USD:')
resp_usd = df.groupby('Responsible').agg({'USD':'sum', 'Id':'count'}).sort_values('USD', ascending=False).head(10)
for i, (resp, row) in enumerate(resp_usd.iterrows(), 1):
    print(f'   {i:2d}. {resp}: ${row["USD"]:,.2f} ({int(row["Id"]):,} opps)')

print(f'\n>>> ANALISIS POR PAIS - TOP RESPONSABLES:')
for market in df['Market'].unique():
    df_market = df[df['Market']==market]
    print(f'\n   [{market}] - {len(df_market)} opps (${df_market["USD"].sum():,.2f})')
    top_resp = df_market.groupby('Responsible').agg({'Id':'count', 'USD':'sum'}).sort_values('Id', ascending=False).head(5)
    for resp, row in top_resp.iterrows():
        print(f'      - {resp}: {int(row["Id"])} opps (${row["USD"]:,.2f})')

print(f'\n>>> ANALISIS POR KPI Y PAIS:')
for kpi in df['KPI'].unique():
    df_kpi = df[df['KPI']==kpi]
    print(f'\n   [{kpi}] - {len(df_kpi)} opps (${df_kpi["USD"].sum():,.2f})')
    by_country = df_kpi.groupby('Market').agg({'Id':'count', 'USD':'sum'}).sort_values('Id', ascending=False)
    for market, row in by_country.iterrows():
        print(f'      - {market}: {int(row["Id"])} opps (${row["USD"]:,.2f})')

print(f'\n>>> DISTRIBUCION POR STAGE:')
stage_counts = df['Stage'].value_counts()
for stage, count in stage_counts.items():
    pct = count/len(df)*100
    usd = df[df['Stage']==stage]['USD'].sum()
    print(f'   - {stage}: {count} opps ({pct:.1f}%) - ${usd:,.2f}')

print(f'\n>>> TOP 10 CLIENTES:')
cliente_counts = df.groupby('Customer').agg({'Id':'count', 'USD':'sum'}).sort_values('Id', ascending=False).head(10)
for i, (cliente, row) in enumerate(cliente_counts.iterrows(), 1):
    print(f'   {i:2d}. {cliente}: {int(row["Id"])} opps (${row["USD"]:,.2f})')

print('\n' + '='*80)
print('ANALISIS COMPLETADO')
print('='*80)
