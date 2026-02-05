"""
Análisis de Oportunidades de Salesforce
========================================
Genera KPIs por responsable, país y categorías de KPI
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

# Configuración para mostrar todas las columnas
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_rows', 100)

def cargar_datos():
    """Cargar el archivo CSV de oportunidades"""
    archivo = "20260203_Detailed Opportunity Records.csv"
    df = pd.read_csv(archivo, encoding='utf-8')
    
    # Limpiar columnas de fecha
    df['CreatedDate'] = pd.to_datetime(df['CreatedDate'], errors='coerce')
    df['CloseDate'] = pd.to_datetime(df['CloseDate'], errors='coerce')
    
    # Limpiar columna USD (convertir a numérico)
    df['USD'] = pd.to_numeric(df['USD'], errors='coerce').fillna(0)
    
    # Limpiar valores nulos en Responsible
    df['Responsible'] = df['Responsible'].fillna('Sin Asignar')
    
    return df

def generar_resumen_general(df):
    """Genera estadísticas generales del dataset"""
    print("\n" + "="*80)
    print("📊 RESUMEN GENERAL DEL DATASET")
    print("="*80)
    
    print(f"\n📈 Total de Oportunidades: {len(df):,}")
    print(f"💰 Valor Total USD: ${df['USD'].sum():,.2f}")
    print(f"📅 Rango de Fechas de Creación: {df['CreatedDate'].min()} a {df['CreatedDate'].max()}")
    print(f"📅 Rango de Fechas de Cierre: {df['CloseDate'].min()} a {df['CloseDate'].max()}")
    
    print("\n📌 Distribución por Categoría de KPI:")
    kpi_counts = df['KPI'].value_counts()
    for kpi, count in kpi_counts.items():
        pct = count/len(df)*100
        usd_sum = df[df['KPI']==kpi]['USD'].sum()
        print(f"   • {kpi}: {count:,} oportunidades ({pct:.1f}%) - USD ${usd_sum:,.2f}")
    
    print("\n🌎 Distribución por Región:")
    region_counts = df['Region'].value_counts()
    for region, count in region_counts.items():
        pct = count/len(df)*100
        usd_sum = df[df['Region']==region]['USD'].sum()
        print(f"   • {region}: {count:,} oportunidades ({pct:.1f}%) - USD ${usd_sum:,.2f}")

def analisis_por_responsable(df):
    """Análisis detallado por responsable"""
    print("\n" + "="*80)
    print("👤 ANÁLISIS POR RESPONSABLE")
    print("="*80)
    
    # Agrupar por responsable
    resumen = df.groupby('Responsible').agg({
        'Id': 'count',
        'USD': ['sum', 'mean'],
        'KPI': lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else 'N/A'
    }).round(2)
    
    resumen.columns = ['Total_Oportunidades', 'Total_USD', 'Promedio_USD', 'KPI_Principal']
    resumen = resumen.sort_values('Total_Oportunidades', ascending=False)
    
    print("\n📋 Resumen por Responsable (Top 20):")
    print("-"*80)
    
    for idx, (responsable, row) in enumerate(resumen.head(20).iterrows()):
        print(f"\n{idx+1}. {responsable}")
        print(f"   📊 Total Oportunidades: {row['Total_Oportunidades']:,}")
        print(f"   💵 Valor Total USD: ${row['Total_USD']:,.2f}")
        print(f"   📈 Promedio USD: ${row['Promedio_USD']:,.2f}")
        print(f"   🏷️ KPI Principal: {row['KPI_Principal']}")
        
        # Distribución por KPI para este responsable
        kpi_dist = df[df['Responsible']==responsable]['KPI'].value_counts()
        print(f"   📌 Distribución KPI: ", end="")
        print(", ".join([f"{k}: {v}" for k, v in kpi_dist.items()]))
        
        # Países trabajados
        paises = df[df['Responsible']==responsable]['Market'].unique()
        print(f"   🌎 Países: {', '.join(paises)}")
    
    return resumen

def analisis_por_pais(df):
    """Análisis detallado por país/mercado"""
    print("\n" + "="*80)
    print("🌍 ANÁLISIS POR PAÍS/MERCADO")
    print("="*80)
    
    # Agrupar por mercado
    resumen_pais = df.groupby('Market').agg({
        'Id': 'count',
        'USD': ['sum', 'mean'],
        'Responsible': 'nunique',
        'Customer': 'nunique',
        'KPI': lambda x: ', '.join(x.unique())
    }).round(2)
    
    resumen_pais.columns = ['Total_Oportunidades', 'Total_USD', 'Promedio_USD', 
                            'Num_Responsables', 'Num_Clientes', 'KPIs']
    resumen_pais = resumen_pais.sort_values('Total_Oportunidades', ascending=False)
    
    print("\n📋 Resumen por País:")
    print("-"*80)
    
    for idx, (pais, row) in enumerate(resumen_pais.iterrows()):
        print(f"\n{idx+1}. {pais}")
        print(f"   📊 Total Oportunidades: {row['Total_Oportunidades']:,}")
        print(f"   💵 Valor Total USD: ${row['Total_USD']:,.2f}")
        print(f"   📈 Promedio USD por Oportunidad: ${row['Promedio_USD']:,.2f}")
        print(f"   👥 Número de Responsables: {row['Num_Responsables']}")
        print(f"   🏢 Número de Clientes: {row['Num_Clientes']}")
        print(f"   🏷️ KPIs: {row['KPIs']}")
        
        # Top responsables por país
        top_resp = df[df['Market']==pais].groupby('Responsible')['Id'].count().sort_values(ascending=False).head(3)
        print(f"   👤 Top Responsables: ", end="")
        print(", ".join([f"{r}: {c}" for r, c in top_resp.items()]))
        
        # Top clientes por país
        top_clientes = df[df['Market']==pais].groupby('Customer')['Id'].count().sort_values(ascending=False).head(3)
        print(f"   🏢 Top Clientes: ", end="")
        print(", ".join([f"{c}: {n}" for c, n in top_clientes.items()]))
    
    return resumen_pais

def analisis_por_kpi(df):
    """Análisis detallado por categoría de KPI"""
    print("\n" + "="*80)
    print("📊 ANÁLISIS POR CATEGORÍA DE KPI")
    print("="*80)
    
    for kpi in df['KPI'].unique():
        df_kpi = df[df['KPI']==kpi]
        
        print(f"\n{'='*60}")
        print(f"🏷️ KPI: {kpi}")
        print(f"{'='*60}")
        print(f"   📊 Total Oportunidades: {len(df_kpi):,}")
        print(f"   💵 Valor Total USD: ${df_kpi['USD'].sum():,.2f}")
        print(f"   📈 Promedio USD: ${df_kpi['USD'].mean():,.2f}")
        
        # Por país
        print("\n   🌎 Por País:")
        by_country = df_kpi.groupby('Market').agg({'Id': 'count', 'USD': 'sum'}).sort_values('Id', ascending=False)
        for pais, row in by_country.iterrows():
            print(f"      • {pais}: {row['Id']} opps (${row['USD']:,.2f})")
        
        # Por responsable (top 5)
        print("\n   👤 Top 5 Responsables:")
        by_resp = df_kpi.groupby('Responsible').agg({'Id': 'count', 'USD': 'sum'}).sort_values('Id', ascending=False).head(5)
        for resp, row in by_resp.iterrows():
            print(f"      • {resp}: {row['Id']} opps (${row['USD']:,.2f})")
        
        # Por stage
        print("\n   📍 Por Etapa (Stage):")
        by_stage = df_kpi['Stage'].value_counts()
        for stage, count in by_stage.items():
            print(f"      • {stage}: {count}")

def analisis_cruzado_responsable_pais(df):
    """Análisis cruzado entre responsable y país"""
    print("\n" + "="*80)
    print("🔀 MATRIZ RESPONSABLE vs PAÍS")
    print("="*80)
    
    matriz = pd.crosstab(df['Responsible'], df['Market'], values=df['Id'], aggfunc='count', margins=True)
    print("\nMatriz de Oportunidades (Responsable x País):")
    print(matriz.to_string())
    
    # Matriz de valor USD
    print("\n\nMatriz de Valor USD (Responsable x País):")
    matriz_usd = pd.crosstab(df['Responsible'], df['Market'], values=df['USD'], aggfunc='sum', margins=True)
    matriz_usd = matriz_usd.round(2)
    print(matriz_usd.to_string())

def analisis_por_stage(df):
    """Análisis por etapa del proceso"""
    print("\n" + "="*80)
    print("📍 ANÁLISIS POR ETAPA (STAGE)")
    print("="*80)
    
    stage_summary = df.groupby('Stage').agg({
        'Id': 'count',
        'USD': ['sum', 'mean']
    }).round(2)
    
    stage_summary.columns = ['Total_Oportunidades', 'Total_USD', 'Promedio_USD']
    stage_summary = stage_summary.sort_values('Total_Oportunidades', ascending=False)
    
    print("\n📋 Resumen por Etapa:")
    for stage, row in stage_summary.iterrows():
        pct = row['Total_Oportunidades']/len(df)*100
        print(f"   • {stage}: {row['Total_Oportunidades']:,} ({pct:.1f}%) - ${row['Total_USD']:,.2f}")
    
    return stage_summary

def analisis_por_cliente(df):
    """Análisis por cliente"""
    print("\n" + "="*80)
    print("🏢 ANÁLISIS POR CLIENTE")
    print("="*80)
    
    cliente_summary = df.groupby('Customer').agg({
        'Id': 'count',
        'USD': ['sum', 'mean'],
        'Market': lambda x: ', '.join(x.unique()),
        'Responsible': 'nunique'
    }).round(2)
    
    cliente_summary.columns = ['Total_Oportunidades', 'Total_USD', 'Promedio_USD', 'Países', 'Num_Responsables']
    cliente_summary = cliente_summary.sort_values('Total_Oportunidades', ascending=False)
    
    print("\n📋 Top 15 Clientes:")
    for idx, (cliente, row) in enumerate(cliente_summary.head(15).iterrows()):
        print(f"\n{idx+1}. {cliente}")
        print(f"   📊 Oportunidades: {row['Total_Oportunidades']:,}")
        print(f"   💵 Valor Total: ${row['Total_USD']:,.2f}")
        print(f"   📈 Promedio: ${row['Promedio_USD']:,.2f}")
        print(f"   🌎 Países: {row['Países']}")
        print(f"   👥 Responsables: {row['Num_Responsables']}")
    
    return cliente_summary

def generar_kpis_ejecutivos(df):
    """Genera KPIs ejecutivos consolidados"""
    print("\n" + "="*80)
    print("📊 KPIs EJECUTIVOS CONSOLIDADOS")
    print("="*80)
    
    # KPIs generales
    total_opps = len(df)
    total_usd = df['USD'].sum()
    avg_usd = df['USD'].mean()
    
    # Oportunidades con valor > 0
    opps_con_valor = df[df['USD'] > 0]
    pct_con_valor = len(opps_con_valor)/total_opps*100
    
    # Por tipo de KPI
    dc002_nb = df[df['KPI']=='DC002 NB']
    dc002_churn = df[df['KPI']=='DC002 CHURN']
    dc004 = df[df['KPI']=='DC004']
    
    print("\n🎯 MÉTRICAS GENERALES:")
    print(f"   • Total Oportunidades: {total_opps:,}")
    print(f"   • Valor Total USD: ${total_usd:,.2f}")
    print(f"   • Promedio USD por Oportunidad: ${avg_usd:,.2f}")
    print(f"   • Oportunidades con Valor > $0: {len(opps_con_valor):,} ({pct_con_valor:.1f}%)")
    
    print("\n🏷️ POR TIPO DE KPI:")
    print(f"   • DC002 NB: {len(dc002_nb):,} opps - ${dc002_nb['USD'].sum():,.2f}")
    print(f"   • DC002 CHURN: {len(dc002_churn):,} opps - ${dc002_churn['USD'].sum():,.2f}")
    print(f"   • DC004: {len(dc004):,} opps - ${dc004['USD'].sum():,.2f}")
    
    print("\n👥 TOP 5 RESPONSABLES POR VOLUMEN:")
    top_resp_vol = df.groupby('Responsible')['Id'].count().sort_values(ascending=False).head(5)
    for i, (resp, count) in enumerate(top_resp_vol.items(), 1):
        usd = df[df['Responsible']==resp]['USD'].sum()
        print(f"   {i}. {resp}: {count:,} opps (${usd:,.2f})")
    
    print("\n💰 TOP 5 RESPONSABLES POR VALOR USD:")
    top_resp_usd = df.groupby('Responsible')['USD'].sum().sort_values(ascending=False).head(5)
    for i, (resp, usd) in enumerate(top_resp_usd.items(), 1):
        count = len(df[df['Responsible']==resp])
        print(f"   {i}. {resp}: ${usd:,.2f} ({count:,} opps)")
    
    print("\n🌎 RESUMEN POR PAÍS:")
    for pais in df['Market'].unique():
        df_pais = df[df['Market']==pais]
        print(f"   • {pais}: {len(df_pais):,} opps - ${df_pais['USD'].sum():,.2f}")

def exportar_a_excel(df, resumen_resp, resumen_pais, resumen_cliente, resumen_stage):
    """Exporta los análisis a Excel"""
    output_file = "Analisis_Oportunidades_KPI.xlsx"
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Datos originales
        df.to_excel(writer, sheet_name='Datos_Originales', index=False)
        
        # Resumen por responsable
        resumen_resp.to_excel(writer, sheet_name='Por_Responsable')
        
        # Resumen por país
        resumen_pais.to_excel(writer, sheet_name='Por_País')
        
        # Resumen por cliente
        resumen_cliente.to_excel(writer, sheet_name='Por_Cliente')
        
        # Resumen por stage
        resumen_stage.to_excel(writer, sheet_name='Por_Stage')
        
        # Matriz cruzada
        matriz = pd.crosstab(df['Responsible'], df['Market'], values=df['Id'], aggfunc='count', margins=True)
        matriz.to_excel(writer, sheet_name='Matriz_Resp_País')
        
        matriz_usd = pd.crosstab(df['Responsible'], df['Market'], values=df['USD'], aggfunc='sum', margins=True)
        matriz_usd.to_excel(writer, sheet_name='Matriz_USD_Resp_País')
    
    print(f"\n✅ Análisis exportado a: {output_file}")

def main():
    """Función principal"""
    print("\n" + "🔷"*40)
    print("ANÁLISIS DE OPORTUNIDADES DE SALESFORCE")
    print("Enfoque: KPIs por Responsable, País y Categoría")
    print(f"Fecha de análisis: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("🔷"*40)
    
    # Cargar datos
    print("\n📂 Cargando datos...")
    df = cargar_datos()
    print(f"   ✓ {len(df):,} registros cargados")
    
    # Ejecutar análisis
    generar_resumen_general(df)
    resumen_resp = analisis_por_responsable(df)
    resumen_pais = analisis_por_pais(df)
    analisis_por_kpi(df)
    analisis_cruzado_responsable_pais(df)
    resumen_stage = analisis_por_stage(df)
    resumen_cliente = analisis_por_cliente(df)
    generar_kpis_ejecutivos(df)
    
    # Exportar a Excel
    try:
        exportar_a_excel(df, resumen_resp, resumen_pais, resumen_cliente, resumen_stage)
    except ImportError:
        print("\n⚠️ Para exportar a Excel, instale openpyxl: pip install openpyxl")
    
    print("\n" + "="*80)
    print("✅ ANÁLISIS COMPLETADO")
    print("="*80)

if __name__ == "__main__":
    main()
