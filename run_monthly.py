"""
Run Monthly - Script Principal para Proceso Mensual
====================================================
Ejecuta el proceso mensual de consolidación y generación de snapshots.
"""

import sys
import logging
from pathlib import Path
from datetime import datetime
import argparse

# Agregar src al path
sys.path.insert(0, str(Path(__file__).parent))

from src.config import (
    RAW_DIR, PROCESSED_DIR, SNAPSHOTS_DIR, MONTHLY_REPORTS_DIR,
    ensure_directories, get_current_month_str, get_current_date_str, COLUMNS
)
from src.data_loader import DataLoader, load_opportunities
from src.change_detector import ChangeDetector, compare_datasets
from src.metrics import MetricsCalculator
from src.report_generator import ExcelReportGenerator
from src.html_report_generator import HTMLReportGenerator

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(f'monthly_run_{get_current_date_str()}.log')
    ]
)
logger = logging.getLogger(__name__)


def main():
    """Función principal del proceso mensual"""
    parser = argparse.ArgumentParser(description='Proceso Mensual de KPIs')
    parser.add_argument('--month', '-m', type=str, help='Mes a procesar (YYYY-MM)')
    args = parser.parse_args()
    
    month = args.month or get_current_month_str()
    
    print("\n" + "="*70)
    print("📅 PROCESO MENSUAL DE KPIs - OPORTUNIDADES SALESFORCE")
    print("="*70)
    print(f"📆 Mes: {month}")
    print(f"🕐 Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*70 + "\n")
    
    # Asegurar directorios
    ensure_directories()
    
    # 1. Cargar el archivo más reciente
    print("📂 PASO 1: Cargando datos actuales...")
    loader = DataLoader()
    
    current_file = loader.get_latest_file(RAW_DIR)
    if current_file is None:
        # Buscar en directorio raíz
        project_root = Path(__file__).parent
        csv_files = list(project_root.glob("*.csv"))
        if csv_files:
            current_file = max(csv_files, key=lambda x: x.stat().st_mtime)
    
    if current_file is None:
        logger.error("No se encontró archivo CSV")
        print("❌ ERROR: No se encontró ningún archivo CSV")
        return 1
    
    print(f"   📄 Archivo: {current_file.name}")
    current_df = loader.load_csv(current_file)
    print(f"   ✅ Registros: {len(current_df):,}")
    
    # 2. Crear snapshot mensual
    print("\n📸 PASO 2: Creando snapshot mensual...")
    snapshot_file = SNAPSHOTS_DIR / f"{month}_snapshot.csv"
    SNAPSHOTS_DIR.mkdir(parents=True, exist_ok=True)
    
    # Agregar metadatos al snapshot
    snapshot_df = current_df.copy()
    snapshot_df['_snapshot_date'] = datetime.now()
    snapshot_df['_snapshot_month'] = month
    snapshot_df.to_csv(snapshot_file, index=False)
    print(f"   ✅ Snapshot guardado: {snapshot_file.name}")
    
    # 3. Comparar con mes anterior
    print("\n🔄 PASO 3: Comparando con mes anterior...")
    previous_snapshots = sorted(SNAPSHOTS_DIR.glob("*_snapshot.csv"))
    previous_snapshots = [s for s in previous_snapshots if s != snapshot_file]
    
    comparison = None
    if previous_snapshots:
        previous_snapshot = previous_snapshots[-1]
        print(f"   📄 Snapshot anterior: {previous_snapshot.name}")
        
        previous_df = loader.load_csv(previous_snapshot)
        detector = ChangeDetector()
        comparison = detector.compare(current_df, previous_df)
        
        print(f"\n   📊 Cambios mes a mes:")
        print(f"      • Nuevas: {comparison.summary['new_count']}")
        print(f"      • Eliminadas: {comparison.summary['removed_count']}")
        print(f"      • Modificadas: {comparison.summary['changed_count']}")
        print(f"      • Cambio USD: ${comparison.summary['usd_change']:,.2f}")
    else:
        print("   ⚠️ No hay snapshot anterior para comparar")
    
    # 4. Generar reporte mensual Excel
    print("\n📊 PASO 4: Generando reporte mensual Excel...")
    MONTHLY_REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    
    monthly_excel = MONTHLY_REPORTS_DIR / f"{month}_monthly_report.xlsx"
    
    import pandas as pd
    with pd.ExcelWriter(monthly_excel, engine='openpyxl') as writer:
        # Resumen ejecutivo mensual
        metrics = MetricsCalculator(current_df)
        summary = metrics.get_summary()
        
        summary_data = [
            ['REPORTE MENSUAL', month],
            ['Fecha de Generación', datetime.now().strftime('%Y-%m-%d')],
            ['', ''],
            ['RESUMEN DEL MES', ''],
            ['Total Oportunidades', summary['total_opportunities']],
            ['Valor Total USD', f"${summary['total_usd']:,.2f}"],
            ['Promedio USD', f"${summary['avg_usd']:,.2f}"],
            ['Responsables', summary['unique_responsibles']],
            ['Países', summary['unique_markets']],
            ['Clientes', summary['unique_customers']],
            ['', ''],
            ['ALERTAS', ''],
            ['Oportunidades Estancadas', summary['stagnant_count']],
            ['Oportunidades En Riesgo', summary['at_risk_count']],
        ]
        
        if comparison:
            summary_data.extend([
                ['', ''],
                ['CAMBIOS VS MES ANTERIOR', ''],
                ['Nuevas', comparison.summary['new_count']],
                ['Eliminadas', comparison.summary['removed_count']],
                ['Modificadas', comparison.summary['changed_count']],
                ['Cambio USD', f"${comparison.summary['usd_change']:,.2f}"],
            ])
        
        pd.DataFrame(summary_data, columns=['Métrica', 'Valor']).to_excel(
            writer, sheet_name='Resumen', index=False
        )
        
        # Por responsable
        metrics.get_responsible_summary_df().to_excel(
            writer, sheet_name='Por_Responsable', index=False
        )
        
        # Por país
        metrics.get_market_summary_df().to_excel(
            writer, sheet_name='Por_País', index=False
        )
        
        # Por KPI
        kpi_data = []
        for kpi, data in metrics.get_kpi_metrics().items():
            kpi_data.append({
                'KPI': kpi,
                'Oportunidades': data['count'],
                'Total_USD': data['total_usd'],
                'Promedio_USD': data['avg_usd']
            })
        pd.DataFrame(kpi_data).to_excel(writer, sheet_name='Por_KPI', index=False)
        
        # Por stage
        metrics.get_stage_distribution().to_excel(
            writer, sheet_name='Por_Stage', index=False
        )
        
        # Cambios (si hay comparación)
        if comparison:
            changes_df = comparison.get_changes_df()
            if len(changes_df) > 0:
                changes_df.to_excel(writer, sheet_name='Cambios', index=False)
            
            if len(comparison.new_opportunities) > 0:
                comparison.new_opportunities.to_excel(
                    writer, sheet_name='Nuevas', index=False
                )
            
            if len(comparison.removed_opportunities) > 0:
                comparison.removed_opportunities.to_excel(
                    writer, sheet_name='Eliminadas', index=False
                )
        
        # Datos completos
        current_df.to_excel(writer, sheet_name='Datos', index=False)
    
    print(f"   ✅ Reporte mensual: {monthly_excel}")
    
    # 5. Generar reporte HTML mensual
    print("\n🌐 PASO 5: Generando reporte HTML mensual...")
    try:
        html_generator = HTMLReportGenerator(current_df, comparison)
        html_report = MONTHLY_REPORTS_DIR / f"{month}_monthly_report.html"
        html_generator.generate_executive_report(html_report)
        print(f"   ✅ Reporte HTML: {html_report}")
    except Exception as e:
        logger.warning(f"No se pudo generar HTML: {e}")
        print(f"   ⚠️ Reporte HTML no generado: {e}")
    
    # 6. Estadísticas históricas
    print("\n📈 PASO 6: Análisis histórico de snapshots...")
    all_snapshots = sorted(SNAPSHOTS_DIR.glob("*_snapshot.csv"))
    
    if len(all_snapshots) > 1:
        print(f"   📊 Snapshots disponibles: {len(all_snapshots)}")
        
        historical_data = []
        for snap in all_snapshots:
            snap_df = pd.read_csv(snap)
            snap_month = snap.stem.replace('_snapshot', '')
            
            historical_data.append({
                'Mes': snap_month,
                'Oportunidades': len(snap_df),
                'Total_USD': snap_df[COLUMNS['usd']].sum() if COLUMNS['usd'] in snap_df.columns else 0,
                'Responsables': snap_df[COLUMNS['responsible']].nunique() if COLUMNS['responsible'] in snap_df.columns else 0
            })
        
        historical_df = pd.DataFrame(historical_data)
        historical_file = MONTHLY_REPORTS_DIR / f"{month}_historical.xlsx"
        historical_df.to_excel(historical_file, index=False)
        print(f"   ✅ Histórico: {historical_file}")
        
        # Mostrar tendencia
        print("\n   📈 Tendencia histórica:")
        for _, row in historical_df.iterrows():
            print(f"      {row['Mes']}: {row['Oportunidades']:,} opps - ${row['Total_USD']:,.0f}")
    
    # Resumen final
    print("\n" + "="*70)
    print("✅ PROCESO MENSUAL COMPLETADO")
    print("="*70)
    print(f"\n📁 Archivos generados:")
    print(f"   • Snapshot: {snapshot_file}")
    print(f"   • Reporte Excel: {monthly_excel}")
    
    print(f"\n📊 Métricas del mes {month}:")
    print(f"   • Oportunidades: {summary['total_opportunities']:,}")
    print(f"   • Valor USD: ${summary['total_usd']:,.2f}")
    print(f"   • Responsables: {summary['unique_responsibles']}")
    print(f"   • Países: {summary['unique_markets']}")
    
    print("\n" + "="*70 + "\n")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
