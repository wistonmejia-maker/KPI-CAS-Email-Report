"""
Run Weekly - Script Principal para Proceso Semanal
===================================================
Ejecuta el proceso semanal de análisis de oportunidades.
"""

import sys
import logging
from pathlib import Path
from datetime import datetime
import argparse

# Agregar src al path
sys.path.insert(0, str(Path(__file__).parent))

from src.config import (
    RAW_DIR, WEEKLY_REPORTS_DIR, EMAIL_REPORTS_DIR,
    ensure_directories, get_current_week_str, get_current_date_str
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
        logging.FileHandler(f'weekly_run_{get_current_date_str()}.log')
    ]
)
logger = logging.getLogger(__name__)


def main():
    """Función principal del proceso semanal"""
    parser = argparse.ArgumentParser(description='Proceso Semanal de KPIs')
    parser.add_argument('--file', '-f', type=str, help='Archivo CSV a procesar (opcional)')
    parser.add_argument('--no-compare', action='store_true', help='No comparar con archivo anterior')
    parser.add_argument('--no-emails', action='store_true', help='No generar emails por responsable')
    parser.add_argument('--no-html', action='store_true', help='No generar reporte HTML')
    args = parser.parse_args()
    
    print("\n" + "="*70)
    print("🚀 PROCESO SEMANAL DE KPIs - OPORTUNIDADES SALESFORCE")
    print("="*70)
    print(f"📅 Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📆 Semana: {get_current_week_str()}")
    print("="*70 + "\n")
    
    # Asegurar que existan los directorios
    ensure_directories()
    
    # 1. Cargar datos
    print("📂 PASO 1: Cargando datos...")
    loader = DataLoader()
    
    if args.file:
        current_file = Path(args.file)
    else:
        current_file = loader.get_latest_file(RAW_DIR)
        if current_file is None:
            # Buscar en directorio raíz del proyecto
            project_root = Path(__file__).parent
            csv_files = list(project_root.glob("*.csv"))
            if csv_files:
                current_file = max(csv_files, key=lambda x: x.stat().st_mtime)
    
    if current_file is None:
        logger.error("No se encontró ningún archivo CSV para procesar")
        print("\n❌ ERROR: No se encontró ningún archivo CSV")
        print("   Coloque el archivo CSV en la carpeta 'data/raw/' o especifique con --file")
        return 1
    
    print(f"   📄 Archivo: {current_file.name}")
    current_df = loader.load_csv(current_file)
    print(f"   ✅ Registros cargados: {len(current_df):,}")
    
    # 2. Comparar con período anterior (si existe)
    comparison = None
    if not args.no_compare:
        print("\n🔄 PASO 2: Comparando con período anterior...")
        previous_file = loader.get_previous_file(current_file, RAW_DIR)
        
        if previous_file is None:
            # Buscar en processed
            from src.config import PROCESSED_DIR
            previous_files = list(PROCESSED_DIR.glob("*.csv"))
            if previous_files:
                previous_file = max(previous_files, key=lambda x: x.stat().st_mtime)
        
        if previous_file:
            print(f"   📄 Archivo anterior: {previous_file.name}")
            previous_df = loader.load_csv(previous_file)
            
            detector = ChangeDetector()
            comparison = detector.compare(current_df, previous_df)
            
            print(f"\n   📊 Resumen de cambios:")
            print(f"      • Nuevas oportunidades: {comparison.summary['new_count']}")
            print(f"      • Oportunidades eliminadas: {comparison.summary['removed_count']}")
            print(f"      • Oportunidades modificadas: {comparison.summary['changed_count']}")
            print(f"      • Cambio en USD: ${comparison.summary['usd_change']:,.2f}")
        else:
            print("   ⚠️ No se encontró archivo anterior para comparar")
            print("   ℹ️ Este es el primer archivo procesado")
    
    # 3. Calcular métricas
    print("\n📊 PASO 3: Calculando métricas...")
    metrics = MetricsCalculator(current_df)
    summary = metrics.get_summary()
    
    print(f"   • Total oportunidades: {summary['total_opportunities']:,}")
    print(f"   • Valor total USD: ${summary['total_usd']:,.2f}")
    print(f"   • Oportunidades estancadas: {summary['stagnant_count']}")
    print(f"   • Oportunidades en riesgo: {summary['at_risk_count']}")
    
    # 4. Generar reporte Excel
    print("\n📁 PASO 4: Generando reporte Excel...")
    excel_generator = ExcelReportGenerator(current_df, comparison)
    excel_report = excel_generator.generate_weekly_report()
    print(f"   ✅ Reporte Excel: {excel_report}")
    
    # 5. Generar reportes por responsable
    if not args.no_emails:
        print("\n📧 PASO 5: Generando reportes por responsable...")
        resp_reports = excel_generator.generate_responsible_reports()
        print(f"   ✅ Generados {len(resp_reports)} reportes individuales")
    
    # 6. Generar reporte HTML ejecutivo
    if not args.no_html:
        print("\n🌐 PASO 6: Generando reporte HTML ejecutivo...")
        html_generator = HTMLReportGenerator(current_df, comparison)
        
        try:
            html_report = html_generator.generate_executive_report()
            print(f"   ✅ Reporte HTML: {html_report}")
        except Exception as e:
            logger.warning(f"No se pudo generar reporte HTML: {e}")
            print(f"   ⚠️ No se pudo generar HTML (instale matplotlib: pip install matplotlib)")
        
        # Emails individuales
        print("\n📨 PASO 7: Generando emails HTML por responsable...")
        try:
            week_dir = EMAIL_REPORTS_DIR / get_current_week_str()
            week_dir.mkdir(parents=True, exist_ok=True)
            
            from src.config import COLUMNS
            email_count = 0
            for responsible in current_df[COLUMNS['responsible']].unique():
                if responsible and responsible != 'Sin Asignar':
                    try:
                        html_generator.generate_responsible_email(
                            responsible, 
                            week_dir / f"{responsible.replace('/', '_')}_email.html"
                        )
                        email_count += 1
                    except Exception as e:
                        logger.warning(f"Error generando email para {responsible}: {e}")
            
            print(f"   ✅ Generados {email_count} emails HTML")
        except Exception as e:
            logger.warning(f"Error generando emails: {e}")
    
    # 8. Resumen final
    print("\n" + "="*70)
    print("✅ PROCESO COMPLETADO EXITOSAMENTE")
    print("="*70)
    print(f"\n📁 Archivos generados en:")
    print(f"   • Reportes Excel: {WEEKLY_REPORTS_DIR}")
    print(f"   • Reportes HTML:  {EMAIL_REPORTS_DIR}")
    
    print(f"\n📊 Resumen del análisis:")
    print(f"   • Total oportunidades: {summary['total_opportunities']:,}")
    print(f"   • Valor total: ${summary['total_usd']:,.2f}")
    print(f"   • Responsables: {summary['unique_responsibles']}")
    print(f"   • Países: {summary['unique_markets']}")
    
    if comparison:
        print(f"\n🔄 Cambios detectados:")
        print(f"   • +{comparison.summary['new_count']} nuevas")
        print(f"   • -{comparison.summary['removed_count']} eliminadas")
        print(f"   • {comparison.summary['changed_count']} modificadas")
    
    print("\n" + "="*70 + "\n")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
