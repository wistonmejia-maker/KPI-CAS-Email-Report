"""
Configuración del Sistema de Seguimiento de KPIs
================================================
"""

from pathlib import Path
from datetime import datetime

# =============================================================================
# RUTAS DEL PROYECTO
# =============================================================================
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
RAW_DIR = DATA_DIR / "raw"
PROCESSED_DIR = DATA_DIR / "processed"
SNAPSHOTS_DIR = DATA_DIR / "snapshots"
REPORTS_DIR = BASE_DIR / "reports"
WEEKLY_REPORTS_DIR = REPORTS_DIR / "weekly"
MONTHLY_REPORTS_DIR = REPORTS_DIR / "monthly"
EMAIL_REPORTS_DIR = REPORTS_DIR / "emails"
TEMPLATES_DIR = BASE_DIR / "templates"

# =============================================================================
# COLUMNAS DEL DATASET
# =============================================================================
COLUMNS = {
    'id': 'Id',
    'link': 'Link',
    'kpi': 'KPI',
    'responsible': 'Responsible',
    'region': 'Region',
    'market': 'Market',
    'site': 'Site',
    'usd': 'USD',
    'siterra_project': 'Siterra Project',
    'customer': 'Customer',
    'product': 'Product',
    'stage': 'Stage',
    'created_date': 'CreatedDate',
    'close_date': 'CloseDate',
    'revision': 'Revision',
    'description': 'Descripcion'
}

# Columnas clave para identificar una oportunidad
KEY_COLUMNS = ['Id']

# Columnas a monitorear para cambios
TRACKED_COLUMNS = ['Stage', 'Responsible', 'USD', 'CloseDate', 'KPI']

# =============================================================================
# CONFIGURACIÓN DE ANÁLISIS
# =============================================================================

# Días sin cambio para considerar una oportunidad "estancada"
STAGNANT_DAYS_THRESHOLD = 30

# Días antes de vencimiento para alertar
WARNING_DAYS_BEFORE_CLOSE = 7

# Etapas ordenadas para medir avance/retroceso
STAGE_ORDER = [
    'Identify the opportunity',
    'Customer Analysis',
    'NDD/RFI',
    'NTP/RFI',
    'Application approved',
    'Financial Analysis',
    'Work Needs Analysis',
    'Tenant Lease',
    'Ground Lease Agreement',
    'TLA Signature',
    'Client Approval',
    'Work Execution',
    'Construction',
    'Service Delivery Analysis',
    'Ready to Bill',
    'Reported to Finance',
    'Proceed with Billing Changes',
    'Equipment Removal',
    'Customer Notification',
    'Cancelado'
]

# =============================================================================
# CONFIGURACIÓN DE REPORTES
# =============================================================================

# Formato de fecha para nombres de archivo
DATE_FORMAT = "%Y%m%d"
WEEK_FORMAT = "%Y-W%W"
MONTH_FORMAT = "%Y-%m"

# Colores para reportes HTML
COLORS = {
    'primary': '#1a73e8',
    'success': '#34a853',
    'warning': '#fbbc04',
    'danger': '#ea4335',
    'info': '#4285f4',
    'light': '#f8f9fa',
    'dark': '#202124'
}

# =============================================================================
# CATEGORÍAS DE KPI - Data Cleansing Definitions
# =============================================================================
KPI_CATEGORIES = {
    # Aging Control - Oportunidades antiguas sin movimiento
    'DC001 NB': {
        'name': 'Aging Control (NB)',
        'description': 'Opportunities Created >9 months (revision 0 only)',
        'type': 'aging',
        'severity': 'high',
        'thresholds': {'green': 10, 'yellow': 15},  # % based
        'is_percentage': True
    },
    'DC001 CHURN': {
        'name': 'Aging Control (Churn)',
        'description': 'Opportunities Created >12 months (revision 0 only)',
        'type': 'aging',
        'severity': 'high',
        'thresholds': {'green': 100, 'yellow': 500},  # absolute
        'is_percentage': False
    },
    
    # Expired Opportunities - Forecast vencido
    'DC002 NB': {
        'name': 'Expired Opportunities (NB)',
        'description': 'Forecast date older than current month closure (follows exchange rate calendar)',
        'type': 'expired',
        'severity': 'high',
        'thresholds': {'green': 15, 'yellow': 50},  # absolute
        'is_percentage': False
    },
    'DC002 CHURN': {
        'name': 'Expired Opportunities (Churn)',
        'description': 'Forecast date older than current month closure (follows full month)',
        'type': 'expired',
        'severity': 'medium',
        'thresholds': {'green': 15, 'yellow': 50},  # absolute
        'is_percentage': False
    },
    
    # On Hold
    'DC003': {
        'name': 'On Hold',
        'description': 'On hold opportunities (all revisions). See Concepts & Refresh Schedule',
        'type': 'operational',
        'severity': 'medium',
        'thresholds': {'green': 0.5, 'yellow': 1},  # % based
        'is_percentage': True
    },
    
    # Revenue Issues
    'DC004': {
        'name': 'Reported to Finance w/o Revenue',
        'description': 'By the rule, opportunities without revenue should go straight to ready to bill',
        'type': 'data_quality',
        'severity': 'low',
        'thresholds': {'green': 50, 'yellow': 100},  # absolute
        'is_percentage': False
    },
    
    # Process Anomalies
    'DC005': {
        'name': 'Conversion w/o Sales Process',
        'description': 'Opp. converted in less than X days (2 days for Collocation, 1 day for BTS) / Opts. created in the last 30 days',
        'type': 'process',
        'severity': 'medium',
        'thresholds': {'green': 1.5, 'yellow': 3},  # % based
        'is_percentage': True
    },
    
    # Change Management
    'DC007': {
        'name': 'Change Management',
        'description': 'Opportunities created by other areas (not created by sales). Considering: Collo, BTS, Churn (excl. subtype additional equipment and lease rent reduction). Opts. created in the last 30 days',
        'type': 'process',
        'severity': 'low',
        'thresholds': {'green': 5, 'yellow': 10},  # absolute
        'is_percentage': False
    },
    
    # Aging in Finance
    'DC008': {
        'name': 'Aging Reported to Finance',
        'description': 'Opportunities with more than 30 days in Reported to Finance. Excl. ON HOLD optys.',
        'type': 'aging',
        'severity': 'high',
        'thresholds': {'green': 15, 'yellow': 30},  # absolute
        'is_percentage': False
    },
    
    # Amount Zero
    'DC010': {
        'name': 'Amount Zero',
        'description': 'Opportunities with Amount = 0. All Revisions. Excl. ToP = TRUE. Excl. Sales Deal, Amd. w/o Rev. Primarily Stages in TI, Renewal',
        'type': 'data_quality',
        'severity': 'high',
        'thresholds': {'green': 0, 'yellow': 15},  # 0 = green, >0 and <15 = yellow, >=15 = red
        'is_percentage': False
    },
    
    # Roles & Responsibilities
    'DC011': {
        'name': 'Actual Roles & Responsibilities',
        'description': '# of opts. that changed to Actual. If the opt. was already in Actual and changed, it will also appear here. According to R&R, only GBS can make these changes, last 30 days',
        'type': 'process',
        'severity': 'low',
        'thresholds': {'green': 0, 'yellow': 1},  # 0 = green, any > 0 = red
        'is_percentage': False
    }
}

# Mapeo de severidad a colores
KPI_SEVERITY_COLORS = {
    'high': 'danger',    # Rojo - Acción urgente
    'medium': 'warning', # Amarillo - En seguimiento
    'low': 'success'     # Verde - OK
}

# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================

def get_current_date_str():
    """Retorna fecha actual en formato para archivos"""
    return datetime.now().strftime(DATE_FORMAT)

def get_current_week_str():
    """Retorna semana actual en formato YYYY-WWW"""
    return datetime.now().strftime(WEEK_FORMAT)

def get_current_month_str():
    """Retorna mes actual en formato YYYY-MM"""
    return datetime.now().strftime(MONTH_FORMAT)

def ensure_directories():
    """Crea los directorios necesarios si no existen"""
    for directory in [RAW_DIR, PROCESSED_DIR, SNAPSHOTS_DIR, 
                      WEEKLY_REPORTS_DIR, MONTHLY_REPORTS_DIR, 
                      EMAIL_REPORTS_DIR, TEMPLATES_DIR]:
        directory.mkdir(parents=True, exist_ok=True)
