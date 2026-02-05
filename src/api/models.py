"""
Pydantic Models for API Request/Response
=========================================
"""

from pydantic import BaseModel, Field
from datetime import datetime
from typing import Optional, Dict, Any
from enum import Enum


class JobStatusEnum(str, Enum):
    """Estados posibles de un job de análisis"""
    PENDING = "PENDING"
    RUNNING = "RUNNING"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"


# =============================================================================
# REQUEST MODELS
# =============================================================================

class AnalysisRequest(BaseModel):
    """Request para ejecutar un análisis"""
    file_path: Optional[str] = Field(
        default=None,
        description="Ruta al CSV. Si es None, usa el archivo más reciente"
    )
    compare_with_previous: bool = Field(
        default=True,
        description="Comparar con snapshot anterior"
    )
    generate_html: bool = Field(
        default=True,
        description="Generar reporte HTML ejecutivo"
    )
    generate_emails: bool = Field(
        default=False,
        description="Generar emails individuales por responsable"
    )
    region: Optional[str] = Field(
        default="CAS",
        description="Filtro de región (ej: CAS, Brasil, Mexico, All)"
    )


# =============================================================================
# RESPONSE MODELS
# =============================================================================

class JobStatusResponse(BaseModel):
    """Respuesta con estado del job"""
    job_id: str
    status: JobStatusEnum
    created_at: datetime
    started_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None
    progress: Optional[str] = None
    error: Optional[str] = None


class KPISummary(BaseModel):
    """Resumen de un KPI"""
    code: str
    name: str
    count: int
    severity: str


class MarketSummary(BaseModel):
    """Resumen de un mercado"""
    name: str
    count: int
    total_usd: float


class ChangesSummary(BaseModel):
    """Resumen de cambios detectados"""
    new_count: int
    removed_count: int
    changed_count: int
    unchanged_count: int
    usd_change: float


class AnalysisResult(BaseModel):
    """Resultado completo del análisis"""
    job_id: str
    status: JobStatusEnum
    
    # Métricas principales
    total_opportunities: int
    total_usd: float
    total_responsibles: int
    total_markets: int
    stagnant_count: int
    at_risk_count: int
    
    # Desgloses
    by_kpi: list[KPISummary]
    by_market: list[MarketSummary]
    
    # Cambios (si aplica)
    changes: Optional[ChangesSummary] = None
    
    # Archivos generados
    excel_report_path: Optional[str] = None
    html_report_path: Optional[str] = None
    
    # Metadata
    analysis_date: datetime
    data_file: str


class HealthResponse(BaseModel):
    """Respuesta del health check"""
    status: str = "healthy"
    version: str = "1.0.0"
    timestamp: datetime
