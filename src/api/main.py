"""
KPI CAS API - FastAPI Application
==================================
REST API para análisis remoto de oportunidades Salesforce.

Inicio:
    uvicorn src.api.main:app --reload --port 8000

Documentación:
    http://localhost:8000/docs (Swagger UI)
    http://localhost:8000/redoc (ReDoc)
"""

import sys
from pathlib import Path
from datetime import datetime
import logging
import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

# Agregar path del proyecto
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.api.routes.analysis import router as analysis_router
from src.api.models import HealthResponse

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# =============================================================================
# CREAR APLICACIÓN
# =============================================================================

app = FastAPI(
    title="KPI CAS API",
    description="""
    API REST para análisis de oportunidades de Salesforce.
    
    ## Funcionalidades
    
    * **Ejecutar análisis** - Procesar CSV de oportunidades
    * **Consultar estado** - Verificar progreso de jobs
    * **Obtener resultados** - Métricas, KPIs, cambios detectados
    
    ## Uso típico
    
    1. `POST /api/v1/analysis/run` - Iniciar análisis
    2. `GET /api/v1/analysis/{job_id}` - Consultar estado
    3. `GET /api/v1/analysis/{job_id}/result` - Obtener resultado
    """,
    version="1.0.0",
    contact={
        "name": "KPI CAS Team"
    }
)

# =============================================================================
# MIDDLEWARE
# =============================================================================

# CORS para permitir requests desde cualquier origen (desarrollo)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =============================================================================
# ROUTES
# =============================================================================

# Incluir router de análisis
app.include_router(analysis_router)

# Mount static files (Frontend)
# ERROR: En Vercel, Python no debe servir estáticos (lo hace el CDN).
# Descomentar solo para desarrollo local si no usas live server.
if os.getenv("VERCEL") != "1":
    try:
        app.mount("/", StaticFiles(directory="public", html=True), name="public")
    except Exception:
        pass


@app.get("/", tags=["Root"])
async def root():
    """Endpoint raíz con información básica"""
    return {
        "name": "KPI CAS API",
        "version": "1.0.0",
        "docs": "/docs",
        "health": "/api/v1/health"
    }


@app.get("/api/v1/health", response_model=HealthResponse, tags=["Health"])
async def health_check():
    """
    Health check endpoint.
    Útil para verificar que el servidor está corriendo.
    """
    return HealthResponse(
        status="healthy",
        version="1.0.0",
        timestamp=datetime.now()
    )


# =============================================================================
# EVENTOS DE CICLO DE VIDA
# =============================================================================

@app.on_event("startup")
async def startup_event():
    """Ejecutar al iniciar el servidor"""
    logger.info("🚀 KPI CAS API iniciada")
    logger.info("📚 Documentación disponible en: /docs")


@app.on_event("shutdown")
async def shutdown_event():
    """Ejecutar al detener el servidor"""
    logger.info("👋 KPI CAS API detenida")


# =============================================================================
# MAIN (para desarrollo)
# =============================================================================

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "src.api.main:app",
        host="0.0.0.0",
        port=8000,
        reload=True
    )
