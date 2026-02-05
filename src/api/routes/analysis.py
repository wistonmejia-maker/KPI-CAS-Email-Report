"""
Analysis Routes - Endpoints para ejecutar y consultar análisis
===============================================================
"""

import sys
from pathlib import Path
from datetime import datetime
from typing import Optional
import threading
import logging

from fastapi import APIRouter, HTTPException, BackgroundTasks, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse
import shutil
import os

# Agregar path del proyecto para imports
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

from src.api.models import (
    AnalysisRequest, JobStatusResponse, AnalysisResult,
    KPISummary, MarketSummary, ChangesSummary, JobStatusEnum
)
from src.api.jobs import job_manager, JobStatus
from src.config import KPI_CATEGORIES, COLUMNS, SNAPSHOTS_DIR
from src.data_loader import DataLoader, load_opportunities
from src.change_detector import compare_datasets
from src.metrics import MetricsCalculator
from src.report_generator import generate_weekly_report
# Reemplazamos el generador antiguo por el nuevo que tiene el branding correcto
# Como generar_resumen_email.py está en la raíz, lo importamos dinámicamente o asumiendo que está en el path
try:
    from generar_resumen_email import generar_html_profesional, calcular_deltas
except ImportError:
    # Fallback por si la estructura de importación falla en dev vs prod
    import sys
    sys.path.append(str(Path(__file__).parent.parent.parent.parent))
    from generar_resumen_email import generar_html_profesional, calcular_deltas
    
from src.infographic.visual_card import generate_executive_card

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/v1/analysis", tags=["Analysis"])


def run_analysis_task(job_id: str, request: AnalysisRequest):
    """
    Ejecuta el análisis en background.
    Esta función corre en un thread separado.
    """
    try:
        job_manager.update_status(job_id, JobStatus.RUNNING, progress="Cargando datos...")
        
        loader = DataLoader()
        
        # 1. Cargar datos
        if request.file_path:
            df = loader.load_csv(Path(request.file_path))
            data_file = request.file_path
        else:
            latest_file = loader.get_latest_file()
            if latest_file is None:
                raise FileNotFoundError("No se encontraron archivos CSV")
            df = loader.load_csv(latest_file)
            data_file = str(latest_file)
        
        job_manager.update_status(job_id, JobStatus.RUNNING, progress="Calculando métricas...")
        
        # 2. Calcular métricas
        metrics = MetricsCalculator(df)
        summary = metrics.get_summary()
        
        # 3. Comparar con anterior (si aplica)
        changes_summary = None
        comparison = None
        if request.compare_with_previous:
            job_manager.update_status(job_id, JobStatus.RUNNING, progress="Comparando con período anterior...")
            try:
                # Buscar snapshot anterior
                snapshot_files = list(SNAPSHOTS_DIR.glob("*.csv"))
                if snapshot_files:
                    snapshot_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                    previous_df = loader.load_csv(snapshot_files[0])
                    comparison = compare_datasets(df, previous_df)
                    changes_summary = {
                        "new_count": comparison.summary.get('new_count', 0),
                        "removed_count": comparison.summary.get('removed_count', 0),
                        "changed_count": comparison.summary.get('changed_count', 0),
                        "unchanged_count": comparison.summary.get('unchanged_count', 0),
                        "usd_change": comparison.summary.get('usd_change', 0)
                    }
            except Exception as e:
                logger.warning(f"No se pudo comparar con anterior: {e}")
        
        # 4. Generar reportes
        excel_path = None
        html_path = None
        
        job_manager.update_status(job_id, JobStatus.RUNNING, progress="Generando reportes...")
        
        try:
            excel_path = str(generate_weekly_report(df, comparison))
        except Exception as e:
            logger.warning(f"Error generando Excel: {e}")
        
        if request.generate_html:
            try:
                html_path = str(generate_executive_html(df, comparison))
            except Exception as e:
                logger.warning(f"Error generando HTML: {e}")
        
        # 5. Construir resultado
        by_kpi = []
        kpi_counts = df[COLUMNS['kpi']].value_counts().to_dict()
        for kpi_code, count in kpi_counts.items():
            kpi_info = KPI_CATEGORIES.get(kpi_code, {})
            by_kpi.append({
                "code": kpi_code,
                "name": kpi_info.get('name', 'Unknown'),
                "count": count,
                "severity": kpi_info.get('severity', 'medium')
            })
        
        by_market = []
        market_groups = df.groupby(COLUMNS['market'])
        for market, group in market_groups:
            by_market.append({
                "name": market,
                "count": len(group),
                "total_usd": float(group[COLUMNS['usd']].sum())
            })
        by_market.sort(key=lambda x: x['count'], reverse=True)
        
        result = {
            "job_id": job_id,
            "status": "COMPLETED",
            "total_opportunities": summary['total_opportunities'],
            "total_usd": summary['total_usd'],
            "total_responsibles": summary['unique_responsibles'],
            "total_markets": summary['unique_markets'],
            "stagnant_count": summary['stagnant_count'],
            "at_risk_count": summary['at_risk_count'],
            "by_kpi": by_kpi,
            "by_market": by_market,
            "changes": changes_summary,
            "excel_report_path": excel_path,
            "html_report_path": html_path,
            "analysis_date": datetime.now().isoformat(),
            "data_file": data_file
        }
        
        job_manager.set_result(job_id, result)
        logger.info(f"Job {job_id} completado exitosamente")
        
    except Exception as e:
        logger.error(f"Job {job_id} falló: {e}")
        job_manager.update_status(job_id, JobStatus.FAILED, error=str(e))


@router.post("/run", response_model=JobStatusResponse)
async def run_analysis(request: AnalysisRequest, background_tasks: BackgroundTasks):
    """
    Inicia un análisis de oportunidades de forma asíncrona.
    
    Retorna un job_id para consultar el estado y resultado.
    """
    # Crear job
    job = job_manager.create_job(request.model_dump())
    
    # Ejecutar en background thread
    thread = threading.Thread(
        target=run_analysis_task,
        args=(job.job_id, request)
    )
    thread.start()
    
    return JobStatusResponse(
        job_id=job.job_id,
        status=JobStatusEnum(job.status.value),
        created_at=job.created_at
    )


@router.get("/{job_id}", response_model=JobStatusResponse)
async def get_job_status(job_id: str):
    """
    Consulta el estado de un job de análisis.
    """
    job = job_manager.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} no encontrado")
    
    return JobStatusResponse(
        job_id=job.job_id,
        status=JobStatusEnum(job.status.value),
        created_at=job.created_at,
        started_at=job.started_at,
        completed_at=job.completed_at,
        progress=job.progress,
        error=job.error
    )


@router.get("/{job_id}/result")
async def get_job_result(job_id: str):
    """
    Obtiene el resultado completo de un análisis.
    
    Solo disponible cuando el job está en estado COMPLETED.
    """
    job = job_manager.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} no encontrado")
    
    if job.status == JobStatus.PENDING:
        raise HTTPException(status_code=202, detail="Job pendiente de ejecución")
    
    if job.status == JobStatus.RUNNING:
        raise HTTPException(status_code=202, detail=f"Job en progreso: {job.progress}")
    
    if job.status == JobStatus.FAILED:
        raise HTTPException(status_code=500, detail=f"Job falló: {job.error}")
    
    return job.result


@router.get("/", response_model=list[JobStatusResponse])
async def list_jobs(limit: int = 10):
    """
    Lista los jobs más recientes.
    """
    jobs = job_manager.list_jobs(limit)
    return [
        JobStatusResponse(
            job_id=job.job_id,
            status=JobStatusEnum(job.status.value),
            created_at=job.created_at,
            started_at=job.started_at,
            completed_at=job.completed_at,
            progress=job.progress,
            error=job.error
        )
        for job in jobs
    ]


@router.get("/{job_id}/card", response_class=FileResponse)
async def get_job_card(job_id: str):
    """
    Genera y retorna la tarjeta ejecutiva (PNG) para un análisis completado.
    """
    job = job_manager.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"Job {job_id} no encontrado")
    
    if job.status != JobStatus.COMPLETED:
        raise HTTPException(status_code=400, detail="El análisis debe estar completado para generar la tarjeta")
            
    if not job.result:
        raise HTTPException(status_code=500, detail="El job está completado pero no tiene resultados")
    
    try:
        # Generar tarjeta
        image_path = generate_executive_card(job.result, session_id=job_id)
        
        return FileResponse(
            path=image_path,
            media_type="image/png",
            filename=f"kpi_card_{job_id}.png"
        )
    except Exception as e:
        logger.error(f"Error generando tarjeta para job {job_id}: {e}")
        raise HTTPException(status_code=500, detail=f"Error generando visual: {str(e)}")

@router.post("/upload-and-analyze", response_class=HTMLResponse)
async def upload_and_analyze(
    current_file: UploadFile = File(...),
    previous_file: Optional[UploadFile] = File(None)
):
    """
    Recibe archivos CSV (actual y opcional previo), ejecuta el análisis 
    y retorna el reporte HTML directamente.
    
    Ideal para uso "stateless" en vercel.
    """
    try:
        # 1. Definir directorio temporal
        # En Vercel solo /tmp es escribible
        temp_dir = Path("/tmp") if os.path.exists("/tmp") else Path("temp_uploads")
        temp_dir.mkdir(exist_ok=True)
        
        current_path = temp_dir / f"current_{datetime.now().timestamp()}.csv"
        previous_path = None
        
        # 2. Guardar archivo actual
        with current_path.open("wb") as buffer:
            shutil.copyfileobj(current_file.file, buffer)
            
        # 3. Guardar archivo previo (si existe)
        if previous_file:
            previous_path = temp_dir / f"previous_{datetime.now().timestamp()}.csv"
            with previous_path.open("wb") as buffer:
                shutil.copyfileobj(previous_file.file, buffer)
                
        # 4. Cargar datos
        loader = DataLoader()
        df = loader.load_csv(current_path)
        
        # 5. Comparar (si aplica)
        # 5. Calcular Deltas (usando la lógica de generar_resumen_email)
        deltas = None
        if previous_path:
            try:
                previous_df = loader.load_csv(previous_path)
                deltas = calcular_deltas(df, previous_df)
            except Exception as e:
                logger.warning(f"Error calculando deltas: {e}")
                # Si falla, calculamos deltas vacios (None como segundo argumento retorna estructura vacía/default)
                deltas = calcular_deltas(df, None)
        else:
            deltas = calcular_deltas(df, None)
                
        # 6. Generar HTML (Profesional / Email)
        # generar_html_profesional devuelve el string directamente
        html_content = generar_html_profesional(df, deltas)
        
        # 7. Limpieza
        try:
            current_path.unlink()
            if previous_path:
                previous_path.unlink()
        except Exception:
            pass
            
        return HTMLResponse(content=html_content)

    except Exception as e:
        import traceback
        error_msg = f"Error en upload_and_analyze: {str(e)}\n{traceback.format_exc()}"
        logger.error(error_msg)
        print(error_msg)
        
        # Fallback logging to file
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(error_msg)
            
        raise HTTPException(status_code=500, detail=f"Error interno: {str(e)}")
