"""
Job Manager - Gestión de tareas de análisis en memoria
=======================================================
"""

import uuid
import threading
from datetime import datetime
from typing import Dict, Optional, Any
from dataclasses import dataclass, field
from enum import Enum
import logging

logger = logging.getLogger(__name__)


class JobStatus(str, Enum):
    PENDING = "PENDING"
    RUNNING = "RUNNING"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"


@dataclass
class Job:
    """Representa un job de análisis"""
    job_id: str
    status: JobStatus
    created_at: datetime
    started_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None
    progress: Optional[str] = None
    error: Optional[str] = None
    result: Optional[Dict[str, Any]] = None
    request_params: Dict[str, Any] = field(default_factory=dict)


class JobManager:
    """
    Gestor de jobs en memoria.
    Thread-safe para operaciones concurrentes.
    """
    
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        """Singleton pattern"""
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._jobs: Dict[str, Job] = {}
                    cls._instance._jobs_lock = threading.Lock()
        return cls._instance
    
    def create_job(self, request_params: Dict[str, Any] = None) -> Job:
        """Crea un nuevo job con estado PENDING"""
        job_id = str(uuid.uuid4())[:8]
        job = Job(
            job_id=job_id,
            status=JobStatus.PENDING,
            created_at=datetime.now(),
            request_params=request_params or {}
        )
        
        with self._jobs_lock:
            self._jobs[job_id] = job
        
        logger.info(f"Job creado: {job_id}")
        return job
    
    def get_job(self, job_id: str) -> Optional[Job]:
        """Obtiene un job por ID"""
        with self._jobs_lock:
            return self._jobs.get(job_id)
    
    def update_status(self, job_id: str, status: JobStatus, 
                      progress: str = None, error: str = None) -> bool:
        """Actualiza el estado de un job"""
        with self._jobs_lock:
            job = self._jobs.get(job_id)
            if not job:
                return False
            
            job.status = status
            if progress:
                job.progress = progress
            if error:
                job.error = error
            
            if status == JobStatus.RUNNING and not job.started_at:
                job.started_at = datetime.now()
            elif status in (JobStatus.COMPLETED, JobStatus.FAILED):
                job.completed_at = datetime.now()
            
            logger.info(f"Job {job_id} actualizado a {status}")
            return True
    
    def set_result(self, job_id: str, result: Dict[str, Any]) -> bool:
        """Establece el resultado de un job completado"""
        with self._jobs_lock:
            job = self._jobs.get(job_id)
            if not job:
                return False
            
            job.result = result
            job.status = JobStatus.COMPLETED
            job.completed_at = datetime.now()
            return True
    
    def list_jobs(self, limit: int = 10) -> list[Job]:
        """Lista los jobs más recientes"""
        with self._jobs_lock:
            sorted_jobs = sorted(
                self._jobs.values(),
                key=lambda j: j.created_at,
                reverse=True
            )
            return sorted_jobs[:limit]
    
    def cleanup_old_jobs(self, max_age_hours: int = 24):
        """Elimina jobs más antiguos que max_age_hours"""
        cutoff = datetime.now()
        with self._jobs_lock:
            old_jobs = [
                job_id for job_id, job in self._jobs.items()
                if (cutoff - job.created_at).total_seconds() > max_age_hours * 3600
            ]
            for job_id in old_jobs:
                del self._jobs[job_id]
                logger.info(f"Job eliminado por antigüedad: {job_id}")


# Instancia global
job_manager = JobManager()
