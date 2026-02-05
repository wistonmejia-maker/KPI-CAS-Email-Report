"""
Data Loader - Carga y Validación de Datos
==========================================
Módulo para cargar archivos CSV de Salesforce con validación y normalización.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple, List
import logging

from .config import (
    RAW_DIR, PROCESSED_DIR, COLUMNS, KEY_COLUMNS,
    DATE_FORMAT, get_current_date_str
)

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DataLoader:
    """Clase para cargar y validar datos de oportunidades"""
    
    def __init__(self):
        self.required_columns = list(COLUMNS.values())
    
    def load_csv(self, filepath: Path, encoding: str = 'utf-8') -> pd.DataFrame:
        """
        Carga un archivo CSV y realiza validación básica
        
        Args:
            filepath: Ruta al archivo CSV
            encoding: Codificación del archivo
            
        Returns:
            DataFrame con los datos cargados y procesados
        """
        logger.info(f"Cargando archivo: {filepath}")
        
        if not filepath.exists():
            raise FileNotFoundError(f"Archivo no encontrado: {filepath}")
        
        # Cargar CSV
        df = pd.read_csv(filepath, encoding=encoding)
        
        # Validar columnas
        self._validate_columns(df)
        
        # Procesar datos
        df = self._process_dataframe(df)
        
        logger.info(f"Archivo cargado exitosamente: {len(df)} registros")
        return df
    
    def _validate_columns(self, df: pd.DataFrame) -> None:
        """Valida que el DataFrame tenga las columnas requeridas"""
        missing_columns = set(self.required_columns) - set(df.columns)
        if missing_columns:
            logger.warning(f"Columnas faltantes: {missing_columns}")
    
    def _process_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Procesa y normaliza el DataFrame"""
        df = df.copy()
        
        # Convertir fechas
        date_columns = [COLUMNS['created_date'], COLUMNS['close_date']]
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Convertir USD a numérico
        if COLUMNS['usd'] in df.columns:
            df[COLUMNS['usd']] = pd.to_numeric(df[COLUMNS['usd']], errors='coerce').fillna(0)
        
        # Limpiar responsables vacíos
        if COLUMNS['responsible'] in df.columns:
            df[COLUMNS['responsible']] = df[COLUMNS['responsible']].fillna('Sin Asignar')
        
        # Agregar columna de fecha de carga
        df['_load_date'] = datetime.now()
        
        return df
    
    def get_latest_file(self, directory: Path = None) -> Optional[Path]:
        """
        Obtiene el archivo más reciente en el directorio
        
        Args:
            directory: Directorio donde buscar (default: RAW_DIR)
            
        Returns:
            Path al archivo más reciente o None
        """
        if directory is None:
            directory = RAW_DIR
        
        csv_files = list(directory.glob("*.csv"))
        if not csv_files:
            return None
        
        # Ordenar por fecha de modificación
        csv_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return csv_files[0]
    
    def get_previous_file(self, current_file: Path, directory: Path = None) -> Optional[Path]:
        """
        Obtiene el archivo anterior al archivo actual
        
        Args:
            current_file: Archivo actual
            directory: Directorio donde buscar
            
        Returns:
            Path al archivo anterior o None
        """
        if directory is None:
            directory = RAW_DIR
        
        csv_files = list(directory.glob("*.csv"))
        csv_files = [f for f in csv_files if f != current_file]
        
        if not csv_files:
            return None
        
        # Ordenar por fecha de modificación
        csv_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return csv_files[0]
    
    def archive_file(self, filepath: Path, processed: bool = True) -> Path:
        """
        Mueve un archivo procesado al directorio de archivado
        
        Args:
            filepath: Archivo a archivar
            processed: Si va a PROCESSED_DIR (True) o permanece
            
        Returns:
            Nueva ruta del archivo
        """
        if not processed:
            return filepath
        
        dest_dir = PROCESSED_DIR
        dest_dir.mkdir(parents=True, exist_ok=True)
        
        new_path = dest_dir / filepath.name
        filepath.rename(new_path)
        
        logger.info(f"Archivo archivado: {new_path}")
        return new_path


def load_opportunities(filepath: Optional[Path] = None) -> pd.DataFrame:
    """
    Función de conveniencia para cargar oportunidades
    
    Args:
        filepath: Ruta al archivo (opcional, usa el más reciente si no se especifica)
        
    Returns:
        DataFrame con oportunidades
    """
    loader = DataLoader()
    
    if filepath is None:
        filepath = loader.get_latest_file()
        if filepath is None:
            raise FileNotFoundError("No se encontraron archivos CSV en el directorio")
    
    return loader.load_csv(filepath)


def get_file_pairs() -> List[Tuple[Path, Path]]:
    """
    Obtiene pares de archivos (actual, anterior) para comparación
    
    Returns:
        Lista de tuplas (archivo_actual, archivo_anterior)
    """
    loader = DataLoader()
    csv_files = list(RAW_DIR.glob("*.csv"))
    csv_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    
    pairs = []
    for i in range(len(csv_files) - 1):
        pairs.append((csv_files[i], csv_files[i + 1]))
    
    return pairs
