"""
Change Detector - Detección de Cambios entre Períodos
======================================================
Módulo para detectar y clasificar cambios entre dos datasets de oportunidades.
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
import logging

from .config import (
    COLUMNS, KEY_COLUMNS, TRACKED_COLUMNS, STAGE_ORDER,
    STAGNANT_DAYS_THRESHOLD, WARNING_DAYS_BEFORE_CLOSE
)

logger = logging.getLogger(__name__)


@dataclass
class ChangeRecord:
    """Representa un cambio detectado en una oportunidad"""
    opportunity_id: str
    field: str
    old_value: any
    new_value: any
    change_type: str  # 'new', 'removed', 'modified', 'stage_advance', 'stage_regress'
    responsible: str
    market: str
    
    def to_dict(self) -> dict:
        return {
            'Id': self.opportunity_id,
            'Campo': self.field,
            'Valor_Anterior': self.old_value,
            'Valor_Nuevo': self.new_value,
            'Tipo_Cambio': self.change_type,
            'Responsable': self.responsible,
            'País': self.market
        }


@dataclass
class ComparisonResult:
    """Resultado de la comparación entre dos datasets"""
    new_opportunities: pd.DataFrame
    removed_opportunities: pd.DataFrame
    changes: List[ChangeRecord]
    unchanged: pd.DataFrame
    summary: Dict
    comparison_date: datetime = field(default_factory=datetime.now)
    
    def get_changes_df(self) -> pd.DataFrame:
        """Convierte la lista de cambios a DataFrame"""
        if not self.changes:
            return pd.DataFrame()
        return pd.DataFrame([c.to_dict() for c in self.changes])
    
    def get_changes_by_responsible(self) -> Dict[str, List[ChangeRecord]]:
        """Agrupa cambios por responsable"""
        by_responsible = {}
        for change in self.changes:
            if change.responsible not in by_responsible:
                by_responsible[change.responsible] = []
            by_responsible[change.responsible].append(change)
        return by_responsible
    
    def get_changes_by_market(self) -> Dict[str, List[ChangeRecord]]:
        """Agrupa cambios por mercado/país"""
        by_market = {}
        for change in self.changes:
            if change.market not in by_market:
                by_market[change.market] = []
            by_market[change.market].append(change)
        return by_market


class ChangeDetector:
    """Clase para detectar cambios entre dos datasets de oportunidades"""
    
    def __init__(self, id_column: str = None, tracked_columns: List[str] = None):
        """
        Inicializa el detector de cambios
        
        Args:
            id_column: Columna que identifica únicamente cada oportunidad
            tracked_columns: Columnas a monitorear para cambios
        """
        self.id_column = id_column or COLUMNS['id']
        self.tracked_columns = tracked_columns or TRACKED_COLUMNS
    
    def compare(self, current_df: pd.DataFrame, previous_df: pd.DataFrame) -> ComparisonResult:
        """
        Compara dos datasets y detecta todos los cambios
        
        Args:
            current_df: Dataset actual (más reciente)
            previous_df: Dataset anterior
            
        Returns:
            ComparisonResult con todos los cambios detectados
        """
        logger.info("Iniciando comparación de datasets...")
        logger.info(f"  Dataset actual: {len(current_df)} registros")
        logger.info(f"  Dataset anterior: {len(previous_df)} registros")
        
        # Obtener IDs únicos
        current_ids = set(current_df[self.id_column].unique())
        previous_ids = set(previous_df[self.id_column].unique())
        
        # Detectar nuevos y eliminados
        new_ids = current_ids - previous_ids
        removed_ids = previous_ids - current_ids
        common_ids = current_ids & previous_ids
        
        # DataFrames de nuevos y eliminados
        new_opportunities = current_df[current_df[self.id_column].isin(new_ids)].copy()
        removed_opportunities = previous_df[previous_df[self.id_column].isin(removed_ids)].copy()
        
        # Detectar cambios en oportunidades existentes
        changes = []
        unchanged_ids = []
        
        for opp_id in common_ids:
            current_row = current_df[current_df[self.id_column] == opp_id].iloc[0]
            previous_row = previous_df[previous_df[self.id_column] == opp_id].iloc[0]
            
            opp_changes = self._detect_row_changes(opp_id, current_row, previous_row)
            
            if opp_changes:
                changes.extend(opp_changes)
            else:
                unchanged_ids.append(opp_id)
        
        unchanged = current_df[current_df[self.id_column].isin(unchanged_ids)].copy()
        
        # Generar resumen
        summary = self._generate_summary(
            current_df, previous_df,
            new_opportunities, removed_opportunities,
            changes, unchanged
        )
        
        logger.info(f"Comparación completada:")
        logger.info(f"  - Nuevas: {len(new_opportunities)}")
        logger.info(f"  - Eliminadas: {len(removed_opportunities)}")
        logger.info(f"  - Con cambios: {len(set(c.opportunity_id for c in changes))}")
        logger.info(f"  - Sin cambios: {len(unchanged)}")
        
        return ComparisonResult(
            new_opportunities=new_opportunities,
            removed_opportunities=removed_opportunities,
            changes=changes,
            unchanged=unchanged,
            summary=summary
        )
    
    def _detect_row_changes(self, opp_id: str, current: pd.Series, 
                            previous: pd.Series) -> List[ChangeRecord]:
        """Detecta cambios entre dos versiones de la misma oportunidad"""
        changes = []
        responsible = current.get(COLUMNS['responsible'], 'Sin Asignar')
        market = current.get(COLUMNS['market'], 'Sin País')
        
        for column in self.tracked_columns:
            if column not in current.index or column not in previous.index:
                continue
            
            current_val = current[column]
            previous_val = previous[column]
            
            # Manejar NaN
            if pd.isna(current_val) and pd.isna(previous_val):
                continue
            
            # Detectar cambio
            if not self._values_equal(current_val, previous_val):
                change_type = self._classify_change(column, current_val, previous_val)
                
                changes.append(ChangeRecord(
                    opportunity_id=str(opp_id),
                    field=column,
                    old_value=self._format_value(previous_val),
                    new_value=self._format_value(current_val),
                    change_type=change_type,
                    responsible=responsible,
                    market=market
                ))
        
        return changes
    
    def _values_equal(self, val1, val2) -> bool:
        """Compara dos valores manejando tipos especiales"""
        if pd.isna(val1) and pd.isna(val2):
            return True
        if pd.isna(val1) or pd.isna(val2):
            return False
        
        # Para fechas, comparar solo la fecha (ignorar hora)
        if isinstance(val1, (datetime, pd.Timestamp)):
            val1 = pd.Timestamp(val1).date()
        if isinstance(val2, (datetime, pd.Timestamp)):
            val2 = pd.Timestamp(val2).date()
        
        return val1 == val2
    
    def _format_value(self, value) -> str:
        """Formatea un valor para mostrar"""
        if pd.isna(value):
            return "N/A"
        if isinstance(value, (datetime, pd.Timestamp)):
            return pd.Timestamp(value).strftime('%Y-%m-%d')
        if isinstance(value, float):
            return f"{value:,.2f}"
        return str(value)
    
    def _classify_change(self, column: str, new_val, old_val) -> str:
        """Clasifica el tipo de cambio"""
        if column == COLUMNS['stage']:
            return self._classify_stage_change(new_val, old_val)
        elif column == COLUMNS['usd']:
            if pd.to_numeric(new_val, errors='coerce') > pd.to_numeric(old_val, errors='coerce'):
                return 'value_increase'
            else:
                return 'value_decrease'
        elif column == COLUMNS['responsible']:
            return 'reassignment'
        elif column == COLUMNS['close_date']:
            return 'reschedule'
        else:
            return 'modified'
    
    def _classify_stage_change(self, new_stage: str, old_stage: str) -> str:
        """Clasifica si un cambio de stage es avance o retroceso"""
        try:
            new_idx = STAGE_ORDER.index(new_stage) if new_stage in STAGE_ORDER else -1
            old_idx = STAGE_ORDER.index(old_stage) if old_stage in STAGE_ORDER else -1
            
            if new_idx > old_idx:
                return 'stage_advance'
            elif new_idx < old_idx:
                return 'stage_regress'
            else:
                return 'stage_change'
        except:
            return 'stage_change'
    
    def _generate_summary(self, current_df, previous_df, new_opps, 
                          removed_opps, changes, unchanged) -> Dict:
        """Genera un resumen estadístico de la comparación"""
        usd_col = COLUMNS['usd']
        
        # Contar cambios únicos por oportunidad
        unique_changed = len(set(c.opportunity_id for c in changes))
        
        # Cambios por tipo
        changes_by_type = {}
        for change in changes:
            change_type = change.change_type
            changes_by_type[change_type] = changes_by_type.get(change_type, 0) + 1
        
        # Cambios por responsable
        changes_by_responsible = {}
        for change in changes:
            resp = change.responsible
            changes_by_responsible[resp] = changes_by_responsible.get(resp, 0) + 1
        
        # Cambios por país
        changes_by_market = {}
        for change in changes:
            market = change.market
            changes_by_market[market] = changes_by_market.get(market, 0) + 1
        
        return {
            'total_current': len(current_df),
            'total_previous': len(previous_df),
            'new_count': len(new_opps),
            'removed_count': len(removed_opps),
            'changed_count': unique_changed,
            'unchanged_count': len(unchanged),
            'total_changes': len(changes),
            'new_usd': new_opps[usd_col].sum() if len(new_opps) > 0 else 0,
            'removed_usd': removed_opps[usd_col].sum() if len(removed_opps) > 0 else 0,
            'current_total_usd': current_df[usd_col].sum(),
            'previous_total_usd': previous_df[usd_col].sum(),
            'usd_change': current_df[usd_col].sum() - previous_df[usd_col].sum(),
            'changes_by_type': changes_by_type,
            'changes_by_responsible': changes_by_responsible,
            'changes_by_market': changes_by_market
        }


def compare_datasets(current_df: pd.DataFrame, 
                     previous_df: pd.DataFrame) -> ComparisonResult:
    """
    Función de conveniencia para comparar dos datasets
    
    Args:
        current_df: Dataset actual
        previous_df: Dataset anterior
        
    Returns:
        ComparisonResult con los cambios detectados
    """
    detector = ChangeDetector()
    return detector.compare(current_df, previous_df)
