"""
Metrics - Cálculo de KPIs y Métricas
=====================================
Módulo para calcular métricas de rendimiento y KPIs de oportunidades.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from dataclasses import dataclass
import logging

from .config import (
    COLUMNS, STAGE_ORDER, STAGNANT_DAYS_THRESHOLD,
    WARNING_DAYS_BEFORE_CLOSE, KPI_CATEGORIES
)

logger = logging.getLogger(__name__)


@dataclass
class ResponsibleMetrics:
    """Métricas de un responsable"""
    name: str
    total_opportunities: int
    total_usd: float
    avg_usd: float
    markets: List[str]
    kpis: Dict[str, int]
    stages: Dict[str, int]
    stagnant_count: int
    at_risk_count: int
    
    def to_dict(self) -> dict:
        return {
            'Responsable': self.name,
            'Total_Oportunidades': self.total_opportunities,
            'Total_USD': self.total_usd,
            'Promedio_USD': self.avg_usd,
            'Países': ', '.join(self.markets),
            'Oportunidades_Estancadas': self.stagnant_count,
            'Oportunidades_En_Riesgo': self.at_risk_count
        }


@dataclass
class MarketMetrics:
    """Métricas de un mercado/país"""
    name: str
    total_opportunities: int
    total_usd: float
    avg_usd: float
    responsibles: List[str]
    top_responsible: str
    kpis: Dict[str, int]
    stages: Dict[str, int]
    
    def to_dict(self) -> dict:
        return {
            'País': self.name,
            'Total_Oportunidades': self.total_opportunities,
            'Total_USD': self.total_usd,
            'Promedio_USD': self.avg_usd,
            'Num_Responsables': len(self.responsibles),
            'Top_Responsable': self.top_responsible
        }


class MetricsCalculator:
    """Clase para calcular métricas y KPIs"""
    
    def __init__(self, df: pd.DataFrame):
        """
        Inicializa el calculador con un DataFrame de oportunidades
        
        Args:
            df: DataFrame con oportunidades
        """
        self.df = df.copy()
        self._prepare_data()
    
    def _prepare_data(self):
        """Prepara los datos para cálculos"""
        # Asegurar tipos correctos
        self.df[COLUMNS['usd']] = pd.to_numeric(self.df[COLUMNS['usd']], errors='coerce').fillna(0)
        self.df[COLUMNS['close_date']] = pd.to_datetime(self.df[COLUMNS['close_date']], errors='coerce')
        self.df[COLUMNS['created_date']] = pd.to_datetime(self.df[COLUMNS['created_date']], errors='coerce')
        
        # Calcular días desde creación
        today = datetime.now()
        self.df['_days_since_creation'] = (today - self.df[COLUMNS['created_date']]).dt.days
        
        # Calcular días hasta cierre
        self.df['_days_to_close'] = (self.df[COLUMNS['close_date']] - today).dt.days
        
        # Marcar estancadas y en riesgo
        self.df['_is_stagnant'] = self.df['_days_since_creation'] > STAGNANT_DAYS_THRESHOLD
        self.df['_is_at_risk'] = self.df['_days_to_close'] < WARNING_DAYS_BEFORE_CLOSE
    
    def get_summary(self) -> Dict:
        """Obtiene resumen general de métricas"""
        return {
            'total_opportunities': len(self.df),
            'total_usd': float(self.df[COLUMNS['usd']].sum()),
            'avg_usd': float(self.df[COLUMNS['usd']].mean()),
            'unique_responsibles': self.df[COLUMNS['responsible']].nunique(),
            'unique_markets': self.df[COLUMNS['market']].nunique(),
            'unique_customers': self.df[COLUMNS['customer']].nunique(),
            'stagnant_count': int(self.df['_is_stagnant'].sum()),
            'at_risk_count': int(self.df['_is_at_risk'].sum()),
            'by_kpi': self.df.groupby(COLUMNS['kpi']).size().to_dict(),
            'by_stage': self.df.groupby(COLUMNS['stage']).size().to_dict(),
            'by_market': self.df.groupby(COLUMNS['market']).size().to_dict()
        }
    
    def get_responsible_metrics(self) -> List[ResponsibleMetrics]:
        """Calcula métricas por responsable"""
        metrics = []
        
        for responsible, group in self.df.groupby(COLUMNS['responsible']):
            metrics.append(ResponsibleMetrics(
                name=responsible,
                total_opportunities=len(group),
                total_usd=float(group[COLUMNS['usd']].sum()),
                avg_usd=float(group[COLUMNS['usd']].mean()),
                markets=group[COLUMNS['market']].unique().tolist(),
                kpis=group[COLUMNS['kpi']].value_counts().to_dict(),
                stages=group[COLUMNS['stage']].value_counts().to_dict(),
                stagnant_count=int(group['_is_stagnant'].sum()),
                at_risk_count=int(group['_is_at_risk'].sum())
            ))
        
        # Ordenar por total de oportunidades
        metrics.sort(key=lambda x: x.total_opportunities, reverse=True)
        return metrics
    
    def get_market_metrics(self) -> List[MarketMetrics]:
        """Calcula métricas por mercado/país"""
        metrics = []
        
        for market, group in self.df.groupby(COLUMNS['market']):
            resp_counts = group[COLUMNS['responsible']].value_counts()
            top_resp = resp_counts.index[0] if len(resp_counts) > 0 else 'N/A'
            
            metrics.append(MarketMetrics(
                name=market,
                total_opportunities=len(group),
                total_usd=float(group[COLUMNS['usd']].sum()),
                avg_usd=float(group[COLUMNS['usd']].mean()),
                responsibles=group[COLUMNS['responsible']].unique().tolist(),
                top_responsible=top_resp,
                kpis=group[COLUMNS['kpi']].value_counts().to_dict(),
                stages=group[COLUMNS['stage']].value_counts().to_dict()
            ))
        
        # Ordenar por total de oportunidades
        metrics.sort(key=lambda x: x.total_opportunities, reverse=True)
        return metrics
    
    def get_kpi_metrics(self) -> Dict[str, Dict]:
        """Calcula métricas por categoría de KPI"""
        result = {}
        
        for kpi, group in self.df.groupby(COLUMNS['kpi']):
            result[kpi] = {
                'count': len(group),
                'total_usd': float(group[COLUMNS['usd']].sum()),
                'avg_usd': float(group[COLUMNS['usd']].mean()),
                'by_market': group[COLUMNS['market']].value_counts().to_dict(),
                'by_stage': group[COLUMNS['stage']].value_counts().to_dict(),
                'category_info': KPI_CATEGORIES.get(kpi, {'name': kpi, 'type': 'unknown'})
            }
        
        return result
    
    def get_stage_distribution(self) -> pd.DataFrame:
        """Obtiene distribución por stage"""
        stage_df = self.df.groupby(COLUMNS['stage']).agg({
            COLUMNS['id']: 'count',
            COLUMNS['usd']: ['sum', 'mean']
        }).round(2)
        
        stage_df.columns = ['Count', 'Total_USD', 'Avg_USD']
        stage_df['Percentage'] = (stage_df['Count'] / len(self.df) * 100).round(1)
        
        # Ordenar por stage order si es posible
        try:
            stage_df = stage_df.reindex([s for s in STAGE_ORDER if s in stage_df.index])
        except:
            pass
        
        return stage_df.reset_index()
    
    def get_opportunities_to_update(self) -> pd.DataFrame:
        """
        Obtiene oportunidades que requieren atención
        (estancadas, en riesgo, o con alertas)
        """
        # Filtrar oportunidades que necesitan atención
        needs_attention = self.df[
            (self.df['_is_stagnant']) | 
            (self.df['_is_at_risk']) |
            (self.df['_days_to_close'] < 0)  # Vencidas
        ].copy()
        
        # Agregar columna de razón
        needs_attention['Razon_Alerta'] = ''
        needs_attention.loc[needs_attention['_days_to_close'] < 0, 'Razon_Alerta'] = 'VENCIDA'
        needs_attention.loc[
            (needs_attention['_is_at_risk']) & (needs_attention['_days_to_close'] >= 0), 
            'Razon_Alerta'
        ] = 'PROX_VENCER'
        needs_attention.loc[needs_attention['_is_stagnant'], 'Razon_Alerta'] += ' ESTANCADA'
        
        # Ordenar por prioridad
        needs_attention = needs_attention.sort_values(
            ['_days_to_close', '_days_since_creation'],
            ascending=[True, False]
        )
        
        return needs_attention
    
    def get_responsible_summary_df(self) -> pd.DataFrame:
        """Retorna DataFrame con resumen por responsable"""
        metrics = self.get_responsible_metrics()
        return pd.DataFrame([m.to_dict() for m in metrics])
    
    def get_market_summary_df(self) -> pd.DataFrame:
        """Retorna DataFrame con resumen por mercado"""
        metrics = self.get_market_metrics()
        return pd.DataFrame([m.to_dict() for m in metrics])
    
    def get_opportunities_for_responsible(self, responsible: str) -> pd.DataFrame:
        """Obtiene oportunidades de un responsable específico"""
        return self.df[self.df[COLUMNS['responsible']] == responsible].copy()
    
    def get_opportunities_for_market(self, market: str) -> pd.DataFrame:
        """Obtiene oportunidades de un mercado específico"""
        return self.df[self.df[COLUMNS['market']] == market].copy()


def calculate_metrics(df: pd.DataFrame) -> MetricsCalculator:
    """
    Función de conveniencia para crear un calculador de métricas
    
    Args:
        df: DataFrame con oportunidades
        
    Returns:
        MetricsCalculator configurado
    """
    return MetricsCalculator(df)
