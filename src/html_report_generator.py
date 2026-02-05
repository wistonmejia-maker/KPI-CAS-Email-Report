"""
HTML Report Generator - Generación de Reportes HTML/PDF con Gráficos
=====================================================================
Módulo para generar reportes ejecutivos en HTML con visualizaciones.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional
import logging
import base64
from io import BytesIO

try:
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.use('Agg')  # Backend sin GUI
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

from .config import (
    COLUMNS, EMAIL_REPORTS_DIR, COLORS,
    KPI_CATEGORIES, KPI_SEVERITY_COLORS,
    get_current_date_str, get_current_week_str
)
from .metrics import MetricsCalculator
from .change_detector import ComparisonResult

logger = logging.getLogger(__name__)


class HTMLReportGenerator:
    """Generador de reportes HTML con gráficos"""
    
    def __init__(self, df: pd.DataFrame, comparison: Optional[ComparisonResult] = None):
        """
        Inicializa el generador
        
        Args:
            df: DataFrame con oportunidades
            comparison: Resultado de comparación (opcional)
        """
        self.df = df.copy()
        self.comparison = comparison
        self.metrics = MetricsCalculator(df)
        self.charts = {}
    
    def generate_executive_report(self, output_path: Optional[Path] = None) -> Path:
        """
        Genera reporte ejecutivo HTML
        
        Args:
            output_path: Ruta de salida
            
        Returns:
            Path al archivo generado
        """
        if output_path is None:
            filename = f"{get_current_week_str()}_executive_report.html"
            output_path = EMAIL_REPORTS_DIR / filename
        
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Generar gráficos
        if MATPLOTLIB_AVAILABLE:
            self._generate_charts()
        
        # Generar HTML
        html_content = self._build_html()
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        logger.info(f"Reporte HTML generado: {output_path}")
        return output_path
    
    def generate_responsible_email(self, responsible: str, 
                                   output_path: Optional[Path] = None) -> Path:
        """
        Genera correo HTML para un responsable específico
        
        Args:
            responsible: Nombre del responsable
            output_path: Ruta de salida
            
        Returns:
            Path al archivo generado
        """
        if output_path is None:
            safe_name = "".join(c for c in responsible if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filename = f"{get_current_week_str()}_{safe_name}_email.html"
            output_path = EMAIL_REPORTS_DIR / filename
        
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Filtrar datos del responsable
        resp_df = self.df[self.df[COLUMNS['responsible']] == responsible]
        resp_metrics = MetricsCalculator(resp_df)
        
        # Generar HTML
        html_content = self._build_responsible_email(responsible, resp_df, resp_metrics)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return output_path
    
    def _generate_charts(self):
        """Genera gráficos para el reporte"""
        plt.style.use('ggplot')
        
        # Gráfico 1: Distribución por País
        self.charts['by_market'] = self._create_pie_chart(
            self.df[COLUMNS['market']].value_counts(),
            'Distribución por País'
        )
        
        # Gráfico 2: Distribución por KPI
        self.charts['by_kpi'] = self._create_bar_chart(
            self.df[COLUMNS['kpi']].value_counts().head(10),
            'Oportunidades por KPI',
            'KPI',
            'Cantidad'
        )
        
        # Gráfico 3: Top Responsables
        resp_counts = self.df[COLUMNS['responsible']].value_counts().head(10)
        self.charts['top_responsible'] = self._create_horizontal_bar(
            resp_counts,
            'Top 10 Responsables',
            'Oportunidades'
        )
        
        # Gráfico 4: Valor USD por País
        usd_by_market = self.df.groupby(COLUMNS['market'])[COLUMNS['usd']].sum().sort_values(ascending=False)
        self.charts['usd_by_market'] = self._create_bar_chart(
            usd_by_market,
            'Valor USD por País',
            'País',
            'USD'
        )
        
        # Gráfico 5: Distribución por Stage
        stage_counts = self.df[COLUMNS['stage']].value_counts().head(10)
        self.charts['by_stage'] = self._create_horizontal_bar(
            stage_counts,
            'Distribución por Etapa',
            'Oportunidades'
        )
    
    def _create_pie_chart(self, data: pd.Series, title: str) -> str:
        """Crea gráfico de pastel y retorna como base64"""
        fig, ax = plt.subplots(figsize=(8, 6))
        
        colors = plt.cm.Set3(np.linspace(0, 1, len(data)))
        wedges, texts, autotexts = ax.pie(
            data.values, 
            labels=data.index,
            autopct='%1.1f%%',
            colors=colors,
            startangle=90
        )
        
        ax.set_title(title, fontsize=14, fontweight='bold')
        plt.tight_layout()
        
        return self._fig_to_base64(fig)
    
    def _create_bar_chart(self, data: pd.Series, title: str, 
                          xlabel: str, ylabel: str) -> str:
        """Crea gráfico de barras y retorna como base64"""
        fig, ax = plt.subplots(figsize=(10, 6))
        
        colors = plt.cm.Blues(np.linspace(0.4, 0.8, len(data)))
        bars = ax.bar(range(len(data)), data.values, color=colors)
        
        ax.set_xticks(range(len(data)))
        ax.set_xticklabels(data.index, rotation=45, ha='right')
        ax.set_xlabel(xlabel)
        ax.set_ylabel(ylabel)
        ax.set_title(title, fontsize=14, fontweight='bold')
        
        # Agregar valores sobre las barras
        for bar, val in zip(bars, data.values):
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{val:,.0f}', ha='center', va='bottom', fontsize=9)
        
        plt.tight_layout()
        return self._fig_to_base64(fig)
    
    def _create_horizontal_bar(self, data: pd.Series, title: str, 
                               xlabel: str) -> str:
        """Crea gráfico de barras horizontales"""
        fig, ax = plt.subplots(figsize=(10, 6))
        
        colors = plt.cm.Greens(np.linspace(0.4, 0.8, len(data)))
        y_pos = range(len(data))
        
        bars = ax.barh(y_pos, data.values, color=colors)
        ax.set_yticks(y_pos)
        ax.set_yticklabels(data.index)
        ax.set_xlabel(xlabel)
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.invert_yaxis()
        
        # Agregar valores
        for bar, val in zip(bars, data.values):
            width = bar.get_width()
            ax.text(width + max(data.values)*0.01, bar.get_y() + bar.get_height()/2.,
                   f'{val:,.0f}', ha='left', va='center', fontsize=9)
        
        plt.tight_layout()
        return self._fig_to_base64(fig)
    
    def _fig_to_base64(self, fig) -> str:
        """Convierte figura matplotlib a string base64"""
        buffer = BytesIO()
        fig.savefig(buffer, format='png', dpi=100, bbox_inches='tight')
        plt.close(fig)
        buffer.seek(0)
        return base64.b64encode(buffer.read()).decode()
    
    def _build_html(self) -> str:
        """Construye el HTML completo del reporte"""
        summary = self.metrics.get_summary()
        
        html = f'''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte Ejecutivo - KPIs Oportunidades</title>
    <style>
        {self._get_styles()}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📊 Reporte Ejecutivo de Oportunidades</h1>
            <p class="subtitle">Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
        </header>
        
        <!-- Tarjetas de Resumen -->
        <section class="cards-section">
            <div class="card primary">
                <div class="card-value">{summary['total_opportunities']:,}</div>
                <div class="card-label">Total Oportunidades</div>
            </div>
            <div class="card success">
                <div class="card-value">${summary['total_usd']:,.0f}</div>
                <div class="card-label">Valor Total USD</div>
            </div>
            <div class="card warning">
                <div class="card-value">{summary['stagnant_count']}</div>
                <div class="card-label">Estancadas (+30 días)</div>
            </div>
            <div class="card danger">
                <div class="card-value">{summary['at_risk_count']}</div>
                <div class="card-label">En Riesgo (-7 días)</div>
            </div>
        </section>
        
        {self._build_changes_section() if self.comparison else ''}
        
        <!-- Gráficos -->
        <section class="charts-section">
            <h2>📈 Visualizaciones</h2>
            <div class="charts-grid">
                {self._build_chart_html('by_market', 'Distribución por País')}
                {self._build_chart_html('by_kpi', 'Por Categoría KPI')}
                {self._build_chart_html('top_responsible', 'Top Responsables')}
                {self._build_chart_html('usd_by_market', 'Valor por País')}
            </div>
        </section>
        
        <!-- Tabla Responsables -->
        <section class="table-section">
            <h2>👤 Resumen por Responsable</h2>
            {self._build_responsible_table()}
        </section>
        
        <!-- Tabla KPIs por Severidad -->
        <section class="table-section">
            <h2>🎯 KPIs de Data Cleansing</h2>
            {self._build_kpi_severity_table()}
        </section>
        
        <!-- Tabla Países -->
        <section class="table-section">
            <h2>🌍 Resumen por País</h2>
            {self._build_market_table()}
        </section>
        
        <!-- Oportunidades que requieren atención -->
        <section class="table-section">
            <h2>⚠️ Oportunidades que Requieren Atención</h2>
            {self._build_attention_table()}
        </section>
        
        <footer>
            <p>Sistema de Seguimiento de KPIs - Oportunidades Salesforce</p>
        </footer>
    </div>
</body>
</html>'''
        return html
    
    def _get_styles(self) -> str:
        """Retorna los estilos CSS"""
        return '''
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        header {
            background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        header h1 { font-size: 2.5em; margin-bottom: 10px; }
        header .subtitle { opacity: 0.9; font-size: 1.1em; }
        
        .cards-section {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }
        
        .card {
            background: white;
            border-radius: 15px;
            padding: 25px;
            text-align: center;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s;
        }
        
        .card:hover { transform: translateY(-5px); }
        
        .card-value {
            font-size: 2.5em;
            font-weight: bold;
            margin-bottom: 10px;
        }
        
        .card-label { color: #666; font-size: 0.95em; }
        
        .card.primary .card-value { color: #1a73e8; }
        .card.success .card-value { color: #34a853; }
        .card.warning .card-value { color: #fbbc04; }
        .card.danger .card-value { color: #ea4335; }
        
        .charts-section, .table-section {
            padding: 30px;
        }
        
        .charts-section h2, .table-section h2 {
            color: #202124;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #1a73e8;
        }
        
        .charts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 20px;
        }
        
        .chart-container {
            background: #f8f9fa;
            border-radius: 15px;
            padding: 20px;
            text-align: center;
        }
        
        .chart-container img {
            max-width: 100%;
            height: auto;
            border-radius: 10px;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }
        
        th, td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #e0e0e0;
        }
        
        th {
            background: #1a73e8;
            color: white;
            font-weight: 600;
        }
        
        tr:nth-child(even) { background: #f8f9fa; }
        tr:hover { background: #e3f2fd; }
        
        .alert-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85em;
        }
        
        .severity-high { background: #ffebee; color: #c62828; font-weight: 600; }
        .severity-medium { background: #fff8e1; color: #f57c00; font-weight: 600; }
        .severity-low { background: #e8f5e9; color: #2e7d32; font-weight: 600; }
        
        .kpi-badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 15px;
            font-size: 0.8em;
            font-weight: 600;
        }
        
        .alert-danger { background: #ffebee; color: #c62828; }
        .alert-warning { background: #fff8e1; color: #f57c00; }
        
        .changes-section {
            padding: 30px;
            background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
        }
        
        .changes-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin-top: 20px;
        }
        
        .change-card {
            background: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 3px 10px rgba(0,0,0,0.1);
        }
        
        .change-value { font-size: 2em; font-weight: bold; color: #1a73e8; }
        .change-label { color: #666; margin-top: 5px; }
        
        footer {
            background: #202124;
            color: white;
            text-align: center;
            padding: 20px;
        }
        
        @media print {
            body { background: white; padding: 0; }
            .container { box-shadow: none; }
            .card:hover { transform: none; }
        }
        '''
    
    def _build_chart_html(self, chart_key: str, title: str) -> str:
        """Construye HTML para un gráfico"""
        if chart_key not in self.charts:
            return f'<div class="chart-container"><p>Gráfico no disponible: {title}</p></div>'
        
        return f'''
        <div class="chart-container">
            <h3>{title}</h3>
            <img src="data:image/png;base64,{self.charts[chart_key]}" alt="{title}">
        </div>
        '''
    
    def _build_changes_section(self) -> str:
        """Construye sección de cambios"""
        if not self.comparison:
            return ''
        
        s = self.comparison.summary
        return f'''
        <section class="changes-section">
            <h2>🔄 Cambios vs Período Anterior</h2>
            <div class="changes-grid">
                <div class="change-card">
                    <div class="change-value" style="color: #34a853;">+{s['new_count']}</div>
                    <div class="change-label">Nuevas</div>
                </div>
                <div class="change-card">
                    <div class="change-value" style="color: #ea4335;">-{s['removed_count']}</div>
                    <div class="change-label">Eliminadas</div>
                </div>
                <div class="change-card">
                    <div class="change-value" style="color: #fbbc04;">{s['changed_count']}</div>
                    <div class="change-label">Modificadas</div>
                </div>
                <div class="change-card">
                    <div class="change-value">${s['usd_change']:,.0f}</div>
                    <div class="change-label">Cambio USD</div>
                </div>
            </div>
        </section>
        '''
    
    def _build_responsible_table(self) -> str:
        """Construye tabla de responsables"""
        df_resp = self.metrics.get_responsible_summary_df().head(15)
        
        rows = ''
        for _, row in df_resp.iterrows():
            rows += f'''
            <tr>
                <td>{row['Responsable']}</td>
                <td>{row['Total_Oportunidades']:,}</td>
                <td>${row['Total_USD']:,.2f}</td>
                <td>{row['Países']}</td>
                <td>{row['Oportunidades_Estancadas']}</td>
                <td>{row['Oportunidades_En_Riesgo']}</td>
            </tr>
            '''
        
        return f'''
        <table>
            <thead>
                <tr>
                    <th>Responsable</th>
                    <th>Oportunidades</th>
                    <th>Total USD</th>
                    <th>Países</th>
                    <th>Estancadas</th>
                    <th>En Riesgo</th>
            </thead>
            <tbody>{rows}</tbody>
        </table>
        '''
    
    def _build_kpi_severity_table(self) -> str:
        """Construye tabla de KPIs por severidad con colores"""
        # Contar oportunidades por KPI
        kpi_counts = self.df[COLUMNS['kpi']].value_counts().to_dict()
        total_opps = len(self.df)
        
        rows = ''
        for kpi_code, kpi_info in KPI_CATEGORIES.items():
            count = kpi_counts.get(kpi_code, 0)
            
            # Calcular valor para comparar con umbral
            thresholds = kpi_info.get('thresholds', {'green': 0, 'yellow': 0})
            is_percentage = kpi_info.get('is_percentage', False)
            
            if is_percentage and total_opps > 0:
                value = (count / total_opps) * 100
                display_value = f"{count} ({value:.1f}%)"
            else:
                value = count
                display_value = str(count)
            
            # Determinar color basado en umbrales
            if value <= thresholds['green']:
                severity_class = 'severity-low'
                indicator = '🟢'
            elif value <= thresholds['yellow']:
                severity_class = 'severity-medium'
                indicator = '🟡'
            else:
                severity_class = 'severity-high'
                indicator = '🔴'
            
            rows += f'''
            <tr>
                <td><span class="kpi-badge">{kpi_code}</span></td>
                <td>{kpi_info['name']}</td>
                <td><span class="alert-badge {severity_class}">{indicator} {display_value}</span></td>
                <td>{kpi_info['type'].title()}</td>
            </tr>
            '''
        
        return f'''
        <table>
            <thead>
                <tr>
                    <th>KPI</th>
                    <th>Nombre</th>
                    <th>Cantidad</th>
                    <th>Tipo</th>
                </tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>
        '''
    
    def _build_market_table(self) -> str:
        """Construye tabla de mercados"""
        df_market = self.metrics.get_market_summary_df()
        
        rows = ''
        for _, row in df_market.iterrows():
            rows += f'''
            <tr>
                <td>{row['País']}</td>
                <td>{row['Total_Oportunidades']:,}</td>
                <td>${row['Total_USD']:,.2f}</td>
                <td>{row['Num_Responsables']}</td>
                <td>{row['Top_Responsable']}</td>
            </tr>
            '''
        
        return f'''
        <table>
            <thead>
                <tr>
                    <th>País</th>
                    <th>Oportunidades</th>
                    <th>Total USD</th>
                    <th>Responsables</th>
                    <th>Top Responsable</th>
                </tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>
        '''
    
    def _build_attention_table(self) -> str:
        """Construye tabla de oportunidades que requieren atención"""
        attention = self.metrics.get_opportunities_to_update().head(20)
        
        if len(attention) == 0:
            return '<p style="color: #34a853; font-size: 1.2em;">✅ No hay oportunidades que requieran atención inmediata.</p>'
        
        rows = ''
        for _, row in attention.iterrows():
            alert_class = 'alert-danger' if row.get('_days_to_close', 0) < 0 else 'alert-warning'
            rows += f'''
            <tr>
                <td>{row.get(COLUMNS['id'], 'N/A')}</td>
                <td>{row.get(COLUMNS['responsible'], 'N/A')}</td>
                <td>{row.get(COLUMNS['market'], 'N/A')}</td>
                <td>{row.get(COLUMNS['stage'], 'N/A')}</td>
                <td>${row.get(COLUMNS['usd'], 0):,.2f}</td>
                <td><span class="alert-badge {alert_class}">{row.get('Razon_Alerta', 'N/A')}</span></td>
            </tr>
            '''
        
        return f'''
        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Responsable</th>
                    <th>País</th>
                    <th>Stage</th>
                    <th>USD</th>
                    <th>Alerta</th>
                </tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>
        '''
    
    def _build_responsible_email(self, responsible: str, 
                                  df: pd.DataFrame, 
                                  metrics: MetricsCalculator) -> str:
        """Construye email HTML para un responsable"""
        summary = metrics.get_summary()
        attention = metrics.get_opportunities_to_update()
        
        html = f'''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Reporte Semanal - {responsible}</title>
    <style>
        body {{ font-family: Arial, sans-serif; background: #f5f5f5; padding: 20px; }}
        .email-container {{ max-width: 700px; margin: 0 auto; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 5px 20px rgba(0,0,0,0.1); }}
        .header {{ background: linear-gradient(135deg, #1a73e8, #0d47a1); color: white; padding: 30px; text-align: center; }}
        .content {{ padding: 30px; }}
        .stats {{ display: flex; justify-content: space-around; margin: 20px 0; }}
        .stat {{ text-align: center; padding: 15px; }}
        .stat-value {{ font-size: 2em; font-weight: bold; color: #1a73e8; }}
        .stat-label {{ color: #666; }}
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th, td {{ padding: 10px; border-bottom: 1px solid #eee; text-align: left; }}
        th {{ background: #f8f9fa; }}
        .alert {{ background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; }}
        .footer {{ background: #f8f9fa; padding: 20px; text-align: center; color: #666; }}
    </style>
</head>
<body>
    <div class="email-container">
        <div class="header">
            <h1>📊 Reporte Semanal</h1>
            <p>Hola {responsible},</p>
            <p>Aquí está el resumen de tus oportunidades</p>
        </div>
        
        <div class="content">
            <div class="stats">
                <div class="stat">
                    <div class="stat-value">{summary['total_opportunities']}</div>
                    <div class="stat-label">Oportunidades</div>
                </div>
                <div class="stat">
                    <div class="stat-value">${summary['total_usd']:,.0f}</div>
                    <div class="stat-label">Valor Total</div>
                </div>
                <div class="stat">
                    <div class="stat-value">{summary['at_risk_count']}</div>
                    <div class="stat-label">En Riesgo</div>
                </div>
            </div>
            
            {f'<div class="alert"><strong>⚠️ Tienes {len(attention)} oportunidades que requieren atención</strong></div>' if len(attention) > 0 else ''}
            
            <h3>Distribución por País</h3>
            <table>
                <tr><th>País</th><th>Oportunidades</th><th>USD</th></tr>
                {''.join(f"<tr><td>{m}</td><td>{c}</td><td>${df[df[COLUMNS['market']]==m][COLUMNS['usd']].sum():,.2f}</td></tr>" for m, c in df[COLUMNS['market']].value_counts().items())}
            </table>
            
            {self._build_attention_table_simple(attention) if len(attention) > 0 else ''}
        </div>
        
        <div class="footer">
            <p>Sistema de Seguimiento de KPIs</p>
            <p>{datetime.now().strftime('%Y-%m-%d')}</p>
        </div>
    </div>
</body>
</html>'''
        return html
    
    def _build_attention_table_simple(self, df: pd.DataFrame) -> str:
        """Tabla simple de oportunidades por actualizar"""
        if len(df) == 0:
            return ''
        
        rows = ''
        for _, row in df.head(10).iterrows():
            rows += f'''
            <tr>
                <td>{row.get(COLUMNS['id'], 'N/A')}</td>
                <td>{row.get(COLUMNS['market'], 'N/A')}</td>
                <td>{row.get(COLUMNS['stage'], 'N/A')}</td>
                <td>{row.get('Razon_Alerta', 'N/A')}</td>
            </tr>
            '''
        
        return f'''
        <h3>⚠️ Oportunidades por Actualizar</h3>
        <table>
            <tr><th>ID</th><th>País</th><th>Stage</th><th>Alerta</th></tr>
            {rows}
        </table>
        '''


def generate_executive_html(df: pd.DataFrame,
                           comparison: Optional[ComparisonResult] = None,
                           output_path: Optional[Path] = None) -> Path:
    """
    Genera reporte ejecutivo HTML
    
    Args:
        df: DataFrame con oportunidades
        comparison: Comparación con período anterior
        output_path: Ruta de salida
        
    Returns:
        Path al archivo generado
    """
    generator = HTMLReportGenerator(df, comparison)
    return generator.generate_executive_report(output_path)


def generate_responsible_emails(df: pd.DataFrame,
                               output_dir: Optional[Path] = None) -> List[Path]:
    """
    Genera emails HTML para todos los responsables
    
    Args:
        df: DataFrame con oportunidades
        output_dir: Directorio de salida
        
    Returns:
        Lista de paths generados
    """
    generator = HTMLReportGenerator(df)
    files = []
    
    for responsible in df[COLUMNS['responsible']].unique():
        if pd.isna(responsible) or responsible == 'Sin Asignar':
            continue
        
        if output_dir:
            safe_name = "".join(c for c in responsible if c.isalnum() or c in (' ', '-', '_')).rstrip()
            output_path = output_dir / f"{safe_name}_email.html"
        else:
            output_path = None
        
        files.append(generator.generate_responsible_email(responsible, output_path))
    
    return files
