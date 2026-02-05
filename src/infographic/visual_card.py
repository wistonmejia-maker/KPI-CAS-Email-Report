"""
Visual Card Service - Generación de tarjetas visuales PNG
==========================================================
Genera imágenes PNG a partir de datos de análisis para compartir
en Teams, WhatsApp, presentaciones, etc.
"""

import os
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any
import logging
import base64
from io import BytesIO

# Agregar path del proyecto
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.config import COLORS, KPI_CATEGORIES, KPI_SEVERITY_COLORS

logger = logging.getLogger(__name__)

# Directorio para guardar assets generados
ASSETS_DIR = Path(__file__).parent.parent.parent / "dist" / "assets"
ASSETS_DIR.mkdir(parents=True, exist_ok=True)

# Verificar si imgkit está disponible
try:
    import imgkit
    IMGKIT_AVAILABLE = True
    logger.info("imgkit disponible")
except ImportError:
    IMGKIT_AVAILABLE = False
    logger.warning("imgkit no disponible, usando fallback matplotlib")

# Verificar si wkhtmltoimage está instalado
def check_wkhtmltoimage() -> bool:
    """Verifica si wkhtmltoimage está instalado en el sistema"""
    import shutil
    return shutil.which('wkhtmltoimage') is not None

WKHTMLTOIMAGE_AVAILABLE = check_wkhtmltoimage()


class VisualCardService:
    """Servicio para generar tarjetas visuales"""
    
    def __init__(self):
        self.use_imgkit = IMGKIT_AVAILABLE and WKHTMLTOIMAGE_AVAILABLE
        if not self.use_imgkit:
            logger.info("Usando matplotlib como fallback para generación de imágenes")
    
    def generate_executive_card(self, data: Dict[str, Any], 
                                 session_id: Optional[str] = None) -> Path:
        """
        Genera una tarjeta ejecutiva como imagen PNG
        
        Args:
            data: Diccionario con métricas del análisis
            session_id: ID único para el archivo
            
        Returns:
            Path al archivo PNG generado
        """
        if session_id is None:
            session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        output_path = ASSETS_DIR / f"{session_id}_executive_card.png"
        
        if self.use_imgkit:
            return self._generate_with_imgkit(data, output_path)
        else:
            return self._generate_with_matplotlib(data, output_path)
    
    def _generate_with_imgkit(self, data: Dict[str, Any], output_path: Path) -> Path:
        """Genera imagen usando imgkit (requiere wkhtmltoimage)"""
        html = self._render_card_html(data)
        
        options = {
            'format': 'png',
            'width': '600',
            'quality': '100',
            'enable-local-file-access': None
        }
        
        imgkit.from_string(html, str(output_path), options=options)
        logger.info(f"Executive card generada: {output_path}")
        return output_path
    
    def _generate_with_matplotlib(self, data: Dict[str, Any], output_path: Path) -> Path:
        """Genera imagen usando matplotlib con diseño Premium Dark Glass"""
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        from matplotlib.collections import PatchCollection
        
        # Colores y Estilos
        BG_COLOR = '#0f172a'  # Slate 900
        CARD_BG = '#1e293b'   # Slate 800
        TEXT_WHITE = '#f8fafc'
        TEXT_GRAY = '#94a3b8'
        ACCENT = '#3b82f6'    # Blue 500
        
        # Configurar figura de alta resolución
        plt.style.use('dark_background')
        fig, ax = plt.subplots(figsize=(12, 8), facecolor=BG_COLOR)
        ax.set_facecolor(BG_COLOR)
        ax.set_xlim(0, 12)
        ax.set_ylim(0, 8)
        ax.axis('off')
        
        # --- HEADER ---
        # Logo placeholder (círculo)
        ax.add_patch(mpatches.Circle((0.8, 7.2), 0.4, color=ACCENT, alpha=0.8))
        ax.text(0.8, 7.2, "KPI", color='white', ha='center', va='center', fontweight='bold', fontsize=10)
        
        # Título y Subtítulo
        ax.text(1.5, 7.3, "EXECUTIVE DASHBOARD", color=TEXT_WHITE, fontsize=22, fontweight='bold', fontname='Arial')
        week_str = datetime.now().strftime('%Y-W%W')
        ax.text(1.5, 6.9, f"Semana {week_str} • Salesforce Data", color=TEXT_GRAY, fontsize=12)
        
        # Fecha generación
        ax.text(11.5, 7.1, datetime.now().strftime("%d %b %Y"), color=TEXT_GRAY, ha='right', fontsize=11)
        
        # Línea separadora
        ax.plot([0.5, 11.5], [6.5, 6.5], color=CARD_BG, linewidth=2)
        
        # --- SCORE CARD (Gauge Simulado) ---
        # Calculamos un "health score" simple
        total = data.get('total_opportunities', 1)
        stagnant = data.get('stagnant_count', 0)
        risk = data.get('at_risk_count', 0)
        health_score = max(0, 100 - int(((stagnant + risk) / total) * 100))
        
        score_color = '#22c55e' if health_score >= 80 else '#f59e0b' if health_score >= 60 else '#ef4444'
        
        # Círculo base del gauge
        ax.add_patch(mpatches.Wedge((10.5, 5.2), 1.2, 0, 180, color=CARD_BG, width=0.3))
        # Arco de progreso
        ax.add_patch(mpatches.Wedge((10.5, 5.2), 1.2, 180 - (health_score * 1.8), 180, color=score_color, width=0.3))
        
        ax.text(10.5, 5.3, f"{health_score}%", color='white', fontsize=20, fontweight='bold', ha='center')
        ax.text(10.5, 4.8, "HEALTH SCORE", color=TEXT_GRAY, fontsize=9, ha='center')

        # --- KEY METRICS ROW ---
        def draw_metric_card(x, y, label, value, sublabel, color_bar):
            # Fondo tarjeta
            rect = mpatches.FancyBboxPatch((x, y), 2.5, 1.5, boxstyle="round,pad=0.1", 
                                           facecolor=CARD_BG, edgecolor='none')
            ax.add_patch(rect)
            # Barra lateral color
            ax.add_patch(mpatches.Rectangle((x, y), 0.1, 1.7, color=color_bar))
            # Textos
            ax.text(x + 0.3, y + 1.1, label, color=TEXT_GRAY, fontsize=9)
            ax.text(x + 0.3, y + 0.5, value, color='white', fontsize=18, fontweight='bold')
            ax.text(x + 0.3, y + 0.2, sublabel, color=color_bar, fontsize=9)

        total_fmt = f"{total:,}"
        usd_fmt = f"${data.get('total_usd', 0):,.0f}"
        
        draw_metric_card(0.5, 4.5, "TOTAL OPORTUNIDADES", total_fmt, "Active Pipeline", ACCENT)
        draw_metric_card(3.3, 4.5, "VALOR TOTAL (USD)", usd_fmt, "Forecasted", '#10b981') # Emerald
        draw_metric_card(6.1, 4.5, "ESTANCADAS (>30d)", f"{stagnant:,}", f"{int(stagnant/total*100)}% del total", '#f59e0b') # Amber
        
        # --- CHANGES SECTION ---
        changes = data.get('changes') or {}
        new_cnt = 0
        rem_cnt = 0
        mod_cnt = 0
        if changes:
            new_cnt = changes.get('new_count', 0)
            rem_cnt = changes.get('removed_count', 0)
            mod_cnt = changes.get('changed_count', 0)
        
        ax.text(0.5, 3.8, "CAMBIOS SEMANALES", color=TEXT_WHITE, fontsize=12, fontweight='bold')
        
        # Grid de cambios
        ax.text(0.5, 3.2, f"🚀 {new_cnt} Nuevas", color='#4ade80', fontsize=11)
        ax.text(2.5, 3.2, f"🗑️ {rem_cnt} Eliminadas", color='#f87171', fontsize=11)
        ax.text(4.5, 3.2, f"✏️ {mod_cnt} Modificadas", color='#fbbf24', fontsize=11)
        
        # --- KPI GRID ---
        ax.text(0.5, 2.5, "TOP KPIs POR SEVERIDAD", color=TEXT_WHITE, fontsize=12, fontweight='bold')
        
        by_kpi = data.get('by_kpi', [])
        # Agrupar grid 2 filas x 3 columnas en orden
        for i, kpi in enumerate(by_kpi[:6]):
            col = i % 3
            row = i // 3
            
            x_pos = 0.5 + (col * 3.8)
            y_pos = 1.8 - (row * 0.8)
            
            sev = kpi.get('severity', 'medium')
            colors = {'high': '#f87171', 'medium': '#fbbf24', 'low': '#4ade80'}
            c = colors.get(sev, '#94a3b8')
            icon = {'high': '🔴', 'medium': '🟡', 'low': '🟢'}.get(sev, '⚪')
            
            # KPI Box minimalista
            rect = mpatches.FancyBboxPatch((x_pos, y_pos), 3.5, 0.6, boxstyle="round,pad=0.05", 
                                           facecolor=CARD_BG, edgecolor=c, linewidth=1, alpha=0.5)
            ax.add_patch(rect)
            
            ax.text(x_pos + 0.2, y_pos + 0.35, f"{icon} {kpi.get('code')}", color='white', fontsize=10, fontweight='bold')
            ax.text(x_pos + 0.2, y_pos + 0.15, kpi.get('name')[:25], color=TEXT_GRAY, fontsize=8)
            ax.text(x_pos + 3.2, y_pos + 0.25, str(kpi.get('count')), color='white', fontsize=12, fontweight='bold', ha='right')

        # --- FOOTER ---
        ax.text(6, 0.2, "Generado por KPI CAS System • Confidential", color='#475569', ha='center', fontsize=9)

        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor=BG_COLOR)
        plt.close()
        
        logger.info(f"Executive card generada (matplotlib v2): {output_path}")
        return output_path
    
    def _render_card_html(self, data: Dict[str, Any]) -> str:
        """Renderiza el HTML de la tarjeta ejecutiva"""
        # ... (código existente, no se usa con matplotlib) ...
        # (Mantener implementación original por si instalan wkhtmltoimage después)
        total = data.get('total_opportunities', 0)
        usd = data.get('total_usd', 0)
        stagnant = data.get('stagnant_count', 0)
        at_risk = data.get('at_risk_count', 0)
        
        changes = data.get('changes') or {}
        new_count = changes.get('new_count', 0) if changes else 0
        removed = changes.get('removed_count', 0) if changes else 0
        changed = changes.get('changed_count', 0) if changes else 0
        
        # KPIs HTML
        kpis_html = ""
        for kpi in data.get('by_kpi', [])[:6]:
            severity = kpi.get('severity', 'medium')
            color = {'high': '#ef4444', 'medium': '#f59e0b', 'low': '#22c55e'}.get(severity, '#888')
            indicator = {'high': '🔴', 'medium': '🟡', 'low': '🟢'}.get(severity, '⚪')
            kpis_html += f'''
            <div style="display:inline-block; margin:5px 10px; padding:5px 12px; 
                        background:{color}22; border-radius:15px; color:{color};">
                {indicator} {kpi.get('code', '')}: {kpi.get('count', 0)}
            </div>
            '''
        
        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                * {{ margin: 0; padding: 0; box-sizing: border-box; }}
                body {{
                    font-family: 'Segoe UI', Arial, sans-serif;
                    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
                    color: white;
                    padding: 20px;
                    width: 600px;
                }}
                .header {{
                    text-align: center;
                    margin-bottom: 20px;
                }}
                .header h1 {{ font-size: 28px; margin-bottom: 5px; }}
                .header .subtitle {{ color: #888; font-size: 14px; }}
                .metrics {{
                    display: flex;
                    justify-content: space-around;
                    margin: 20px 0;
                }}
                .metric-box {{
                    text-align: center;
                    padding: 15px 20px;
                    border-radius: 12px;
                    min-width: 120px;
                }}
                .metric-box.green {{ background: #22c55e; }}
                .metric-box.blue {{ background: #3b82f6; }}
                .metric-box.orange {{ background: #f59e0b; }}
                .metric-box.red {{ background: #ef4444; }}
                .metric-value {{ font-size: 24px; font-weight: bold; }}
                .metric-label {{ font-size: 11px; opacity: 0.9; margin-top: 5px; }}
                .changes {{
                    display: flex;
                    justify-content: center;
                    gap: 30px;
                    margin: 15px 0;
                    font-size: 14px;
                }}
                .changes .new {{ color: #22c55e; }}
                .changes .removed {{ color: #ef4444; }}
                .changes .modified {{ color: #f59e0b; }}
                .kpis {{
                    background: rgba(255,255,255,0.05);
                    border-radius: 12px;
                    padding: 15px;
                    margin-top: 15px;
                }}
                .kpis h3 {{ text-align: center; margin-bottom: 10px; font-size: 16px; }}
                .kpi-grid {{ text-align: center; }}
                .footer {{
                    text-align: center;
                    margin-top: 20px;
                    color: #666;
                    font-size: 11px;
                }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>📊 KPI DASHBOARD</h1>
                <div class="subtitle">Semana {datetime.now().strftime('%Y-W%W')}</div>
            </div>
            
            <div class="metrics">
                <div class="metric-box green">
                    <div class="metric-value">{total:,}</div>
                    <div class="metric-label">TOTAL</div>
                </div>
                <div class="metric-box blue">
                    <div class="metric-value">${usd:,.0f}</div>
                    <div class="metric-label">USD TOTAL</div>
                </div>
                <div class="metric-box orange">
                    <div class="metric-value">{stagnant:,}</div>
                    <div class="metric-label">ESTANCADAS</div>
                </div>
                <div class="metric-box red">
                    <div class="metric-value">{at_risk:,}</div>
                    <div class="metric-label">EN RIESGO</div>
                </div>
            </div>
            
            <div class="changes">
                <span class="new">✅ +{new_count} Nuevas</span>
                <span class="removed">❌ -{removed} Eliminadas</span>
                <span class="modified">🔄 {changed} Modificadas</span>
            </div>
            
            <div class="kpis">
                <h3>KPIs por Severidad</h3>
                <div class="kpi-grid">
                    {kpis_html}
                </div>
            </div>
            
            <div class="footer">
                Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}
            </div>
        </body>
        </html>
        '''
    
    def get_card_as_base64(self, data: Dict[str, Any]) -> str:
        """Genera la tarjeta y retorna como base64 para embedding"""
        output_path = self.generate_executive_card(data)
        
        with open(output_path, 'rb') as f:
            return base64.b64encode(f.read()).decode('utf-8')


# Función de conveniencia
def generate_executive_card(data: Dict[str, Any], 
                            session_id: Optional[str] = None) -> Path:
    """
    Genera una tarjeta ejecutiva PNG
    
    Args:
        data: Diccionario con métricas
        session_id: ID único opcional
        
    Returns:
        Path al archivo generado
    """
    service = VisualCardService()
    return service.generate_executive_card(data, session_id)
