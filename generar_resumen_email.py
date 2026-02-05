"""
Resumen de Oportunidades para Email
Diseño basado en Dashboard Ejecutivo
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

# Configuración
DATA_DIR = Path(__file__).parent / "data" / "raw"

# Descripciones de KPIs
KPI_DESCRIPTIONS = {
    'DC001 NB': ('AGING CONTROL', 'Opps creadas > 9 meses (rev 0)'),
    'DC001 CHURN': ('AGING CONTROL', 'Opps creadas > 12 meses (rev 0)'),
    'DC002 NB': ('EXPIRED OPPS', 'Forecast anterior al cierre (exchange rate calendar)'),
    'DC002 CHURN': ('EXPIRED OPPS', 'Forecast anterior al cierre (full month)'),
    'DC003': ('ON HOLD', 'Opps en hold - Tab Concepts & Refresh Schedule'),
    'DC004': ('FINANCE W/O REVENUE', 'Opps sin revenue que deben ir a ready to bill'),
    'DC005': ('CONVERSION W/O SALES', 'Convertidas en < X días (Collo 2d, BTS 1d)'),
    'DC007': ('CHANGE MANAGEMENT', 'Opps creadas por otras áreas (no ventas)'),
    'DC008': ('AGING REPORTED TO FINANCE', 'Opps > 30 días en reported to finance'),
    'DC010': ('AMOUNT ZERO', 'Opps con Amount = 0, Excl ToP=TRUE, Sales Deal'),
    'DC011': ('ROLES & RESPONSIBILITIES', 'Opps que cambiaron a Actual (últimos 30 días)'),
}


def cargar_ultimo_csv():
    """Carga el archivo CSV más reciente"""
    csv_files = sorted(DATA_DIR.glob("*.csv"), reverse=True)
    if not csv_files:
        csv_files = sorted(Path(__file__).parent.glob("*.csv"), reverse=True)
    
    if not csv_files:
        print("❌ No se encontró ningún archivo CSV")
        return None
    
    print(f"📂 Cargando: {csv_files[0].name}")
    return pd.read_csv(csv_files[0], encoding='utf-8')


def cargar_csv_anterior():
    """Carga el segundo archivo CSV más reciente para comparación"""
    csv_files = sorted(DATA_DIR.glob("*.csv"), reverse=True)
    if not csv_files:
        csv_files = sorted(Path(__file__).parent.glob("*.csv"), reverse=True)
    
    if len(csv_files) < 2:
        print("ℹ️ No hay archivo anterior para comparar")
        return None
    
    print(f"📂 Comparando con: {csv_files[1].name}")
    return pd.read_csv(csv_files[1], encoding='utf-8')


def calcular_deltas(df_actual, df_anterior):
    """Calcula los deltas entre el archivo actual y el anterior"""
    if df_anterior is None:
        return {
            'total': None,
            'por_responsable': {},
            'por_pais': {},
            'por_kpi': {},
            'por_resp_kpi': {}
        }
    
    # Excluir Brasil y Mexico de ambos
    df_actual = df_actual[~df_actual['Market'].isin(['Brasil', 'Mexico'])]
    df_anterior = df_anterior[~df_anterior['Market'].isin(['Brasil', 'Mexico'])]
    
    # Delta total
    total_actual = len(df_actual)
    total_anterior = len(df_anterior)
    delta_total = total_actual - total_anterior
    
    # Delta por responsable
    resp_actual = df_actual.groupby('Responsible')['Id'].count()
    resp_anterior = df_anterior.groupby('Responsible')['Id'].count()
    delta_resp = {}
    for resp in resp_actual.index:
        ant = resp_anterior.get(resp, 0)
        delta_resp[resp] = resp_actual[resp] - ant
    
    # Delta por país
    pais_actual = df_actual.groupby('Market')['Id'].count()
    pais_anterior = df_anterior.groupby('Market')['Id'].count()
    delta_pais = {}
    for pais in pais_actual.index:
        ant = pais_anterior.get(pais, 0)
        delta_pais[pais] = pais_actual[pais] - ant
    
    # Delta por KPI
    kpi_actual = df_actual.groupby('KPI')['Id'].count()
    kpi_anterior = df_anterior.groupby('KPI')['Id'].count()
    delta_kpi = {}
    for kpi in kpi_actual.index:
        ant = kpi_anterior.get(kpi, 0)
        delta_kpi[kpi] = kpi_actual[kpi] - ant
    
    # Delta por Responsable×KPI
    resp_kpi_actual = df_actual.groupby(['KPI', 'Responsible'])['Id'].count()
    resp_kpi_anterior = df_anterior.groupby(['KPI', 'Responsible'])['Id'].count()
    delta_resp_kpi = {}
    for (kpi, resp) in resp_kpi_actual.index:
        ant = resp_kpi_anterior.get((kpi, resp), 0)
        if kpi not in delta_resp_kpi:
            delta_resp_kpi[kpi] = {}
        delta_resp_kpi[kpi][resp] = resp_kpi_actual[(kpi, resp)] - ant
    
    return {
        'total': delta_total,
        'por_responsable': delta_resp,
        'por_pais': delta_pais,
        'por_kpi': delta_kpi,
        'por_resp_kpi': delta_resp_kpi
    }


def formato_delta(delta, invertir=False):
    """Genera HTML para mostrar un delta con color y flecha estilo badge"""
    if delta is None:
        return ''
    
    if delta == 0:
        return ''
    
    # Para oportunidades: aumentar es malo (rojo), disminuir es bueno (verde)
    # invertir=True para casos donde aumentar es bueno
    if invertir:
        color = '#059669' if delta > 0 else '#E21F26'  # Verde si sube, Rojo ATC si baja
    else:
        color = '#E21F26' if delta > 0 else '#059669'  # Rojo ATC si sube, verde si baja
    
    arrow = '↑' if delta > 0 else '↓'
    sign = '+' if delta > 0 else ''
    
    return f'<span style="color:{color}; font-size:9px; font-weight:600;">{sign}{delta}{arrow}</span>'


def generar_html_profesional(df, deltas=None):
    """
    Genera HTML con diseño VisActor para email ejecutivo.
    Incluye matriz Responsable×KPI con semáforo, CTA y comparativa con datos anteriores.
    Compatible con Outlook: 600px fijo, CSS inline, tablas para layout.
    """
    if deltas is None:
        deltas = {'total': None, 'por_responsable': {}, 'por_pais': {}, 'por_kpi': {}}
    fecha = datetime.now().strftime("%d/%m/%Y")
    trimestre = f"Q{(datetime.now().month - 1) // 3 + 1} {datetime.now().year}"
    
    # Excluir Brasil y Mexico
    df = df[~df['Market'].isin(['Brasil', 'Mexico'])]
    
    total_opps = len(df)
    total_responsables = df['Responsible'].nunique()
    total_paises = df['Market'].nunique()
    total_kpis = df['KPI'].nunique()
    
    # ===========================================
    # UMBRALES DE REFERENCIA POR KPI (semáforo)
    # ===========================================
    KPI_THRESHOLDS = {
        'DC001 NB':    {'green': 50, 'yellow': 100},
        'DC001 CHURN': {'green': 100, 'yellow': 500},
        'DC002 NB':    {'green': 15, 'yellow': 50},
        'DC002 CHURN': {'green': 15, 'yellow': 50},
        'DC003':       {'green': 5, 'yellow': 10},
        'DC004':       {'green': 50, 'yellow': 100},
        'DC005':       {'green': 15, 'yellow': 30},
        'DC007':       {'green': 5, 'yellow': 10},
        'DC008':       {'green': 15, 'yellow': 30},
        'DC010':       {'green': 1, 'yellow': 15},
        'DC011':       {'green': 1, 'yellow': 1},
    }
    
    def get_kpi_color(kpi_name, count):
        """Retorna color semáforo según umbrales."""
        thresholds = KPI_THRESHOLDS.get(kpi_name, {'green': 20, 'yellow': 50})
        if count < thresholds['green']:
            return '#10b981'  # Verde
        elif count < thresholds['yellow']:
            return '#f59e0b'  # Amarillo
        else:
            return '#ef4444'  # Rojo
    
    # Datos por KPI
    kpis_data = df.groupby('KPI').agg({'Id': 'count'}).rename(columns={'Id': 'Total'}).reset_index()
    kpis_data = kpis_data.sort_values('Total', ascending=False)
    
    # Datos por País
    paises_data = df.groupby('Market').agg({'Id': 'count'}).rename(columns={'Id': 'Oportunidades'}).reset_index()
    paises_data = paises_data.sort_values('Oportunidades', ascending=False)
    max_pais = paises_data['Oportunidades'].max() if len(paises_data) > 0 else 1
    
    # Oportunidades por Responsable (Todos)
    opps_by_resp = df.groupby('Responsible')['Id'].count().sort_values(ascending=False)
    max_opps_resp = opps_by_resp.max() if len(opps_by_resp) > 0 else 1
    
    # Matriz Responsable × KPI (Top 5 responsables × Top 4 KPIs)
    top_responsables = df['Responsible'].value_counts().head(5).index.tolist()
    top_kpis = df['KPI'].value_counts().head(4).index.tolist()
    matriz_data = df[df['Responsible'].isin(top_responsables) & df['KPI'].isin(top_kpis)]
    pivot = matriz_data.groupby(['Responsible', 'KPI']).size().unstack(fill_value=0)
    
    # =================== HTML CON DISEÑO VISACTOR ===================
    html = f'''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KPI CAS - Reporte Ejecutivo</title>
</head>
<body style="margin:0; padding:0; background-color:#f8fafc; font-family:'Segoe UI', Arial, Helvetica, sans-serif;">
    <!-- Container principal 600px -->
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#f8fafc;">
        <tr>
            <td align="center" style="padding:24px 16px;">
                <table role="presentation" width="600" cellspacing="0" cellpadding="0" border="0" style="background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 4px 12px rgba(0,0,0,0.08);">
                    
                    <!-- HEADER -->
                    <tr>
                        <td style="background:linear-gradient(135deg, #003764 0%, #001a30 100%); padding:40px 32px; text-align:center;">
                            <p style="margin:0 0 12px 0; font-size:12px; letter-spacing:3px; text-transform:uppercase; color:rgba(255,255,255,0.6); font-weight:600;">
                                American Tower • Data Cleansing • {trimestre}
                            </p>
                            <h1 style="margin:0 0 8px 0; font-size:56px; font-weight:800; color:#ffffff; line-height:1;">
                                {total_opps:,}
                            </h1>
                            <p style="margin:0 0 12px 0; font-size:18px; color:rgba(255,255,255,0.9); font-weight:600;">
                                Oportunidades Activas
                            </p>
                            <p style="margin:0; font-size:14px; color:rgba(255,255,255,0.65);">
                                {formato_delta(deltas['total']) if deltas['total'] is not None else ''} vs reporte anterior
                            </p>
                        </td>
                    </tr>
                    
                    <!-- KPI CARDS ROW -->
                    <tr>
                        <td style="padding:32px 32px 32px 32px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                <tr>
                                    <!-- Card: Países -->
                                    <td width="33%" style="padding:0 8px 0 0;">
                                        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#ffffff; border:1px solid #e2e8f0; border-radius:12px; box-shadow:0 2px 4px rgba(0,0,0,0.04);">
                                            <tr>
                                                <td style="padding:24px 16px; text-align:center;">
                                                    <div style="margin-bottom:8px;">
                                                        <span style="display:inline-block; width:32px; height:32px; line-height:32px; background-color:#f1f5f9; border-radius:50%; color:#003764; font-size:16px;">🌐</span>
                                                    </div>
                                                    <p style="margin:0 0 2px 0; font-size:10px; text-transform:uppercase; color:#64748b; letter-spacing:1px; font-weight:700;">Opps por Pais</p>
                                                    <p style="margin:0; font-size:28px; font-weight:800; color:#003764;">{total_paises}</p>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <!-- Card: Responsables -->
                                    <td width="33%" style="padding:0 4px;">
                                        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#ffffff; border:1px solid #e2e8f0; border-radius:12px; box-shadow:0 2px 4px rgba(0,0,0,0.04);">
                                            <tr>
                                                <td style="padding:24px 16px; text-align:center;">
                                                    <div style="margin-bottom:8px;">
                                                        <span style="display:inline-block; width:32px; height:32px; line-height:32px; background-color:#f1f5f9; border-radius:50%; color:#003764; font-size:16px;">👥</span>
                                                    </div>
                                                    <p style="margin:0 0 2px 0; font-size:10px; text-transform:uppercase; color:#64748b; letter-spacing:1px; font-weight:700;">Opps por Resp.</p>
                                                    <p style="margin:0; font-size:28px; font-weight:800; color:#003764;">{total_responsables}</p>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <!-- Card: KPIs Activos -->
                                    <td width="33%" style="padding:0 0 0 8px;">
                                        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#ffffff; border:1px solid #e2e8f0; border-radius:12px; box-shadow:0 2px 4px rgba(0,0,0,0.04);">
                                            <tr>
                                                <td style="padding:24px 16px; text-align:center;">
                                                    <div style="margin-bottom:8px;">
                                                        <span style="display:inline-block; width:32px; height:32px; line-height:32px; background-color:#f1f5f9; border-radius:50%; color:#003764; font-size:16px;">📊</span>
                                                    </div>
                                                    <p style="margin:0 0 2px 0; font-size:10px; text-transform:uppercase; color:#64748b; letter-spacing:1px; font-weight:700;">Opps por KPI</p>
                                                    <p style="margin:0; font-size:28px; font-weight:800; color:#003764;">{total_kpis}</p>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    
                    <!-- SECCION: OPORTUNIDADES POR RESPONSABLE -->
                    <tr>
                        <td style="padding:0 32px 32px 32px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-bottom:2px solid #e2e8f0; margin-bottom:16px;">
                                <tr>
                                    <td style="padding-bottom:12px;">
                                        <p style="margin:0; font-size:16px; font-weight:700; color:#0f172a;">
                                            👤 Distribución de Oportunidades por Responsable
                                        </p>
                                        <p style="margin:4px 0 0 0; font-size:12px; color:#64748b;">
                                            Oportunidades asignadas por analista
                                        </p>
                                    </td>
                                </tr>
                            </table>
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
'''
    
    # Generar filas de Oportunidades por Responsable
    for resp_name, opps_count in opps_by_resp.items():
        bar_width = (opps_count / max_opps_resp) * 100 if max_opps_resp > 0 else 0
        resp_delta = deltas['por_responsable'].get(resp_name, 0)
        delta_html = formato_delta(resp_delta) if resp_delta != 0 else ''
        
        # Obtener KPIs de este responsable
        resp_kpis = df[df['Responsible'] == resp_name].groupby('KPI')['Id'].count().sort_values(ascending=False)
        kpi_badges = ''
        for kpi_name, kpi_count in resp_kpis.items():
            # Color basado en si es CHURN o no (Colores ATC)
            badge_color = '#E21F26' if 'CHURN' in kpi_name else '#003764'
            kpi_short = kpi_name.replace(' NB', '').replace(' CHURN', 'C')
            kpi_badges += f'<span style="display:inline-block; background-color:{badge_color}; color:#fff; font-size:8px; padding:1px 4px; border-radius:3px; margin-right:2px;">{kpi_short}:{kpi_count}</span>'
        
        html += f'''                                <tr>
                                    <td style="padding:6px 0;">
                                        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                            <tr>
                                                <td style="padding:0 0 4px 0;">
                                                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                                        <tr>
                                                            <td style="font-size:12px; color:#374151;">
                                                                {resp_name} {kpi_badges}
                                                            </td>
                                                            <td style="font-size:12px; font-weight:600; color:#0f172a; text-align:right; vertical-align:top;">{opps_count} {delta_html}</td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#f1f5f9; border-radius:10px; overflow:hidden;">
                                                        <tr>
                                                            <td style="height:10px; width:{bar_width:.0f}%; background:linear-gradient(90deg, #003764 0%, #00569c 100%); border-radius:10px;"></td>
                                                            <td style="height:10px;"></td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
'''
    
    html += '''                            </table>
                        </td>
                    </tr>
                    
                    <!-- SECCION: POR PAIS -->
                    <tr>
                        <td style="padding:0 32px 32px 32px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-bottom:2px solid #e2e8f0; margin-bottom:16px;">
                                <tr>
                                    <td style="padding-bottom:12px;">
                                        <p style="margin:0; font-size:16px; font-weight:700; color:#0f172a;">
                                            🌍 Distribución de Oportunidades por País
                                        </p>
                                        <p style="margin:4px 0 0 0; font-size:12px; color:#64748b;">
                                            Cobertura geográfica del pipeline
                                        </p>
                                    </td>
                                </tr>
                            </table>
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
'''
    
    # Generar filas de países - usar imágenes de banderas para compatibilidad con Outlook
    COUNTRY_CODES = {
        'Chile': 'cl',
        'Peru': 'pe',
        'Colombia': 'co',
        'Paraguay': 'py',
        'Costa Rica': 'cr',
        'Argentina': 'ar',
        'Ecuador': 'ec',
        'Uruguay': 'uy',
        'Bolivia': 'bo',
        'Venezuela': 've',
        'Panama': 'pa',
        'Guatemala': 'gt',
        'El Salvador': 'sv',
        'Honduras': 'hn',
        'Nicaragua': 'ni',
        'Dominican Republic': 'do',
        'Puerto Rico': 'pr',
    }
    
    for _, pais_row in paises_data.iterrows():
        pais = pais_row['Market']
        opps = pais_row['Oportunidades']
        percent = (opps / total_opps) * 100 if total_opps > 0 else 0
        bar_width = (opps / max_pais) * 100 if max_pais > 0 else 0
        pais_delta = deltas['por_pais'].get(pais, 0)
        delta_html = formato_delta(pais_delta) if pais_delta != 0 else ''
        country_code = COUNTRY_CODES.get(pais, 'xx')
        flag_img = f'<img src="https://flagcdn.com/20x15/{country_code}.png" width="20" height="15" alt="{pais}" style="vertical-align:middle; margin-right:4px;">'
        
        html += f'''                                <tr>
                                    <td style="padding:6px 0; border-bottom:1px solid #f1f5f9;">
                                        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                            <tr>
                                                <td width="110" style="font-size:12px; font-weight:500; color:#0f172a;">{flag_img}{pais}</td>
                                                <td width="40" style="font-size:14px; font-weight:700; color:#0f172a; text-align:center;">{opps}</td>
                                                <td style="padding:0 8px;">
                                                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#f1f5f9; border-radius:10px; overflow:hidden;">
                                                        <tr>
                                                            <td style="height:10px; width:{bar_width:.0f}%; background:linear-gradient(90deg, #003764 0%, #00569c 100%); border-radius:10px;"></td>
                                                            <td style="height:10px;"></td>
                                                        </tr>
                                                    </table>
                                                </td>
                                                <td width="45" style="font-size:11px; font-weight:600; color:#10b981; text-align:right;">{percent:.1f}%</td>
                                                <td width="50" style="text-align:right;">{delta_html}</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
'''
    
    html += '''                            </table>
                        </td>
                    </tr>
                    
                    <!-- SECCION: GRID DE KPIs -->
                    <tr>
                        <td style="padding:0 32px 24px 32px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-bottom:2px solid #e2e8f0; margin-bottom:16px;">
                                <tr>
                                    <td style="padding-bottom:12px;">
                                        <p style="margin:0; font-size:16px; font-weight:700; color:#0f172a;">
                                            🎯 KPIs de Control de Datos
                                        </p>
                                        <p style="margin:4px 0 0 0; font-size:12px; color:#64748b;">
                                            Oportunidades por tipo de problema
                                        </p>
                                    </td>
                                </tr>
                            </table>
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
'''
    
    # Separar KPIs normales y CHURN
    kpis_df = kpis_data.copy()
    kpis_normal = kpis_df[~kpis_df['KPI'].str.contains('CHURN')].sort_values('KPI').to_dict('records')
    kpis_churn = kpis_df[kpis_df['KPI'].str.contains('CHURN')].sort_values('KPI').to_dict('records')
    
    # Función para generar tarjeta KPI
    def generar_tarjeta_kpi(kpi_record, df_data, deltas_data):
        kpi_name = kpi_record['KPI']
        kpi_total = kpi_record['Total']
        kpi_info = KPI_DESCRIPTIONS.get(kpi_name, ('', ''))
        kpi_subtitle = kpi_info[0]
        
        # Para DC011 usar columna "User" en lugar de "Responsible"
        group_column = 'User' if kpi_name == 'DC011' else 'Responsible'
        kpi_df = df_data[df_data['KPI'] == kpi_name]
        top_resp_kpi = kpi_df.groupby(group_column)['Id'].count().sort_values(ascending=False)
        max_resp_kpi = top_resp_kpi.max() if len(top_resp_kpi) > 0 else 1
        
        resp_html = ''
        for resp_name, resp_count in top_resp_kpi.items():
            resp_delta = deltas_data['por_resp_kpi'].get(kpi_name, {}).get(resp_name, 0)
            delta_resp_html = formato_delta(resp_delta) if resp_delta != 0 else ''
            bar_pct = (resp_count / max_resp_kpi) * 100 if max_resp_kpi > 0 else 0
            # Color de barra y degradado (Navy ATC para normal)
            bar_color_start = '#003764'
            bar_color_end = '#00569c'
            
            # Si es CHURN, usar Rojo ATC
            if 'CHURN' in kpi_name:
                bar_color_start = '#E21F26'
                bar_color_end = '#f14b51'

            resp_html += f'''<div style="padding:4px 0;">
                                                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                                                <tr>
                                                                    <td style="font-size:10px; color:#374151; padding-bottom:2px;">{resp_name}</td>
                                                                    <td style="font-size:10px; font-weight:600; color:#0f172a; text-align:right; padding-bottom:2px;">{resp_count} {delta_resp_html}</td>
                                                                </tr>
                                                            </table>
                                                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#f1f5f9; border-radius:6px; overflow:hidden;">
                                                                <tr>
                                                                    <td style="height:6px; width:{bar_pct:.0f}%; background:linear-gradient(90deg, {bar_color_start} 0%, {bar_color_end} 100%); border-radius:6px;"></td>
                                                                    <td style="height:6px;"></td>
                                                                </tr>
                                                            </table>
                                                        </div>'''
        
        # Colores de borde vibrantes para destacar las tarjetas
        border_color = '#dc2626' if 'CHURN' in kpi_name else '#4f46e5' 
        text_accent = border_color
        
        kpi_delta = deltas_data['por_kpi'].get(kpi_name, 0)
        delta_html = formato_delta(kpi_delta) if kpi_delta != 0 else ''
        
        return f'''<td width="50%" style="padding:4px; vertical-align:top;">
                                        <table role="presentation" width="100%" height="100%" cellspacing="0" cellpadding="0" border="0" style="background-color:#ffffff; border:1px solid #e2e8f0; border-top:4px solid {border_color}; border-radius:6px;">
                                            <tr>
                                                <td style="padding:10px; vertical-align:top;">
                                                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                                                        <tr>
                                                            <td>
                                                                <p style="margin:0 0 2px 0; font-size:11px; font-weight:600; color:#0f172a;">{kpi_name}</p>
                                                                <p style="margin:0; font-size:8px; color:#94a3b8; text-transform:uppercase;">{kpi_subtitle}</p>
                                                            </td>
                                                            <td style="text-align:right; vertical-align:top;">
                                                                <p style="margin:0; font-size:20px; font-weight:700; color:{text_accent};">{kpi_total}</p>
                                                                <p style="margin:0;">{delta_html}</p>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <div style="margin-top:6px; border-top:1px solid #e2e8f0; padding-top:6px;">
                                                        {resp_html}
                                                    </div>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>'''
    
    # Grid de KPIs normales (2 columnas)
    for i in range(0, len(kpis_normal), 2):
        html += '                                <tr>\n'
        for j in range(2):
            if i + j < len(kpis_normal):
                html += generar_tarjeta_kpi(kpis_normal[i + j], df, deltas) + '\n'
            else:
                html += '                                    <td width="50%"></td>\n'
        html += '                                </tr>\n'
    
    html += '''                            </table>
                        </td>
                    </tr>
                    
                    <!-- SECCION: KPIs CHURN -->
                    <tr>
                        <td style="padding:0 32px 32px 32px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-bottom:2px solid #fee2e2; margin-bottom:16px;">
                                <tr>
                                    <td style="padding-bottom:12px;">
                                        <p style="margin:0; font-size:16px; font-weight:700; color:#E21F26;">
                                            ⚠️ KPIs de Riesgo (CHURN)
                                        </p>
                                        <p style="margin:4px 0 0 0; font-size:12px; color:#f14b51;">
                                            Oportunidades en riesgo de pérdida
                                        </p>
                                    </td>
                                </tr>
                            </table>
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
'''
    
    # Grid de KPIs CHURN (2 columnas)
    for i in range(0, len(kpis_churn), 2):
        html += '                                <tr>\n'
        for j in range(2):
            if i + j < len(kpis_churn):
                html += generar_tarjeta_kpi(kpis_churn[i + j], df, deltas) + '\n'
            else:
                html += '                                    <td width="50%"></td>\n'
        html += '                                </tr>\n'
    
    html += f'''                            </table>
                        </td>
                    </tr>
                    
                     <!-- CALL TO ACTION -->
                    <tr>
                        <td style="padding:0 32px 32px 32px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background:linear-gradient(135deg, #003764 0%, #002544 100%); border-radius:12px; overflow:hidden; box-shadow:0 4px 12px rgba(0, 55, 100, 0.25);">
                                <tr>
                                    <td style="padding:24px; text-align:center;">
                                        <p style="margin:0 0 6px 0; font-size:18px; font-weight:800; color:#ffffff; letter-spacing:0.5px;">
                                            ⚡ Gestión en Salesforce
                                        </p>
                                        <p style="margin:0 0 16px 0; font-size:13px; color:rgba(255,255,255,0.95);">
                                            Mantén el pipeline saludable actualizando fechas y estados directamente
                                        </p>
                                        <a href="https://amtowerus2.lightning.force.com" style="display:inline-block; padding:12px 32px; background:#E21F26; color:#ffffff; font-size:14px; font-weight:700; text-decoration:none; border-radius:8px; box-shadow:0 4px 8px rgba(226, 31, 38, 0.3);">
                                            📱 Ir a Salesforce
                                        </a>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    
                    <!-- FOOTER -->
                    <tr>
                        <td style="background-color:#f8fafc; padding:16px 24px; text-align:center; border-top:1px solid #e2e8f0;">
                            <p style="margin:0; font-size:11px; color:#94a3b8;">
                                Generado automáticamente • {fecha} • KPI CAS System
                            </p>
                        </td>
                    </tr>
                    
                </table>
            </td>
        </tr>
    </table>
</body>
</html>'''
    
    return html


def main():
    print("\n" + "=" * 60)
    print("🚀 GENERADOR DE RESUMEN PARA EMAIL")
    print("   Diseño: Dashboard Ejecutivo + Comparativa")
    print("=" * 60 + "\n")
    
    # Cargar datos actual y anterior
    df = cargar_ultimo_csv()
    if df is None:
        return
    
    df_anterior = cargar_csv_anterior()
    
    # Calcular deltas
    deltas = calcular_deltas(df, df_anterior)
    
    # Generar HTML profesional con deltas
    html = generar_html_profesional(df, deltas)
    
    # Guardar
    output_html = Path(__file__).parent / "reports" / "emails" / "resumen_email.html"
    output_html.parent.mkdir(parents=True, exist_ok=True)
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"✅ Archivo HTML guardado: {output_html}")
    print("\n💡 Abre el archivo en tu navegador, selecciona todo (Ctrl+A)")
    print("   y copia (Ctrl+C) para pegarlo en Outlook con formato.\n")


if __name__ == "__main__":
    main()
