# Sistema de Seguimiento de KPIs - Oportunidades Salesforce

Sistema para gestionar y analizar oportunidades de Salesforce con reportes semanales, mensuales y detección de cambios.

## 📋 Características

- ✅ **Carga de datos**: Importa CSVs de Salesforce con validación
- ✅ **Detección de cambios**: Compara períodos y detecta nuevas, eliminadas y modificadas
- ✅ **Métricas**: Calcula KPIs por responsable, país, cliente y etapa
- ✅ **Reportes Excel**: Múltiples hojas con análisis detallado
- ✅ **Reportes HTML**: Visualizaciones ejecutivas con gráficos
- ✅ **Emails**: Reportes individuales por responsable

## 🚀 Instalación

```bash
# Instalar dependencias
pip install -r requirements.txt
```

## 📁 Estructura del Proyecto

```
KPI CAS/
├── data/
│   ├── raw/                    # CSVs descargados de Salesforce
│   ├── processed/              # Archivos procesados
│   └── snapshots/              # Cortes mensuales
├── reports/
│   ├── weekly/                 # Reportes semanales
│   ├── monthly/                # Reportes mensuales
│   └── emails/                 # Correos HTML por responsable
├── src/                        # Código fuente
│   ├── config.py              # Configuración
│   ├── data_loader.py         # Carga de datos
│   ├── change_detector.py     # Detección de cambios
│   ├── metrics.py             # Cálculo de métricas
│   ├── report_generator.py    # Reportes Excel
│   └── html_report_generator.py # Reportes HTML
├── run_weekly.py              # Script semanal
├── run_monthly.py             # Script mensual
└── requirements.txt           # Dependencias
```

## 📖 Uso

### Proceso Semanal

1. **Descargar CSV de Salesforce** y guardarlo en `data/raw/` o en la raíz del proyecto

2. **Ejecutar el proceso semanal**:
   ```bash
   python run_weekly.py
   ```

   Opciones:
   ```bash
   # Especificar archivo
   python run_weekly.py --file "archivo.csv"
   
   # Sin comparación con período anterior
   python run_weekly.py --no-compare
   
   # Sin generar emails individuales
   python run_weekly.py --no-emails
   
   # Sin reporte HTML
   python run_weekly.py --no-html
   ```

3. **Revisar los reportes generados** en `reports/weekly/`

### Proceso Mensual

1. **Ejecutar al final del mes**:
   ```bash
   python run_monthly.py
   ```

   Opciones:
   ```bash
   # Especificar mes
   python run_monthly.py --month 2026-02
   ```

2. **Revisar resultados** en:
   - `data/snapshots/` - Snapshot mensual
   - `reports/monthly/` - Reportes del mes

## 📊 Reportes Generados

### Reporte Semanal Excel
- **Resumen**: Métricas generales y alertas
- **Por_Responsable**: Detalle por cada responsable
- **Por_País**: Detalle por mercado
- **Por_KPI**: Distribución por categoría
- **Por_Stage**: Distribución por etapa
- **Por_Actualizar**: Oportunidades que requieren atención
- **Cambios**: Lista de cambios detectados (si hay comparación)
- **Datos_Completos**: Todas las oportunidades

### Reporte HTML Ejecutivo
- Dashboard visual con gráficos
- Tarjetas de métricas clave
- Tablas de responsables y países
- Alertas visuales
- Ideal para adjuntar en correos

### Emails por Responsable
- Resumen personalizado
- Lista de oportunidades
- Alertas específicas
- Formato listo para copiar/pegar en Outlook

## 🔄 Flujo de Trabajo Recomendado

### Semanal (cada lunes)
1. Descargar CSV actualizado de Salesforce
2. Colocar en `data/raw/` con nombre: `YYYYMMDD_opportunities.csv`
3. Ejecutar `python run_weekly.py`
4. Revisar reporte Excel y HTML
5. Enviar emails a responsables (copiar HTML o adjuntar Excel)

### Mensual (primer día del mes)
1. Ejecutar `python run_monthly.py`
2. Revisar snapshot y comparativa
3. Archivar reportes del mes anterior

## ⚙️ Configuración

Editar `src/config.py` para ajustar:
- **STAGNANT_DAYS_THRESHOLD**: Días para considerar oportunidad estancada (default: 30)
- **WARNING_DAYS_BEFORE_CLOSE**: Días de alerta antes de vencimiento (default: 7)
- **STAGE_ORDER**: Orden de etapas para medir avance/retroceso

## 📈 Métricas Disponibles

### Por Oportunidad
- Días sin cambio
- Cambio de stage (avance/retroceso)
- Riesgo de vencimiento

### Por Responsable
- Total oportunidades
- Valor cartera (USD)
- Oportunidades estancadas
- Oportunidades en riesgo

### Por País
- Volumen de oportunidades
- Valor pipeline (USD)
- Top responsables

## 🔍 Detección de Cambios

El sistema detecta automáticamente:
- **Nuevas oportunidades**: No existían en el período anterior
- **Oportunidades cerradas**: Ya no aparecen
- **Cambios de stage**: Avance o retroceso en el proceso
- **Reasignaciones**: Cambio de responsable
- **Cambios de valor**: Modificaciones en USD
- **Reprogramaciones**: Cambio de fecha de cierre

## 📝 Notas

- El primer archivo procesado será la línea base (sin comparación)
- Los snapshots mensuales permiten análisis histórico
- Los reportes HTML requieren matplotlib para gráficos
- Todos los archivos se generan con fechas para trazabilidad

## 🆘 Solución de Problemas

### "No se encontró archivo CSV"
- Verificar que el archivo esté en `data/raw/` o especificar con `--file`

### "matplotlib no disponible"
- Ejecutar: `pip install matplotlib`
- Los reportes HTML se generarán sin gráficos

### "openpyxl no disponible"
- Ejecutar: `pip install openpyxl`
- Necesario para reportes Excel

---

**Desarrollado para el seguimiento de KPIs CAS**
