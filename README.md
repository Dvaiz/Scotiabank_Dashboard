# FASTCO × Scotiabank — Dashboard Operacional

Dashboard interactivo de gestión y análisis operacional para la cuenta Scotiabank de FASTCO. Visualiza métricas de productividad, facturación, bitácora cliente, ejecutivos y correlación entre fuentes de datos.

## Estructura del Proyecto

```
├── dashboard_scotiabank_fastco.html   # Dashboard principal (abrir con doble-click)
├── LOGO1.png                          # Logo oficial FASTCO
├── README.md                          # Este archivo
├── Data/
│   ├── generate_dashboard.py          # Script ETL que regenera el dashboard
│   ├── template.html                  # Template HTML base para inyección de datos
│   ├── DATA_HISTORIA_DETALLE.xlsx     # Gestiones, compromisos, montos por ejecutivo/periodo
│   ├── Bitacora.xlsx                  # Bitácora del cliente (operaciones diarias por producto)
│   ├── CARGAS.xlsx                    # Registros y gestiones por periodo (eficiencia)
│   ├── Detalle_Ejecutivo_2026.xlsx    # Detalle mensual por ejecutivo (monto, cantidad, tipo)
│   ├── FACTURACION_PROVISIONES.xlsx   # Facturación real vs provisiones por OT/mes
│   ├── Modelo_Facturacion.xlsx        # Tarifas C/IVA por producto para proyección
│   └── RELACION_BITACORA_OT.txt      # Mapeo de productos bitácora ↔ OT facturación
```

## Requisitos

- **Python 3.10+**
- **pandas** (`pip install pandas`)
- **openpyxl** (`pip install openpyxl`)

## Uso

### Regenerar el Dashboard

```powershell
cd Data
python generate_dashboard.py
```

El script:
1. Lee los 6 archivos de datos Excel/txt
2. Procesa y cruza la información (ETL)
3. Inyecta los datos como constantes JS en `template.html`
4. Genera `dashboard_scotiabank_fastco.html` en la raíz del proyecto

### Visualizar

Abrir `dashboard_scotiabank_fastco.html` con doble-click en cualquier navegador moderno. No requiere servidor web.

## Pestañas del Dashboard

| Pestaña | Descripción |
|---------|-------------|
| **Macro** | Resumen ejecutivo con KPIs principales y tendencias globales |
| **Campañas** | Desglose por tipo de campaña (Avance, TDC, PER, Plan Zero, etc.) |
| **Bitácora** | Operaciones del cliente por día/mes, cruce seguro, composición |
| **Facturación** | Facturación histórica, mensual, proyección por bitácora × tarifas |
| **Oportunidades** | Palancas de crecimiento y simulador de escenarios |
| **Ejecutivos** | Ranking, headcount, análisis de tiempos (TMO), productividad |
| **Correlación** | Cruce entre datos FASTCO vs cliente, facturación vs provisión, HC |
| **Diagnóstico** | Análisis Situación–Complicación–Resolución con datos reales |

## Fuentes de Datos

| Archivo | Contenido | Métricas clave |
|---------|-----------|----------------|
| `DATA_HISTORIA_DETALLE.xlsx` | Gestiones CRM | Gestiones, compromisos, montos, contactabilidad, HC |
| `Bitacora.xlsx` | Reporte del cliente | Operaciones por producto/día, montos, cruce seguro |
| `CARGAS.xlsx` | Asignación de registros | Registros cargados vs gestionados (eficiencia) |
| `Detalle_Ejecutivo_2026.xlsx` | Performance individual | Monto, cantidad, tipo por ejecutivo/mes |
| `FACTURACION_PROVISIONES.xlsx` | Billing | Provisión vs facturación real por OT |
| `Modelo_Facturacion.xlsx` | Tarifas vigentes | Precio C/IVA por producto para proyección |
| `RELACION_BITACORA_OT.txt` | Mapeo | Relación columna bitácora ↔ producto/OT |

## Modelo de Facturación

- **TDC**: Se factura por cantidad de operaciones (ventas), NO por monto
  - TC Titular: $28.9K/op
  - TC Adicional: $33.8K/op
- **Avance Externo**: $15.4K/op
- **Pago Flexible**: $15.2K/op
- **Plan Zero PER**: $38.6K/op
- **Costos fijos**: $8M Backoffice + $10M WEB por mes

## Tecnologías

- **Frontend**: HTML5, CSS3, Chart.js 4.4.1 (CDN)
- **ETL**: Python 3 + pandas + openpyxl
- **Fuentes**: Plus Jakarta Sans, DM Mono, Syne (Google Fonts CDN), Calibri (títulos)
- **Sin servidor**: Funciona 100% offline via `file://`

## Notas

- Al actualizar cualquier archivo Excel, ejecutar `python Data/generate_dashboard.py` para reflejar los cambios.
- El template (`Data/template.html`) debe mantenerse sincronizado con el dashboard antes de regenerar. El script lo usa como base para inyectar datos.
- Los gráficos de Bitácora se muestran por día al seleccionar un mes específico en el filtro.
