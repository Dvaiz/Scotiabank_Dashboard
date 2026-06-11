# FASTCO x Scotiabank - Dashboard Operacional

Dashboard ejecutivo-operacional para seguimiento de productividad, gestion comercial, bitacora cliente, facturacion y correlaciones para la cuenta Scotiabank.

## 1) Objetivo

El proyecto genera un entregable HTML unico y portable (`dashboard_scotiabank_fastco.html`) a partir de multiples fuentes de datos. El dashboard no requiere servidor para visualizarse y se abre directamente en navegador.

## 2) Arquitectura (SQL-first)

La logica actual prioriza lectura desde SQL Server y mantiene fallback a Excel cuando la conexion SQL no esta disponible.

Flujo resumido:
1. Extrae datos desde SQL Server (si hay conectividad/driver).
2. Si una fuente falla, usa archivo local equivalente en `Data/`.
3. Normaliza y agrega metricas en Python (`Data/generate_dashboard.py`).
4. Inyecta constantes JS en `Data/template.html`.
5. Genera `dashboard_scotiabank_fastco.html` en la raiz.

## 3) Estructura del repositorio

```
.
├── Data/
│   ├── generate_dashboard.py
│   └── template.html
├── create_flow_diagram.py
├── flujo_operacional_scotiabank.png
├── dashboard_scotiabank_fastco.html
├── LOGO1.png
└── README.md
```

## 4) Inputs del ETL

### 4.1 Fuentes SQL principales

`Data/generate_dashboard.py` consulta principalmente:
- `ALERTAS.dbo.BITACORA`
- `PCVMEZA.QFASTCO_INFORMES.dbo.TBL_CIERRE_CALIDAD`
- `BASE_CARGAS.DBO.TBL_CARGAS_POR_PRODUCTO`
- `BASE_REPORTES.dbo.v3_Informe_x_ejecutivos_producto_dia`
- `COMISIONES.dbo.TBL_VENTAS_PERIODO`
- `ALERTAS.dbo.MAPA`

### 4.2 Fallback locales (cuando SQL no responde)

En `Data/`, el script puede usar:
- `DATA_HISTORIA_DETALLE.xlsx`
- `BITACORA.xlsx`
- `CARGAS.xlsx`
- `DETALLE_EJECUTIVO.xlsx`
- `FACTURACION_PROVISIONES.xlsx`
- `CALIDAD.xlsx`

### 4.3 Inputs embebidos en codigo

Se eliminaron dependencias manuales para:
- `RELACION_BITACORA_OT.txt` (hoy embebido como `OT_MAP_EMBEDDED`)
- `MODELO_FACTURACION.xlsx` (tarifas embebidas en `FACT_PROY_TARIFFS`)

## 5) Variables de entorno SQL

Variables soportadas:
- `SCOTIA_SQL_SERVER` (default: `192.168.100.136`)
- `SCOTIA_SQL_DATABASE` (default: `ALERTAS`)
- `SCOTIA_SQL_CONNECTION_STRING` (opcional; si existe, tiene prioridad)

El script prueba drivers ODBC compatibles y utiliza trusted connection.

## 6) Stack tecnologico

- Python 3.11+
- pandas
- numpy
- pyodbc (SQL Server)
- openpyxl (fallback Excel)
- requests (UF API)
- prophet (opcional, para una de las metodologias de proyeccion)
- Frontend: HTML + CSS + JavaScript + Chart.js (CDN)

## 7) Como ejecutar

Desde la raiz del proyecto:

```powershell
& ".\.venv\Scripts\python.exe" .\Data\generate_dashboard.py
```

Salida esperada:
- Archivo generado/actualizado: `dashboard_scotiabank_fastco.html`
- Mensajes de estado por cada bloque de carga

## 8) Pestañas del dashboard

- Macro
- Campanas
- Mapa
- Facturacion
- Oportunidades
- Ejecutivos
- Correlacion
- Diagnostico

Notas funcionales recientes:
- Campanas: en los comparativos "Gestion por Mes" y "Gestion por Dia" la seleccion de KPI es single-select.
- Campanas: cada grafico comparativo incluye una tabla lateral con Actual, Anterior y Prom. 3M.
- Mapa: nombres de campana/cartera mostrados desde origen SQL y etiquetados como `CARTERA | OT`.

## 9) Logica funcional clave

- Correlacion: cruza FASTCO, cliente y facturacion por producto/OT y mes.
- Ejecutivos: ranking, headcount, tiempos (incluyendo conversion robusta de unidad para `TURNO`).
- Facturacion: mezcla historico + proyeccion en base a UF y series de actividad.
- Campanas: comparativos mensuales y diarios con filtros globales y tabla lateral de lectura rapida.
- Mapa: visualizacion con etiquetas de campana alineadas al origen SQL (`CARTERA | OT`).

## 10) Diagrama de arquitectura

El diagrama se genera con:

```powershell
& ".\.venv\Scripts\python.exe" .\create_flow_diagram.py
```

Archivo de salida:
- `flujo_operacional_scotiabank.png`

## 11) Mantenimiento

- Si cambias logica visual o de componentes, editar `Data/template.html` y luego regenerar.
- Si cambias logica ETL o fuentes, editar `Data/generate_dashboard.py`.
- Mantener sincronizados `README.md` y `create_flow_diagram.py` cuando cambie la arquitectura.
