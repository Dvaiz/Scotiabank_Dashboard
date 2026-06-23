"""
FASTCO x Scotiabank — Generador automatico de Dashboard
=========================================================
Lee los archivos Excel de la carpeta Data/ y regenera el HTML
del dashboard con datos actualizados.

USO:
    python Data/generate_dashboard.py

REQUISITOS:
    pip install openpyxl pandas pyodbc

RESULTADO:
    Regenera dashboard_scotiabank_fastco.html con datos frescos.
"""

import json
import os
import re
import sys
from pathlib import Path
from datetime import date, timedelta, datetime

try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("ERROR: Requiere pandas. Instala con: pip install pandas openpyxl")
    sys.exit(1)

try:
    import pyodbc
    _HAS_PYODBC = True
except ImportError:
    _HAS_PYODBC = False
    print("  WARN: pyodbc no instalado. Se usara fallback a Excel. Instala con: pip install pyodbc")

try:
    import requests
    _HAS_REQUESTS = True
except ImportError:
    _HAS_REQUESTS = False
    print("  WARN: requests no instalado. Fetch UF deshabilitado. Instala con: pip install requests")

try:
    from prophet import Prophet
    import warnings as _warnings
    _warnings.filterwarnings('ignore')
    _HAS_PROPHET = True
except ImportError:
    _HAS_PROPHET = False
    print("  WARN: prophet no instalado. Proyección usará solo 2 métodos. Instala con: pip install prophet")

# Paths
SCRIPT_DIR = Path(__file__).parent
PROJECT_DIR = SCRIPT_DIR.parent
DATA_DIR = SCRIPT_DIR
TEMPLATE_PATH = DATA_DIR / "template.html"
OUTPUT_PATH = PROJECT_DIR / "dashboard_scotiabank_fastco.html"

SQL_SERVER = os.getenv("SCOTIA_SQL_SERVER", "192.168.100.136")
SQL_DATABASE = os.getenv("SCOTIA_SQL_DATABASE", "ALERTAS")
SQL_CONNECTION_STRING = os.getenv("SCOTIA_SQL_CONNECTION_STRING", "")
SQL_DRIVER_CANDIDATES = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
]

OT_MAP_EMBEDDED = [
        {'prod': 'PLAN ZERO PER', 'ot': 1320030},
        {'prod': 'TC ADICIONAL', 'ot': 1220056},
        {'prod': 'TC TITULAR', 'ot': 1220055},
        {'prod': 'PAGO FLEXIBLE', 'ot': 1220007},
        {'prod': 'AVANCE EXTERNO', 'ot': 1220005},
]

QUERY_BITACORA = """
SELECT *
FROM [ALERTAS].[dbo].[BITACORA] WITH (NOLOCK)
"""

QUERY_CALIDAD = """
SELECT
        ID_GESTION,
        RUT_AGENTE,
        USUARIO_AGENTE,
        FECHA_GESTION,
        CICLO,
        CAST(NOTA AS DECIMAL(10, 2)) AS NOTA,
        ID_PRODUCTO,
        PRODUCTO,
        RUT_DEUDOR
FROM PCVMEZA.QFASTCO_INFORMES.dbo.TBL_CIERRE_CALIDAD WITH (NOLOCK)
WHERE MES >= CONVERT(VARCHAR(6), DATEADD(month, -12, GETDATE()), 112)
    AND ID_PRODUCTO IN (325, 569, 444, 761, 740, 762, 760, 585, 784, 333)
"""

QUERY_CARGAS = """
SELECT *
FROM BASE_CARGAS.DBO.TBL_CARGAS_POR_PRODUCTO A WITH (NOLOCK)
LEFT JOIN BASE_REPORTES.DBO.v_Productos_Agrupados B WITH (NOLOCK)
        ON A.ID_PRODUCTO = B.ID_PRODUCTO
WHERE B.NOMBRE_PRODUCTO LIKE '%SCOTIA%'
    AND ID_CARGA_ARCHIVO IN (1, 7)
    AND PERIODO >= CONVERT(VARCHAR(6), DATEADD(month, -12, GETDATE()), 112)
    AND A.ID_PRODUCTO NOT IN (763,735,685,36,27,321,221,872,222,201,848,742,786,792)
"""

QUERY_DATA_HISTORIA = """
SELECT
        C.OT,
        A.PERIODO,
        A.FECHA,
        A.CODIGO_USUARIO,
        A.RUT_EJECUTIVO,
        A.NOMBRE_EJECUTIVO,
        A.ID_PRODUCTO,
        A.NOMBRE_PRODUCTO,
        A.GESTIONES_Q,
        A.EFECTIVOS_Q,
        A.COMPROMISOS_GENERADOS_Q,
        A.COMPROMISO_GENERADOS_MONTO,
        A.TURNO,
        A.HABLADO,
        A.Disponible,
        A.TMO,
        CASE
                WHEN A.ID_PRODUCTO IN (785,784,740,444,499) THEN 'TDC'
                WHEN A.ID_PRODUCTO IN (569,761,762,589,760,588) THEN 'REFINANCIAMIENTO'
                WHEN A.ID_PRODUCTO IN (585,325,333) THEN 'AVANCE'
        END AS TIPO
FROM [BASE_REPORTES].[dbo].[v3_Informe_x_ejecutivos_producto_dia] A WITH (NOLOCK)
LEFT JOIN BASE_REPORTES.DBO.v_Productos_Agrupados B WITH (NOLOCK)
        ON A.ID_PRODUCTO = B.ID_PRODUCTO
INNER JOIN NOMINA.DBO.OT C
        ON A.ID_PRODUCTO = C.ID_PRODUCTO
WHERE A.PERIODO >= CONVERT(VARCHAR(6), DATEADD(month, -12, GETDATE()), 112)
    AND C.OT IN (1220022,1220005,1220006,1220007,1220056,1320025,1220055,1320030)
    AND A.ID_PRODUCTO NOT IN (763,879,872,685,786,321,288,257)
"""

QUERY_DETALLE_EJECUTIVO = """
SELECT *
FROM [COMISIONES].[dbo].[TBL_VENTAS_PERIODO] WITH (NOLOCK)
WHERE PERIODO >= CONVERT(VARCHAR(6), DATEADD(month, -12, GETDATE()), 112)
    AND TIPO IN ('refinanciamiento', 'avance', 'TDC')
"""

QUERY_FACTURACION = """
SELECT *
FROM [ALERTAS].[dbo].[MAPA] WITH (NOLOCK)
"""

# Fixed tariff model embedded to avoid runtime dependency on MODELO_FACTURACION.xlsx.
FACT_PROY_TARIFFS = {
    'AVANCE EXTERNO': 15438.801055999998,
    'SEG. DESGRAVAMEN': 12973.782399999996,
    'PAGO FLEXIBLE': 15197.569789499998,
    'PLAN ZERO OPERA': 24123.12665,
    'PLAN ZERO PER': 38597.00264,
    'PER': 12244.00714,
    'TC TITULAR': 28947.751979999997,
    'TC ADICIONAL': 33772.377309999996,
    'TC DIGITAL': 33772.377309999996,
    'TC TSYS': 0.0,
}
FACT_PROY_BACKOFFICE = 8000000.0
FACT_PROY_WEB_EXTRA = 10000000.0


def _get_sql_connection_string():
    if SQL_CONNECTION_STRING:
        return SQL_CONNECTION_STRING
    if not _HAS_PYODBC:
        return None
    available = {driver.lower(): driver for driver in pyodbc.drivers()}
    selected = None
    for driver in SQL_DRIVER_CANDIDATES:
        if driver.lower() in available:
            selected = available[driver.lower()]
            break
    if not selected:
        print("  WARN: No se encontro driver ODBC SQL Server compatible. Se usara fallback a Excel.")
        return None
    return (
        f"DRIVER={{{selected}}};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={SQL_DATABASE};"
        "Trusted_Connection=yes;"
        "TrustServerCertificate=yes;"
    )


def _read_sql_df(query, source_name):
    conn_str = _get_sql_connection_string()
    if not conn_str:
        return None
    try:
        with pyodbc.connect(conn_str, timeout=60) as conn:
            return pd.read_sql(query, conn)
    except Exception as exc:
        print(f"  WARN: SQL no disponible para {source_name}. Fallback a Excel. Detalle: {exc}")
        return None


def _load_bitacora_source_df():
    df = _read_sql_df(QUERY_BITACORA, 'BITACORA')
    if df is None:
        path = DATA_DIR / "BITACORA.xlsx"
        df = pd.read_excel(path, engine='openpyxl')
    df.columns = [str(c).replace('\xa0', ' ').strip() for c in df.columns]
    return df


def _find_column(df, *patterns):
    upper_map = {str(c).strip().upper(): c for c in df.columns}
    for pattern in patterns:
        for upper_name, original in upper_map.items():
            if pattern.upper() in upper_name:
                return original
    return None


def _rolling_12m_cutoff_ym():
    """Return inclusive YYYY-MM cutoff for rolling 12 months."""
    return (pd.Timestamp.today().replace(day=1) - pd.DateOffset(months=12)).strftime('%Y-%m')


def safe_int(val):
    """Convert to int safely, handling NaN/None."""
    if pd.isna(val):
        return 0
    return int(val)


def safe_float(val):
    """Convert to float safely."""
    if pd.isna(val):
        return 0.0
    return float(val)


# ============================================================
# 1. RELACION_BITACORA_OT.txt
# ============================================================
def load_ot_map():
    colors = {
        'AVANCE EXTERNO': '#00c4b4',
        'PAGO FLEXIBLE': '#f5a623',
        'TC TITULAR': '#3d7cf4',
        'TC ADICIONAL': '#e8382a',
        'PLAN ZERO PER': '#27c47a',
    }
    result = []
    for row in OT_MAP_EMBEDDED:
        prod = row['prod'].strip()
        ot = int(row['ot'])
        result.append({'prod': prod, 'ot': ot, 'color': colors.get(prod, '#8892b0')})
    return result


# ============================================================
# 2. DATA_HISTORIA_DETALLE.xlsx
#    Cols: OT, PERIODO, FECHA, CODIGO_USUARIO, NOMBRE_EJECUTIVO,
#          GESTIONES_Q, EFECTIVOS_Q, COMPROMISOS_GENERADOS_Q,
#          COMPROMISO_GENERADOS_MONTO, TIPO
# ============================================================
def load_data_historia():
    df = _read_sql_df(QUERY_DATA_HISTORIA, 'DATA_HISTORIA_DETALLE')
    if df is None:
        path = DATA_DIR / "DATA_HISTORIA_DETALLE.xlsx"
        df = pd.read_excel(path, engine='openpyxl')

    df['PERIODO'] = df['PERIODO'].astype(int)
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['CODIGO_USUARIO'] = df['CODIGO_USUARIO'].astype(str).str.strip().str.upper()
    df['GESTIONES_Q'] = pd.to_numeric(df['GESTIONES_Q'], errors='coerce').fillna(0).astype(int)
    df['EFECTIVOS_Q'] = pd.to_numeric(df['EFECTIVOS_Q'], errors='coerce').fillna(0).astype(int)
    df['COMPROMISOS_GENERADOS_Q'] = pd.to_numeric(df['COMPROMISOS_GENERADOS_Q'], errors='coerce').fillna(0).astype(int)
    df['COMPROMISO_GENERADOS_MONTO'] = pd.to_numeric(df['COMPROMISO_GENERADOS_MONTO'], errors='coerce').fillna(0).astype(int)
    df['TIPO'] = df['TIPO'].astype(str).str.strip().str.upper()
    df = df[~df['TIPO'].isin(['NAN', 'NAT', ''])]

    # DH_PIVOT: group by PERIODO + TIPO
    grouped = df.groupby(['PERIODO', 'TIPO']).agg({
        'GESTIONES_Q': 'sum',
        'EFECTIVOS_Q': 'sum',
        'COMPROMISOS_GENERADOS_Q': 'sum',
        'COMPROMISO_GENERADOS_MONTO': 'sum',
    }).reset_index()

    dh_pivot = []
    for _, row in grouped.iterrows():
        dh_pivot.append({
            'p': int(row['PERIODO']),
            't': row['TIPO'],
            'g': int(row['GESTIONES_Q']),
            'e': int(row['EFECTIVOS_Q']),
            'c': int(row['COMPROMISOS_GENERADOS_Q']),
            'm': int(row['COMPROMISO_GENERADOS_MONTO']),
        })

    periodos = sorted(set(str(r['p']) for r in dh_pivot))

    meses = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC']
    per_labels = {}
    for p in periodos:
        yr = p[2:4]
        mm = int(p[4:6])
        per_labels[p] = f"{meses[mm-1]}-{yr}"

    # DH_DAILY: group by FECHA + TIPO for campañas daily comparison (with campaign filter support)
    df['_FECHA_STR'] = pd.to_datetime(df['FECHA'], errors='coerce').dt.strftime('%Y%m%d')
    df_d = df[df['_FECHA_STR'].notna() & (df['_FECHA_STR'] != 'NaT')]
    daily_grp = df_d.groupby(['_FECHA_STR', 'TIPO']).agg({
        'GESTIONES_Q': 'sum',
        'EFECTIVOS_Q': 'sum',
        'COMPROMISOS_GENERADOS_Q': 'sum',
        'COMPROMISO_GENERADOS_MONTO': 'sum',
    }).reset_index()
    dh_daily = sorted([{
        'd': row['_FECHA_STR'],
        't': row['TIPO'],
        'g': int(row['GESTIONES_Q']),
        'e': int(row['EFECTIVOS_Q']),
        'c': int(row['COMPROMISOS_GENERADOS_Q']),
        'm': int(row['COMPROMISO_GENERADOS_MONTO']),
    } for _, row in daily_grp.iterrows()], key=lambda x: x['d'])

    return dh_pivot, periodos, per_labels, df, dh_daily


# ============================================================
# 3. CARGAS.xlsx
#    Cols: PERIODO, nREGISTROS (aggregate by period)
# ============================================================
def load_cargas(dh_pivot):
    df = _read_sql_df(QUERY_CARGAS, 'CARGAS')
    if df is None:
        path = DATA_DIR / "CARGAS.xlsx"
        df = pd.read_excel(path, engine='openpyxl')

    df['PERIODO'] = pd.to_numeric(df['PERIODO'], errors='coerce').fillna(0).astype(int)
    reg_col = _find_column(df, 'nREGISTROS', 'REGISTROS')
    if not reg_col:
        raise KeyError('No se encontro columna nREGISTROS/REGISTROS en CARGAS')
    df[reg_col] = pd.to_numeric(df[reg_col], errors='coerce').fillna(0).astype(int)
    agg = df.groupby('PERIODO')[reg_col].sum().reset_index()

    # Get gestiones from DH_PIVOT by period
    gest_by_per = {}
    for r in dh_pivot:
        p = r['p']
        gest_by_per[p] = gest_by_per.get(p, 0) + r['g']

    result = []
    for _, row in agg.iterrows():
        p = int(row['PERIODO'])
        if p > 0:
            result.append({
                'p': p,
                'reg': int(row[reg_col]),
                'gest': gest_by_per.get(p, 0),
            })

    return sorted(result, key=lambda x: x['p'])


# ============================================================
# 4. BITACORA.xlsx — DAILY data → aggregate monthly
#    Cols: FECHA, AVANCE EXTERNO Nro. Op., AVANCE EXTERNO Monto MM$,
#          SEGUROS DESGRAVAMEN AVANCE Nro. Op., ...% CRUCE SEG,
#          PAGO FLEXIBLE Nro. Op., PAGO FLEXIBLE Monto MM$,
#          PLAN ZERO Nro. Op., PER Nro. Op.,
#          TC TITULAR, TC ADICIONAL, TC DIGITAL
# ============================================================
def load_bitacora():
    df = _load_bitacora_source_df()

    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df = df.dropna(subset=['FECHA'])
    cutoff_date = pd.Timestamp.today().replace(day=1) - pd.DateOffset(months=12)
    df = df[df['FECHA'] >= cutoff_date]
    df['MES'] = df['FECHA'].dt.to_period('M')

    # Map columns by pattern matching
    col_av_o = next((c for c in df.columns if 'AVANCE EXTERNO' in c and 'Nro' in c), None)
    col_av_m = next((c for c in df.columns if 'AVANCE EXTERNO' in c and 'Monto' in c), None)
    col_seg_o = next((c for c in df.columns if 'SEGUROS' in c and 'Nro' in c), None)
    col_pf_o = next((c for c in df.columns if 'PAGO FLEXIBLE' in c and 'Nro' in c), None)
    col_pf_m = next((c for c in df.columns if 'PAGO FLEXIBLE' in c and 'Monto' in c), None)
    col_pz_o = next((c for c in df.columns if 'PLAN ZERO' in c and 'Nro' in c), None)
    col_per = next((c for c in df.columns if c.startswith('PER') and 'Nro' in c), None)
    col_tc_t = next((c for c in df.columns if 'TC' in c and 'TITULAR' in c), None)
    col_tc_a = next((c for c in df.columns if 'TC' in c and 'ADICIONAL' in c), None)
    col_tc_d = next((c for c in df.columns if 'TC' in c and 'DIGITAL' in c), None)

    # Fill NaN for numeric columns
    num_cols = [col_av_o, col_av_m, col_seg_o, col_pf_o, col_pf_m, col_pz_o, col_per, col_tc_t, col_tc_a, col_tc_d]
    for c in num_cols:
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    # Group by month
    result = []
    for mes, grp in df.groupby('MES'):
        dias = grp['FECHA'].dt.date.nunique()
        av_o = safe_int(grp[col_av_o].sum()) if col_av_o else 0
        av_m = safe_int(grp[col_av_m].sum()) if col_av_m else 0
        seg_o = safe_int(grp[col_seg_o].sum()) if col_seg_o else 0
        pf_o = safe_int(grp[col_pf_o].sum()) if col_pf_o else 0
        pf_m = safe_int(grp[col_pf_m].sum()) if col_pf_m else 0
        pz_o = safe_int(grp[col_pz_o].sum()) if col_pz_o else 0
        per = safe_int(grp[col_per].sum()) if col_per else 0
        tc_t = safe_int(grp[col_tc_t].sum()) if col_tc_t else 0
        tc_a = safe_int(grp[col_tc_a].sum()) if col_tc_a else 0
        tc_d = safe_int(grp[col_tc_d].sum()) if col_tc_d else 0

        cruce = round((seg_o / av_o * 100), 1) if av_o > 0 else 0.0

        result.append({
            'm': str(mes),
            'd': dias,
            'av_o': av_o, 'av_m': av_m, 'seg_o': seg_o,
            'pf_o': pf_o, 'pf_m': pf_m,
            'pz_o': pz_o, 'per': per,
            'tc_t': tc_t, 'tc_a': tc_a, 'tc_d': tc_d,
            'cruce': cruce,
        })

    return sorted(result, key=lambda x: x['m'])


def load_bitacora_daily():
    """Load daily bitacora rows for drill-down in correlación tab."""
    df = _load_bitacora_source_df()
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df = df.dropna(subset=['FECHA'])
    cutoff_date = pd.Timestamp.today().replace(day=1) - pd.DateOffset(months=12)
    df = df[df['FECHA'] >= cutoff_date]

    col_av_o = next((c for c in df.columns if 'AVANCE EXTERNO' in c and 'Nro' in c), None)
    col_av_m = next((c for c in df.columns if 'AVANCE EXTERNO' in c and 'Monto' in c), None)
    col_pf_o = next((c for c in df.columns if 'PAGO FLEXIBLE' in c and 'Nro' in c), None)
    col_pf_m = next((c for c in df.columns if 'PAGO FLEXIBLE' in c and 'Monto' in c), None)
    col_pz_o = next((c for c in df.columns if 'PLAN ZERO' in c and 'Nro' in c), None)
    col_per = next((c for c in df.columns if c.startswith('PER') and 'Nro' in c), None)
    col_tc_t = next((c for c in df.columns if 'TC' in c and 'TITULAR' in c), None)
    col_tc_a = next((c for c in df.columns if 'TC' in c and 'ADICIONAL' in c), None)
    col_tc_d = next((c for c in df.columns if 'TC' in c and 'DIGITAL' in c), None)
    col_seg_o = next((c for c in df.columns if 'SEGUROS' in c and 'Nro' in c), None)

    num_cols = [col_av_o, col_av_m, col_pf_o, col_pf_m, col_pz_o, col_per, col_tc_t, col_tc_a, col_tc_d, col_seg_o]
    for c in num_cols:
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    result = []
    for _, row in df.iterrows():
        fecha = row['FECHA'].strftime('%Y-%m-%d')
        mes = row['FECHA'].strftime('%Y-%m')
        av_o = safe_int(row[col_av_o]) if col_av_o else 0
        seg_o = safe_int(row[col_seg_o]) if col_seg_o else 0
        cruce = round((seg_o / av_o * 100), 1) if av_o > 0 else 0.0
        result.append({
            'f': fecha,
            'm': mes,
            'av_o': av_o,
            'av_m': safe_int(row[col_av_m]) if col_av_m else 0,
            'pf_o': safe_int(row[col_pf_o]) if col_pf_o else 0,
            'pf_m': safe_int(row[col_pf_m]) if col_pf_m else 0,
            'pz_o': safe_int(row[col_pz_o]) if col_pz_o else 0,
            'per': safe_int(row[col_per]) if col_per else 0,
            'tc_t': safe_int(row[col_tc_t]) if col_tc_t else 0,
            'tc_a': safe_int(row[col_tc_a]) if col_tc_a else 0,
            'tc_d': safe_int(row[col_tc_d]) if col_tc_d else 0,
            'seg_o': seg_o,
            'cruce': cruce,
        })

    return sorted(result, key=lambda x: x['f'])


# ============================================================
# 5. DETALLE_EJECUTIVO.xlsx (sheet 'base')
#    Cols: CODIGO_USUARIO, PERIODO, MONTO, Q, TIPO
# ============================================================
def load_ejecutivos():
    df = _read_sql_df(QUERY_DETALLE_EJECUTIVO, 'DETALLE_EJECUTIVO')
    if df is None:
        path = DATA_DIR / "DETALLE_EJECUTIVO.xlsx"
        df = pd.read_excel(path, engine='openpyxl', sheet_name='base')

    df['CODIGO_USUARIO'] = df['CODIGO_USUARIO'].astype(str).str.strip().str.upper()
    df['PERIODO'] = df['PERIODO'].astype(int).astype(str)
    df['MONTO'] = pd.to_numeric(df['MONTO'], errors='coerce').fillna(0).astype(int)
    df['Q'] = pd.to_numeric(df['Q'], errors='coerce').fillna(0).astype(int)
    df['TIPO'] = df['TIPO'].astype(str).str.strip().str.upper()
    df = df[~df['TIPO'].isin(['NAN', 'NAT', ''])]

    # Group by user + period + tipo → compromisos (Q) and monto
    grouped = df.groupby(['CODIGO_USUARIO', 'PERIODO', 'TIPO']).agg({
        'Q': 'sum',
        'MONTO': 'sum',
    }).reset_index()

    ej_monthly = []
    for _, row in grouped.iterrows():
        ej_monthly.append({
            'name': row['CODIGO_USUARIO'],
            'p': row['PERIODO'],
            't': row['TIPO'],
            'g': 0,  # enriched later from DATA_HISTORIA
            'e': 0,  # enriched later
            'c': int(row['Q']),
            'm': int(row['MONTO']),
        })

    # TOP10 by monto
    top_agg = df.groupby('CODIGO_USUARIO').agg({'Q': 'sum', 'MONTO': 'sum'}).reset_index()
    top_agg = top_agg.sort_values('MONTO', ascending=False).head(15)

    top10 = []
    for _, row in top_agg.iterrows():
        top10.append({
            'name': row['CODIGO_USUARIO'],
            'g': 0,
            't': 0,
            'm': int(row['MONTO']),
        })

    return ej_monthly, top10


def load_ej_tiempos(dh_df):
    """Generate headcount and time analysis from DATA_HISTORIA_DETALLE per user per month."""
    if dh_df is None:
        return [], []

    df = dh_df.copy()
    df['CODIGO_USUARIO'] = df['CODIGO_USUARIO'].astype(str).str.strip().str.upper()
    df['PERIODO'] = df['PERIODO'].astype(str)
    df['FECHA_STR'] = df['FECHA'].dt.strftime('%Y-%m-%d')
    turno_raw = pd.to_numeric(df['TURNO'], errors='coerce').fillna(0)
    turno_pos = turno_raw[turno_raw > 0]
    if turno_pos.empty:
        df['TURNO_H'] = turno_raw
    else:
        p95 = float(turno_pos.quantile(0.95))
        # SQL source now sends TURNO in hours; keep compatibility with older seconds/ms inputs.
        if p95 > 10000:
            df['TURNO_H'] = turno_raw / 3600000
        elif p95 > 100:
            df['TURNO_H'] = turno_raw / 3600
        else:
            df['TURNO_H'] = turno_raw
    df['HABLADO'] = pd.to_numeric(df['HABLADO'], errors='coerce').fillna(0)
    df['Disponible'] = pd.to_numeric(df['Disponible'], errors='coerce').fillna(0)
    df['TMO'] = pd.to_numeric(df['TMO'], errors='coerce').fillna(0)

    # --- Headcount by month (distinct users per day, averaged or total) ---
    hc_daily = df.groupby(['PERIODO', 'FECHA_STR'])['CODIGO_USUARIO'].nunique().reset_index()
    hc_daily.columns = ['p', 'd', 'hc']
    hc_monthly = hc_daily.groupby('p').agg(
        hc_avg=('hc', 'mean'),
        hc_max=('hc', 'max'),
        dias=('d', 'nunique')
    ).reset_index()

    ej_hc = []
    for _, row in hc_monthly.iterrows():
        ej_hc.append({
            'p': row['p'],
            'avg': round(row['hc_avg'], 1),
            'max': int(row['hc_max']),
            'dias': int(row['dias']),
        })
    # Also add daily headcount detail
    ej_hc_daily = hc_daily.to_dict('records')

    # --- Time analysis per user per month ---
    time_agg = df.groupby(['CODIGO_USUARIO', 'PERIODO']).agg(
        hablado=('HABLADO', 'sum'),
        disponible=('Disponible', 'sum'),
        tmo_avg=('TMO', 'mean'),
        turno_h=('TURNO_H', 'sum'),
        gestiones=('GESTIONES_Q', 'sum'),
        dias_trabajados=('FECHA_STR', 'nunique'),
    ).reset_index()

    ej_tiempos = []
    for _, row in time_agg.iterrows():
        ej_tiempos.append({
            'name': row['CODIGO_USUARIO'],
            'p': row['PERIODO'],
            'hab': round(row['hablado'], 2),
            'disp': round(row['disponible'], 2),
            'tmo': round(row['tmo_avg'], 1),
            'turno': round(row['turno_h'], 1),
            'gest': int(row['gestiones']),
            'dias': int(row['dias_trabajados']),
        })

    return ej_hc, ej_tiempos


def enrich_ejecutivos(ej_monthly, top10, dh_df):
    """Enrich with gestiones/efectivos from DATA_HISTORIA.
    
    DATA_HISTORIA TIPO is the source of truth for campaign classification.
    Detalle_Ejecutivo compromisos/monto are distributed proportionally
    across the TIPOs that DATA_HISTORIA assigns to each user+period.
    """
    if dh_df is None:
        return ej_monthly, top10

    dh_user = dh_df.groupby(['CODIGO_USUARIO', 'PERIODO', 'TIPO']).agg({
        'GESTIONES_Q': 'sum',
        'EFECTIVOS_Q': 'sum',
    }).reset_index()
    dh_user['CODIGO_USUARIO'] = dh_user['CODIGO_USUARIO'].str.strip().str.upper()
    dh_user['PERIODO'] = dh_user['PERIODO'].astype(str)

    # lookup: (user, period, tipo) -> (g, e)
    lookup = {}
    for _, row in dh_user.iterrows():
        key = (row['CODIGO_USUARIO'], row['PERIODO'], row['TIPO'])
        lookup[key] = (int(row['GESTIONES_Q']), int(row['EFECTIVOS_Q']))

    # Build DH types per user+period with gestiones for proportional split
    # dh_user_tipos[(user, period)] = {tipo: gestiones, ...}
    dh_user_tipos = {}
    for _, row in dh_user.iterrows():
        up_key = (row['CODIGO_USUARIO'], row['PERIODO'])
        if up_key not in dh_user_tipos:
            dh_user_tipos[up_key] = {}
        dh_user_tipos[up_key][row['TIPO']] = int(row['GESTIONES_Q'])

    # Aggregate Detalle entries by user+period (ignore Detalle's TIPO)
    detalle_agg = {}
    for entry in ej_monthly:
        up_key = (entry['name'], entry['p'])
        if up_key not in detalle_agg:
            detalle_agg[up_key] = {'c': 0, 'm': 0}
        detalle_agg[up_key]['c'] += entry['c']
        detalle_agg[up_key]['m'] += entry['m']

    # Rebuild ej_monthly from DATA_HISTORIA types with proportional split
    ej_monthly_new = []
    processed_ups = set()

    for up_key, tipos_gest in dh_user_tipos.items():
        processed_ups.add(up_key)
        total_gest = sum(tipos_gest.values())
        det = detalle_agg.get(up_key, {'c': 0, 'm': 0})

        for tipo, gest in tipos_gest.items():
            # Proportional split of compromisos/monto based on gestiones
            proportion = gest / total_gest if total_gest > 0 else 0
            c_split = round(det['c'] * proportion)
            m_split = round(det['m'] * proportion)
            g, e = lookup.get((up_key[0], up_key[1], tipo), (0, 0))

            ej_monthly_new.append({
                'name': up_key[0],
                'p': up_key[1],
                't': tipo,
                'g': g,
                'e': e,
                'c': c_split,
                'm': m_split,
            })

    # Keep Detalle entries that have no DH data (user+period not in DH)
    for entry in ej_monthly:
        up_key = (entry['name'], entry['p'])
        if up_key not in processed_ups:
            ej_monthly_new.append(entry)

    ej_monthly = ej_monthly_new

    # Top10 totals
    dh_top = dh_df.groupby('CODIGO_USUARIO').agg({
        'GESTIONES_Q': 'sum',
        'EFECTIVOS_Q': 'sum',
    }).reset_index()
    dh_top['CODIGO_USUARIO'] = dh_top['CODIGO_USUARIO'].str.strip().str.upper()

    for entry in top10:
        match = dh_top[dh_top['CODIGO_USUARIO'] == entry['name']]
        if not match.empty:
            g = int(match.iloc[0]['GESTIONES_Q'])
            e = int(match.iloc[0]['EFECTIVOS_Q'])
            entry['g'] = g
            entry['t'] = round((e / g * 100), 1) if g > 0 else 0

    return ej_monthly, top10


# ============================================================
# 6. FACTURACION_PROVISIONES.xlsx → CORR_DATA
#    Cols: OT, MES DEL SERVICIO, ANO DEL SERVICIO,
#          IMPTE PROVISION, IMPTE FACT. REAL, CLIENTE
# ============================================================
def load_facturacion(ot_map):
    df = _read_sql_df(QUERY_FACTURACION, 'FACTURACION_PROVISIONES')
    if df is None:
        path = DATA_DIR / "FACTURACION_PROVISIONES.xlsx"
        df = pd.read_excel(path, engine='openpyxl')
    df.columns = [str(c).replace('\xa0', ' ').strip() for c in df.columns]

    # Filter to Scotiabank
    if 'CLIENTE' in df.columns:
        df = df[df['CLIENTE'].astype(str).str.upper().str.contains('SCOTIABANK|SCOTIA', na=False)]

    prov_col = _find_column(df, 'IMPTE PROVISION', 'IMPTE_PROVISION')
    real_col = _find_column(df, 'IMPTE FACT. REAL', 'IMPTE_FACT_REAL')
    mes_col = _find_column(df, 'MES DEL SERVICIO', 'MES_DEL_SERVICIO')
    ano_col = _find_column(df, 'AÑO DEL SERVICIO', 'ANO DEL SERVICIO', 'ANO_DEL_SERVICIO', 'AÑO_DEL_SERVICIO')
    cartera_col = _find_column(df, 'CARTERA', 'NOMBRE CARTERA', 'CAMPAÑA', 'CAMPANA')
    if not prov_col or not real_col or not mes_col or not ano_col:
        print(f"  WARN: cols disponibles FACTURACION: {list(df.columns)}")
        return None

    df['OT'] = pd.to_numeric(df['OT'], errors='coerce').fillna(0).astype(int)
    df[prov_col] = pd.to_numeric(df[prov_col], errors='coerce').fillna(0)
    df[real_col] = pd.to_numeric(df[real_col], errors='coerce').fillna(0)

    def parse_mes(val):
        val = str(val).strip()
        m = re.match(r'^(\d+)', val)
        return int(m.group(1)) if m else 0

    df['MES_NUM'] = df[mes_col].apply(parse_mes)
    df['ANO_NUM'] = pd.to_numeric(df[ano_col], errors='coerce').fillna(0).astype(int)
    df['YYYY_MM'] = df.apply(
        lambda r: f"{int(r['ANO_NUM'])}-{int(r['MES_NUM']):02d}"
        if r['ANO_NUM'] > 0 and r['MES_NUM'] > 0 else '', axis=1
    )
    df = df[df['YYYY_MM'] != '']
    cutoff_ym = _rolling_12m_cutoff_ym()
    df = df[df['YYYY_MM'] >= cutoff_ym]

    # Filter to known OTs
    known_ots = {item['ot'] for item in ot_map}
    corr_df = df[df['OT'].isin(known_ots)]

    group_cols = ['OT', 'YYYY_MM']
    if cartera_col:
        group_cols.append(cartera_col)

    corr_grouped = corr_df.groupby(group_cols).agg({
        prov_col: 'sum',
        real_col: 'sum',
    }).reset_index()

    ot_to_prod = {item['ot']: item['prod'] for item in ot_map}

    corr_data = []
    for _, row in corr_grouped.iterrows():
        ot = int(row['OT'])
        cartera_name = str(row[cartera_col]).strip() if cartera_col and pd.notna(row[cartera_col]) else ''
        prod_name = ot_to_prod.get(ot, f'OT_{ot}')
        campaign_label = f"{cartera_name} | OT {ot}" if cartera_name else f"{prod_name} | OT {ot}"
        corr_data.append({
            'pr': prod_name,
            'camp': campaign_label,
            'ot': ot,
            'm': row['YYYY_MM'],
            'g': 0, 'e': 0, 'c': 0, 'md': 0, 'ej': 0,
            'pv': int(row[prov_col]),
            'fr': int(row[real_col]),
            'oc': 0, 'mc': 0,
        })

    return corr_data


def enrich_corr_data(corr_data, dh_df, bita_monthly):
    """Enrich CORR_DATA with DATA_HISTORIA and BITACORA data."""
    if not corr_data:
        return corr_data

    if dh_df is not None:
        dh_df['YYYY_MM'] = dh_df['PERIODO'].apply(
            lambda p: f"{str(p)[:4]}-{str(p)[4:6]}" if len(str(p)) == 6 else ''
        )
        dh_ot = dh_df.groupby(['OT', 'YYYY_MM']).agg({
            'GESTIONES_Q': 'sum',
            'EFECTIVOS_Q': 'sum',
            'COMPROMISOS_GENERADOS_Q': 'sum',
            'COMPROMISO_GENERADOS_MONTO': 'sum',
            'CODIGO_USUARIO': 'nunique',
        }).reset_index()

        for entry in corr_data:
            match = dh_ot[(dh_ot['OT'] == entry['ot']) & (dh_ot['YYYY_MM'] == entry['m'])]
            if not match.empty:
                row = match.iloc[0]
                entry['g'] = int(row['GESTIONES_Q'])
                entry['e'] = int(row['EFECTIVOS_Q'])
                entry['c'] = int(row['COMPROMISOS_GENERADOS_Q'])
                entry['md'] = int(row['COMPROMISO_GENERADOS_MONTO'])
                entry['ej'] = int(row['CODIGO_USUARIO'])

    if bita_monthly:
        bita_map = {b['m']: b for b in bita_monthly}
        for entry in corr_data:
            bm = bita_map.get(entry['m'])
            if bm:
                if entry['pr'] == 'AVANCE EXTERNO':
                    entry['oc'] = bm['av_o']
                    entry['mc'] = bm['av_m']
                elif entry['pr'] == 'PAGO FLEXIBLE':
                    entry['oc'] = bm['pf_o']
                    entry['mc'] = bm['pf_m']
                elif entry['pr'] == 'TC TITULAR':
                    entry['oc'] = bm['tc_t']
                    entry['mc'] = bm['tc_t']
                elif entry['pr'] == 'TC ADICIONAL':
                    entry['oc'] = bm['tc_a']
                    entry['mc'] = bm['tc_a']
                elif entry['pr'] == 'PLAN ZERO PER':
                    entry['oc'] = bm['pz_o']
                    entry['mc'] = bm.get('per', 0)

    return corr_data


# ============================================================
# 7b. UF DEL DÍA — Banco Central de Chile
# ============================================================
_BCCH_USER = os.getenv("BCCH_USER", "clorcat@fastcogroup.com")
_BCCH_PASS = os.getenv("BCCH_PASS", "Fastco2025")

def fetch_uf_bcch():
    """
    Fetch today's UF value from BCCh REST API.
    Falls back to mindicador.cl (public) if BCCh fails.
    Returns (uf_value: float, uf_date: str).
    """
    if not _HAS_REQUESTS:
        return None, None

    # Try BCCh REST API (look up to 7 days back in case today isn't published yet)
    base_url = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"
    series = "F073.UFF.PRE.Z.D"
    today = date.today()
    for delta in range(8):
        check_date = today - timedelta(days=delta)
        date_str = check_date.strftime("%Y-%m-%d")
        params = {
            "function": "GetSeries",
            "user": _BCCH_USER,
            "pass": _BCCH_PASS,
            "timeseries": series,
            "firstdate": date_str,
            "lastdate": date_str,
        }
        try:
            resp = requests.get(base_url, params=params, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                obs = data.get("Series", {}).get("Obs", [])
                if obs:
                    raw = str(obs[0].get("value", "")).strip()
                    try:
                        # BCCh returns standard decimal notation e.g. "40763.26"
                        uf_val = float(raw)
                    except ValueError:
                        # Fallback: Chilean thousands format "40.844,79"
                        uf_val = float(raw.replace(".", "").replace(",", "."))
                    if 10000 < uf_val < 200000:
                        print(f"  [UF] {uf_val:,.2f} ({date_str}) — BCCh API")
                        return uf_val, date_str
        except Exception:
            pass

    # Fallback: mindicador.cl (public API, no auth required)
    try:
        resp = requests.get("https://mindicador.cl/api/uf", timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            series_data = data.get("serie", [])
            if series_data:
                raw = str(series_data[0].get("valor", "")).replace(",", ".")
                uf_val = float(raw)
                uf_date = str(series_data[0].get("fecha", ""))[:10]
                print(f"  [UF] {uf_val:,.2f} ({uf_date}) — mindicador.cl (fallback)")
                return uf_val, uf_date
    except Exception:
        pass

    print("  [UF] No se pudo obtener valor de UF. Proyección de facturación deshabilitada.")
    return None, None


# ============================================================
# 7c. PROYECCIÓN FACTURACIÓN BASADA EN UF (PAGO FLEXIBLE + AVANCE)
# ============================================================
def load_bita_facturacion_uf(uf_value):
    """
    Calcula billing mensual real e histórico desde BITACORA.xlsx usando la UF del día.

    Fórmulas:
      PAGO FLEXIBLE : pf_qty × 0.315 × UF + pf_monto_CLP × 0.01
      AVANCE EXTERNO: av_qty × 0.32  × UF + av_monto_CLP × 0.0125

    Luego proyecta el mes en curso usando los mismos 3 métodos del
    modelo proyeccion_marzo_2026.py (ritmo actual + mes anterior + Prophet).
    """
    if uf_value is None:
        return None

    df = _load_bitacora_source_df()
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df = df.dropna(subset=['FECHA']).sort_values('FECHA').reset_index(drop=True)
    df['mes'] = df['FECHA'].dt.to_period('M')

    col_pf_o = next((c for c in df.columns if 'PAGO FLEXIBLE' in c and 'Nro' in c), None)
    col_pf_m = next((c for c in df.columns if 'PAGO FLEXIBLE' in c and 'Monto' in c), None)
    col_av_o = next((c for c in df.columns if 'AVANCE EXTERNO' in c and 'Nro' in c), None)
    col_av_m = next((c for c in df.columns if 'AVANCE EXTERNO' in c and 'Monto' in c), None)

    if not all([col_pf_o, col_pf_m, col_av_o, col_av_m]):
        print("  [FACT-UF] No se encontraron columnas PF/AV en BITACORA.xlsx")
        return None

    for c in [col_pf_o, col_pf_m, col_av_o, col_av_m]:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    # --- Billing diario ---
    df['fact_pf'] = df[col_pf_o] * 0.315 * uf_value + df[col_pf_m] * 0.01
    df['fact_av'] = df[col_av_o] * 0.32  * uf_value + df[col_av_m] * 0.0125
    df['fact_tot'] = df['fact_pf'] + df['fact_av']

    # --- Agrupado mensual ---
    monthly = df.groupby('mes').agg(
        dias=('FECHA', 'nunique'),
        pf_o=(col_pf_o, 'sum'),
        pf_m=(col_pf_m, 'sum'),
        av_o=(col_av_o, 'sum'),
        av_m=(col_av_m, 'sum'),
        fact_pf=('fact_pf', 'sum'),
        fact_av=('fact_av', 'sum'),
        fact_tot=('fact_tot', 'sum'),
    ).reset_index()
    monthly['mes_str'] = monthly['mes'].apply(lambda p: str(p))
    monthly['mes_dt']  = monthly['mes'].dt.to_timestamp()

    mes_actual   = df['mes'].max()
    mes_anterior = mes_actual - 1

    # Meses cerrados (todo menos el actual)
    cerrados = monthly[monthly['mes'] < mes_actual]
    dias_promedio_mes = cerrados['dias'].mean() if len(cerrados) > 0 else 20.0

    dias_registrados = int(monthly.loc[monthly['mes'] == mes_actual, 'dias'].values[0]) \
        if mes_actual in monthly['mes'].values else 0

    fraccion = dias_registrados / dias_promedio_mes if dias_promedio_mes > 0 else 0.0

    def _proyectar_col(col_name):
        """Run 3-method adaptive projection on a billing column."""
        real_parcial = float(monthly.loc[monthly['mes'] == mes_actual, col_name].values[0]) \
            if mes_actual in monthly['mes'].values else 0.0

        # Método 1: ritmo del mes actual
        datos_actual = df[df['mes'] == mes_actual]
        prom_diario_actual = datos_actual[col_name].mean() if len(datos_actual) > 0 else 0.0
        dias_base = max(dias_promedio_mes, dias_registrados)
        proy_ritmo = prom_diario_actual * dias_base

        # Método 2: ritmo del mes anterior
        datos_anterior = df[df['mes'] == mes_anterior]
        prom_diario_anterior = datos_anterior[col_name].mean() if len(datos_anterior) > 0 else 0.0
        proy_mes_anterior = prom_diario_anterior * dias_base

        # Método 3: Prophet estacional sobre histórico mensual cerrado
        ic_lower = ic_upper = proy_prophet = proy_ritmo  # defaults de seguridad
        if _HAS_PROPHET and len(cerrados) >= 2:
            historico = cerrados[['mes_dt', col_name]].copy()
            historico.columns = ['ds', 'y']
            historico = historico.dropna()
            try:
                model = Prophet(
                    yearly_seasonality=True,
                    weekly_seasonality=False,
                    daily_seasonality=False,
                    seasonality_mode='multiplicative',
                    changepoint_prior_scale=0.05,
                    interval_width=0.90,
                )
                model.fit(historico)
                future   = model.make_future_dataframe(periods=1, freq='MS')
                forecast = model.predict(future)
                fila     = forecast[forecast['ds'] == mes_actual.to_timestamp()]
                if not fila.empty:
                    proy_prophet = float(fila['yhat'].values[0])
                    ic_lower     = float(fila['yhat_lower'].values[0])
                    ic_upper     = float(fila['yhat_upper'].values[0])
                    if proy_ritmo > 0:
                        proy_prophet = min(proy_prophet, proy_ritmo * 1.25)
            except Exception as e:
                print(f"    [Prophet] Error en {col_name}: {e}")
        elif not _HAS_PROPHET:
            # Sin Prophet instalado, se mantiene proyección con 2 métodos.
            pass

        # Pesos adaptativos (igual que proyeccion_marzo_2026.py)
        peso_prophet      = max(0.0, 0.15 * (1 - fraccion / 0.70)) if fraccion < 0.70 else 0.0
        peso_ritmo_actual = min(0.85, fraccion * 1.20)
        peso_mes_anterior = max(0.05, 1 - peso_ritmo_actual - peso_prophet)
        total_peso = peso_ritmo_actual + peso_mes_anterior + peso_prophet
        if total_peso > 0:
            peso_ritmo_actual /= total_peso
            peso_mes_anterior /= total_peso
            peso_prophet      /= total_peso

        proy_final = (
            peso_ritmo_actual * proy_ritmo +
            peso_mes_anterior * proy_mes_anterior +
            peso_prophet      * proy_prophet
        )
        proy_final = max(proy_final, real_parcial)

        return {
            'real'    : round(real_parcial),
            'proy'    : round(proy_final),
            'ic_low'  : round(max(ic_lower, real_parcial)),
            'ic_high' : round(ic_upper),
            'pesos'   : (round(peso_ritmo_actual, 3), round(peso_mes_anterior, 3), round(peso_prophet, 3)),
        }

    pf_res  = _proyectar_col('fact_pf')
    av_res  = _proyectar_col('fact_av')
    tot_res = _proyectar_col('fact_tot')

    # --- Histórico mensual para el gráfico ---
    hist = []
    for _, row in monthly.iterrows():
        is_current = (row['mes'] == mes_actual)
        hist.append({
            'm'       : row['mes_str'],
            'pf'      : round(row['fact_pf']),
            'av'      : round(row['fact_av']),
            'tot'     : round(row['fact_tot']),
            'current' : is_current,
        })

    return {
        'uf'          : round(uf_value, 2),
        'fraccion'    : round(fraccion, 4),
        'dias_reg'    : dias_registrados,
        'dias_prom'   : round(dias_promedio_mes, 1),
        'mes_actual'  : str(mes_actual),
        'pesos'       : pf_res['pesos'],
        'pf'          : pf_res,
        'av'          : av_res,
        'tot'         : tot_res,
        'hist'        : hist,
    }


# ============================================================
# 7d. PROYECCIÓN FACTURACIÓN from BITACORA + MODELO_FACTURACION
# ============================================================
def load_facturacion_proyeccion():
    """Calculate billing projection per product per month from BITACORA history and tariff model."""
    tariffs = FACT_PROY_TARIFFS
    backoffice = FACT_PROY_BACKOFFICE
    web_extra = FACT_PROY_WEB_EXTRA

    # --- Load Bitacora daily ---
    bdf = _load_bitacora_source_df()
    bdf['FECHA'] = pd.to_datetime(bdf['FECHA'], errors='coerce')
    bdf = bdf.dropna(subset=['FECHA'])
    bdf['YYYY_MM'] = bdf['FECHA'].dt.strftime('%Y-%m')

    # Map Bitacora columns to product names
    col_map = {
        'AVANCE EXTERNO Nro. Op.': 'AVANCE EXTERNO',
        'SEGUROS DESGRAVAMEN AVANCE Nro. Op.': 'SEG. DESGRAVAMEN',
        'PAGO FLEXIBLE Nro. Op.': 'PAGO FLEXIBLE',
        'PLAN ZERO Nro. Op.': 'PLAN ZERO OPERA',
        'PLAN ZERO PER': 'PLAN ZERO PER',
        'PER Nro. Op.': 'PER',
    }
    # TC columns (may have \xa0)
    for c in bdf.columns:
        cn = c.strip().replace('\xa0', ' ').replace('  ', ' ')
        if 'TC' in cn and 'TITULAR' in cn:
            col_map[c] = 'TC TITULAR'
        elif 'TC' in cn and 'ADICIONAL' in cn:
            col_map[c] = 'TC ADICIONAL'
        elif 'TC' in cn and 'DIGITAL' in cn:
            col_map[c] = 'TC DIGITAL'
        elif 'TC' in cn and 'TSYS' in cn:
            col_map[c] = 'TC TSYS'

    # Aggregate by month
    months = sorted(bdf['YYYY_MM'].unique())
    result = []
    for m in months:
        mdata = bdf[bdf['YYYY_MM'] == m]
        dias_trabajados = len(mdata)
        month_entry = {'m': m, 'dias': dias_trabajados, 'prods': [], 'fijos': backoffice + web_extra}
        total_fact = 0
        for col, prod in col_map.items():
            if col not in bdf.columns:
                continue
            ops = int(pd.to_numeric(mdata[col], errors='coerce').fillna(0).sum())
            tarifa = tariffs.get(prod, 0)
            fact = ops * tarifa
            total_fact += fact
            month_entry['prods'].append({
                'prod': prod,
                'ops': ops,
                'fact': round(fact),
            })
        month_entry['subtotal'] = round(total_fact)
        month_entry['total'] = round(total_fact + backoffice + web_extra)
        result.append(month_entry)

    return result


# ============================================================
# 8. Generate JavaScript data block
# ============================================================
def generate_js_data(dh_pivot, periodos, per_labels, cargas, bita, ej_monthly, top10, ot_map, corr_data, bita_daily=None, ej_hc=None, ej_tiempos=None, fact_proy=None, bita_fact_uf=None, dh_daily=None, cal_ciclo=None, cal_prod=None, cal_usr=None, cal_productos=None, cal_ciclos=None):
    generated_at = datetime.now()
    generated_meta = {
        'generated_at_iso': generated_at.isoformat(timespec='seconds'),
        'generated_at': generated_at.strftime('%Y-%m-%d %H:%M:%S'),
    }
    lines = []
    lines.append("// ===================== DATA (auto-generated) =====================")
    lines.append(f"const DASHBOARD_META = {json.dumps(generated_meta)};")
    lines.append(f"const PERIODOS_DH = {json.dumps(periodos)};")
    lines.append(f"const PER_LABELS = {json.dumps(per_labels)};")
    lines.append("const TIPO_COLORS = {REFINANCIAMIENTO:'#3d7cf4',AVANCE:'#00c4b4',TDC:'#e8382a',CONSUMO:'#f5a623'};")
    lines.append(f"\nconst DH_PIVOT = {json.dumps(dh_pivot)};")
    if dh_daily:
        lines.append(f"\nconst DH_DAILY = {json.dumps(dh_daily)};")
    lines.append(f"\nconst CARGAS_RAW = {json.dumps(cargas)};")
    lines.append(f"\nconst BITA_MONTHLY = {json.dumps(bita)};")
    lines.append("const BITA_MONTHS_ALL = BITA_MONTHLY.map(b=>b.m);")
    if bita_daily:
        lines.append(f"\nconst BITA_DAILY = {json.dumps(bita_daily)};")
    lines.append(f"\nconst TOP10 = {json.dumps(top10)};")
    lines.append(f"\nconst EJ_MONTHLY = {json.dumps(ej_monthly)};")
    ej_tipos = sorted(set(e['t'] for e in ej_monthly if e.get('t') and e['t'] != 'CONSUMO'))
    lines.append(f"const EJ_TIPOS = {json.dumps(ej_tipos)};")
    if ej_hc:
        lines.append(f"\nconst EJ_HC = {json.dumps(ej_hc)};")
    if ej_tiempos:
        lines.append(f"\nconst EJ_TIEMPOS = {json.dumps(ej_tiempos)};")
    lines.append(f"\nconst OT_PROD_MAP = {json.dumps(ot_map)};")
    if corr_data:
        lines.append(f"\nconst CORR_DATA = {json.dumps(corr_data)};")
        lines.append("const CORR_MONTHS = [...new Set(CORR_DATA.map(d=>d.m))].sort();")
        lines.append("const CORR_PRODS = OT_PROD_MAP.map(p=>p.prod);")
        lines.append("let activeCorrProd = 'ALL';")
    if fact_proy:
        lines.append(f"\nconst FACT_PROY = {json.dumps(fact_proy)};")
    if bita_fact_uf:
        lines.append(f"\nconst BITA_FACT_UF = {json.dumps(bita_fact_uf)};")
    if cal_ciclo is not None:
        lines.append(f"\nconst CAL_CICLO = {json.dumps(cal_ciclo)};")
        lines.append(f"const CAL_PROD = {json.dumps(cal_prod)};")
        lines.append(f"const CAL_USR = {json.dumps(cal_usr)};")
        lines.append(f"const CAL_PRODUCTOS = {json.dumps(cal_productos)};")
        lines.append(f"const CAL_CICLOS = {json.dumps(cal_ciclos)};")
    lines.append("")
    return '\n'.join(lines)


# ============================================================
# 9. CALIDAD.xlsx
# ============================================================
def load_calidad():
    df = _read_sql_df(QUERY_CALIDAD, 'CALIDAD')
    if df is None:
        path = DATA_DIR / 'CALIDAD.xlsx'
        if not path.exists():
            return None, None, None, None, None
        df = pd.read_excel(path)
    df['FECHA_GESTION'] = pd.to_datetime(df['FECHA_GESTION'], errors='coerce')
    df = df[df['FECHA_GESTION'].notna()].copy()
    df['p'] = df['FECHA_GESTION'].dt.strftime('%Y%m')

    # Simplify product name: strip SCOTIABANK_ prefix for display
    df['prod_short'] = df['PRODUCTO'].str.replace(r'^SCOTIABANK_', '', regex=True)

    # 1. Per-period per-ciclo per-product aggregate (allows product filter in ciclo chart)
    cal_ciclo = (
        df.groupby(['p', 'CICLO', 'prod_short'])['NOTA']
        .agg(n='mean', q='count')
        .reset_index()
        .rename(columns={'CICLO': 'c', 'prod_short': 'prod'})
    )
    cal_ciclo['n'] = cal_ciclo['n'].round(4)
    cal_ciclo_list = cal_ciclo.to_dict('records')

    # 2. Per-period per-product aggregate
    cal_prod = (
        df.groupby(['p', 'prod_short'])['NOTA']
        .agg(n='mean', q='count')
        .reset_index()
        .rename(columns={'prod_short': 'prod'})
    )
    cal_prod['n'] = cal_prod['n'].round(4)
    cal_prod_list = cal_prod.to_dict('records')

    # 3. Per-user per-period per-product aggregate (for ranking with campaign filter)
    cal_usr = (
        df.groupby(['USUARIO_AGENTE', 'p', 'prod_short'])['NOTA']
        .agg(n='mean', q='count')
        .reset_index()
        .rename(columns={'USUARIO_AGENTE': 'u', 'prod_short': 'prod'})
    )
    cal_usr['n'] = cal_usr['n'].round(4)
    cal_usr_list = cal_usr.to_dict('records')

    productos = sorted(df['prod_short'].unique().tolist())
    ciclos = sorted(df['CICLO'].unique().tolist())
    return cal_ciclo_list, cal_prod_list, cal_usr_list, productos, ciclos


# ============================================================
# MAIN
# ============================================================
def main():
    print("=" * 60)
    print("FASTCO x Scotiabank — Generador de Dashboard")
    print("=" * 60)

    # Ensure template exists
    if not TEMPLATE_PATH.exists():
        print("\nCreando template desde HTML existente...")
        create_template_from_html()
        print("Template creado.")

    # Load all sources
    print("\n[1/6] RELACION_BITACORA_OT (embebido)...")
    ot_map = load_ot_map()
    print(f"  OK: {len(ot_map)} productos")

    print("[2/6] DATA_HISTORIA_DETALLE...")
    dh_pivot, periodos, per_labels, dh_df, dh_daily = load_data_historia()
    print(f"  OK: {len(dh_pivot)} registros, periodos: {periodos}")

    print("[3/6] CARGAS...")
    cargas = load_cargas(dh_pivot)
    print(f"  OK: {len(cargas)} periodos de cargas")

    print("[4/6] BITACORA...")
    bita = load_bitacora()
    bita_daily = load_bitacora_daily()
    print(f"  OK: {len(bita)} meses, {len(bita_daily)} días")

    print("[5/6] DETALLE_EJECUTIVO...")
    ej_monthly, top10 = load_ejecutivos()
    ej_monthly, top10 = enrich_ejecutivos(ej_monthly, top10, dh_df)
    ej_hc, ej_tiempos = load_ej_tiempos(dh_df)
    print(f"  OK: {len(ej_monthly)} registros ejecutivos, Top {len(top10)} (monto), HC {len(ej_hc)} meses, Tiempos {len(ej_tiempos)}")

    print("[6/6] FACTURACION_PROVISIONES...")
    corr_data = load_facturacion(ot_map)
    if corr_data:
        corr_data = enrich_corr_data(corr_data, dh_df, bita)
        print(f"  OK: {len(corr_data)} registros correlacion")
    else:
        print("  WARN: usando fallback del template")

    # Billing projection
    fact_proy = load_facturacion_proyeccion()
    if fact_proy:
        print(f"  Proyección: {len(fact_proy)} meses calculados")

    print("[7/8] CALIDAD...")
    cal_result = load_calidad()
    if cal_result[0] is not None:
        cal_ciclo_d, cal_prod_d, cal_usr_d, cal_prods, cal_cicls = cal_result
        print(f"  OK: {len(cal_ciclo_d)} registros ciclo, {len(cal_prod_d)} prod, {len(cal_usr_d)} usr")
    else:
        cal_ciclo_d = cal_prod_d = cal_usr_d = cal_prods = cal_cicls = None
        print("  SKIP: CALIDAD no disponible")

    # UF del día + proyección facturación UF
    print("\n[8/8] Proyección Facturación UF (BCCh)...")
    uf_value, uf_date = fetch_uf_bcch()
    bita_fact_uf = load_bita_facturacion_uf(uf_value)
    if bita_fact_uf:
        pesos = bita_fact_uf['pesos']
        print(f"  UF del día: {bita_fact_uf['uf']:,.2f} ({uf_date})")
        print(f"  Mes actual: {bita_fact_uf['mes_actual']} — avance {bita_fact_uf['fraccion']:.1%} ({bita_fact_uf['dias_reg']} días)")
        print(f"  Pesos - Ritmo: {pesos[0]:.0%} | Mes ant: {pesos[1]:.0%} | Prophet: {pesos[2]:.0%}")
        print(f"  PF real: ${bita_fact_uf['pf']['real']:,.0f}  - proy: ${bita_fact_uf['pf']['proy']:,.0f}")
        print(f"  AV real: ${bita_fact_uf['av']['real']:,.0f}  - proy: ${bita_fact_uf['av']['proy']:,.0f}")
    else:
        print("  WARN: Proyección UF no disponible")

    # Generate
    print("\n[GEN] Generando dashboard...")
    template = TEMPLATE_PATH.read_text(encoding='utf-8')

    js_data = generate_js_data(dh_pivot, periodos, per_labels, cargas, bita,
                               ej_monthly, top10, ot_map, corr_data, bita_daily,
                               ej_hc, ej_tiempos, fact_proy, bita_fact_uf, dh_daily,
                               cal_ciclo_d, cal_prod_d, cal_usr_d, cal_prods, cal_cicls)

    # Inject: replace the data section between markers
    marker_start = "// ===================== DATA"
    marker_end = "// ===================== UTILITY ====================="

    if marker_start in template and marker_end in template:
        start_idx = template.index(marker_start)
        end_idx = template.index(marker_end)
        output = template[:start_idx] + js_data + '\n\n' + template[end_idx:]
    elif '/* __DATA_PLACEHOLDER__ */' in template:
        # Fallback: replace placeholder
        output = template.replace('/* __DATA_PLACEHOLDER__ */', js_data)
    else:
        print("  ERROR: No se encontro marcador en template. Usando template tal cual.")
        output = template

    OUTPUT_PATH.write_text(output, encoding='utf-8')
    size_kb = OUTPUT_PATH.stat().st_size / 1024
    print(f"\n{'='*60}")
    print(f"  DASHBOARD GENERADO EXITOSAMENTE")
    print(f"  Archivo: {OUTPUT_PATH.name}")
    print(f"  Tamano:  {size_kb:.1f} KB")
    print(f"  Datos:   {len(periodos)} periodos, {len(bita)} meses bitacora")
    print(f"           {len(ej_monthly)} reg. ejecutivos, {len(corr_data) if corr_data else 0} corr.")
    print(f"{'='*60}")
    print(f"\n  Abre '{OUTPUT_PATH.name}' con doble-click para ver el dashboard.\n")


def create_template_from_html():
    """Create template.html from current dashboard HTML."""
    content = OUTPUT_PATH.read_text(encoding='utf-8')
    # The template is the full HTML — the data section will be replaced on each run
    TEMPLATE_PATH.write_text(content, encoding='utf-8')


if __name__ == '__main__':
    main()
