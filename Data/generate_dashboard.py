"""
FASTCO x Scotiabank — Generador automatico de Dashboard
=========================================================
Lee los archivos Excel de la carpeta Data/ y regenera el HTML
del dashboard con datos actualizados.

USO:
    python Data/generate_dashboard.py

REQUISITOS:
    pip install openpyxl pandas

RESULTADO:
    Regenera dashboard_scotiabank_fastco.html con datos frescos.
"""

import json
import os
import re
import sys
from pathlib import Path

try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("ERROR: Requiere pandas. Instala con: pip install pandas openpyxl")
    sys.exit(1)

# Paths
SCRIPT_DIR = Path(__file__).parent
PROJECT_DIR = SCRIPT_DIR.parent
DATA_DIR = SCRIPT_DIR
TEMPLATE_PATH = DATA_DIR / "template.html"
OUTPUT_PATH = PROJECT_DIR / "dashboard_scotiabank_fastco.html"


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
    path = DATA_DIR / "RELACION_BITACORA_OT.txt"
    df = pd.read_csv(path, sep='\t')
    colors = {
        'AVANCE EXTERNO': '#00c4b4',
        'PAGO FLEXIBLE': '#f5a623',
        'TC TITULAR': '#3d7cf4',
        'TC ADICIONAL': '#e8382a',
        'PLAN ZERO PER': '#27c47a',
    }
    result = []
    for _, row in df.iterrows():
        prod = row['Producto'].strip()
        ot = int(row['OT'])
        result.append({'prod': prod, 'ot': ot, 'color': colors.get(prod, '#8892b0')})
    return result


# ============================================================
# 2. DATA_HISTORIA_DETALLE.xlsx
#    Cols: OT, PERIODO, FECHA, CODIGO_USUARIO, NOMBRE_EJECUTIVO,
#          GESTIONES_Q, EFECTIVOS_Q, COMPROMISOS_GENERADOS_Q,
#          COMPROMISO_GENERADOS_MONTO, TIPO
# ============================================================
def load_data_historia():
    path = DATA_DIR / "DATA_HISTORIA_DETALLE.xlsx"
    df = pd.read_excel(path, engine='openpyxl')

    df['PERIODO'] = df['PERIODO'].astype(int)
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

    return dh_pivot, periodos, per_labels, df


# ============================================================
# 3. CARGAS.xlsx
#    Cols: PERIODO, nREGISTROS (aggregate by period)
# ============================================================
def load_cargas(dh_pivot):
    path = DATA_DIR / "CARGAS.xlsx"
    df = pd.read_excel(path, engine='openpyxl')

    df['PERIODO'] = pd.to_numeric(df['PERIODO'], errors='coerce').fillna(0).astype(int)
    df['nREGISTROS'] = pd.to_numeric(df['nREGISTROS'], errors='coerce').fillna(0).astype(int)
    agg = df.groupby('PERIODO')['nREGISTROS'].sum().reset_index()

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
                'reg': int(row['nREGISTROS']),
                'gest': gest_by_per.get(p, 0),
            })

    return sorted(result, key=lambda x: x['p'])


# ============================================================
# 4. Bitacora.xlsx — DAILY data → aggregate monthly
#    Cols: FECHA, AVANCE EXTERNO Nro. Op., AVANCE EXTERNO Monto MM$,
#          SEGUROS DESGRAVAMEN AVANCE Nro. Op., ...% CRUCE SEG,
#          PAGO FLEXIBLE Nro. Op., PAGO FLEXIBLE Monto MM$,
#          PLAN ZERO Nro. Op., PER Nro. Op.,
#          TC TITULAR, TC ADICIONAL, TC DIGITAL
# ============================================================
def load_bitacora():
    path = DATA_DIR / "Bitacora.xlsx"
    df = pd.read_excel(path, engine='openpyxl')

    # Clean column names (remove non-breaking spaces)
    df.columns = [c.replace('\xa0', ' ').strip() for c in df.columns]

    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df = df.dropna(subset=['FECHA'])
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
    path = DATA_DIR / "Bitacora.xlsx"
    df = pd.read_excel(path, engine='openpyxl')

    df.columns = [c.replace('\xa0', ' ').strip() for c in df.columns]
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df = df.dropna(subset=['FECHA'])

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
# 5. Detalle_Ejecutivo_2026.xlsx (sheet 'base')
#    Cols: CODIGO_USUARIO, PERIODO, MONTO, Q, TIPO
# ============================================================
def load_ejecutivos():
    path = DATA_DIR / "Detalle_Ejecutivo_2026.xlsx"
    df = pd.read_excel(path, engine='openpyxl', sheet_name='base')

    df['CODIGO_USUARIO'] = df['CODIGO_USUARIO'].astype(str).str.strip().str.upper()
    df['PERIODO'] = df['PERIODO'].astype(int).astype(str)
    df['MONTO'] = pd.to_numeric(df['MONTO'], errors='coerce').fillna(0).astype(int)
    df['Q'] = pd.to_numeric(df['Q'], errors='coerce').fillna(0).astype(int)

    # Group by user + period → compromisos (Q) and monto
    grouped = df.groupby(['CODIGO_USUARIO', 'PERIODO']).agg({
        'Q': 'sum',
        'MONTO': 'sum',
    }).reset_index()

    ej_monthly = []
    for _, row in grouped.iterrows():
        ej_monthly.append({
            'name': row['CODIGO_USUARIO'],
            'p': row['PERIODO'],
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
    # TURNO: convert from ms to hours
    df['TURNO_H'] = pd.to_numeric(df['TURNO'], errors='coerce').fillna(0) / 3600000
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
    """Enrich with gestiones/efectivos from DATA_HISTORIA."""
    if dh_df is None:
        return ej_monthly, top10

    dh_user = dh_df.groupby(['CODIGO_USUARIO', 'PERIODO']).agg({
        'GESTIONES_Q': 'sum',
        'EFECTIVOS_Q': 'sum',
    }).reset_index()
    dh_user['CODIGO_USUARIO'] = dh_user['CODIGO_USUARIO'].str.strip().str.upper()
    dh_user['PERIODO'] = dh_user['PERIODO'].astype(str)

    lookup = {}
    for _, row in dh_user.iterrows():
        key = (row['CODIGO_USUARIO'], row['PERIODO'])
        lookup[key] = (int(row['GESTIONES_Q']), int(row['EFECTIVOS_Q']))

    # Existing entries in EJ_MONTHLY
    existing_keys = set((e['name'], e['p']) for e in ej_monthly)

    for entry in ej_monthly:
        key = (entry['name'], entry['p'])
        if key in lookup:
            entry['g'] = lookup[key][0]
            entry['e'] = lookup[key][1]

    # Add entries from DATA_HISTORIA that don't exist in Detalle_Ejecutivo
    # (periods before 2026 have gestiones but no compromisos/monto from that source)
    for key, (g, e) in lookup.items():
        if key not in existing_keys:
            ej_monthly.append({
                'name': key[0],
                'p': key[1],
                'g': g,
                'e': e,
                'c': 0,
                'm': 0,
            })

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
    path = DATA_DIR / "FACTURACION_PROVISIONES.xlsx"
    df = pd.read_excel(path, engine='openpyxl')

    # Filter to Scotiabank
    if 'CLIENTE' in df.columns:
        df = df[df['CLIENTE'].astype(str).str.upper().str.contains('SCOTIABANK|SCOTIA', na=False)]

    df['OT'] = pd.to_numeric(df['OT'], errors='coerce').fillna(0).astype(int)
    df['IMPTE PROVISION'] = pd.to_numeric(df['IMPTE PROVISION'], errors='coerce').fillna(0)
    df['IMPTE FACT. REAL'] = pd.to_numeric(df['IMPTE FACT. REAL'], errors='coerce').fillna(0)

    # Parse MES DEL SERVICIO (format: "03) MARZO" or "01) ENERO")
    mes_col = next((c for c in df.columns if 'MES DEL SERVICIO' in c.upper() or 'MES' == c.strip().upper()), None)
    ano_col = next((c for c in df.columns if 'AÑO DEL SERVICIO' in c or 'ANO DEL SERVICIO' in c or 'AÑO' in c.upper()), None)

    if not mes_col:
        mes_col = next((c for c in df.columns if 'MES' in c.upper()), None)
    if not ano_col:
        ano_col = next((c for c in df.columns if 'AÑO' in c.upper() or 'ANO' in c.upper()), None)

    if not mes_col or not ano_col:
        print(f"  WARN: cols disponibles: {list(df.columns)}")
        return None

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

    # Filter to known OTs
    known_ots = {item['ot'] for item in ot_map}
    corr_df = df[df['OT'].isin(known_ots)]

    corr_grouped = corr_df.groupby(['OT', 'YYYY_MM']).agg({
        'IMPTE PROVISION': 'sum',
        'IMPTE FACT. REAL': 'sum',
    }).reset_index()

    ot_to_prod = {item['ot']: item['prod'] for item in ot_map}

    corr_data = []
    for _, row in corr_grouped.iterrows():
        ot = int(row['OT'])
        corr_data.append({
            'pr': ot_to_prod.get(ot, f'OT_{ot}'),
            'ot': ot,
            'm': row['YYYY_MM'],
            'g': 0, 'e': 0, 'c': 0, 'md': 0, 'ej': 0,
            'pv': int(row['IMPTE PROVISION']),
            'fr': int(row['IMPTE FACT. REAL']),
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
# 7b. PROYECCIÓN FACTURACIÓN from Bitacora + Modelo_Facturacion
# ============================================================
def load_facturacion_proyeccion():
    """Calculate billing projection per product per month from Bitacora history and tariff model."""
    bita_path = DATA_DIR / "Bitacora.xlsx"
    modelo_path = DATA_DIR / "Modelo_Facturacion.xlsx"
    if not bita_path.exists() or not modelo_path.exists():
        return []

    # --- Load tariffs from Modelo ---
    mdf = pd.read_excel(modelo_path, sheet_name='Bitacora', header=None)
    tariffs = {}
    for i in range(4, 14):
        prod = mdf.iloc[i, 4]
        tarifa_civa = mdf.iloc[i, 9]
        if pd.notna(prod) and str(prod).strip():
            tariffs[str(prod).strip()] = float(tarifa_civa) if pd.notna(tarifa_civa) else 0
    backoffice = float(mdf.iloc[9, 1]) if pd.notna(mdf.iloc[9, 1]) else 8000000
    web_extra = float(mdf.iloc[10, 1]) if pd.notna(mdf.iloc[10, 1]) else 10000000

    # --- Load Bitacora daily ---
    bdf = pd.read_excel(bita_path, engine='openpyxl')
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
def generate_js_data(dh_pivot, periodos, per_labels, cargas, bita, ej_monthly, top10, ot_map, corr_data, bita_daily=None, ej_hc=None, ej_tiempos=None, fact_proy=None):
    lines = []
    lines.append("// ===================== DATA (auto-generated) =====================")
    lines.append(f"const PERIODOS_DH = {json.dumps(periodos)};")
    lines.append(f"const PER_LABELS = {json.dumps(per_labels)};")
    lines.append("const TIPO_COLORS = {REFINANCIAMIENTO:'#3d7cf4',AVANCE:'#00c4b4',TDC:'#e8382a',CONSUMO:'#f5a623'};")
    lines.append(f"\nconst DH_PIVOT = {json.dumps(dh_pivot)};")
    lines.append(f"\nconst CARGAS_RAW = {json.dumps(cargas)};")
    lines.append(f"\nconst BITA_MONTHLY = {json.dumps(bita)};")
    lines.append("const BITA_MONTHS_ALL = BITA_MONTHLY.map(b=>b.m);")
    if bita_daily:
        lines.append(f"\nconst BITA_DAILY = {json.dumps(bita_daily)};")
    lines.append(f"\nconst TOP10 = {json.dumps(top10)};")
    lines.append(f"\nconst EJ_MONTHLY = {json.dumps(ej_monthly)};")
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
    lines.append("")
    return '\n'.join(lines)


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
    print("\n[1/6] RELACION_BITACORA_OT.txt...")
    ot_map = load_ot_map()
    print(f"  OK: {len(ot_map)} productos")

    print("[2/6] DATA_HISTORIA_DETALLE.xlsx...")
    dh_pivot, periodos, per_labels, dh_df = load_data_historia()
    print(f"  OK: {len(dh_pivot)} registros, periodos: {periodos}")

    print("[3/6] CARGAS.xlsx...")
    cargas = load_cargas(dh_pivot)
    print(f"  OK: {len(cargas)} periodos de cargas")

    print("[4/6] Bitacora.xlsx...")
    bita = load_bitacora()
    bita_daily = load_bitacora_daily()
    print(f"  OK: {len(bita)} meses, {len(bita_daily)} días")

    print("[5/6] Detalle_Ejecutivo_2026.xlsx...")
    ej_monthly, top10 = load_ejecutivos()
    ej_monthly, top10 = enrich_ejecutivos(ej_monthly, top10, dh_df)
    ej_hc, ej_tiempos = load_ej_tiempos(dh_df)
    print(f"  OK: {len(ej_monthly)} registros ejecutivos, Top {len(top10)}, HC {len(ej_hc)} meses, Tiempos {len(ej_tiempos)}")

    print("[6/6] FACTURACION_PROVISIONES.xlsx...")
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

    # Generate
    print("\n[GEN] Generando dashboard...")
    template = TEMPLATE_PATH.read_text(encoding='utf-8')

    js_data = generate_js_data(dh_pivot, periodos, per_labels, cargas, bita,
                               ej_monthly, top10, ot_map, corr_data, bita_daily,
                               ej_hc, ej_tiempos, fact_proy)

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
