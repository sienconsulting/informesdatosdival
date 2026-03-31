"""
Expanded Data Processing Script for Valencia Province Tourism Report 2025
Processes all data sources and extracts comprehensive metrics for the PPTX report
"""
import os
import json
import sys
import glob
import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ---- Find data files ----
data_files = {}
for root, dirs, files in os.walk(os.path.expanduser('~')):
    if 'Diputaci' in root and 'Informe datos' in root and 'Datos' in root:
        for f in files:
            data_files[f] = os.path.join(root, f)

# Also find VUT files by path pattern
vut_files = [v for k, v in data_files.items() if 'obtener' in k.lower()]

print(f"Found {len(data_files)} data files")
print(f"Found {len(vut_files)} VUT files")

# ---- Helper functions ----
def fmt_num(n, decimals=0):
    if pd.isna(n):
        return "N/D"
    if decimals == 0:
        return f"{int(round(n)):,}".replace(",", ".")
    return f"{n:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def pct_change(new, old):
    if pd.isna(old) or old == 0:
        return None
    return (new - old) / abs(old) * 100

def fmt_pct(p):
    if p is None or pd.isna(p):
        return "N/D"
    sign = "+" if p > 0 else ""
    return f"{sign}{p:.1f}%"

# ---- SECTION 1: MAESTRO MUNICIPIOS - build lookup ----
df_mun = pd.read_excel(data_files['Maestro_Municipios.xlsx'])
df_com = pd.read_excel(data_files['Maestro_Comarcas.xlsx'])

# Valencia municipalities: prov_cod = 46
mun_46 = df_mun[df_mun['ID_Prov'] == 46].copy()
mun_codes_46 = set(mun_46['ID_Dest'].tolist())
print(f"Valencia municipalities: {len(mun_codes_46)}")

# Build comarca lookup
df_com_val = df_com[df_com['provincia'].str.contains('Valencia', na=False, case=False)].copy()
comarca_lookup = dict(zip(df_com_val['código INE'], df_com_val['comarca']))
print(f"Valencia comarcas: {df_com_val['comarca'].nunique()}")

# Also build mun_name -> comarca lookup for VUT matching
mun_name_comarca = {}
if 'municipio' in [c.lower() for c in df_mun.columns]:
    mun_col = [c for c in df_mun.columns if c.lower() == 'municipio'][0]
    for _, row in mun_46.iterrows():
        comarca_id = row.get('ID_Com', None)
        if comarca_id and comarca_id in comarca_lookup:
            mun_name_comarca[str(row[mun_col]).strip().upper()] = comarca_lookup[comarca_id]

# ---- SECTION 2: TURISMO RECEPTOR (International tourists) ----
print("\n=== Processing RECEPTOR (International) ===")
receptor_files = [v for k, v in data_files.items() if 'receptor' in k.lower() and v.endswith('.xlsx')]
receptor_files_sorted = sorted(receptor_files)

all_receptor = []
for fpath in receptor_files_sorted:
    fname = os.path.basename(fpath)
    xl = pd.ExcelFile(fpath)
    sheets = [s for s in xl.sheet_names if s != 'Notas']
    for sheet in sheets:
        try:
            df = xl.parse(sheet)
            if df.empty or len(df.columns) == 0:
                continue
            if 'prov_dest_cod' in df.columns:
                df_v = df[df['prov_dest_cod'] == 46].copy()
            elif 'mun_dest_cod' in df.columns:
                df_v = df[df['mun_dest_cod'].isin(mun_codes_46)].copy()
            else:
                continue
            if not df_v.empty:
                all_receptor.append(df_v)
        except Exception as e:
            print(f"  Error reading {fname}/{sheet}: {e}")

df_receptor = pd.concat(all_receptor, ignore_index=True) if all_receptor else pd.DataFrame()
print(f"  Total receptor rows for Valencia: {len(df_receptor)}")

if not df_receptor.empty:
    df_receptor['mes_dt'] = pd.to_datetime(df_receptor['mes'].astype(str), format='%Y-%m', errors='coerce')
    df_receptor['year'] = df_receptor['mes_dt'].dt.year
    df_receptor['month'] = df_receptor['mes_dt'].dt.month

    # Aggregate codes to exclude
    aggregate_codes = {0, 10, 11, 12, 20, 30, 40, 50, 60, 70, 80, 90, 100}
    df_rec_total = df_receptor[df_receptor['pais_orig_cod'] == 0].copy()

    # TOP 10 COUNTRIES 2025
    df_rec_countries_2025 = df_receptor[
        (~df_receptor['pais_orig_cod'].isin(aggregate_codes)) &
        (df_receptor['year'] == 2025)
    ].groupby(['pais_orig_cod', 'pais_orig'])['turistas'].sum().reset_index()
    top10_countries_2025 = df_rec_countries_2025.nlargest(10, 'turistas')

    df_rec_countries_2024 = df_receptor[
        (~df_receptor['pais_orig_cod'].isin(aggregate_codes)) &
        (df_receptor['year'] == 2024)
    ].groupby(['pais_orig_cod', 'pais_orig'])['turistas'].sum().reset_index()

    print("  Top 10 countries 2025:")
    for _, row in top10_countries_2025.iterrows():
        prev = df_rec_countries_2024[df_rec_countries_2024['pais_orig_cod'] == row['pais_orig_cod']]['turistas'].sum()
        chg = pct_change(row['turistas'], prev) if prev > 0 else None
        print(f"    {row['pais_orig']}: {fmt_num(row['turistas'])} ({fmt_pct(chg)})")

# ---- SECTION 3: TURISMO INTERNO (Domestic tourists) ----
print("\n=== Processing INTERNO (Domestic) ===")
interno_files = [v for k, v in data_files.items() if 'interno' in k.lower() and v.endswith('.xlsx')]
interno_files_sorted = sorted(interno_files)

all_interno = []
for fpath in interno_files_sorted:
    fname = os.path.basename(fpath)
    xl = pd.ExcelFile(fpath)
    sheets = [s for s in xl.sheet_names if s != 'Notas']
    for sheet in sheets:
        try:
            df = xl.parse(sheet)
            if df.empty or len(df.columns) == 0:
                continue
            if 'prov_dest_cod' in df.columns:
                df_v = df[df['prov_dest_cod'] == 46].copy()
            else:
                continue
            if not df_v.empty:
                all_interno.append(df_v)
        except Exception as e:
            print(f"  Error reading {fname}/{sheet}: {e}")

df_interno = pd.concat(all_interno, ignore_index=True) if all_interno else pd.DataFrame()
print(f"  Total interno rows for Valencia: {len(df_interno)}")

if not df_interno.empty:
    df_interno['mes_dt'] = pd.to_datetime(df_interno['mes'].astype(str), format='%Y-%m', errors='coerce')
    df_interno['year'] = df_interno['mes_dt'].dt.year
    df_interno['month'] = df_interno['mes_dt'].dt.month

    # TOP 10 DOMESTIC ORIGINS 2025
    df_int_origins_2025 = df_interno[df_interno['year'] == 2025].groupby(['prov_orig_cod', 'prov_orig'])['turistas'].sum().reset_index()
    df_int_origins_ext_2025 = df_int_origins_2025[df_int_origins_2025['prov_orig_cod'] != 46]
    top10_origins_2025 = df_int_origins_ext_2025.nlargest(10, 'turistas')

    df_int_origins_2024 = df_interno[df_interno['year'] == 2024].groupby(['prov_orig_cod', 'prov_orig'])['turistas'].sum().reset_index()

    print("  Top 10 domestic origins 2025:")
    for _, row in top10_origins_2025.iterrows():
        prev = df_int_origins_2024[df_int_origins_2024['prov_orig_cod'] == row['prov_orig_cod']]['turistas'].sum()
        chg = pct_change(row['turistas'], prev) if prev > 0 else None
        print(f"    {row['prov_orig']}: {fmt_num(row['turistas'])} ({fmt_pct(chg)})")

# ---- SECTION 4: COMARCA BREAKDOWN ----
print("\n=== Comarca breakdown ===")
comarcas_2024 = {}
comarcas_2025 = {}

if not df_receptor.empty and not df_interno.empty:
    df_rec_2025 = df_receptor[(df_receptor['year'] == 2025) & (df_receptor['pais_orig_cod'] == 0)].copy()
    df_int_2025 = df_interno[df_interno['year'] == 2025].copy()
    df_rec_2024 = df_receptor[(df_receptor['year'] == 2024) & (df_receptor['pais_orig_cod'] == 0)].copy()
    df_int_2024 = df_interno[df_interno['year'] == 2024].copy()

    # Add comarca
    df_rec_2025['comarca'] = df_rec_2025['mun_dest_cod'].map(comarca_lookup)
    df_int_2025['comarca'] = df_int_2025['dest_cod'].map(comarca_lookup)
    df_rec_2024['comarca'] = df_rec_2024['mun_dest_cod'].map(comarca_lookup)
    df_int_2024['comarca'] = df_int_2024['dest_cod'].map(comarca_lookup)

    rec_by_comarca_2025 = df_rec_2025.groupby('comarca')['turistas'].sum()
    int_by_comarca_2025 = df_int_2025.groupby('comarca')['turistas'].sum()
    rec_by_comarca_2024 = df_rec_2024.groupby('comarca')['turistas'].sum()
    int_by_comarca_2024 = df_int_2024.groupby('comarca')['turistas'].sum()

    # Build 2025 dict
    all_comarcas = set(rec_by_comarca_2025.index) | set(int_by_comarca_2025.index)
    for comarca in all_comarcas:
        if comarca and not pd.isna(comarca):
            rec = int(rec_by_comarca_2025.get(comarca, 0))
            intn = int(int_by_comarca_2025.get(comarca, 0))
            rec24 = int(rec_by_comarca_2024.get(comarca, 0))
            int24 = int(int_by_comarca_2024.get(comarca, 0))
            total25 = rec + intn
            total24 = rec24 + int24
            comarcas_2025[comarca] = {
                'receptor': rec,
                'interno': intn,
                'total': total25,
                'pct_internacional': round(rec / total25 * 100, 1) if total25 > 0 else 0
            }
            comarcas_2024[comarca] = {
                'receptor': rec24,
                'interno': int24,
                'total': total24
            }

# ---- SECTION 5: VIAJEROS Y PERNOCTACIONES ----
print("\n=== Viajeros y Pernoctaciones ===")
df_vp = pd.read_csv(data_files['2074.csv'], sep=';')
df_vp_val = df_vp[df_vp['Provincias'].str.contains('Valencia', na=False, case=False)].copy()
df_vp_val['Total_num'] = pd.to_numeric(df_vp_val['Total'].str.replace('.', '').str.replace(',', '.'), errors='coerce')
df_vp_val['year'] = df_vp_val['Periodo'].str[:4].astype(int)
df_vp_val['month'] = df_vp_val['Periodo'].str[5:7].astype(int)

# Annual totals
perc_2025 = df_vp_val[(df_vp_val['year'] == 2025) & (df_vp_val['Viajeros y pernoctaciones'] == 'Pernoctaciones') & (df_vp_val['Residencia: Nivel 1'] == 'Total')]['Total_num'].sum()
perc_2024 = df_vp_val[(df_vp_val['year'] == 2024) & (df_vp_val['Viajeros y pernoctaciones'] == 'Pernoctaciones') & (df_vp_val['Residencia: Nivel 1'] == 'Total')]['Total_num'].sum()
viaj_2025 = df_vp_val[(df_vp_val['year'] == 2025) & (df_vp_val['Viajeros y pernoctaciones'] == 'Viajero') & (df_vp_val['Residencia: Nivel 1'] == 'Total')]['Total_num'].sum()
viaj_2024 = df_vp_val[(df_vp_val['year'] == 2024) & (df_vp_val['Viajeros y pernoctaciones'] == 'Viajero') & (df_vp_val['Residencia: Nivel 1'] == 'Total')]['Total_num'].sum()

print(f"  Pernoctaciones 2025: {fmt_num(perc_2025)} ({fmt_pct(pct_change(perc_2025, perc_2024))})")
print(f"  Viajeros 2025: {fmt_num(viaj_2025)} ({fmt_pct(pct_change(viaj_2025, viaj_2024))})")

# Monthly pernoctaciones 2024 and 2025
perc_monthly_2025 = df_vp_val[
    (df_vp_val['year'] == 2025) & (df_vp_val['Viajeros y pernoctaciones'] == 'Pernoctaciones') & (df_vp_val['Residencia: Nivel 1'] == 'Total')
].groupby('month')['Total_num'].sum().sort_index()

perc_monthly_2024 = df_vp_val[
    (df_vp_val['year'] == 2024) & (df_vp_val['Viajeros y pernoctaciones'] == 'Pernoctaciones') & (df_vp_val['Residencia: Nivel 1'] == 'Total')
].groupby('month')['Total_num'].sum().sort_index()

print("  Monthly pernoctaciones 2025:", {int(k): int(v) for k, v in perc_monthly_2025.items()})
print("  Monthly pernoctaciones 2024:", {int(k): int(v) for k, v in perc_monthly_2024.items()})

# ---- SECTION 6: ESTANCIA MEDIA ----
print("\n=== Estancia Media ===")
df_est = pd.read_csv(data_files['56942.csv'], sep=';')
df_est_val = df_est[df_est['Provincias de destino'].str.contains('Valencia', na=False, case=False)].copy()
df_est_val['Total_num'] = pd.to_numeric(df_est_val['Total'].str.replace(',', '.'), errors='coerce')

month_map_es = {
    'Total': 0, 'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4, 'Mayo': 5, 'Junio': 6,
    'Julio': 7, 'Agosto': 8, 'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12
}

# Total average stay
est_total = df_est_val[(df_est_val['Meses'] == 'Total') & (df_est_val['Procedencia de los viajeros'] == 'Total')]['Total_num'].mean()
print(f"  Estancia media total: {est_total:.2f}")

# National vs international estancia media
# Filter for 'Residentes en España' and 'No residentes en España'
est_nacional_total = df_est_val[
    (df_est_val['Meses'] == 'Total') &
    (df_est_val['Procedencia de los viajeros'].str.contains('Residentes en Espa', na=False))
]['Total_num'].mean()

est_internacional_total = df_est_val[
    (df_est_val['Meses'] == 'Total') &
    (df_est_val['Procedencia de los viajeros'].str.contains('No residentes', na=False))
]['Total_num'].mean()

print(f"  Estancia media nacional: {est_nacional_total:.2f}")
print(f"  Estancia media internacional: {est_internacional_total:.2f}")

# Monthly estancia media (Total)
est_monthly = {}
for _, row in df_est_val[(df_est_val['Procedencia de los viajeros'] == 'Total')].iterrows():
    mes = row['Meses']
    if mes in month_map_es and month_map_es[mes] > 0:
        est_monthly[month_map_es[mes]] = row['Total_num']

print("  Monthly estancia media:", est_monthly)

# ---- SECTION 7: HOTEL OCCUPANCY ----
print("\n=== Hotel Occupancy ===")
df_hocomp = pd.read_csv(data_files['Encuestaocupacion_Hoteles_INE.csv'], sep=';')
df_hocomp_val = df_hocomp[df_hocomp['Provincias'].str.contains('Valencia', na=False, case=False)].copy()
df_hocomp_val['Total_num'] = pd.to_numeric(df_hocomp_val['Total'].str.replace('.','').str.replace(',','.'), errors='coerce')
df_hocomp_val['year'] = df_hocomp_val['Periodo'].str[:4].astype(int)
df_hocomp_val['month'] = df_hocomp_val['Periodo'].str[5:7].astype(int)

occ_metric = 'Grado de ocupación por plazas'

# 2025 monthly occupancy
occ_2025 = df_hocomp_val[
    (df_hocomp_val['year'] == 2025) &
    (df_hocomp_val['Establecimientos y personal empleado (plazas)'] == occ_metric)
].groupby('month')['Total_num'].mean().sort_index()

# 2024 monthly occupancy
occ_2024 = df_hocomp_val[
    (df_hocomp_val['year'] == 2024) &
    (df_hocomp_val['Establecimientos y personal empleado (plazas)'] == occ_metric)
].groupby('month')['Total_num'].mean().sort_index()

occ_2025_avg = occ_2025.mean()
occ_2024_avg = occ_2024.mean()
print(f"  Hotel occupancy 2025 avg: {occ_2025_avg:.1f}%")
print(f"  Hotel occupancy 2024 avg: {occ_2024_avg:.1f}%")
print("  Monthly 2025:", {int(k): round(v,1) for k,v in occ_2025.items()})
print("  Monthly 2024:", {int(k): round(v,1) for k,v in occ_2024.items()})

# Apartment occupancy
df_apt = pd.read_csv(data_files['Encuestaocupacion_Apartamentos_INE.csv'], sep=';')
df_apt_val = df_apt[df_apt['Provincias'].str.contains('Valencia', na=False, case=False)].copy()
df_apt_val['Total_num'] = pd.to_numeric(df_apt_val['Total'].str.replace('.','').str.replace(',','.'), errors='coerce')
df_apt_val['year'] = df_apt_val['Periodo'].str[:4].astype(int)

apt_occ_2025 = df_apt_val[
    (df_apt_val['year'] == 2025) &
    (df_apt_val['Establecimientos y personal empleado (plazas)'].str.contains('ocupaci', na=False, case=False))
]['Total_num'].mean()
print(f"  Apartment occupancy 2025 avg: {apt_occ_2025:.1f}%" if not pd.isna(apt_occ_2025) else "  Apartment occupancy: N/D")

# ---- SECTION 8: GVA SUPPLY DATA ----
print("\n=== GVA Supply Data ===")

# Hotels
df_hoteles = pd.read_csv(data_files['Hoteles_GVA.csv'], encoding='latin-1', sep=';')
df_hoteles_val = df_hoteles[df_hoteles['Cod. Provincia'].astype(str) == '46'].copy()
df_hoteles_val = df_hoteles_val[df_hoteles_val['Estado'] == 'ALTA'].copy()
df_hoteles_val['Habitaciones'] = pd.to_numeric(df_hoteles_val['Habitaciones'], errors='coerce')
df_hoteles_val['Plazas'] = pd.to_numeric(df_hoteles_val['Plazas'], errors='coerce')

print(f"  Total hotels Valencia: {len(df_hoteles_val)}")

hotels_by_comarca = df_hoteles_val.groupby('Comarca').agg(
    establecimientos=('Signatura', 'count'),
    habitaciones=('Habitaciones', 'sum'),
    plazas=('Plazas', 'sum')
).sort_values('establecimientos', ascending=False)

hotels_by_cat = df_hoteles_val.groupby('Categoría').agg(
    establecimientos=('Signatura', 'count'),
    plazas=('Plazas', 'sum')
).sort_values('establecimientos', ascending=False)

# Top municipalities by hotels
hotels_by_mun = df_hoteles_val.groupby('Municipio').agg(
    establecimientos=('Signatura', 'count'),
    plazas=('Plazas', 'sum')
).sort_values('plazas', ascending=False).head(10)
print("  Top 10 municipalities by hotel beds:")
print(hotels_by_mun.to_string())

# ---- SECTION 9: CAMPINGS ----
print("\n=== Campings ===")
df_camp = None
for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
    try:
        df_camp = pd.read_csv(data_files['Campings_GVA.csv'], encoding=enc, sep=';')
        prov_col = [c for c in df_camp.columns if 'rovincia' in c or 'prov' in c.lower()]
        if prov_col:
            # Filter for Valencia
            df_camp_val = df_camp[df_camp[prov_col[0]].astype(str).str.contains('46|Valencia|VALENCIA', na=False, case=False)].copy()
            df_camp = df_camp_val
            print(f"  Campings ({enc}): {len(df_camp)} Valencia records")
            print("  Columns:", list(df_camp.columns)[:10])
            break
    except Exception as e:
        continue

campings_by_comarca = {}
camping_total = 0
camping_plazas_total = 0
if df_camp is not None and not df_camp.empty:
    # Find comarca column
    com_col = [c for c in df_camp.columns if 'omarca' in c or 'Comarca' in c]
    plazas_col = [c for c in df_camp.columns if 'laza' in c and 'total' in c.lower()]
    if not plazas_col:
        plazas_col = [c for c in df_camp.columns if 'laza' in c]

    camping_total = len(df_camp)
    if plazas_col:
        df_camp['plazas_num'] = pd.to_numeric(df_camp[plazas_col[0]], errors='coerce')
        camping_plazas_total = int(df_camp['plazas_num'].sum())

    # Group by comarca if col exists
    if com_col:
        camp_by_com = df_camp.groupby(com_col[0]).agg(
            count=('Provincia', 'count'),
            plazas=('plazas_num', 'sum')
        ).sort_values('count', ascending=False) if plazas_col else df_camp.groupby(com_col[0]).size().sort_values(ascending=False)
        print("  Campings by comarca (top 10):")
        print(camp_by_com.head(10).to_string())
        for com, row in camp_by_com.iterrows():
            if com and not pd.isna(com):
                if hasattr(row, 'count'):
                    campings_by_comarca[str(com)] = {'count': int(row['count']), 'plazas': int(row['plazas']) if not pd.isna(row['plazas']) else 0}
                else:
                    campings_by_comarca[str(com)] = {'count': int(row), 'plazas': 0}
    else:
        # Group by Municipio
        print("  No comarca column, grouping by municipio:")
        mun_col_c = [c for c in df_camp.columns if 'unicip' in c]
        if mun_col_c and plazas_col:
            by_mun = df_camp.groupby(mun_col_c[0]).agg(count=('Provincia','count'), plazas=('plazas_num','sum')).sort_values('count', ascending=False)
            print(by_mun.head(10).to_string())

print(f"  Total campings: {camping_total}, total plazas: {camping_plazas_total}")

# ---- SECTION 10: CASAS RURALES ----
print("\n=== Casas Rurales ===")
df_cr = None
for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
    try:
        df_cr = pd.read_csv(data_files['Casasrurales_GVA.csv'], encoding=enc, sep=';')
        prov_col = [c for c in df_cr.columns if 'rovincia' in c or 'prov' in c.lower()]
        if prov_col:
            df_cr_val = df_cr[df_cr[prov_col[0]].astype(str).str.contains('46|Valencia|VALENCIA', na=False, case=False)].copy()
            df_cr = df_cr_val
            print(f"  Casas Rurales ({enc}): {len(df_cr)} Valencia records")
            print("  Columns:", list(df_cr.columns)[:10])
            break
    except Exception as e:
        continue

casasrurales_by_comarca = {}
cr_total = 0
cr_plazas_total = 0
if df_cr is not None and not df_cr.empty:
    # Find plazas column - could be 'N° plazas totales' or similar
    plazas_cr_col = [c for c in df_cr.columns if 'laza' in c and 'total' in c.lower()]
    if not plazas_cr_col:
        plazas_cr_col = [c for c in df_cr.columns if 'laza' in c]

    cr_total = len(df_cr)
    if plazas_cr_col:
        df_cr['plazas_num'] = pd.to_numeric(df_cr[plazas_cr_col[0]], errors='coerce')
        cr_plazas_total = int(df_cr['plazas_num'].sum())

    # Group by comarca or municipio
    com_col_cr = [c for c in df_cr.columns if 'omarca' in c]
    mun_col_cr = [c for c in df_cr.columns if 'unicip' in c]

    if mun_col_cr and plazas_cr_col:
        # Group by municipality, try to map to comarca
        cr_by_mun = df_cr.groupby(mun_col_cr[0]).agg(
            count=('Provincia', 'count'),
            plazas=('plazas_num', 'sum')
        ).sort_values('count', ascending=False)
        print("  Casa Rurales by municipio (top 15):")
        print(cr_by_mun.head(15).to_string())

        # Map municipality to comarca using mun_name_comarca lookup
        df_cr['comarca_mapped'] = df_cr[mun_col_cr[0]].astype(str).str.strip().str.upper().map(mun_name_comarca)
        cr_com_grp = df_cr.groupby('comarca_mapped').agg(
            count=('Provincia', 'count'),
            plazas=('plazas_num', 'sum')
        ).sort_values('count', ascending=False)
        print("  Casa Rurales by comarca (top 15):")
        print(cr_com_grp.head(15).to_string())
        for com, row in cr_com_grp.iterrows():
            if com and not pd.isna(com):
                casasrurales_by_comarca[str(com)] = {'count': int(row['count']), 'plazas': int(row['plazas']) if not pd.isna(row['plazas']) else 0}

print(f"  Total casas rurales: {cr_total}, total plazas: {cr_plazas_total}")

# ---- SECTION 11: ALBERGUES ----
print("\n=== Albergues ===")
df_alb = None
for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
    try:
        df_alb = pd.read_csv(data_files['Albergues_GVA.csv'], encoding=enc, sep=';')
        prov_col = [c for c in df_alb.columns if 'rovincia' in c or 'prov' in c.lower()]
        if prov_col:
            df_alb_val = df_alb[df_alb[prov_col[0]].astype(str).str.contains('46|Valencia|VALENCIA', na=False, case=False)].copy()
            df_alb = df_alb_val
            print(f"  Albergues ({enc}): {len(df_alb)} Valencia records")
            print("  Columns:", list(df_alb.columns)[:10])
            break
    except Exception as e:
        continue

alb_total = 0
alb_plazas_total = 0
if df_alb is not None and not df_alb.empty:
    alb_total = len(df_alb)
    # Find total plazas col
    plazas_alb_col = [c for c in df_alb.columns if 'laza' in c]
    if plazas_alb_col:
        # Sum all plazas columns
        total_plazas = 0
        for pc in plazas_alb_col:
            df_alb[pc] = pd.to_numeric(df_alb[pc], errors='coerce')
            total_plazas += df_alb[pc].sum()
        alb_plazas_total = int(total_plazas)
    print(f"  Albergues total: {alb_total}, plazas: {alb_plazas_total}")

# ---- SECTION 12: VUT (Tourism Apartments) ----
print("\n=== VUT Processing ===")
all_vut = []
err_count = 0
for f in vut_files:
    try:
        df_vut = pd.read_csv(f, encoding='cp1252', sep=';', low_memory=False)
        if 'Provincia' in df_vut.columns:
            df_vut_val = df_vut[df_vut['Provincia'].astype(str).str.contains('VALENCIA|Valencia', na=False)]
            if not df_vut_val.empty:
                all_vut.append(df_vut_val)
        elif any('rovincia' in c for c in df_vut.columns):
            pc = [c for c in df_vut.columns if 'rovincia' in c][0]
            df_vut_val = df_vut[df_vut[pc].astype(str).str.contains('VALENCIA|Valencia', na=False)]
            if not df_vut_val.empty:
                all_vut.append(df_vut_val)
    except Exception as e:
        err_count += 1

df_vut_all = pd.concat(all_vut, ignore_index=True) if all_vut else pd.DataFrame()
print(f"  Total VUT Valencia: {len(df_vut_all)} (errors: {err_count})")

vut_total = 0
vut_plazas = 0
vut_by_mun_dict = {}
vut_by_comarca_dict = {}

if not df_vut_all.empty:
    # Find plazas column
    plazas_vut_col = [c for c in df_vut_all.columns if 'laza' in c and 'total' in c.lower()]
    if not plazas_vut_col:
        plazas_vut_col = [c for c in df_vut_all.columns if 'laza' in c]

    vut_total = len(df_vut_all)
    if plazas_vut_col:
        df_vut_all['plazas_num'] = pd.to_numeric(df_vut_all[plazas_vut_col[0]], errors='coerce')
        vut_plazas = int(df_vut_all['plazas_num'].sum())

    # By municipality
    mun_vut_col = [c for c in df_vut_all.columns if 'unicip' in c.lower()]
    if mun_vut_col and plazas_vut_col:
        vut_by_mun = df_vut_all.groupby(mun_vut_col[0]).agg(
            count=('plazas_num', 'count'),
            plazas=('plazas_num', 'sum')
        ).sort_values('count', ascending=False).head(15)
        print("  Top 15 VUT municipalities:")
        print(vut_by_mun.to_string())
        for mun, row in vut_by_mun.iterrows():
            vut_by_mun_dict[str(mun)] = {'count': int(row['count']), 'plazas': int(row['plazas']) if not pd.isna(row['plazas']) else 0}

    # By comarca (map via mun_name_comarca)
    if mun_vut_col:
        df_vut_all['comarca_mapped'] = df_vut_all[mun_vut_col[0]].astype(str).str.strip().str.upper().map(mun_name_comarca)
        vut_by_com = df_vut_all.groupby('comarca_mapped').agg(
            count=('plazas_num', 'count'),
            plazas=('plazas_num', 'sum')
        ).sort_values('count', ascending=False)
        print("  VUT by comarca:")
        print(vut_by_com.head(12).to_string())
        for com, row in vut_by_com.iterrows():
            if com and not pd.isna(com):
                vut_by_comarca_dict[str(com)] = {'count': int(row['count']), 'plazas': int(row['plazas']) if not pd.isna(row['plazas']) else 0}

print(f"  VUT total: {vut_total}, plazas: {vut_plazas}")

# ---- SECTION 13: SS AFFILIATES ----
print("\n=== SS Affiliates ===")
df_ss = pd.read_excel(data_files['Afiliados_SS.xlsx'])
print(df_ss.to_string())

ss_data = {}
for _, row in df_ss.iterrows():
    fecha = str(row['Fecha '])[:10]
    cat = str(row['Categoría']).strip()
    if fecha not in ss_data:
        ss_data[fecha] = {}
    ss_data[fecha][cat] = int(row['Afiliados'])

# Calculate seasonality
all_ss_vals = {}
for fecha, cats in ss_data.items():
    for cat, val in cats.items():
        if cat not in all_ss_vals:
            all_ss_vals[cat] = []
        all_ss_vals[cat].append(val)

# Peak/trough ratio (July vs January)
q1_total = sum(ss_data.get('2025-01-01', {}).values())
q3_total = sum(ss_data.get('2025-07-01', {}).values())
q4_total = sum(ss_data.get('2025-10-01', {}).values())
seasonality_ratio = round(q3_total / q1_total * 100, 1) if q1_total > 0 else None
print(f"  Q1 total SS: {q1_total}, Q3 total SS: {q3_total}")
print(f"  Seasonality ratio (Q3/Q1): {seasonality_ratio}")

# ---- BUILD FINAL METRICS DICT ----
print("\n=== Building final metrics ===")
metrics = {}

# Totals
total_rec_2025 = 0
total_rec_2024 = 0
total_int_2025 = 0
total_int_2024 = 0

if not df_receptor.empty:
    total_rec_2025 = int(df_receptor[(df_receptor['year'] == 2025) & (df_receptor['pais_orig_cod'] == 0)]['turistas'].sum())
    total_rec_2024 = int(df_receptor[(df_receptor['year'] == 2024) & (df_receptor['pais_orig_cod'] == 0)]['turistas'].sum())

if not df_interno.empty:
    total_int_2025 = int(df_interno[df_interno['year'] == 2025]['turistas'].sum())
    total_int_2024 = int(df_interno[df_interno['year'] == 2024]['turistas'].sum())

total_2025 = total_rec_2025 + total_int_2025
total_2024 = total_rec_2024 + total_int_2024

metrics['total_turistas_2025'] = total_2025
metrics['total_turistas_2024'] = total_2024
metrics['chg_total'] = pct_change(total_2025, total_2024)
metrics['total_receptor_2025'] = total_rec_2025
metrics['total_receptor_2024'] = total_rec_2024
metrics['chg_receptor'] = pct_change(total_rec_2025, total_rec_2024)
metrics['total_interno_2025'] = total_int_2025
metrics['total_interno_2024'] = total_int_2024
metrics['chg_interno'] = pct_change(total_int_2025, total_int_2024)

# Historical series
hist_rec = {}
hist_int = {}
if not df_receptor.empty:
    rec_by_year = df_receptor[df_receptor['pais_orig_cod'] == 0].groupby('year')['turistas'].sum()
    for y, v in rec_by_year.items():
        hist_rec[int(y)] = int(v)
if not df_interno.empty:
    int_by_year = df_interno.groupby('year')['turistas'].sum()
    for y, v in int_by_year.items():
        hist_int[int(y)] = int(v)
metrics['historico_receptor'] = hist_rec
metrics['historico_interno'] = hist_int

# Monthly data
if not df_receptor.empty:
    df_rec_total_data = df_receptor[df_receptor['pais_orig_cod'] == 0]
    rec_2025_mo = df_rec_total_data[df_rec_total_data['year'] == 2025].groupby('month')['turistas'].sum()
    rec_2024_mo = df_rec_total_data[df_rec_total_data['year'] == 2024].groupby('month')['turistas'].sum()
    metrics['receptor_mensual_2025'] = {int(k): int(v) for k, v in rec_2025_mo.items()}
    metrics['receptor_mensual_2024'] = {int(k): int(v) for k, v in rec_2024_mo.items()}

if not df_interno.empty:
    int_2025_mo = df_interno[df_interno['year'] == 2025].groupby('month')['turistas'].sum()
    int_2024_mo = df_interno[df_interno['year'] == 2024].groupby('month')['turistas'].sum()
    metrics['interno_mensual_2025'] = {int(k): int(v) for k, v in int_2025_mo.items()}
    metrics['interno_mensual_2024'] = {int(k): int(v) for k, v in int_2024_mo.items()}

# TOP 10 countries
if not df_receptor.empty:
    top10_c = []
    for _, row in top10_countries_2025.iterrows():
        prev = df_rec_countries_2024[df_rec_countries_2024['pais_orig_cod'] == row['pais_orig_cod']]['turistas'].sum()
        top10_c.append({
            'pais': row['pais_orig'],
            'turistas': int(row['turistas']),
            'chg': pct_change(row['turistas'], prev) if prev > 0 else None
        })
    metrics['top5_paises_2025'] = top10_c[:5]
    metrics['top10_paises_2025'] = top10_c

# TOP 10 domestic origins
if not df_interno.empty:
    top10_o = []
    for _, row in top10_origins_2025.iterrows():
        prev = df_int_origins_2024[df_int_origins_2024['prov_orig_cod'] == row['prov_orig_cod']]['turistas'].sum()
        top10_o.append({
            'origen': row['prov_orig'],
            'turistas': int(row['turistas']),
            'chg': pct_change(row['turistas'], prev) if prev > 0 else None
        })
    metrics['top5_origenes_nacionales_2025'] = top10_o[:5]
    metrics['top10_origenes_nacionales_2025'] = top10_o

# Comarca data
metrics['turistas_por_comarca_2025'] = comarcas_2025
metrics['turistas_por_comarca_2024'] = comarcas_2024

# Pernoctaciones
metrics['pernoctaciones_2025'] = int(perc_2025) if not pd.isna(perc_2025) else 0
metrics['pernoctaciones_2024'] = int(perc_2024) if not pd.isna(perc_2024) else 0
metrics['chg_pernoctaciones'] = pct_change(perc_2025, perc_2024)
metrics['viajeros_2025'] = int(viaj_2025) if not pd.isna(viaj_2025) else 0
metrics['viajeros_2024'] = int(viaj_2024) if not pd.isna(viaj_2024) else 0
metrics['pernoctaciones_mensual_2025'] = {int(k): int(v) for k, v in perc_monthly_2025.items() if not pd.isna(v)}
metrics['pernoctaciones_mensual_2024'] = {int(k): int(v) for k, v in perc_monthly_2024.items() if not pd.isna(v)}

# Estancia media
metrics['estancia_media_total'] = round(float(est_total), 2) if not pd.isna(est_total) else 2.24
metrics['estancia_media_nacional'] = round(float(est_nacional_total), 2) if not pd.isna(est_nacional_total) else None
metrics['estancia_media_internacional'] = round(float(est_internacional_total), 2) if not pd.isna(est_internacional_total) else None
metrics['estancia_media_mensual'] = {int(k): round(float(v), 2) for k, v in est_monthly.items() if not pd.isna(v)}

# Hotel occupancy
metrics['ocupacion_hoteles_2025_avg'] = round(occ_2025_avg, 1) if not pd.isna(occ_2025_avg) else None
metrics['ocupacion_hoteles_2024_avg'] = round(occ_2024_avg, 1) if not pd.isna(occ_2024_avg) else None
metrics['chg_ocupacion'] = pct_change(occ_2025_avg, occ_2024_avg) if not pd.isna(occ_2025_avg) else None
metrics['ocupacion_hoteles_mensual_2025'] = {int(k): round(v, 1) for k, v in occ_2025.items() if not pd.isna(v)}
metrics['ocupacion_hoteles_mensual_2024'] = {int(k): round(v, 1) for k, v in occ_2024.items() if not pd.isna(v)}

# Supply - hotels
metrics['hoteles_total'] = int(len(df_hoteles_val))
metrics['hoteles_habitaciones'] = int(df_hoteles_val['Habitaciones'].sum())
metrics['hoteles_plazas'] = int(df_hoteles_val['Plazas'].sum())

cat_dict = {}
for cat, row in hotels_by_cat.iterrows():
    cat_dict[str(cat)] = {'establecimientos': int(row['establecimientos']), 'plazas': int(row['plazas'])}
metrics['hoteles_por_categoria'] = cat_dict

comarca_hotels = {}
for comarca, row in hotels_by_comarca.iterrows():
    comarca_hotels[str(comarca)] = {
        'establecimientos': int(row['establecimientos']),
        'habitaciones': int(row['habitaciones']) if not pd.isna(row['habitaciones']) else 0,
        'plazas': int(row['plazas']) if not pd.isna(row['plazas']) else 0
    }
metrics['hoteles_por_comarca'] = comarca_hotels

# Supply - VUT
metrics['vut_total'] = vut_total
metrics['vut_plazas'] = vut_plazas
metrics['vut_por_municipio'] = vut_by_mun_dict
metrics['vut_por_comarca'] = vut_by_comarca_dict

# Supply - Campings
metrics['campings_total'] = camping_total
metrics['campings_plazas'] = camping_plazas_total
metrics['campings_por_comarca'] = campings_by_comarca

# Supply - Casas Rurales
metrics['casasrurales_total'] = cr_total
metrics['casasrurales_plazas'] = cr_plazas_total
metrics['casasrurales_por_comarca'] = casasrurales_by_comarca

# Supply - Albergues
metrics['albergues_total'] = alb_total
metrics['albergues_plazas'] = alb_plazas_total

# Employment
metrics['afiliados_ss'] = ss_data
metrics['seasonality_ratio_ss'] = seasonality_ratio

# Save to JSON
output_path = os.path.join(os.path.dirname(data_files['Hoteles_GVA.csv']), '..', '..', 'metrics_dival.json')
output_path = os.path.normpath(output_path)
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(metrics, f, ensure_ascii=False, indent=2, default=str)
print(f"\n\nMetrics saved to: {output_path}")
print("\n=== FINAL KEY METRICS ===")
print(f"  Total turistas 2025: {fmt_num(metrics['total_turistas_2025'])}")
print(f"  Turistas internacionales: {fmt_num(metrics['total_receptor_2025'])} ({fmt_pct(metrics['chg_receptor'])})")
print(f"  Turistas nacionales: {fmt_num(metrics['total_interno_2025'])} ({fmt_pct(metrics['chg_interno'])})")
print(f"  Pernoctaciones: {fmt_num(metrics['pernoctaciones_2025'])}")
print(f"  Estancia media: {metrics['estancia_media_total']}")
print(f"  Hoteles: {metrics['hoteles_total']} ({fmt_num(metrics['hoteles_plazas'])} plazas)")
print(f"  VUTs: {metrics['vut_total']} ({fmt_num(metrics['vut_plazas'])} plazas)")
print(f"  Campings: {metrics['campings_total']} ({fmt_num(metrics['campings_plazas'])} plazas)")
print(f"  Casas rurales: {metrics['casasrurales_total']} ({fmt_num(metrics['casasrurales_plazas'])} plazas)")
print(f"  Albergues: {metrics['albergues_total']} ({fmt_num(metrics['albergues_plazas'])} plazas)")
