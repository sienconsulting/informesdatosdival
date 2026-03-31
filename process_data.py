"""
Data processing script for Valencia Province Tourism Report 2025
Processes all data sources and extracts metrics for use in PPTX
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

# ---- SECTION 2: TURISMO RECEPTOR (International tourists) ----
print("\n=== Processing RECEPTOR (International) ===")
receptor_files = [v for k, v in data_files.items() if 'receptor' in k.lower() and v.endswith('.xlsx')]
receptor_files_sorted = sorted(receptor_files)

all_receptor = []
for fpath in receptor_files_sorted:
    fname = os.path.basename(fpath)
    # Get year from filename
    year_str = None
    for part in fname.split('_'):
        if '2019' <= part[:4] <= '2026':
            year_str = part[:4]
            break

    xl = pd.ExcelFile(fpath)
    # Skip 'Notas' sheet
    sheets = [s for s in xl.sheet_names if s != 'Notas']

    for sheet in sheets:
        try:
            df = xl.parse(sheet)
            if df.empty or len(df.columns) == 0:
                continue
            # Filter for Valencia destination province (prov_dest_cod == 46)
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
    # Parse date
    df_receptor['mes_dt'] = pd.to_datetime(df_receptor['mes'].astype(str), format='%Y-%m', errors='coerce')
    df_receptor['year'] = df_receptor['mes_dt'].dt.year
    df_receptor['month'] = df_receptor['mes_dt'].dt.month

    # Unique pais values
    print("  Unique countries (sample):", df_receptor['pais_orig'].unique()[:10])

    # Filter for Total (pais_orig_cod == 0 means all countries combined)
    df_rec_total = df_receptor[df_receptor['pais_orig_cod'] == 0].copy()

    # Annual totals by year
    rec_annual = df_rec_total.groupby('year')['turistas'].sum()
    print("  Annual receptor totals:")
    for y, v in rec_annual.items():
        print(f"    {y}: {fmt_num(v)}")

    # Monthly 2025
    rec_2025_monthly = df_rec_total[df_rec_total['year'] == 2025].groupby('month')['turistas'].sum()
    rec_2024_monthly = df_rec_total[df_rec_total['year'] == 2024].groupby('month')['turistas'].sum()

    # Top countries 2025 (exclude aggregated)
    # Exclude codes 0 (Total), 10 (Total Europa), 11 (Total UE), etc.
    aggregate_codes = {0, 10, 11, 12, 20, 30, 40, 50, 60, 70, 80, 90, 100}
    df_rec_countries = df_receptor[
        (~df_receptor['pais_orig_cod'].isin(aggregate_codes)) &
        (df_receptor['year'] == 2025)
    ].groupby(['pais_orig_cod', 'pais_orig'])['turistas'].sum().reset_index()
    top5_countries_2025 = df_rec_countries.nlargest(5, 'turistas')

    # Same for 2024
    df_rec_countries_2024 = df_receptor[
        (~df_receptor['pais_orig_cod'].isin(aggregate_codes)) &
        (df_receptor['year'] == 2024)
    ].groupby(['pais_orig_cod', 'pais_orig'])['turistas'].sum().reset_index()

    print("  Top 5 countries 2025:")
    for _, row in top5_countries_2025.iterrows():
        # Find 2024 value
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
            # For interno: we want tourists going TO Valencia (prov_dest_cod == 46)
            # columns: mes, mun_orig_cod, mun_orig, dest_cod, dest, turistas, prov_orig_cod, prov_orig, prov_dest_cod, prov_dest
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

    # Annual totals
    int_annual = df_interno.groupby('year')['turistas'].sum()
    print("  Annual interno totals:")
    for y, v in int_annual.items():
        print(f"    {y}: {fmt_num(v)}")

    # Monthly 2025/2024
    int_2025_monthly = df_interno[df_interno['year'] == 2025].groupby('month')['turistas'].sum()
    int_2024_monthly = df_interno[df_interno['year'] == 2024].groupby('month')['turistas'].sum()

    # Top origin communities for 2025
    df_int_origins = df_interno[df_interno['year'] == 2025].groupby(['prov_orig_cod', 'prov_orig'])['turistas'].sum().reset_index()
    # Exclude Valencia itself (code 46) and aggregate
    df_int_origins_ext = df_int_origins[df_int_origins['prov_orig_cod'] != 46]
    top5_origins_2025 = df_int_origins_ext.nlargest(5, 'turistas')

    df_int_origins_2024 = df_interno[df_interno['year'] == 2024].groupby(['prov_orig_cod', 'prov_orig'])['turistas'].sum().reset_index()

    print("  Top 5 domestic origins 2025:")
    for _, row in top5_origins_2025.iterrows():
        prev = df_int_origins_2024[df_int_origins_2024['prov_orig_cod'] == row['prov_orig_cod']]['turistas'].sum()
        chg = pct_change(row['turistas'], prev) if prev > 0 else None
        print(f"    {row['prov_orig']}: {fmt_num(row['turistas'])} ({fmt_pct(chg)})")

# ---- SECTION 4: COMARCA BREAKDOWN ----
print("\n=== Comarca breakdown ===")
# Merge turistas with comarca lookup
if not df_receptor.empty and not df_interno.empty:
    df_rec_2025 = df_receptor[(df_receptor['year'] == 2025) & (df_receptor['pais_orig_cod'] == 0)].copy()
    df_int_2025 = df_interno[df_interno['year'] == 2025].copy()

    # Add comarca
    df_rec_2025['comarca'] = df_rec_2025['mun_dest_cod'].map(comarca_lookup)
    df_int_2025['comarca'] = df_int_2025['dest_cod'].map(comarca_lookup)

    rec_by_comarca = df_rec_2025.groupby('comarca')['turistas'].sum().sort_values(ascending=False)
    int_by_comarca = df_int_2025.groupby('comarca')['turistas'].sum().sort_values(ascending=False)

    print("  Top 10 comarcas by international tourists:")
    print(rec_by_comarca.head(10).to_string())
    print("  Top 10 comarcas by domestic tourists:")
    print(int_by_comarca.head(10).to_string())

# ---- SECTION 5: VIAJEROS Y PERNOCTACIONES ----
print("\n=== Viajeros y Pernoctaciones ===")
df_vp = pd.read_csv(data_files['2074.csv'], sep=';')
df_vp_val = df_vp[df_vp['Provincias'].str.contains('Valencia', na=False, case=False)].copy()
df_vp_val['Total_num'] = pd.to_numeric(df_vp_val['Total'].str.replace('.', '').str.replace(',', '.'), errors='coerce')
df_vp_val['year'] = df_vp_val['Periodo'].str[:4].astype(int)
df_vp_val['month'] = df_vp_val['Periodo'].str[5:7].astype(int)

# Overnight stays 2025 (total)
perc_2025 = df_vp_val[
    (df_vp_val['year'] == 2025) &
    (df_vp_val['Viajeros y pernoctaciones'] == 'Pernoctaciones') &
    (df_vp_val['Residencia: Nivel 1'] == 'Total')
]['Total_num'].sum()

perc_2024 = df_vp_val[
    (df_vp_val['year'] == 2024) &
    (df_vp_val['Viajeros y pernoctaciones'] == 'Pernoctaciones') &
    (df_vp_val['Residencia: Nivel 1'] == 'Total')
]['Total_num'].sum()

viaj_2025 = df_vp_val[
    (df_vp_val['year'] == 2025) &
    (df_vp_val['Viajeros y pernoctaciones'] == 'Viajero') &
    (df_vp_val['Residencia: Nivel 1'] == 'Total')
]['Total_num'].sum()

viaj_2024 = df_vp_val[
    (df_vp_val['year'] == 2024) &
    (df_vp_val['Viajeros y pernoctaciones'] == 'Viajero') &
    (df_vp_val['Residencia: Nivel 1'] == 'Total')
]['Total_num'].sum()

print(f"  Pernoctaciones 2025: {fmt_num(perc_2025)} ({fmt_pct(pct_change(perc_2025, perc_2024))})")
print(f"  Viajeros 2025: {fmt_num(viaj_2025)} ({fmt_pct(pct_change(viaj_2025, viaj_2024))})")

# Monthly pernoctaciones 2025
perc_monthly_2025 = df_vp_val[
    (df_vp_val['year'] == 2025) &
    (df_vp_val['Viajeros y pernoctaciones'] == 'Pernoctaciones') &
    (df_vp_val['Residencia: Nivel 1'] == 'Total')
].groupby('month')['Total_num'].sum().sort_index()

# ---- SECTION 6: ESTANCIA MEDIA ----
print("\n=== Estancia Media ===")
df_est = pd.read_csv(data_files['56942.csv'], sep=';')
df_est_val = df_est[df_est['Provincias de destino'].str.contains('Valencia', na=False, case=False)].copy()
df_est_val['Total_num'] = pd.to_numeric(df_est_val['Total'].str.replace(',', '.'), errors='coerce')

# Total average stay (all months combined)
est_total = df_est_val[df_est_val['Meses'] == 'Total']
print(df_est_val.head(10)[['Provincias de destino','Procedencia de los viajeros','Meses','Total']].to_string())

# ---- SECTION 7: HOTEL OCCUPANCY ----
print("\n=== Hotel Occupancy ===")
df_hocomp = pd.read_csv(data_files['Encuestaocupacion_Hoteles_INE.csv'], sep=';')
df_hocomp_val = df_hocomp[df_hocomp['Provincias'].str.contains('Valencia', na=False, case=False)].copy()
df_hocomp_val['Total_num'] = pd.to_numeric(df_hocomp_val['Total'].str.replace('.','').str.replace(',','.'), errors='coerce')
df_hocomp_val['year'] = df_hocomp_val['Periodo'].str[:4].astype(int)
df_hocomp_val['month'] = df_hocomp_val['Periodo'].str[5:7].astype(int)

# Occupancy rate by month 2025
occ_metric = 'Grado de ocupación por plazas'
occ_2025 = df_hocomp_val[
    (df_hocomp_val['year'] == 2025) &
    (df_hocomp_val['Establecimientos y personal empleado (plazas)'] == occ_metric)
].groupby('month')['Total_num'].mean().sort_index()

# Annual avg occupancy
occ_2025_avg = occ_2025.mean()
occ_2024_avg = df_hocomp_val[
    (df_hocomp_val['year'] == 2024) &
    (df_hocomp_val['Establecimientos y personal empleado (plazas)'] == occ_metric)
]['Total_num'].mean()
print(f"  Hotel occupancy 2025 avg: {occ_2025_avg:.1f}% ({fmt_pct(pct_change(occ_2025_avg, occ_2024_avg))})")
print("  Monthly occupancy 2025:")
print(occ_2025.to_string())

# Also get apartment and rural tourism
df_apt = pd.read_csv(data_files['Encuestaocupacion_Apartamentos_INE.csv'], sep=';')
df_apt_val = df_apt[df_apt['Provincias'].str.contains('Valencia', na=False, case=False)].copy()
df_apt_val['Total_num'] = pd.to_numeric(df_apt_val['Total'].str.replace('.','').str.replace(',','.'), errors='coerce')
df_apt_val['year'] = df_apt_val['Periodo'].str[:4].astype(int)

df_rural = pd.read_csv(data_files['Encuestaocupacion_Turismorural_INE.csv'], sep=';', encoding='utf-8-sig')
# The rural CSV has a different province column name - find it dynamically
rural_prov_col = [c for c in df_rural.columns if 'rovincias' in c or 'rovincia' in c][0]
df_rural_val = df_rural[df_rural[rural_prov_col].str.contains('Valencia', na=False, case=False)].copy()
df_rural_val['Total_num'] = pd.to_numeric(df_rural_val['Total'].str.replace('.','').str.replace(',','.'), errors='coerce')
df_rural_val['year'] = df_rural_val['Periodo'].str[:4].astype(int)

# ---- SECTION 8: GVA SUPPLY DATA ----
print("\n=== GVA Supply Data ===")

# Hotels
df_hoteles = pd.read_csv(data_files['Hoteles_GVA.csv'], encoding='latin-1', sep=';')
df_hoteles_val = df_hoteles[df_hoteles['Cod. Provincia'].astype(str) == '46'].copy()
df_hoteles_val = df_hoteles_val[df_hoteles_val['Estado'] == 'ALTA'].copy()
df_hoteles_val['Habitaciones'] = pd.to_numeric(df_hoteles_val['Habitaciones'], errors='coerce')
df_hoteles_val['Plazas'] = pd.to_numeric(df_hoteles_val['Plazas'], errors='coerce')

print(f"  Total hotels Valencia: {len(df_hoteles_val)}")
print(f"  Total rooms: {df_hoteles_val['Habitaciones'].sum():.0f}")
print(f"  Total beds: {df_hoteles_val['Plazas'].sum():.0f}")

hotels_by_comarca = df_hoteles_val.groupby('Comarca').agg(
    establecimientos=('Signatura', 'count'),
    habitaciones=('Habitaciones', 'sum'),
    plazas=('Plazas', 'sum')
).sort_values('establecimientos', ascending=False)
print("  Top 10 comarcas by hotels:")
print(hotels_by_comarca.head(10).to_string())

hotels_by_cat = df_hoteles_val.groupby('Categoría').agg(
    establecimientos=('Signatura', 'count'),
    plazas=('Plazas', 'sum')
).sort_values('establecimientos', ascending=False)
print("  Hotels by category:")
print(hotels_by_cat.to_string())

# Campings
try:
    df_camp = pd.read_csv(data_files['Campings_GVA.csv'], encoding='utf-8', sep=';', nrows=5)
    print("  Campings cols:", list(df_camp.columns))
except:
    try:
        df_camp = pd.read_csv(data_files['Campings_GVA.csv'], encoding='latin-1', sep=';', nrows=5)
        print("  Campings cols (latin1):", list(df_camp.columns))
    except Exception as e:
        print(f"  Campings error: {e}")

# Try to load all campings
for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
    try:
        df_camp = pd.read_csv(data_files['Campings_GVA.csv'], encoding=enc, sep=';')
        # Find provincia column
        prov_col = [c for c in df_camp.columns if 'rovincia' in c or 'prov' in c.lower()]
        print(f"  Campings ({enc}): {len(df_camp)} rows, prov col: {prov_col}")
        if prov_col:
            df_camp_val = df_camp[df_camp[prov_col[0]].astype(str).str.contains('46|Valencia|VALENCIA', na=False)]
            print(f"  Valencia campings: {len(df_camp_val)}")
            break
    except Exception as e:
        continue

# Rural houses
for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
    try:
        df_cr = pd.read_csv(data_files['Casasrurales_GVA.csv'], encoding=enc, sep=';')
        prov_col = [c for c in df_cr.columns if 'rovincia' in c or 'prov' in c.lower()]
        print(f"  CasasRurales ({enc}): {len(df_cr)} rows, prov col: {prov_col}")
        if prov_col:
            df_cr_val = df_cr[df_cr[prov_col[0]].astype(str).str.contains('46|Valencia|VALENCIA', na=False)]
            print(f"  Valencia rural houses: {len(df_cr_val)}")
            break
    except Exception as e:
        continue

# VUT
print("\n  Processing VUT files...")
vut_list = []
for f in vut_files[:10]:  # Test first 10
    try:
        df_vut = pd.read_csv(f, encoding='cp1252', sep=';', low_memory=False)
        df_vut_val = df_vut[df_vut['Provincia'].astype(str).str.contains('VALENCIA|Valencia', na=False)]
        if not df_vut_val.empty:
            vut_list.append(df_vut_val)
    except Exception as e:
        pass

# Load all VUT
all_vut = []
err_count = 0
for f in vut_files:
    try:
        df_vut = pd.read_csv(f, encoding='cp1252', sep=';', low_memory=False)
        if 'Provincia' in df_vut.columns:
            df_vut_val = df_vut[df_vut['Provincia'].astype(str).str.contains('VALENCIA|Valencia', na=False)]
            if not df_vut_val.empty:
                all_vut.append(df_vut_val)
    except Exception as e:
        err_count += 1

df_vut_all = pd.concat(all_vut, ignore_index=True) if all_vut else pd.DataFrame()
print(f"  Total VUT Valencia: {len(df_vut_all)} (errors: {err_count})")
if not df_vut_all.empty:
    df_vut_all['Plazas totales'] = pd.to_numeric(df_vut_all['Plazas totales'], errors='coerce')
    print(f"  VUT total plazas: {df_vut_all['Plazas totales'].sum():.0f}")
    if 'Municipio' in df_vut_all.columns:
        vut_by_mun = df_vut_all.groupby('Municipio').agg(
            vuts=('Provincia', 'count'),
            plazas=('Plazas totales', 'sum')
        ).sort_values('vuts', ascending=False)
        print("  Top 10 municipalities by VUT:")
        print(vut_by_mun.head(10).to_string())

# ---- SECTION 9: SS AFFILIATES ----
print("\n=== SS Affiliates ===")
df_ss = pd.read_excel(data_files['Afiliados_SS.xlsx'])
print(df_ss.to_string())

# ---- BUILD FINAL METRICS DICT ----
print("\n=== Building final metrics ===")
metrics = {}

# Total tourists 2025
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

print(f"  Total tourists 2025: {fmt_num(total_2025)}")
print(f"  Total tourists 2024: {fmt_num(total_2024)}")
print(f"  Change: {fmt_pct(metrics['chg_total'])}")
print(f"  International 2025: {fmt_num(total_rec_2025)}")
print(f"  Domestic 2025: {fmt_num(total_int_2025)}")

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

# Pernoctaciones
metrics['pernoctaciones_2025'] = int(perc_2025) if not pd.isna(perc_2025) else 0
metrics['pernoctaciones_2024'] = int(perc_2024) if not pd.isna(perc_2024) else 0
metrics['chg_pernoctaciones'] = pct_change(perc_2025, perc_2024)
metrics['viajeros_2025'] = int(viaj_2025) if not pd.isna(viaj_2025) else 0
metrics['viajeros_2024'] = int(viaj_2024) if not pd.isna(viaj_2024) else 0

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

# Pernoctaciones mensual
perc_monthly = df_vp_val[
    (df_vp_val['year'] == 2025) &
    (df_vp_val['Viajeros y pernoctaciones'] == 'Pernoctaciones') &
    (df_vp_val['Residencia: Nivel 1'] == 'Total')
].groupby('month')['Total_num'].sum()
metrics['pernoctaciones_mensual_2025'] = {int(k): int(v) for k, v in perc_monthly.items() if not pd.isna(v)}

# Top countries and origins
if not df_receptor.empty:
    top5_c = []
    for _, row in top5_countries_2025.iterrows():
        prev = df_rec_countries_2024[df_rec_countries_2024['pais_orig_cod'] == row['pais_orig_cod']]['turistas'].sum()
        top5_c.append({
            'pais': row['pais_orig'],
            'turistas': int(row['turistas']),
            'chg': pct_change(row['turistas'], prev) if prev > 0 else None
        })
    metrics['top5_paises_2025'] = top5_c

if not df_interno.empty:
    top5_o = []
    for _, row in top5_origins_2025.iterrows():
        prev = df_int_origins_2024[df_int_origins_2024['prov_orig_cod'] == row['prov_orig_cod']]['turistas'].sum()
        top5_o.append({
            'origen': row['prov_orig'],
            'turistas': int(row['turistas']),
            'chg': pct_change(row['turistas'], prev) if prev > 0 else None
        })
    metrics['top5_origenes_nacionales_2025'] = top5_o

# Comarca breakdown
if not df_receptor.empty and not df_interno.empty:
    comarcas_dict = {}
    for comarca, rec_val in rec_by_comarca.items():
        int_val = int_by_comarca.get(comarca, 0) if comarca in int_by_comarca else 0
        if comarca and not pd.isna(comarca):
            comarcas_dict[comarca] = {
                'receptor': int(rec_val),
                'interno': int(int_val) if not pd.isna(int_val) else 0,
                'total': int(rec_val) + (int(int_val) if not pd.isna(int_val) else 0)
            }
    metrics['turistas_por_comarca_2025'] = comarcas_dict

# Hotel occupancy
metrics['ocupacion_hoteles_2025_avg'] = round(occ_2025_avg, 1) if not pd.isna(occ_2025_avg) else None
metrics['ocupacion_hoteles_2024_avg'] = round(occ_2024_avg, 1) if not pd.isna(occ_2024_avg) else None
metrics['chg_ocupacion'] = pct_change(occ_2025_avg, occ_2024_avg) if not pd.isna(occ_2025_avg) else None
metrics['ocupacion_hoteles_mensual_2025'] = {int(k): round(v, 1) for k, v in occ_2025.items() if not pd.isna(v)}

# Supply data
metrics['hoteles_total'] = int(len(df_hoteles_val))
metrics['hoteles_habitaciones'] = int(df_hoteles_val['Habitaciones'].sum())
metrics['hoteles_plazas'] = int(df_hoteles_val['Plazas'].sum())

# By category
cat_dict = {}
for cat, row in hotels_by_cat.iterrows():
    cat_dict[str(cat)] = {'establecimientos': int(row['establecimientos']), 'plazas': int(row['plazas'])}
metrics['hoteles_por_categoria'] = cat_dict

# By comarca
comarca_hotels = {}
for comarca, row in hotels_by_comarca.iterrows():
    comarca_hotels[str(comarca)] = {
        'establecimientos': int(row['establecimientos']),
        'habitaciones': int(row['habitaciones']) if not pd.isna(row['habitaciones']) else 0,
        'plazas': int(row['plazas']) if not pd.isna(row['plazas']) else 0
    }
metrics['hoteles_por_comarca'] = comarca_hotels

if not df_vut_all.empty:
    metrics['vut_total'] = int(len(df_vut_all))
    metrics['vut_plazas'] = int(df_vut_all['Plazas totales'].sum()) if not df_vut_all.empty else 0

# SS Affiliates
ss_data = {}
for _, row in df_ss.iterrows():
    fecha = str(row['Fecha '])[:10]
    cat = str(row['Categoría']).strip()
    if fecha not in ss_data:
        ss_data[fecha] = {}
    ss_data[fecha][cat] = int(row['Afiliados'])
metrics['afiliados_ss'] = ss_data

# Estancia media
est_val_total = df_est_val[
    (df_est_val['Meses'] == 'Total') &
    (df_est_val['Procedencia de los viajeros'] == 'Total')
]['Total_num'].mean()
metrics['estancia_media_total'] = round(float(est_val_total), 2) if not pd.isna(est_val_total) else None
print(f"  Estancia media total: {metrics['estancia_media_total']}")

# Save to JSON
output_path = os.path.join(os.path.dirname(data_files['Hoteles_GVA.csv']), '..', '..', 'metrics_dival.json')
output_path = os.path.normpath(output_path)
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(metrics, f, ensure_ascii=False, indent=2, default=str)
print(f"\n\nMetrics saved to: {output_path}")
print("\n=== FINAL KEY METRICS ===")
print(f"  Total turistas 2025: {fmt_num(metrics['total_turistas_2025'])}")
print(f"  Variación vs 2024: {fmt_pct(metrics['chg_total'])}")
print(f"  Turistas internacionales: {fmt_num(metrics['total_receptor_2025'])} ({fmt_pct(metrics['chg_receptor'])})")
print(f"  Turistas nacionales: {fmt_num(metrics['total_interno_2025'])} ({fmt_pct(metrics['chg_interno'])})")
print(f"  Pernoctaciones: {fmt_num(metrics['pernoctaciones_2025'])}")
print(f"  Viajeros: {fmt_num(metrics['viajeros_2025'])}")
print(f"  Hoteles activos: {metrics['hoteles_total']}")
print(f"  Plazas hoteleras: {fmt_num(metrics['hoteles_plazas'])}")
if 'vut_total' in metrics:
    print(f"  VUT registradas: {fmt_num(metrics['vut_total'])}")
    print(f"  Plazas VUT: {fmt_num(metrics['vut_plazas'])}")
