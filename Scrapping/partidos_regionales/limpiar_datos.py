"""
Clean and aggregate ONPE election data from mesa-level to district-level.
Creates readable CSVs with standardized columns.
"""
import pandas as pd
import unicodedata
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

BASE = r'C:/Users/Usuario/Documents/GitHub/Forest_Peru/Scrapping/partidos_regionales'


def fix_ubigeo_col(df):
    """Fix BOM in UBIGEO column name."""
    cols = list(df.columns)
    cols[0] = 'UBIGEO'
    df.columns = cols
    return df


def aggregate_to_district(df, year, election_type):
    """Aggregate mesa-level data to district-level: sum votes by party per district."""
    # Filter only valid rows
    if election_type == 'municipal':
        valid = df[df['TIPO_ELECCION'] == 'MUNICIPAL DISTRITAL'].copy()
    else:
        valid = df[df['TIPO_ELECCION'].str.contains('GOBERNADOR', na=False)].copy()

    if len(valid) == 0:
        print(f"  WARNING: no valid rows for {election_type}")
        return None

    # Drop rows with NaN in key columns
    valid = valid.dropna(subset=['UBIGEO', 'AGRUPACION_POLITICA', 'VOTOS_OBTENIDOS'])

    # Ensure numeric
    valid['VOTOS_OBTENIDOS'] = pd.to_numeric(valid['VOTOS_OBTENIDOS'], errors='coerce').fillna(0).astype(int)

    # Aggregate to district level
    agg = valid.groupby(
        ['UBIGEO', 'DEPARTAMENTO', 'PROVINCIA', 'DISTRITO', 'AGRUPACION_POLITICA'],
        as_index=False
    ).agg({'VOTOS_OBTENIDOS': 'sum'})

    # Rename for clarity
    agg.columns = ['ubigeo', 'region', 'provincia', 'distrito',
                    'organizacion_politica', 'total_votos']

    # Sort
    agg = agg.sort_values(['ubigeo', 'organizacion_politica']).reset_index(drop=True)

    return agg


# === Process Municipal Distrital files ===
muni_files = {
    2006: f'{BASE}/ERM2006_Municipal_Distrital.csv',
    2010: f'{BASE}/ERM2010_Municipal_Distrital.csv',
    2014: f'{BASE}/ERM2014_Municipal_Distrital.csv',
}

for year, path in muni_files.items():
    print(f'\n{"="*60}')
    print(f'  Processing {year} Municipal Distrital')
    print(f'{"="*60}')
    df = pd.read_csv(path, sep=';', encoding='latin-1')
    df = fix_ubigeo_col(df)
    print(f'  Raw rows: {len(df):,}')

    agg = aggregate_to_district(df, year, 'municipal')
    if agg is not None:
        out_path = f'{BASE}/distrital_{year}.csv'
        agg.to_csv(out_path, index=False, encoding='latin-1')
        n_dist = agg['ubigeo'].nunique()
        n_part = agg['organizacion_politica'].nunique()
        print(f'  Districts: {n_dist:,}')
        print(f'  Parties: {n_part}')
        print(f'  Rows: {len(agg):,}')
        print(f'  Saved: {out_path}')


# === Process Regional/Gobernador files ===
reg_files = {
    2006: f'{BASE}/ERM2006_Gobernador_Vicegobernador.csv',
    2010: f'{BASE}/ERM2010_Gobernador_Vicegobernador.csv',
    2014: f'{BASE}/ERM2014_Gobernador_Vicegobernador.csv',
}

for year, path in reg_files.items():
    print(f'\n{"="*60}')
    print(f'  Processing {year} Regional/Gobernador')
    print(f'{"="*60}')
    df = pd.read_csv(path, sep=';', encoding='latin-1')
    df = fix_ubigeo_col(df)
    print(f'  Raw rows: {len(df):,}')

    agg = aggregate_to_district(df, year, 'regional')
    if agg is not None:
        out_path = f'{BASE}/regional_{year}.csv'
        agg.to_csv(out_path, index=False, encoding='latin-1')
        n_dept = agg['region'].nunique()
        n_part = agg['organizacion_politica'].nunique()
        print(f'  Departments: {n_dept}')
        print(f'  Parties: {n_part}')
        print(f'  Rows: {len(agg):,}')
        print(f'  Saved: {out_path}')


# === 2002: check what's in the file ===
print(f'\n{"="*60}')
print(f'  Checking 2002 file')
print(f'{"="*60}')
df02 = pd.read_csv(f'{BASE}/ERM2002_Municipal_Distrital.csv', sep=';', encoding='latin-1')
df02 = fix_ubigeo_col(df02)
print(f'  Raw rows: {len(df02):,}')
print(f'  TIPO_ELECCION values: {df02["TIPO_ELECCION"].value_counts().to_dict()}')

# Check if it has municipal distrital data
if 'MUNICIPAL DISTRITAL' in df02['TIPO_ELECCION'].values:
    agg = aggregate_to_district(df02, 2002, 'municipal')
    if agg is not None:
        out_path = f'{BASE}/distrital_2002.csv'
        agg.to_csv(out_path, index=False, encoding='latin-1')
        print(f'  Saved: {out_path}')
else:
    print('  NOTE: File only contains GOBERNADOR data, NOT municipal distrital.')
    print('  This file is mislabeled on datosabiertos.gob.pe.')
    # Save as regional 2002
    agg = aggregate_to_district(df02, 2002, 'regional')
    if agg is not None:
        out_path = f'{BASE}/regional_2002.csv'
        agg.to_csv(out_path, index=False, encoding='latin-1')
        print(f'  Saved as regional: {out_path}')

print(f'\n{"="*60}')
print('  DONE')
print(f'{"="*60}')
