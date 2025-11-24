import pandas as pd          # Manejo de tablas (CSV, Excel)
import geopandas as gpd      # Leer shapefiles y datos geográficos
from openpyxl import load_workbook  # Para editar Excel
from openpyxl.styles import PatternFill, Font, Alignment  # Estilos
from openpyxl.utils import get_column_letter             # Ancho columnas
from openpyxl.worksheet.table import Table, TableStyleInfo  # Tablas Excel


# ===============================
# RUTAS DE ARCHIVOS (INPUT)
# ===============================

# Ruta del dataset de denuncias
ruta_csv = r"C:\era_data2\DATASET_Denuncias_Policiales_Enero 2018 a Octubre 2025.csv"

# Ruta del shapefile distrital del INEI
ruta_shp = r"C:\Users\Usuario\Downloads\Distrital INEI 2023 geogpsperu SuyoPomalia\Distrital INEI 2023 geogpsperu SuyoPomalia.shp"


# ===============================
# 1. CARGAR SHAPEFILE
# ===============================

# Leer shapefile con geopandas
shp = gpd.read_file(ruta_shp)

# Asegurar que el UBIGEO tenga siempre 6 dígitos (ej: "10101" -> "010203")
shp["UBIGEO"] = shp["UBIGEO"].astype(str).str.zfill(6)

# Seleccionar únicamente las columnas relevantes del shapefile
shp_simpl = shp[["UBIGEO", "DEPARTAMEN", "PROVINCIA", "DISTRITO"]].copy()

# Renombrar columnas para que coincidan con el CSV
shp_simpl.columns = ["ubigeo", "departamento_shp", "provincia_shp", "distrito_shp"]


# ===============================
# 2. CARGAR CSV (DENUNCIAS)
# ===============================

# Leer el CSV completo
df = pd.read_csv(ruta_csv, low_memory=False)

# Quitar espacios extra en los nombres de columnas
df.columns = df.columns.str.strip()

# Renombrar columnas para tener nombres cortos y consistentes
df = df.rename(columns={
    "UBIGEO_HECHO": "ubigeo",
    "ANIO": "año",
    "DPTO_HECHO_NEW": "departamento",
    "PROV_HECHO": "provincia",
    "DIST_HECHO": "distrito",
    "P_MODALIDADES": "modalidad",
    "cantidad": "cantidad"
})

# Asegurar formato estándar del UBIGEO en el CSV
df["ubigeo"] = df["ubigeo"].astype(str).str.zfill(6)


# ===============================
# 3. UNIR SHAPEFILE + CSV
# ===============================

# Hacemos un merge por UBIGEO para obtener nombre de dep/prov/dist oficial INEI
df = df.merge(shp_simpl, on="ubigeo", how="left")

# Preferir los nombres del shapefile (oficiales del INEI) si existen
df["departamento"] = df["departamento_shp"].fillna(df["departamento"])
df["provincia"] = df["provincia_shp"].fillna(df["provincia"])
df["distrito"] = df["distrito_shp"].fillna(df["distrito"])

# Eliminar columnas auxiliares
df = df.drop(columns=["departamento_shp", "provincia_shp", "distrito_shp"])

def crear_excel_formateado(nombre_salida, dataframe):
    
    # Guardar Excel inicial
    dataframe.to_excel(nombre_salida, index=False, engine="openpyxl")
    wb = load_workbook(nombre_salida)
    ws = wb.active

    # Formato del encabezado
    header_fill = PatternFill("solid", fgColor="305496")   # Color azul
    header_font = Font(color="FFFFFF", bold=True)          # Letras blancas negrita
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Aplicar formato a cada columna del encabezado
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    # Crear tabla estilo Excel
    data_ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    tabla = Table(displayName="Tabla_Datos", ref=data_ref)
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tabla)

    # Autoajustar anchos de columna
    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        max_len = max(len(str(ws.cell(row=r, column=col).value or "")) for r in range(1, ws.max_row + 1))
        ws.column_dimensions[col_letter].width = max_len + 2

    # Congelar primera fila
    ws.freeze_panes = "A2"

    wb.save(nombre_salida)
    print("✔ Archivo creado:", nombre_salida)


# ===============================
# 4. COLAPSO 1 (ANUAL–DISTRITAL–MODALIDAD)
# ===============================
# ¿Qué hace este colapso?
# 👉 Suma las denuncias por:
#    - año
#    - departamento
#    - provincia
#    - distrito
#    - modalidad
#    - ubigeo
#
# Este es el nivel más detallado.
# Sirve para mapas por año o análisis completos.

colapso1 = (
    df.groupby(["año", "departamento", "provincia", "distrito", "modalidad", "ubigeo"])["cantidad"]
    .sum().reset_index()
)
crear_excel_formateado(r"C:\era_data2\colapso1.xlsx", colapso1)


# ===============================
# 5. COLAPSO 2 (TOTAL ANUAL)
# ===============================
# ¿Qué hace este colapso?
# 👉 Suma todas las denuncias de cada año.
# permite ver tendencias criminales de 2018–2025.

colapso2 = df.groupby("año")["cantidad"].sum().reset_index()
crear_excel_formateado(r"C:\era_data2\colapso2.xlsx", colapso2)


# ===============================
# 6. COLAPSO 3 (AÑO × MODALIDAD)
# ===============================
# ¿Qué hace este colapso?
# 👉 Suma denuncias por año y tipo de delito.
# Permite analizar qué delitos suben/bajan cada año.

colapso3 = df.groupby(["año", "modalidad"])["cantidad"].sum().reset_index()
crear_excel_formateado(r"C:\era_data2\colapso3.xlsx", colapso3)


colapso4 = (
    df.groupby(["año", "departamento", "provincia", "distrito", "ubigeo"])["cantidad"]
    .sum().reset_index()
)
crear_excel_formateado(r"C:\era_data2\colapso4.xlsx", colapso4)


colapso5 = (
    df.groupby(["departamento", "provincia", "distrito", "ubigeo"])["cantidad"]
    .sum().reset_index()
)
crear_excel_formateado(r"C:\era_data2\colapso5.xlsx", colapso5)


# ===============================
# 7. GANADOR CRIMINAL (DELITO PREDOMINANTE)
# ===============================
# ¿Qué hace este colapso?
# 👉 Calcula el delito con MAYOR número de denuncias acumuladas
#    entre 2018–2025 para cada distrito (UBIGEO).
#
# Esta tabla es la base del MAPA CRIMINAL.

temp = df.groupby(["ubigeo", "modalidad", "departamento", "provincia", "distrito"])["cantidad"].sum().reset_index()

# Elegimos el delito con mayor cantidad por distrito
ganador = temp.loc[temp.groupby("ubigeo")["cantidad"].idxmax()]
ganador = ganador.rename(columns={"modalidad": "modalidad_ganadora"})

crear_excel_formateado(r"C:\era_data2\ganador_criminal.xlsx", ganador)

# ===============================
# FINAL
# ===============================
print("\n✔ TODOS LOS COLAPSOS GENERADOS CORRECTAMENTE:")
print("   - colapso1.xlsx → año/provincia/distrito/modalidad")
print("   - colapso2.xlsx → totales por año")
print("   - colapso3.xlsx → año × modalidad")
print("   - ganador_criminal.xlsx → delito predominante por distrito\n")
