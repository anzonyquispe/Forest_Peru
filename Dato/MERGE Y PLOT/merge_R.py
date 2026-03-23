import pandas as pd
import unicodedata

def quitar_tildes(texto):
    if isinstance(texto, str):
        texto = unicodedata.normalize('NFKD', texto)
        texto = texto.encode('ASCII', 'ignore').decode('utf-8')
        return texto
    return texto

def normalizar(texto):
    if isinstance(texto, str):
        texto = quitar_tildes(texto)
        return texto.lower().strip()
    return texto

ruta_excel = r"C:\era_data2\orden_partidos.xlsx"
ruta_res = r"C:\Users\Usuario\Documents\GitHub\Forest_Peru\Scrapping\resultados.xlsx"


df_orden = pd.read_excel(ruta_excel)
df_res = pd.read_excel(ruta_res)


df_orden.rename(columns={"organizacion_politica": "partido"}, inplace=True)
df_res.rename(columns={"organizacion_politica": "partido"}, inplace=True)

# Guardar el texto original
df_orden["partido_original"] = df_orden["partido"]
df_res["partido_original"] = df_res["partido"]


df_orden["partido_norm"] = df_orden["partido"].apply(normalizar)
df_res["partido_norm"] = df_res["partido"].apply(normalizar)

df_orden["region_norm"] = df_orden["region"].apply(normalizar)
df_res["region_norm"] = df_res["region"].apply(normalizar)

# ---------------------------------------------
# 5. LISTA DE EQUIVALENCIAS
# ---------------------------------------------
equivalencias = {
    "alianza para el progreso": "alianza para el progreso",
    "alianza por el progreso": "alianza para el progreso",

    "victoria amazonense": "victoria amazonense",
    "movimiento regional victoria amazonense": "victoria amazonense",

    "avanza pais - partido de integracion social": "avanza pais",
    "avanza pais partido de integracion social": "avanza pais",
    "avanza pais - partidode integracion social": "avanza pais",
    "avanza pais": "avanza pais",

    "frente amplio para el desarrollo del pueblo": "fadep",
    "frente amplio para el desarrollo del pueblo – fadep": "fadep",

    "alianza gobierno unidad y accion": "agua",
    "alianza gobierno unidad y accion agua": "agua",
    "alianza gobierno unidad y accion - agua": "agua",

    "alianza para nuestro desarrollo": "alianza para nuestro desarrollo",
    "alianza por nuestro desarrollo": "alianza para nuestro desarrollo",

    "movimiento regional construyendo": "construyendo",
    "construyendo": "construyendo",

    "movimiento accion nacionalista peruano": "mana",
    "movimiento accionnacionalista peruano": "mana"
}

def aplicar_equiv(x):
    return equivalencias.get(x, x)

df_orden["partido_std"] = df_orden["partido_norm"].apply(aplicar_equiv)
df_res["partido_std"] = df_res["partido_norm"].apply(aplicar_equiv)

# ---------------------------------------------
# 6. MERGE
# ---------------------------------------------
df_merge = pd.merge(
    df_res,
    df_orden,
    on=["partido_std", "region_norm"],
    how="left",
    indicator=True
)

# ---------------------------------------------
# 7. CREAR match_status (sin FutureWarning)
# ---------------------------------------------
df_merge["_merge"] = df_merge["_merge"].astype(str)

df_merge["match_status"] = df_merge["_merge"].map({
    "both": "match",
    "left_only": "no_match",
    "right_only": "no_match"
})

df_final = df_merge[[
    "ubigeo",
    "region_x",
    "provincia",
    "distrito",
    "partido_original_x",
    "total_votos",
    "orden_aparicion",     
    "match_status"
]].rename(columns={
    "region_x": "region",
    "partido_original_x": "organizacion_politica"
})

total_filas = len(df_final)
total_match = (df_final["match_status"] == "match").sum()
total_no_match = (df_final["match_status"] == "no_match").sum()


ruta_salida = r"C:\era_data2\merge_partidos_final.xlsx"
ruta_no_match = r"C:\era_data2\no_match.xlsx"

with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
    sheet_name = "resultados"
    df_final.to_excel(writer, index=False, sheet_name=sheet_name)

    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]

    n_rows, n_cols = df_final.shape

    header_format = workbook.add_format({
        "bold": True,
        "text_wrap": True,
        "valign": "top",
        "align": "center",
        "bg_color": "#D9E1F2",  
        "border": 1
    })


    for col_num, col_name in enumerate(df_final.columns):
        worksheet.write(0, col_num, col_name, header_format)
        max_len = max(
            df_final[col_name].astype(str).map(len).max(),
            len(col_name)
        ) + 2
        worksheet.set_column(col_num, col_num, max_len)


    match_format = workbook.add_format({
        "bg_color": "#C6EFCE",  
        "border": 1
    })
    no_match_format = workbook.add_format({
        "bg_color": "#FFC7CE",  
        "border": 1
    })

    # Columna de match_status
    col_match = df_final.columns.get_loc("match_status")

    # Formato condicional: MATCH
    worksheet.conditional_format(
        1, col_match, n_rows, col_match,
        {
            "type": "cell",
            "criteria": "==",
            "value": '"match"',
            "format": match_format
        }
    )

    # Formato condicional: NO MATCH
    worksheet.conditional_format(
        1, col_match, n_rows, col_match,
        {
            "type": "cell",
            "criteria": "==",
            "value": '"no_match"',
            "format": no_match_format
        }
    )

    # Filtro en la fila de encabezado (arriba)
    worksheet.autofilter(0, 0, n_rows, n_cols - 1)

    # ------------ RESUMEN EN EL EXCEL -------------
    # Lo ponemos a la derecha de la tabla
    summary_col = n_cols + 1  # columna siguiente a las usadas por la tabla

    resumen_title_format = workbook.add_format({
        "bold": True,
        "bg_color": "#FFF2CC",  # amarillito
        "border": 1,
        "align": "center"
    })
    resumen_label_format = workbook.add_format({
        "bold": True,
        "border": 1
    })

    worksheet.write(0, summary_col, "RESUMEN", resumen_title_format)
    worksheet.set_column(summary_col, summary_col + 1, 18)

    worksheet.write(1, summary_col, "Total filas:", resumen_label_format)
    worksheet.write(1, summary_col + 1, total_filas)

    worksheet.write(2, summary_col, "MATCH:", resumen_label_format)
    worksheet.write(2, summary_col + 1, total_match)

    worksheet.write(3, summary_col, "NO MATCH:", resumen_label_format)
    worksheet.write(3, summary_col + 1, total_no_match)
    # ----------------------------------------------

# --- Archivo solo con NO MATCH ---
df_no_match = df_final[df_final["match_status"] == "no_match"]

with pd.ExcelWriter(ruta_no_match, engine="xlsxwriter") as writer:
    sheet_name = "no_match"
    df_no_match.to_excel(writer, index=False, sheet_name=sheet_name)

    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]

    n_rows, n_cols = df_no_match.shape

    # Encabezado
    header_format = workbook.add_format({
        "bold": True,
        "text_wrap": True,
        "valign": "top",
        "align": "center",
        "bg_color": "#F8CBAD",  # naranjita/rojo suave
        "border": 1
    })

    for col_num, col_name in enumerate(df_no_match.columns):
        worksheet.write(0, col_num, col_name, header_format)
        max_len = max(
            df_no_match[col_name].astype(str).map(len).max(),
            len(col_name)
        ) + 2
        worksheet.set_column(col_num, col_num, max_len)

    # Filtro
    worksheet.autofilter(0, 0, n_rows, n_cols - 1)

# ---------------------------------------------
# 10. IMPRESIÓN DE ESTADÍSTICAS
# ---------------------------------------------
print("\nArchivo final guardado en:\n", ruta_salida)
print("Archivo de no match guardado en:\n", ruta_no_match)
print("\n----------------------------------------")
print("RESULTADO DEL MATCH")
print("----------------------------------------")
print(f"Total filas: {total_filas}")
print(f"MATCH: {total_match}")
print(f"NO MATCH: {total_no_match}")
print("----------------------------------------")
