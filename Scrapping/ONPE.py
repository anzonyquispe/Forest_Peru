# %% [markdown]
# # ONPE ERM2022 – Scraping Actas por Ubigeo
# Notebook para extraer **organización política → votos (y porcentaje)** por **Ubigeo** desde:
# `https://resultadoshistorico.onpe.gob.pe/ERM2022/EleccionesMunicipales/RePro`.

# %%

import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# %%
INPUT_CSV = "ubigeos_distritos.csv"
OUTPUT_CSV = "erm2022_distrital_por_ubigeo.csv"

BASE_URL = (
    "https://resultadoshistorico.onpe.gob.pe/ERM2022/EleccionesMunicipales/"
    "RePro/{dep}/{prov}/{dist}"
)

SLEEP_SECONDS = 0.5   # pausa entre distritos


# %%
def crear_driver(headless=True):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def limpiar_numero(texto: str):
    if not texto:
        return None
    t = (
        texto.replace(",", "")
             .replace(".", "")
             .replace(" ", "")
    )
    return int(t) if t.isdigit() else None


def extraer_tabla_partidos(driver):
    """
    Tabla con cabecera:
    Organización política | Total | % Votos válidos | % Votos emitidos
    Devuelve lista de tuplas (organizacion_politica, total_votos)
    """

    wait = WebDriverWait(driver, 20)

    # tabla que tenga un <th> con "Organización política"
    tabla = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//table[.//th[contains(translate(., "
                " 'óÓáÁéÉíÍúÚñÑ', 'oOaAeEiIuUnN'),"
                " 'ORGANIZACION POLITICA')]]"
            )
        )
    )

    # centrar
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});", tabla
    )

    filas = tabla.find_elements(By.XPATH, ".//tbody/tr")
    resultados = []

    for tr in filas:
        celdas = tr.find_elements(By.TAG_NAME, "td")
        if len(celdas) < 2:
            continue

        org = celdas[0].text.strip()
        total_txt = celdas[1].text.strip()

        if not org:
            continue


        up = org.upper()
        if any(x in up for x in [
            "TOTAL DE VOTOS VÁLIDOS",
            "TOTAL DE VOTOS EMITIDOS",
            "VOTOS EN BLANCO",
            "VOTOS NULOS",
        ]):
            continue

        total = limpiar_numero(total_txt)
        if total is None:
            continue

        resultados.append((org, total))

    return resultados


def obtener_resultados_distrito(driver, cod_dep, cod_prov, ubigeo):
    url = BASE_URL.format(dep=cod_dep, prov=cod_prov, dist=ubigeo)
    print(f"  > {url}")

    try:
        driver.get(url)
    except Exception as e:
        print(f"[ERROR] al abrir {url}: {e}")
        return []

    try:
        filas = extraer_tabla_partidos(driver)
    except Exception as e:
        print(f"[ERROR] al leer tabla en {url}: {e}")
        return []

    resultados = []
    for org, total in filas:
        resultados.append({
            "url": url,
            "organizacion_politica": org,
            "total_votos": total,
        })
    return resultados

# %%
def main():
    ubigeos = pd.read_csv(INPUT_CSV, dtype=str)

    columnas_necesarias = [
        "ubigeo", "cod_dep", "cod_prov",
        "departamento", "provincia", "distrito"
    ]
    for col in columnas_necesarias:
        if col not in ubigeos.columns:
            raise ValueError(f"Falta la columna '{col}' en {INPUT_CSV}")

    driver = crear_driver(headless=True)
    registros = []

    try:
        for _, row in ubigeos.iterrows():
            ubigeo = row["ubigeo"]
            cod_dep = row["cod_dep"]
            cod_prov = row["cod_prov"]
            departamento = row["departamento"]
            provincia = row["provincia"]
            distrito = row["distrito"]

            print(f"\n{ubigeo} - {departamento}/{provincia}/{distrito}")

            resultados = obtener_resultados_distrito(
                driver, cod_dep, cod_prov, ubigeo
            )

            for r in resultados:
                registros.append({
                    "ubigeo": ubigeo,
                    "departamento": departamento,
                    "provincia": provincia,
                    "distrito": distrito,
                    "organizacion_politica": r["organizacion_politica"],
                    "total_votos": r["total_votos"],
                    "url_origen": r["url"],
                })

            time.sleep(SLEEP_SECONDS)

    finally:
        driver.quit()

    df = pd.DataFrame(registros)
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"\nListo -> {OUTPUT_CSV}")


if __name__ == "__main__":
    main()


