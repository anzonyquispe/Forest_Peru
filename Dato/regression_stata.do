* =========================================================
* 1. CONFIGURATION
* =========================================================

* Definir ruta base
global base_path "/Users/karlavega/Proyectos/Colaborativos/Forest_Peru/Scrapping/data"

* Definir archivos por año
local file_2006 "distrital_2006.csv"
local file_2010 "distrital_2010.csv"
local file_2014 "distrital_2014.csv"
local file_2018 "distrital_2018.csv"
local file_2022 "distrital_2022.csv"


* Definir directorio de salida
global output_dir "."
capture mkdir "$output_dir"


* =========================================================
* 2. HELPERS
* =========================================================

*---------------------------------------------------------
* read_csv_safe(path_file)
* Intenta importar primero con UTF-8 y si falla usa latin1
*---------------------------------------------------------
capture program drop read_csv_safe
program define read_csv_safe
    args path_file

    capture noisily import delimited using "`path_file'", clear varnames(1) encoding(UTF-8)
    if _rc {
        import delimited using "`path_file'", clear varnames(1) encoding(ISO-8859-1)
    }
end


*---------------------------------------------------------
* stars(p)
* Devuelve estrellas según p-value
* Uso:
*     local s = stars(0.034)
*---------------------------------------------------------
capture program drop stars
program define stars, rclass
    args p

    if `p' < 0.01 {
        return local stars "***"
    }
    else if `p' < 0.05 {
        return local stars "**"
    }
    else if `p' < 0.10 {
        return local stars "*"
    }
    else {
        return local stars ""
    }
end


*---------------------------------------------------------
* get_cluster_count(df, cluster_var="ubigeo")
* Cuenta número de clusters únicos
* Uso:
*     get_cluster_count ubigeo
*     local n_clusters = r(nclusters)
*---------------------------------------------------------
capture program drop get_cluster_count
program define get_cluster_count, rclass
    args cluster_var

    quietly egen __tag_cluster__ = tag(`cluster_var')
    quietly count if __tag_cluster__ == 1
    return scalar nclusters = r(N)
    drop __tag_cluster__
end


*---------------------------------------------------------
* fit_ols_cluster(formula, data, cluster_var="ubigeo")
* En Stata esto se hace directamente con regress
*
* Ejemplo:
*     regress y x1 x2 i.year, vce(cluster ubigeo)
*
* Si quieres una versión programada:
*---------------------------------------------------------
capture program drop fit_ols_cluster
program define fit_ols_cluster, eclass
    syntax varlist(min=2 fv) [if] [in], CLuster(varname)

    tokenize `varlist'
    local depvar `1'
    macro shift
    local indepvars `*'

    regress `depvar' `indepvars' `if' `in', vce(cluster `cluster')
end


*---------------------------------------------------------
* format_cell(model, varname, decimals=4)
* Formatea coeficiente + estrellas + error estándar
*
* OJO: en Stata esto funciona después de una regresión activa
* o tras estimates restore nombre_modelo
*
* Uso:
*     regress y x1 x2, vce(cluster ubigeo)
*     format_cell x1 4
*     local celda = r(cell)
*---------------------------------------------------------
capture program drop format_cell
program define format_cell, rclass
    args varname decimals

    if "`decimals'" == "" local decimals = 4

    capture scalar __b = _b[`varname']
    if _rc {
        return local cell ""
        exit
    }

    scalar __se = _se[`varname']
    scalar __t  = __b / __se
    scalar __p  = 2 * ttail(e(df_r), abs(__t))

    quietly stars __p
    local st = r(stars)

    local coef_str : display %9.`decimals'f __b
    local se_str   : display %9.`decimals'f __se

    return local cell "`coef_str'`st' (`se_str')"
end


* =========================================================
* 3. LOAD AND STACK DATA
* =========================================================

tempfile master
save `master', emptyok replace

foreach pair in ///
    "2006 distrital_2006.csv" ///
    "2010 distrital_2010.csv" ///
    "2014 distrital_2014.csv" ///
    "2018 distrital_2018.csv" ///
    "2022 distrital_2022.csv" {

    tokenize `"`pair'"'
    local year `1'
    local filename `2'

    di "$base_path/`filename'"

    * Leer CSV
    read_csv_safe "$base_path/`filename'"

    * Pasar nombres de variables a minúsculas
    rename *, lower

    * Crear año
    gen year = `year'

    * Crear orden original dentro del archivo
    gen original_file_order = _n

    * Apilar
    append using `master'
    save `master', replace
}

use `master', clear


* =========================================================
* 4. BASIC CLEANING
* =========================================================

foreach col in region provincia distrito organizacion_politica {
    capture confirm variable `col'
    if !_rc {
        replace `col' = strtrim(`col')
    }
}

* Convertir variables a numéricas si vinieron como string
capture destring ubigeo, replace force
capture destring total_votos, replace force
capture destring year, replace force

* Eliminar observaciones con missing en variables clave
drop if missing(ubigeo)
drop if missing(organizacion_politica)
drop if missing(year)
drop if missing(total_votos)


* =========================================================
* 5. CONSTRUCT POSITION FROM ORIGINAL ORDER WITHIN DISTRICT-YEAR
* =========================================================

sort year ubigeo original_file_order
by year ubigeo: gen position = _n


* =========================================================
* 6. DEPENDENT VARIABLE: WIN
* =========================================================

bysort year ubigeo: egen max_votes = max(total_votos)
gen win = (total_votos == max_votes)


* =========================================================
* 7. MOVEMENT / PARTY INDICATOR
* =========================================================

gen org_upper = upper(organizacion_politica)

gen movement = strpos(org_upper, "MOVIMIENTO") > 0
gen org_type = cond(movement == 1, "Movement", "Party")

* position squared
gen position_sq = position^2


* =========================================================
* 8. SAMPLE RESTRICTIONS
* =========================================================

keep if position <= 15

bysort year position: egen n_pos_year = count(position)
keep if n_pos_year >= 30



* =========================================================
* 9. SUMMARY STATS FOR TABLE FOOTERS
* =========================================================

* Media de la variable dependiente (win)
summ win, meanonly
local depvar_mean = r(mean)

* Número de observaciones
count
local n_obs = r(N)

* Número de clusters (ubigeo)
get_cluster_count ubigeo
local n_clusters = r(nclusters)

* Mostrar resultados
display "Final sample size: " `n_obs'
display "Dependent variable mean: " %9.4f `depvar_mean'
display "Number of clusters: " `n_clusters'

capture drop org_pol_id
encode organizacion_politica, gen(org_pol_id)

* =========================================================
* 10. MODEL SETS
* =========================================================

* -------------------------------
* TABLE 1: No interaction
* -------------------------------
 

* (1) No FE
reghdfe win position, vce(cluster ubigeo)
estimates store m1_1

* (2) District FE
reghdfe win position, absorb(ubigeo) vce(cluster ubigeo)
estimates store m1_2

* (3) District FE + Election year FE
reghdfe win position, absorb(ubigeo year) vce(cluster ubigeo)
estimates store m1_3

* (4) District FE + Election year FE + Organization FE
reghdfe win position, absorb(ubigeo year org_pol_id) vce(cluster ubigeo)
estimates store m1_4


* -------------------------------
* TABLE 2: Interaction with movement
* -------------------------------

* (1)
reghdfe win c.position##i.movement, vce(cluster ubigeo)
estimates store m2_1

* (2)
reghdfe win c.position##i.movement, absorb(ubigeo) vce(cluster ubigeo)
estimates store m2_2

* (3)
reghdfe win c.position##i.movement, absorb(ubigeo year) vce(cluster ubigeo)
estimates store m2_3

* (4)
reghdfe win c.position##i.movement, absorb(ubigeo year org_pol_id) vce(cluster ubigeo)
estimates store m2_4


* -------------------------------
* TABLE 3: Quadratic + interaction
* -------------------------------

* (1)
reghdfe win c.position##i.movement c.position_sq##i.movement, vce(cluster ubigeo)
estimates store m3_1

* (2)
reghdfe win c.position##i.movement c.position_sq##i.movement, absorb(ubigeo) vce(cluster ubigeo)
estimates store m3_2

* (3)
reghdfe win c.position##i.movement c.position_sq##i.movement, absorb(ubigeo year) vce(cluster ubigeo)
estimates store m3_3

* (4)
reghdfe win c.position##i.movement c.position_sq##i.movement, absorb(ubigeo year org_pol_id) vce(cluster ubigeo)
estimates store m3_4






	*Capturar desinteres politico*
	

	*Corrupcion - distrital
	*GDP - nightlights - data armonized nightlights
	*Educación
	*Fenomenos naturales - shock - distrital 













