
clear all
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

*keep if position <= 15

*bysort year position: egen n_pos_year = count(position)
*keep if n_pos_year >= 30



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

*******************************************************
* SAVE ELECTORAL PANEL
*******************************************************

save electoral_panel.dta, replace

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


*Tabla 1

esttab m1_1 m1_2 m1_3 m1_4, ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    label compress nogaps ///
    mtitles("(1)" "(2)" "(3)" "(4)") ///
    title("Table 1. Effect of Position on Winning") ///
    keep(position) ///
    stats(district_fe year_fe org_fe N r2_a, ///
          labels("District FE" "Election year FE" "Organization FE" "Observations" "Adj. R-squared"))


*Tabla 2

esttab m2_1 m2_2 m2_3 m2_4, ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    label compress nogaps ///
    mtitles("(1)" "(2)" "(3)" "(4)") ///
    title("Table 2. Effect of Position by Movement") ///
    keep(position 1.movement 1.movement#c.position) ///
    order(position 1.movement 1.movement#c.position) ///
    varlabels(position "Position" ///
              1.movement "Movement" ///
              1.movement#c.position "Position × Movement") ///
    stats(district_fe year_fe org_fe N r2_a, ///
          labels("District FE" "Election year FE" "Organization FE" "Observations" "Adj. R-squared"))
		  
*Tabla 3

esttab m3_1 m3_2 m3_3 m3_4, ///
    b(%9.3f) se(%9.3f) star(* 0.10 ** 0.05 *** 0.01) ///
    label compress nogaps ///
    mtitles("(1)" "(2)" "(3)" "(4)") ///
    title("Table 3. Nonlinear effects") ///
    keep(position position_sq 1.movement 1.movement#c.position 1.movement#c.position_sq) ///
    order(position position_sq 1.movement 1.movement#c.position 1.movement#c.position_sq) ///
    varlabels(position "Position" ///
              position_sq "Position squared" ///
              1.movement "Movement" ///
              1.movement#c.position "Position × Movement" ///
              1.movement#c.position_sq "Position² × Movement") ///
    stats(district_fe year_fe org_fe N r2_a, ///
          labels("District FE" "Election year FE" "Organization FE" "Observations" "Adj. R-squared"))




* HETEROGENEITY ANALYSIS
* EDUCATION, LITERACY AND AGE (CENSUS 2007)



* 1. BUILD DISTRICT CHARACTERISTICS FROM CENSUS



import excel "censo_distrital_inei.xlsx", ///
firstrow clear

capture destring ubigeo, replace



rename EducaciónInicial EducacionInicial

rename De0a4años age0_4
rename De5a9años age5_9
rename De10a14años age10_14
rename De15a19años age15_19
rename De20a24años age20_24
rename De25a29años age25_29
rename De30a34años age30_34
rename De35a39años age35_39
rename De40a44años age40_44
rename De45a49años age45_49
rename De50a54años age50_54
rename De55a59años age55_59
rename De60a64años age60_64
rename De65a69años age65_69
rename De70a74años age70_74
rename De75a79años age75_79
rename De80a84años age80_84
rename De85a89años age85_89
rename De90a94años age90_94
rename De95a99años age95_99


*TOTAL POPULATION
capture drop population
gen population = ///
age0_4 + age5_9 + age10_14 + age15_19 + ///
age20_24 + age25_29 + age30_34 + age35_39 + ///
age40_44 + age45_49 + age50_54 + age55_59 + ///
age60_64 + age65_69 + age70_74 + age75_79 + ///
age80_84 + age85_89 + age90_94 + age95_99



*EDUCATION
gen educ_superior = ///
SuperiorNoUnivincompleta + ///
SuperiorNoUnivcompleta + ///
SuperiorUnivincompleta + ///
SuperiorUnivcompleta
gen educ_superior_share = ///
educ_superior/population
gen sin_educ_share = ///
SinNivel/population

*LITERACY
gen literacy_share = ///
Sisabeleeryescribir/population
gen illiteracy_share = ///
Nosabeleeryescribir/population

*AGE STRUCTURE
gen young_pop = ///
age0_4 + ///
age5_9 + ///
age10_14 + ///
age15_19 + ///
age20_24 + ///
age25_29
gen young_share = ///
young_pop/population
gen old_pop = ///
age60_64 + ///
age65_69 + ///
age70_74 + ///
age75_79 + ///
age80_84 + ///
age85_89 + ///
age90_94 + ///
age95_99
gen old_share = ///
old_pop/population


*MEDIAN SPLITS
summ educ_superior_share, detail
gen high_educ = educ_superior_share >= r(p50)
summ literacy_share, detail
gen high_literacy = literacy_share >= r(p50)
summ young_share, detail
gen high_young = young_share >= r(p50)
summ old_share, detail
gen high_old = old_share >= r(p50)

*DECILES
xtile educ_decile = educ_superior_share, nq(10)
xtile literacy_decile = literacy_share, nq(10)
xtile young_decile = young_share, nq(10)
xtile old_decile = old_share, nq(10)

*SAVE FOR MERGE
keep ubigeo ///
educ_superior_share ///
sin_educ_share ///
literacy_share ///
illiteracy_share ///
young_share ///
old_share ///
high_educ ///
high_literacy ///
high_young ///
high_old ///
educ_decile ///
literacy_decile ///
young_decile ///
old_decile
save censo_heterogeneity.dta, replace



*******************************************************
* LOAD ELECTORAL PANEL
*******************************************************

use electoral_panel.dta, clear

*******************************************************
* MERGE CENSUS VARIABLES
*******************************************************

merge m:1 ubigeo using censo_heterogeneity.dta

tab _merge

keep if _merge==3
drop _merge


*******************************************************
* TABLE 4
* EDUCATION HETEROGENEITY
*******************************************************

* Continuous measure

reghdfe win ///
    c.position##c.educ_superior_share, ///
    absorb(ubigeo year org_pol_id) ///
    vce(cluster ubigeo)

estimates store edu1

* Above/below median

reghdfe win ///
    c.position##i.high_educ, ///
    absorb(ubigeo year org_pol_id) ///
    vce(cluster ubigeo)

estimates store edu2

esttab edu1 edu2, ///
    b(%9.3f) se(%9.3f) ///
    star(* 0.10 ** 0.05 *** 0.01) ///
    keep(position ///
         educ_superior_share ///
         1.high_educ ///
         c.position#c.educ_superior_share ///
         1.high_educ#c.position) ///
    title("Education Heterogeneity")


*******************************************************
* TABLE 5
* LITERACY HETEROGENEITY
*******************************************************

reghdfe win ///
    c.position##c.literacy_share, ///
    absorb(ubigeo year org_pol_id) ///
    vce(cluster ubigeo)

estimates store lit1

reghdfe win ///
    c.position##i.high_literacy, ///
    absorb(ubigeo year org_pol_id) ///
    vce(cluster ubigeo)

estimates store lit2

esttab lit1 lit2, ///
    b(%9.3f) se(%9.3f) ///
    star(* 0.10 ** 0.05 *** 0.01) ///
    keep(position ///
         literacy_share ///
         1.high_literacy ///
         c.position#c.literacy_share ///
         1.high_literacy#c.position) ///
    title("Literacy Heterogeneity")

*******************************************************
* TABLE 6
* AGE HETEROGENEITY
*******************************************************

reghdfe win ///
    c.position##c.young_share, ///
    absorb(ubigeo year org_pol_id) ///
    vce(cluster ubigeo)

estimates store age1

reghdfe win ///
    c.position##c.old_share, ///
    absorb(ubigeo year org_pol_id) ///
    vce(cluster ubigeo)

estimates store age2

esttab age1 age2, ///
    b(%9.3f) se(%9.3f) ///
    star(* 0.10 ** 0.05 *** 0.01) ///
    keep(position ///
         young_share ///
         old_share ///
         c.position#c.young_share ///
         c.position#c.old_share) ///
    title("Age Heterogeneity")


	
*******************************************************
* SUMMARY TABLE OF INTERACTIONS
*******************************************************

tempname results

postfile `results' ///
str30 interaction ///
double coef ///
double se ///
double t ///
double p ///
using interactions_summary.dta, replace

* Education share
est restore edu1
lincom c.position#c.educ_superior_share
post `results' ///
("Position × Education Share") ///
(r(estimate)) ///
(r(se)) ///
(r(estimate)/r(se)) ///
(r(p))

* High education
est restore edu2
lincom 1.high_educ#c.position
post `results' ///
("Position × High Education") ///
(r(estimate)) ///
(r(se)) ///
(r(estimate)/r(se)) ///
(r(p))

* Literacy share
est restore lit1
lincom c.position#c.literacy_share
post `results' ///
("Position × Literacy Share") ///
(r(estimate)) ///
(r(se)) ///
(r(estimate)/r(se)) ///
(r(p))

* High literacy
est restore lit2
lincom 1.high_literacy#c.position
post `results' ///
("Position × High Literacy") ///
(r(estimate)) ///
(r(se)) ///
(r(estimate)/r(se)) ///
(r(p))

* Young share
est restore age1
lincom c.position#c.young_share
post `results' ///
("Position × Young Share") ///
(r(estimate)) ///
(r(se)) ///
(r(estimate)/r(se)) ///
(r(p))

* Old share
est restore age2
lincom c.position#c.old_share
post `results' ///
("Position × Old Share") ///
(r(estimate)) ///
(r(se)) ///
(r(estimate)/r(se)) ///
(r(p))

postclose `results'

use interactions_summary.dta, clear

gen significance = ""
replace significance = "***" if p < 0.01
replace significance = "**"  if p >= 0.01 & p < 0.05
replace significance = "*"   if p >= 0.05 & p < 0.10

list, clean noobs


*******************************************************
* EXPORT SUMMARY TABLE TO LATEX
*******************************************************

use interactions_summary.dta, clear

gen coef_se = ///
    string(coef,"%9.4f") + ///
    " (" + string(se,"%9.4f") + ")"

gen stars = ""
replace stars = "***" if p < 0.01
replace stars = "**"  if p >= 0.01 & p < 0.05
replace stars = "*"   if p >= 0.05 & p < 0.10

replace coef_se = coef_se + stars

keep interaction coef se t p coef_se

list

ssc install listtex, replace


listtex interaction coef se t p ///
using interactions_summary.tex, ///
replace ///
rstyle(tabular)








