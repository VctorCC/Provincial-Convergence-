
clear all


*WARMING: CHANGE  THE REPOSITORIES!!*
cd "D:\Doctorado\Convergencia\DATOS"

*WARMING II: with the data files, use execute "Principal.dta" from line 770*

*=============================================================================
*                       2000
*=============================================================================
*________PRODUCCIÓN
* Importamos datos excel provinciales y los unimos en una misma carpeta
local provincia "acoruna alava albacete alicante almeria asturias avila badajoz balears barcelona bizkaia burgos caceres cadiz cantabria castellon ceuta ciudadreal cordoba cuenca gipuzkoa girona granada guadalajara huelva huesca jaen larioja laspalmas leon lleida lugo madrid malaga  melilla murcia navarra ourense palencia pontevedra salamanca segovia sevilla soria tarragona tenerife teruel toledo valencia valladolid zamora zaragoza"
foreach p of local provincia{
import excel "Provincias XXI\sh_p_`p'.xls", sheet("Tabla_1") cellrange(B9:AX19) firstrow clear
drop E G I K M O Q S U W Y AA AC AE AG AI AK AM AO AQ AS AU AW
drop if D==""
gen Provincia = "`p'"
tempfile `p'
save ``p''
}
local provincia "acoruna alava albacete alicante almeria asturias avila badajoz balears barcelona bizkaia burgos caceres cadiz cantabria castellon ceuta ciudadreal cordoba cuenca gipuzkoa girona granada guadalajara huelva huesca jaen larioja laspalmas leon lleida lugo madrid malaga  melilla murcia navarra ourense palencia pontevedra salamanca segovia sevilla soria tarragona tenerife teruel toledo valencia valladolid zamora"
foreach p of local provincia{
append using ``p''
}
rename D sector

*Nombramos las variables por el año correspondiente
foreach v of varlist F H J L N P R T V X Z AB AD AF AH AJ AL AN AP AR AT AV {
   local x : variable label `v'
   rename `v' a`x'
}
rename AX a2022 
 
tempfile pib      // Guardamos archivo temporal
save `pib'

* ___________EMPLEO
*Importamos datos excel provinciales y los unimos en una misma carpeta
local provincia "acoruna alava albacete alicante almeria asturias avila badajoz balears barcelona bizkaia burgos caceres cadiz cantabria castellon ceuta ciudadreal cordoba cuenca gipuzkoa girona granada guadalajara huelva huesca jaen larioja laspalmas leon lleida lugo madrid malaga  melilla murcia navarra ourense palencia pontevedra salamanca segovia sevilla soria tarragona tenerife teruel toledo valencia valladolid zamora zaragoza"
foreach p of local provincia{

import excel "Provincias XXI\sh_p_`p'.xls", sheet("Tabla_2") cellrange(B9:AX19) firstrow clear
drop E G I K M O Q S U W Y AA AC AE AG AI AK AM AO AQ AS AU AW
drop if D==""
gen Provincia = "`p'"
tempfile `p'
save ``p''
}
local provincia "acoruna alava albacete alicante almeria asturias avila badajoz balears barcelona bizkaia burgos caceres cadiz cantabria castellon ceuta ciudadreal cordoba cuenca gipuzkoa girona granada guadalajara huelva huesca jaen larioja laspalmas leon lleida lugo madrid malaga  melilla murcia navarra ourense palencia pontevedra salamanca segovia sevilla soria tarragona tenerife teruel toledo valencia valladolid zamora"
foreach p of local provincia{
append using ``p''
}
*Nombramos las variables por el año correspondiente
rename D sector
foreach v of varlist F H J L N P R T V X Z AB AD AF AH AJ AL AN AP AR AT AV {
   local x : variable label `v'
   rename `v' E`x'
}
rename AX E2022 

* Unimos documento PRODUCCION Y EMPLEO
merge m:m Provincia sector using `pib'

bysort Provincia sector: gen n=_n
drop if n==2

*Le damos la forma
reshape long E a, i(Provincia sector) j(año)

 drop DivisionesNACErev2 _merge n
 

sort Provincia
encode Provincia, generate(X)
label values X
encode sector, generate(S)
label values S

gen Sect= S
drop S
*replace a= a*166.386/1000 //miles de euros en millones pesetas 



sort año Provincia S

*tab S if sector= strpos( sector, "PERSONAS") ==> 8
*tab S if strpos( sector, "BRUTO") ==> 9
*tab S if strpos( sector, "bruto") ==> 10


by año Provincia: replace E= E[_n-1] if S== 9
by año Provincia: replace E= E[_n-2] if S== 10
drop if S==8

*Renombramos las Provincias
replace Provincia= "A Coruña" if Provincia== "acoruna"
replace Provincia= "Álava" if Provincia== "alava"
replace Provincia= "Albacete" if Provincia=="albacete"
replace Provincia= "Alicante" if Provincia== "alicante"
replace Provincia= "Almería" if Provincia== "almeria"
replace Provincia= "Asturias" if Provincia=="asturias"
replace Provincia= "Ávila" if Provincia== "avila"

replace Provincia= "Badajoz" if Provincia== "badajoz"
replace Provincia= "Baleares" if Provincia== "balears"
replace Provincia= "Barcelona" if Provincia== "barcelona"
replace Provincia= "Burgos" if Provincia== "burgos"

replace Provincia= "Cáceres" if Provincia== "caceres"
replace Provincia= "Cádiz" if Provincia== "cadiz"
replace Provincia= "Castellón" if Provincia== "castellon"
replace Provincia= "Cantabria" if Provincia== "cantabria"
replace Provincia= "Ceuta" if Provincia== "ceuta"
replace Provincia= "Ciudad Real" if Provincia== "ciudadreal"
replace Provincia= "Córdoba" if Provincia== "cordoba"
replace Provincia= "Cuenca" if Provincia== "cuenca"

replace Provincia= "Girona" if Provincia== "girona"
replace Provincia= "Granada" if Provincia== "granada"
replace Provincia= "Guadalajara" if Provincia== "guadalajara"
replace Provincia= "Guipuzcoa" if Provincia== "gipuzkoa"

replace Provincia= "Huelva" if Provincia== "huelva"
replace Provincia= "Huesca" if Provincia== "huesca"

replace Provincia= "Jaén" if Provincia== "jaen"

replace Provincia= "La Rioja" if Provincia== "larioja"
replace Provincia= "Las Palmas" if Provincia== "laspalmas"
replace Provincia= "León" if Provincia== "leon"
replace Provincia= "Lleida" if Provincia== "lleida"
replace Provincia= "Lugo" if Provincia== "lugo"

replace Provincia= "Málaga" if Provincia== "malaga"
replace Provincia= "Madrid" if Provincia== "madrid"
replace Provincia= "Melilla" if Provincia== "melilla"
replace Provincia= "Murcia" if Provincia== "murcia"

replace Provincia= "Navarra" if Provincia== "navarra"

replace Provincia= "Ourense" if Provincia== "ourense"

replace Provincia= "Palencia" if Provincia== "palencia"
replace Provincia= "Pontevedra" if Provincia== "pontevedra"

replace Provincia= "Salamanca" if Provincia== "salamanca"
replace Provincia= "Santa Cruz de Tenerife" if Provincia== "tenerife"
replace Provincia= "Segovia" if Provincia== "segovia"
replace Provincia= "Sevilla" if Provincia== "sevilla"
replace Provincia= "Soria" if Provincia== "soria"

replace Provincia= "Tarragona" if Provincia== "tarragona"
replace Provincia= "Teruel" if Provincia== "teruel"
replace Provincia= "Toledo" if Provincia== "toledo"
				
replace Provincia= "Valencia" if Provincia== "valencia"
replace Provincia= "Valladolid" if Provincia== "valladolid"
replace Provincia= "Vizcaya" if Provincia== "bizkaia"

replace Provincia= "Zamora" if Provincia=="zamora"
replace Provincia= "Zaragoza" if Provincia=="zaragoza"


rename a VAB  // miles de euros
rename E Empl // miles de personas
replace Empl = Empl*1000   

save "Dosmil.dta", replace


/*================================================================================
*                                Población
===============================================================================*/

import excel "D:\Doctorado\Convergencia\DATOS\Población\Poblacion.xlsx", sheet("Población") firstrow clear
keep A *enero*
																		
drop if A==""
reshape long deenerode, i(A) j(año)
rename deenerode Poblacion
rename A Provincia

drop if año<2000

tempfile pob
save `pob'

use "Dosmil.dta", clear

merge m:1 Provincia año using `pob'

rename X Prov

save "Principal.dta", replace
 

/*================================================================================
*                                Infraestructuras
===============================================================================*/

import delimited "D:\Doctorado\Convergencia\DATOS\K privado-publico\DE_2025_BBDD_Stock_provincias-infraestructuras.csv", delimiter(";") varnames(1) clear 

forvalues i= 6/41 {      //borramos hasta 1999
drop v`i'
}

forvalues i= 42/63 {
local x : variable label v`i'
rename v`i' a`x'
}

rename lugar Provincia

replace Provincia= "Álava" if Provincia== "Araba/Álava"
replace Provincia= "Alicante" if Provincia== "Alacant/Alicante"
replace Provincia= "Baleares" if Provincia== "Illes Balears"
replace Provincia= "Castellón" if Provincia== "Castelló/Castellón"
replace Provincia= "Guipuzcoa" if Provincia== "Gipuzkoa"			
replace Provincia= "Valencia" if Provincia== "València/Valencia"
replace Provincia= "Vizcaya" if Provincia== "Bizkaia"

replace Provincia= "Total Nacional" if Provincia== "España"

drop if strpos( Provincia, "Andalucía")
drop if strpos( Provincia, "Aragón")
drop if strpos( Provincia, "Canarias")
drop if strpos( Provincia, "Castilla")
drop if strpos( Provincia, "Cataluña")
drop if strpos( Provincia, "Ciudades")
drop if strpos( Provincia, "Comuni")
drop if strpos( Provincia, "Extremadura")
drop if strpos( Provincia, "Galicia")
drop if strpos( Provincia, "País Vasco")
drop if strpos( Provincia, "Principado")
drop if strpos( Provincia, "Región")

*-----
encode variable, generate(var)
/* Inversión bruta 
Stock de capital neto riqueza |     
  Stock de capital productivo  */
keep if var==2

encode unidad, generate(und)
/* Miles de euros constantes base 2020  |    
              Miles de euros corrientes |     
Tasa de variación anual real (Tornqvist |     
           Índice de volumen 2020 = 100 */

keep if und==1        //  Miles de euros constantes base 2020

keep if rama== "0. Total" 
*Energia, transporte, administraciones públicas sí que tienen (el resto 0)

drop rama variable unidad var und
*-----

gen i= _n
reshape long a, i(i) j(año)

destring a, generate(Inf) ignore(",") force
drop a i

bysort año activo Provincia: gen n= _n
drop if n==2   //Borramos los repetidos de Cantabria, La Rioja y Baleares
drop n

encode activo, generate(act)
/* 1  Infraestructuras públicas 
   2  Infr. viarias 
   3  Infr. hidráulicas públicas 
   4  Infr. ferroviarias 
   5  Infr. aeroportuarias 
   6  Infr. portuarias 
   7  Infr. urbanas de CC.LL.     */

drop activo
reshape wide Inf, i(año Provincia) j(act)

save "Infraestructuras.dta", replace


/*================================================================================
*                             Kapital Privado
===============================================================================*/

import delimited "D:\Doctorado\Convergencia\DATOS\K privado-publico\DE_2025_BBDD_Stock_provincias-activosbasicos.csv", delimiter(";") varnames(1) clear

forvalues i= 6/41 {      //borramos hasta 1999
drop v`i'
}

forvalues i= 42/63 {
local x : variable label v`i'
rename v`i' a`x'
}

rename lugar Provincia

replace Provincia= "Álava" if Provincia== "Araba/Álava"
replace Provincia= "Alicante" if Provincia== "Alacant/Alicante"

replace Provincia= "Baleares" if Provincia== "Illes Balears"

replace Provincia= "Castellón" if Provincia== "Castelló/Castellón"

replace Provincia= "Guipuzcoa" if Provincia== "Gipuzkoa"
				
replace Provincia= "Valencia" if Provincia== "València/Valencia"
replace Provincia= "Vizcaya" if Provincia== "Bizkaia"

replace Provincia= "Total Nacional" if Provincia== "España"

drop if strpos( Provincia, "Andalucía")
drop if strpos( Provincia, "Aragón")
drop if strpos( Provincia, "Canarias")
drop if strpos( Provincia, "Castilla")
drop if strpos( Provincia, "Cataluña")
drop if strpos( Provincia, "Ciudades")
drop if strpos( Provincia, "Comuni")
drop if strpos( Provincia, "Extremadura")
drop if strpos( Provincia, "Galicia")
drop if strpos( Provincia, "País Vasco")
drop if strpos( Provincia, "Principado")
drop if strpos( Provincia, "Región")

keep if activo== "0. Total"
gen i= _n

reshape long a, i(i) j(año)

bysort año variable unidad activo rama Provincia: gen n= _n
drop if n==2   //Borramos los repetidos de Cantabria, La Rioja y Baleares

destring a, generate(K) ignore(",") force
drop a n
*-----
encode variable, generate(var)
/* Inversión bruta 
Stock de capital neto riqueza |     
  Stock de capital productivo  */
keep if var==2

encode unidad, generate(und)
/* Miles de euros constantes base 2020  |    
              Miles de euros corrientes |     
Tasa de variación anual real (Tornqvist |     
           Índice de volumen 2020 = 100 */

keep if und==1        //  Miles de euros constantes base 2020

drop variable var unidad und activo

 /* 
                               1 Total |     
2 Agricultura, ganadería, silvicultu.. |     
                      3 Otros servicios |     
                            4 Industria |    
                              5 Energía |     
                      6    Manufacturas |     
                        7  Construcción |     
   8  Comercio, transporte y hostelería |     
             9    Comercio y reparación |    
                        10   Transporte |   
                        11   Hostelería |    
        12 Información y comunicaciones |
13 Actividades financieras y de seguros |     
           14 Actividades inmobiliarias |     
           15 Actividades profesionales |    
16 Administración pública, sanidad y .. |     
            17   Administración pública |     
                 18   Educación pública |     
                   19 Sanidad pública |   
				   
  1: Industrias Manufactureras
* 2: Servicios financieros, inmobiliarios, profesionales, admin
* 3: Servicios: Admin Pub, Sanidad, educacion ,artistas y entretenimiento y otros
* 4: Sector Primario
* 5: Servicios: Comercio; rep. de vehículos, transporte, hostelería, Comunicacion
* 6: COnstrucción
* 7: Industria               */  

encode rama, generate(Sect)
label values Sect

 drop if Sect ==9 | Sect ==10 | Sect ==11| Sect ==5
 replace Sect= 9 if Sect== 1      // Total
 replace Sect= 1 if Sect== 6      // Manufacturas
 replace Sect= 6 if Sect==  7    // Construcción
 replace Sect= 7 if Sect==  4    // Industria
 replace Sect= 4 if Sect== 2     // Agricultura

 //Comercio, Comuniación
 bysort Provincia año:  egen Servicios1 = total(K) if Sect== 8 | Sect == 12
 replace K= Servicios1 if Sect== 8 | Sect == 12
 drop if Sect== 12
 replace Sect= 5 if Sect==  8     
  
 // Admin, Otros Admin, Otros 
 bysort Provincia año:  egen Servicios2 = total(K) if Sect== 16 | Sect == 3
 replace K= Servicios2 if Sect== 16 | Sect == 3  
 drop if Sect== 16
 
 // Servicios financieros, inmobiliarios
 bysort Provincia año:  egen Servicios3 = total(K) if Sect== 13 | Sect == 14 | Sect == 15
 replace K= Servicios3 if Sect== 13 | Sect == 14 | Sect == 15
 drop if Sect> 13
 replace Sect= 2 if Sect==13
 
 drop Servicios1 Servicios2 Servicios3 i rama
  
save "Kapital.dta", replace
 
/*================================================================================
*                               K HUMANO
===============================================================================*/
import delimited "D:\Doctorado\Convergencia\DATOS\K Humano\21-23.csv", varnames(1) clear 

*Formación por Tipo de Actividad
rename provincias Provincia
replace Provincia= ïtotalnacional if Provincia=="" & comunidadesyciudadesautãnomas==""
drop if Provincia==""
*replace Provincias= comunidadesyciudadesautãnomas if Provincias==""
drop ïtotalnacional comunidadesyciudadesautãnomas

replace Provincia= "A Coruña" if Provincia== "15 CoruÃ±a, A"
replace Provincia= "Álava" if Provincia== "01 Araba/Ãlava"
replace Provincia= "Albacete" if Provincia=="02 Albacete"
replace Provincia= "Alicante" if Provincia== "03 Alicante/Alacant"
replace Provincia= "Almería" if Provincia== "04 AlmerÃ­a"
replace Provincia= "Asturias" if Provincia=="33 Asturias"
replace Provincia= "Ávila" if Provincia== "05 Ãvila"

replace Provincia= "Badajoz" if Provincia== "06 Badajoz"
replace Provincia= "Baleares" if Provincia== "07 Balears, Illes"
replace Provincia= "Barcelona" if Provincia== "08 Barcelona"
replace Provincia= "Burgos" if Provincia== "09 Burgos"

replace Provincia= "Cáceres" if Provincia== "10 CÃ¡ceres"
replace Provincia= "Cádiz" if Provincia== "11 CÃ¡diz"
replace Provincia= "Castellón" if Provincia== "12 CastellÃ³n/CastellÃ³"
replace Provincia= "Cantabria" if Provincia== "39 Cantabria"
replace Provincia= "Ceuta" if Provincia== "51 Ceuta"
replace Provincia= "Ciudad Real" if Provincia== "13 Ciudad Real"
replace Provincia= "Córdoba" if Provincia== "14 CÃ³rdoba"
replace Provincia= "Cuenca" if Provincia== "16 Cuenca"

replace Provincia= "Girona" if Provincia== "17 Girona"
replace Provincia= "Granada" if Provincia== "18 Granada"
replace Provincia= "Guadalajara" if Provincia== "19 Guadalajara"
replace Provincia= "Guipuzcoa" if Provincia== "20 Gipuzkoa"

replace Provincia= "Huelva" if Provincia== "21 Huelva"
replace Provincia= "Huesca" if Provincia== "22 Huesca"

replace Provincia= "Jaén" if Provincia== "23 JaÃ©n"

replace Provincia= "La Rioja" if Provincia== "26 Rioja, La"
replace Provincia= "Las Palmas" if Provincia== "35 Palmas, Las"
replace Provincia= "León" if Provincia== "24 LeÃ³n"
replace Provincia= "Lleida" if Provincia== "25 Lleida"
replace Provincia= "Lugo" if Provincia== "27 Lugo"

replace Provincia= "Málaga" if Provincia== "29 MÃ¡laga"
replace Provincia= "Madrid" if Provincia== "28 Madrid"
replace Provincia= "Melilla" if Provincia== "52 Melilla"
replace Provincia= "Murcia" if Provincia== "30 Murcia"

replace Provincia= "Navarra" if Provincia== "31 Navarra"

replace Provincia= "Ourense" if Provincia== "32 Ourense"

replace Provincia= "Palencia" if Provincia== "34 Palencia"
replace Provincia= "Pontevedra" if Provincia== "36 Pontevedra"

replace Provincia= "Salamanca" if Provincia== "37 Salamanca"
replace Provincia= "Santa Cruz de Tenerife" if Provincia== "38 Santa Cruz de Tenerife"
replace Provincia= "Segovia" if Provincia== "40 Segovia"
replace Provincia= "Sevilla" if Provincia== "41 Sevilla"
replace Provincia= "Soria" if Provincia== "42 Soria"

replace Provincia= "Tarragona" if Provincia== "43 Tarragona"
replace Provincia= "Teruel" if Provincia== "44 Teruel"
replace Provincia= "Toledo" if Provincia== "45 Toledo"
				
replace Provincia= "Valencia" if Provincia== "46 Valencia/ValÃ¨ncia"
replace Provincia= "Valladolid" if Provincia== "47 Valladolid"
replace Provincia= "Vizcaya" if Provincia== "48 Bizkaia"

replace Provincia= "Zamora" if Provincia=="49 Zamora"
replace Provincia= "Zaragoza" if Provincia=="50 Zaragoza"

rename niveldeestudios edu
rename periodo año
rename total h
destring h,replace ignore(".") 


replace edu= "_Phd" if strpos( edu, "Doctorado")
replace edu= "_Prim" if strpos( edu, "primaria")
replace edu= "_fp" if strpos( edu, "formaciÃ³n profesional")    //EnseÃ±anzas de formaciÃ³n profesional
replace edu= "_GradUinf" if strpos( edu, "hasta 240")          //Grados universitarios de hasta 240 cr
replace edu= "_GradUsup" if strpos( edu, "mÃ¡s de 240")        //Grados universitarios de mÃ¡s de 240
replace edu= "_Master" if strpos( edu, "steres")
replace edu= "_Secund1" if strpos( edu, "Primera etapa")
replace edu= "_Secund2" if strpos( edu, "orientaciÃ³n general")       //Segunda etapa con orientación general
replace edu= "_Secund3" if strpos( edu, "orientaciÃ³n profesional")   //Segunda etapa con orientación profesional
replace edu= "_No" if strpos( edu, "Sin estudios")

reshape wide h, i( Provincia sexo relaciãnconlaactividad año ) j(edu) string

/*relaciãnconlaactividad
                             Estudiante 1 
                              Ocupado/a 2      
         Otra situaciÃ³n de inactividad 3      
                               Parado/a 4        
    Perceptor/a pensiÃ³n de incapacidad 5       
Perceptor/a pensiÃ³n de jubilaciÃ³n, .. 6       
                                  Total 7 */

encode relaciãnconlaactividad, gen(re)
label values re
keep if re == 2  //dejamos solo los ocupados
drop relaciãnconlaactividad re

keep if sexo == "Total"
drop sexo

save "Khumano.dta", replace
*Valores absolutos.Activos por provincia y nivel de formación alcanzado.										
import excel "D:\Doctorado\Convergencia\DATOS\K Humano\Ocupados 2002.xlsx", sheet("tabla-0") cellrange(A7:I59) firstrow clear
rename A Provincia
gen año= 2002

replace Provincia= "A Coruña" if Provincia== "Coruña (A)"
replace Provincia= "Baleares" if Provincia== "Balears (Illes)"
replace Provincia= "Castellón" if Provincia== "Castellón de la Plana"
replace Provincia= "Guipuzcoa" if Provincia== "Guipúzcoa"
replace Provincia= "Las Palmas" if Provincia== "Palmas (Las)"
replace Provincia= "La Rioja" if Provincia== "Rioja (La)"
replace Provincia= "Total Nacional" if Provincia== "Total"
*Ceuta y Melilla 

/*Analfabetos 
Educación primaria 
Educación secundaria.Primera etapa y formación e inserción laboral correspondien
Educación secundaria.Segunda etapa y formación e inserción laboral correspondien  
Formación e inserción laboral con título de secundaria (2ª etapa) 
Educación superior, excepto doctorado 
Doctorado */

destring Analfabetos, replace force
replace Analfabetos= 0 if Analfabetos==.
destring Formacióneinserciónlaboralco, replace force
replace Formacióneinserciónlaboralco= 0 if Formacióneinserciónlaboralco==.


append using "Khumano.dta"

gen h_lowm= (Analfabetos + Educaciónprimaria + EducaciónsecundariaPrimeraeta ) *1000
gen h_mhigh = (EducaciónsecundariaSegundaeta + Formacióneinserciónlaboralco + Educaciónsuperiorexceptodoct + Doctorado) *1000

replace h_lowm = h_No + h_Prim + h_Secund1 if año>= 2021
replace h_mhigh= h_Secund2 + h_Secund3 + h_fp + h_GradUinf + h_GradUsup + h_Master + h_Phd  if año>= 2021
 
keep Provincia año h_lowm h_mhigh

gen Empl_H= h_lowm + h_mhigh            //Ocupados según estos datos

gen h_lmT = h_lowm if Provincia == "Total Nacional"          //Total Nacional
bysort año: egen H_lmT = total(h_lmT)
drop h_lmT

gen h_mhT = h_mhigh if Provincia == "Total Nacional"          //Total Nacional
bysort año: egen H_mhT = total(h_mhT)
drop h_mhT
 
save "Khumano.dta", replace

/*================================================================================
*                               DESEMPLEO
===============================================================================*/
import delimited "PARO\Activo 1 sect.csv", varnames(1) clear 
 
tempfile a1
save `a1'
 
import delimited "PARO\Activo 2 sect.csv", varnames(1) clear 
append using `a1'

rename total activos
  
 tempfile act
 save `act'


import delimited "PARO\Ocupados 1 sect.csv", varnames(1) clear 
 
tempfile o1
save `o1'
 
import delimited "PARO\Ocupados 2 sect.csv", varnames(1) clear 
append using `o1'
  
rename total ocupados

merge m:m ïprovincias sectoreconãmico periodo using `act' 

rename ïprovincias Provincia
rename sectoreconãmico sectores

gen año= substr(periodo,1, 4)
gen quart= substr(periodo,5, 2)

destring activos, replace ignore(".") dpcomma 
destring ocupados,replace ignore(".") dpcomma

bysort Provincia sector año: egen act= mean(activos)
bysort Provincia sector año: egen ocup= mean(ocupados)

keep if quart== "T4"

gen U = 1 - (ocup/ act)

 /*                         Agricultura 1 
                          ConstrucciÃ³n 2      
                              Industria 3     
Parados que buscan primer empleo o ha.. |      
                              Servicios 4      
                                  Total 5      */

drop if _merge == 2      //Elimanos Parados que buscan primer empleo...
encode sectores, generate(Sect)
label values Sect

replace Sect= 9 if Sect== 5      // Total
replace Sect= 7 if Sect== 3    // Industria
replace Sect= 6 if Sect== 2    // Construcción
replace Sect= 5 if Sect== 4     // Servicios (agregados)!!
replace Sect= 4 if Sect== 1     // Agricultura
 
drop periodo ocupados activos quart _merge sectores

replace Provincia= "A Coruña" if Provincia== "15 CoruÃ±a, A"
replace Provincia= "Álava" if Provincia== "01 Araba/Ãlava"
replace Provincia= "Albacete" if Provincia=="02 Albacete"
replace Provincia= "Alicante" if Provincia== "03 Alicante/Alacant"
replace Provincia= "Almería" if Provincia== "04 AlmerÃ­a"
replace Provincia= "Asturias" if Provincia=="33 Asturias"
replace Provincia= "Ávila" if Provincia== "05 Ãvila"

replace Provincia= "Badajoz" if Provincia== "06 Badajoz"
replace Provincia= "Baleares" if Provincia== "07 Balears, Illes"
replace Provincia= "Barcelona" if Provincia== "08 Barcelona"
replace Provincia= "Burgos" if Provincia== "09 Burgos"

replace Provincia= "Cáceres" if Provincia== "10 CÃ¡ceres"
replace Provincia= "Cádiz" if Provincia== "11 CÃ¡diz"
replace Provincia= "Castellón" if Provincia== "12 CastellÃ³n/CastellÃ³"
replace Provincia= "Cantabria" if Provincia== "39 Cantabria"
replace Provincia= "Ceuta" if Provincia== "51 Ceuta"
replace Provincia= "Ciudad Real" if Provincia== "13 Ciudad Real"
replace Provincia= "Córdoba" if Provincia== "14 CÃ³rdoba"
replace Provincia= "Cuenca" if Provincia== "16 Cuenca"

replace Provincia= "Girona" if Provincia== "17 Girona"
replace Provincia= "Granada" if Provincia== "18 Granada"
replace Provincia= "Guadalajara" if Provincia== "19 Guadalajara"
replace Provincia= "Guipuzcoa" if Provincia== "20 Gipuzkoa"

replace Provincia= "Huelva" if Provincia== "21 Huelva"
replace Provincia= "Huesca" if Provincia== "22 Huesca"

replace Provincia= "Jaén" if Provincia== "23 JaÃ©n"

replace Provincia= "La Rioja" if Provincia== "26 Rioja, La"
replace Provincia= "Las Palmas" if Provincia== "35 Palmas, Las"
replace Provincia= "León" if Provincia== "24 LeÃ³n"
replace Provincia= "Lleida" if Provincia== "25 Lleida"
replace Provincia= "Lugo" if Provincia== "27 Lugo"

replace Provincia= "Málaga" if Provincia== "29 MÃ¡laga"
replace Provincia= "Madrid" if Provincia== "28 Madrid"
replace Provincia= "Melilla" if Provincia== "52 Melilla"
replace Provincia= "Murcia" if Provincia== "30 Murcia"

replace Provincia= "Navarra" if Provincia== "31 Navarra"

replace Provincia= "Ourense" if Provincia== "32 Ourense"

replace Provincia= "Palencia" if Provincia== "34 Palencia"
replace Provincia= "Pontevedra" if Provincia== "36 Pontevedra"

replace Provincia= "Salamanca" if Provincia== "37 Salamanca"
replace Provincia= "Santa Cruz de Tenerife" if Provincia== "38 Santa Cruz de Tenerife"
replace Provincia= "Segovia" if Provincia== "40 Segovia"
replace Provincia= "Sevilla" if Provincia== "41 Sevilla"
replace Provincia= "Soria" if Provincia== "42 Soria"

replace Provincia= "Tarragona" if Provincia== "43 Tarragona"
replace Provincia= "Teruel" if Provincia== "44 Teruel"
replace Provincia= "Toledo" if Provincia== "45 Toledo"
				
replace Provincia= "Valencia" if Provincia== "46 Valencia/ValÃ¨ncia"
replace Provincia= "Valladolid" if Provincia== "47 Valladolid"
replace Provincia= "Vizcaya" if Provincia== "48 Bizkaia"

replace Provincia= "Zamora" if Provincia=="49 Zamora"
replace Provincia= "Zaragoza" if Provincia=="50 Zaragoza"

  *           Total Nacional 
gen ut= U if Provincia == "Total Nacional"
bysort Sect año: egen UT = total(ut)
gen uTT = UT if Sect== 9
bysort Provincia año: egen UTT = total(uTT)

destring año, replace                    //Borramos Provincias-CCAA repetidas
bysort Provincia año Sect: gen i=_n
drop if i==2
drop i ut utt

save "Paro.dta", replace

/*================================================================================
*                               NODOS- VECINDAD / SUPERFICIE
===============================================================================*/
*import excel "D:\Doctorado\Convergencia\DATOS\Nodos.xlsx", sheet("Hoja2") firstrow  clear
*rename A Provincia

import excel "D:\Doctorado\Convergencia\DATOS\IGN_INFOGEO_PROVINCIAS.xlsx", sheet("IGN_INFOGEO_PROVINCIAS") cellrange(B1:C53) firstrow clear
rename Nombre Provincia
rename Superficiekm2 Superf
tempfile Super
save `Super'

cd "D:\Doctorado\Convergencia\DATOS\Mapa"
use recintos_provinciales_inspire_canarias_regcan95, clear
replace _ID= _ID + 51
append using recintos_provinciales_inspire_peninbal_etrs89
spset

rename NAMEUNIT Provincia
merge m:1 Provincia using `Super', nogenerate 
replace Provincia= "Alicante" if Provincia== "Alacant/Alicante"
replace Provincia= "Baleares" if Provincia== "Illes Balears"
replace Provincia= "Castellón" if Provincia== "Castelló/Castellón"
replace Provincia= "Guipuzcoa" if Provincia== "Gipuzkoa"
replace Provincia= "Valencia" if Provincia== "València/Valencia"
replace Provincia= "Vizcaya" if Provincia== "Bizkaia"
replace Provincia= "Álava" if Provincia== "Araba/Álava"

drop if strpos( Provincia, "Territorios")
drop if Provincia=="Ceuta"
drop if Provincia=="Melilla"

spmatrix create contiguity W, replace 
spmatrix create idistance A, replace 

tempfile nodos
save `nodos'

/*===============================================================================
                          ANÁLISIS
*==============================================================================*/
cd "D:\Doctorado\Convergencia\DATOS"
use "Principal.dta", clear

*Nodos
global Nd P_*
merge m:1 Provincia using `nodos', nogenerate 
* 
merge m:m Provincia año Sect using "Kapital", nogenerate 
merge m:m Provincia año using "Infraestructuras", nogenerate 
merge m:m Provincia año Sect using "Paro", nogenerate 
merge m:m Provincia año using "Khumano.dta", nogenerate 
drop _merge

*Borramos datos que no necesitamos
drop if strpos( Provincia, "Total Nacional")
drop if strpos( Provincia, "Ceuta")                      //Prov= 17
drop if strpos( Provincia, "Melilla")                    //Prov= 35
drop if año>2022                                  // K no tiene datos para 2022

*-------------------------------------------------------------------------------
*Factores 

gen lnY =  ln(VAB)
gen lnK =  ln(K)
gen lnL =  ln(Empl)
gen lnF =  ln(Inf1)
gen lnH =  ln(h_mhigh)

*-------------------------------------------------------------------------------
*VAB per capita
gen y = VAB/Poblacion
gen ln_y = ln(y)

*VAB por trabajador == productividad
gen p = VAB/Empl
gen ln_p = ln(p)

*Empleo percápita
gen ln_ep = ln(Empl/Poblacion)

*Kapital intensity
gen k= K/Empl
gen ln_k = ln(K/Empl)

*Infraestructuras 
forvalues i= 1/7 {                  //infraestructuras intensity
gen f_`i'= Inf`i'/Empl
gen ln_f`i' = ln(Inf`i'/Empl)
}

forvalues i= 1/7 {                  //infraestructuras/PIB
gen fy_`i'= Inf`i'/VAB
gen ln_fy`i' = ln(Inf`i'/VAB)
}

forvalues i= 1/7 {                  //infraestructuras/PIB
gen fs_`i'= Inf`i'/Superf
gen ln_fs`i' = ln(Inf`i'/Superf)
}

*Desempleo
gen ln_U = ln(U)

*Capital Humano ratio
gen hl = h_lowm/Empl_H
gen hh = h_mhigh/Empl_H
gen ln_hlm = ln(h_lowm/Empl_H)
gen ln_hmh= ln(h_mhigh/Empl_H)

*-------------------------------------------------------------------------------
* Emplep total provincia: 
gen Emp_PT = Empl if Sect == 9                           //Empleado empleado total por Provincia (PIB)
bysort Provincia año: egen Empl_PT = total(Emp_PT)
drop Emp_PT

*Medias nacionales
bysort sector año:  egen Empl_T = total(Empl)            // EMpleado total por sector en España
gen Emp_TT = Empl_T if Sect == 9                         // EMpleado total(todos sect)en España
bysort Provincia año: egen Empl_TT = total(Emp_TT)
drop Emp_TT

bysort sector año:  egen VAB_T = total(VAB)              // VAB España, por sector y  año
bysort sector año:  egen Poblacion_T = total(Poblacion)  // Población España, por año 
bysort sector año:  egen K_T = total(K)                  // K España, por sector y  año

forvalues i= 1/7 {  
bysort sector año:  egen Inf_T`i' = total(Inf`i')                  // Infraestructuras España, por sector y  año
}

* Variables agregadas
gen ln_yT = ln(VAB_T/Poblacion_T)
gen ln_pT = ln(VAB_T/Empl_T)
gen ln_epT = ln(Empl_T/Poblacion_T)
gen ln_kT = ln(K_T/Empl_T)
forvalues i= 1/7 { 
gen ln_fT`i' = ln(Inf_T`i'/Empl_T)
gen ln_fyT`i' = ln(Inf_T`i'/VAB_T)
}
gen ln_UT= ln(UT) 
gen ln_HHT= ln(H_mhT/(H_lmT+H_mhT))
*-------------------------------------------------------------------------------
sort sector Provincia año
*Incremento anual 
by sector Provincia: gen I_y = ln_y - ln_y[_n-1]
by sector Provincia: gen I_p = ln_p - ln_p[_n-1]
by sector Provincia: gen I_ep = ln_ep - ln_ep[_n-1]

by sector Provincia: gen I_yT = ln_yT - ln_yT[_n-1]
by sector Provincia: gen I_pT = ln_pT - ln_pT[_n-1]

*Anterior
by sector Provincia: gen ln_y_1 = ln_y[_n-1]
by sector Provincia: gen ln_p_1 = ln_p[_n-1]
by sector Provincia: gen ln_ep_1 = ln_ep[_n-1]
by sector Provincia: gen ln_k_1 = ln_k[_n-1]
by sector Provincia: gen ln_f1_1 = ln_f1[_n-1]


*Incremento periodo entero
by sector Provincia: gen I_y_22 = ln_y - ln_y[_n-22]
by sector Provincia: gen I_p_22= ln_p - ln_p[_n-22]

*Incremento periodo entero
by sector Provincia: gen O_y = ln_y[_n-22]
by sector Provincia: gen O_p = ln_p[_n-22]

*___Pib per capita
gen I_yyT = I_y - I_yT                              // dependiente
gen lny_yT = ln_y - ln_yT                              // dependiente
by sector Provincia: gen lnyT_y_1 = ln_yT[_n-1] - ln_y[_n-1]   

*___Productividad 
gen I_ppT = I_p - I_pT                              // dependiente
gen lnp_pT = ln_p - ln_pT                              // dependiente
by sector Provincia: gen lnpT_p_1 = ln_pT[_n-1] - ln_p[_n-1] 

*Estimación TFT
gen lnYT= ln(VAB_T)
gen lnKT= ln(K_T) 
gen lnLT= ln(Empl_T) 


*1

*reg lnY lnK lnL if Sect==9, robust
*predict u, resid
*gen A_1 = _b[_cons] + u

reg lnY lnK lnL i.Prov i.año if Sect==9, robust
predict u_fe, resid
predict xb, xb
gen lnA = xb - (_b[lnK]*lnK + _b[lnL]*lnL) + u_fe
*gen TFP_2T = lnYT - (_b[lnK]*lnKT + _b[lnL]*lnLT)

*agregado
bysort sector año:  egen YT_hat = total( K^_b[lnK]*Empl^_b[lnL])
gen AT = VAB_T/YT_hat
gen lnAT = ln(AT)
*-------

*constraint 1 lnK + lnL = 1
*cnsreg lnY lnK lnL i.Prov i.año if Sect==9, constraints(1)robust

*-------
*2
reg lnY lnK lnL lnF i.Prov i.año if Sect==9, robust
predict u_feF, resid
predict xbF, xb
gen lnAF = xbF - (_b[lnK]*lnK + _b[lnL]*lnL + _b[lnF]*lnF) + u_feF

*agregado
bysort sector año:  egen YT_hatF = total( K^_b[lnK]*Empl^_b[lnL]*Inf1^_b[lnF])
gen ATF = VAB_T/YT_hatF
gen lnATF = ln(ATF)
*-------

reg lnY lnK lnL lnF lnH i.Prov i.año if Sect==9, robust

constraint 3 lnK + lnL + lnF + lnH = 1
cnsreg lnY lnK lnL lnF lnH i.Prov i.año if Sect==9, constraints(3)robust

by sector Provincia: gen lnAF_1 = lnAF[_n-1]

*save "Principal.dta", replace

*-------------------------------------------------------------------------------
*SIGMA CONVERGENCIA
*___Pib per capita
gen lny_yT2 = lny_yT^2
bysort sector año:  egen sigma2_y = total(lny_yT^2)
gen sigma_y = (sigma2_y/50)^(1/2)                       //OJO: sin Ceuta y Melilla
graph twoway line sigma_y año if Sect == 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_y_Total.png", as(png) replace

*___Productividad 
gen lnp_pT2 = lnp_pT^2
bysort sector año:  egen sigma2_p = total(lnp_pT^2)
gen sigma_p = (sigma2_p/50)^(1/2)                       //OJO: sin Ceuta y Melilla
graph twoway line sigma_p año if Sect == 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_p_Total.png", as(png) replace

*___Empleo percápita
gen lnep_epT = ln_ep - ln_epT  
gen lnep_epT2 = lnep_epT^2
bysort sector año:  egen sigma2_ep = total(lnep_epT^2)  
gen sigma_ep = (sigma2_ep/50)^(1/2)                      //OJO: sin Ceuta y Melilla
graph twoway line sigma_ep año if Sect == 9       
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_ep_Total.png", as(png) replace
                  
*___Kapital intensity
gen lnk_kT = ln_k - ln_kT  
gen lnk_kT2 = lnk_kT^2
bysort sector año:  egen sigma2_k = total(lnk_kT^2) 
gen sigma_k = (sigma2_k/50)^(1/2)                      //OJO: sin Ceuta y Melilla
graph twoway line sigma_k año if Sect == 9 & año<2022   
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_k_Total.png", as(png) replace

*___Infrastructure intensity
forvalues i= 1/7 { 
gen lnf_fT`i' = ln_f`i' - ln_fT`i'  
gen lnf_fT2`i' = lnf_fT`i'^2
bysort sector año:  egen sigma2_f`i' = total(lnf_fT`i'^2) 
gen sigma_f`i' = (sigma2_f`i'/50)^(1/2)                      //OJO: sin Ceuta y Melilla
graph twoway line sigma_f`i' año if Sect == 9 & año<2022   
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_f`i'.png", as(png) replace
}

forvalues i= 1/7 { 
gen lnf_fyT`i' = ln_fy`i' - ln_fyT`i'  
gen lnf_fyT2`i' = lnf_fyT`i'^2
bysort sector año:  egen sigma2_fy`i' = total(lnf_fyT`i'^2) 
gen sigma_fy`i' = (sigma2_fy`i'/50)^(1/2)                      //OJO: sin Ceuta y Melilla
graph twoway line sigma_fy`i' año if Sect == 9 & año<2022   
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_fy`i'.png", as(png) replace
}

*___Desempleo
gen lnU_UT = ln_U - ln_UT  
bysort sector año:  egen sigma2_U = total(lnU_UT^2) 
gen sigma_U = (sigma2_U/50)^(1/2)                      //OJO: sin Ceuta y Melilla
graph twoway line sigma_U año if Sect == 9 & año>2001 
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_U_Total.png", as(png) replace

*___Capital Humano
gen lnH_HT = ln_hmh - ln_HHT    
bysort sector año:  egen sigma2_H = total(lnH_HT^2) 
gen sigma_H = (sigma2_H/50)^(1/2)                      //OJO: sin Ceuta y Melilla
graph twoway line sigma_H año if Sect == 9 

*___TFP
gen lnA_AT = lnA - lnAT
bysort sector año:  egen sigma2_A = total(lnA_AT^2) 
gen sigma_A = (sigma2_A/50)^(1/2)                      //OJO: sin Ceuta y Melilla
graph twoway line sigma_A año if Sect == 9 & año<2022
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_A_Total.png", as(png) replace

gen lnAF_ATF = lnAF - lnATF
bysort sector año:  egen sigma2_AF = total(lnAF_ATF^2) 
gen sigma_AF = (sigma2_AF/50)^(1/2)                      //OJO: sin Ceuta y Melilla
graph twoway line sigma_AF año if Sect == 9 & año<2022
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_A2_Total.png", as(png) replace

*Sectores
forvalue s= 1/7 {
* 1: Industrias Manufactureras
* 2: Servicios financieros, inmobiliarios, profesionales, admin
* 3: Servicios: Admin Pub, Sanidad, educacion ,artistas y entretenimiento y otros
* 4: Sector Primario
* 5: Servicios: omercio; rep. de vehículos, transporte, hostelería, Comunicacion
* 6: COnstrucción
* 7: Indusstria

graph twoway line sigma_p año if Sect == `s' 
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Sigma_p_`s'.png", as(png) replace
}

*-------------------------------------------------------------------------------
*DESIGUALDAD EN ESTRUCTURA SECTORIAL (ID): 
gen PS= Empl/Empl_PT
gen PS_T= Empl_T/Empl_TT


forvalue s= 1/7 {
gen P_`s' = PS  if Sect == `s'
gen P_`s'T= PS_T if Sect == `s'

bysort Provincia año: egen PS_`s' = total(P_`s')
bysort Provincia año: egen PN_`s' = total(P_`s'T)
drop P_`s' P_`s'T
}

forvalue s= 1/7 {
gen D_`s'= (PS_`s'- PN_`s')^2
}
bysort Provincia año:  gen ID_P = D_2 + D_3 + D_4 + D_5 + D_6 + D_7  //Excluimos Industrias Manufactureras (D_1)
bysort sector año:  egen ID = total(ID_P)
replace ID= ID/50                          //OJO: sin Ceuta y Melilla
graph twoway line ID año if Sect == 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\ID.png", as(png) replace

*Por sector:
forvalue s= 1/7 {
bysort sector año: egen ID_`s' = total(D_`s')
replace ID_`s'= ID_`s'/50

graph twoway line ID_`s' año if Sect == 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\ID_`s'.png", as(png) replace
}

*-------------------------------------------------------------------------------
*DENSIDADES: 
*Simple
*Renta
kdensity ln_y if año== 2000 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dy00.png", as(png) replace
kdensity ln_y if año== 2011 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dy11.png", as(png) replace
kdensity ln_y if año== 2022 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dy22.png", as(png) replace
* Productividad
kdensity ln_p if año== 2000 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dp00.png", as(png) replace
kdensity ln_p if año== 2011 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dp11.png", as(png) replace
kdensity ln_p if año== 2022 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dp22.png", as(png) replace
*Kapital
kdensity ln_k if año== 2000 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dk00.png", as(png) replace
kdensity ln_k if año== 2011 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dk11.png", as(png) replace
kdensity ln_k if año== 2021 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dk21.png", as(png) replace
*Human kapital
 kdensity lnH if año== 2002 & Sect== 9
 graph export "D:\Doctorado\Convergencia\DATOS\Resultados\DH00.png", as(png) replace
 kdensity lnH if año== 2022 & Sect== 9
 graph export "D:\Doctorado\Convergencia\DATOS\Resultados\DH22.png", as(png) replace
 
*Ponderada
*Renta
gen pop = Poblacion/Poblacion_T
kdensity ln_y [aweight=pop] if año== 2000 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dy00w.png", as(png) replace
kdensity ln_y [aweight=pop] if año== 2011 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dy11w.png", as(png) replace
kdensity ln_y [aweight=pop] if año== 2022 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dy22w.png", as(png) replace
* Productividad
kdensity ln_p [aweight=pop] if año== 2000 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dp00w.png", as(png) replace
kdensity ln_p [aweight=pop] if año== 2011 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dp11w.png", as(png) replace
kdensity ln_p [aweight=pop] if año== 2022 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dp22w.png", as(png) replace
* Kapital
kdensity ln_k [aweight=pop] if año== 2000 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dk00w.png", as(png) replace
kdensity ln_k [aweight=pop] if año== 2011 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dk11w.png", as(png) replace
kdensity ln_k [aweight=pop] if año== 2021 & Sect== 9
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\Dk21w.png", as(png) replace

*-------------------------------------------------------------------------------
* REGRESIÓN : BETA CONVERGENCIA
*_ Básico: INCONDICIONADA
eststo clear
reg I_y ln_y_1 if Sect ==9 , robust   // PIB
eststo  v1
reg I_y ln_y_1 if Sect ==9 & año<=2010, robust 
eststo  v2
reg I_y ln_y_1 if Sect ==9 & año>2010, robust 
eststo  v3

reg I_y ln_y_1 i.Prov if Sect ==9 , robust   // PIB
eststo  v7
reg I_y ln_y_1 i.Prov if Sect ==9 & año<=2010, robust 
eststo  v8
reg I_y ln_y_1 i.Prov if Sect ==9 & año>2010, robust 
eststo  v9

esttab using tabla_y.tex , se ar2 replace booktabs title ("v1" "v7" "v2" "v8" "v3" "v9") mtitle ("v1" "v7" "v2" "v8" "v3" "v9")    keep(ln_y_1)
eststo clear

reg I_p ln_p_1 if Sect ==9  , robust   // productividada
eststo  v4
reg I_p ln_p_1 if Sect ==9 & año<=2010, robust 
eststo  v5
reg I_p ln_p_1 if Sect ==9 & año>2010, robust 
eststo  v6

reg I_ep ln_ep_1 if Sect ==9 , robust   // empleo/poblacion


*_ Básico: CONDICIONADA
reg I_p ln_p_1  i.Prov if Sect ==9  , robust   // productividada
eststo  v10
reg I_p ln_p_1 i.Prov  if Sect ==9 & año<=2010, robust 
eststo  v11
reg I_p ln_p_1 i.Prov  if Sect ==9 & año>2010, robust 
eststo  v12

esttab using tabla_p.tex , se ar2 replace booktabs title ("v1" "v7" "v2" "v8" "v3" "v9") mtitle ("v1" "v7" "v2" "v8" "v3" "v9")    keep(ln_p_1)

reg I_y ln_y_1 i.Prov i.año if Sect ==9 , robust  // PIB
reg I_p ln_p_1 i.Prov i.año if Sect ==9 , robust  // productividada


*con factores

reg I_y ln_y_1 ln_k_1 ln_f1_1 lnAF_1 if Sect ==9 , robust   // PIB


*_Spatial effects 
preserve
keep if Sect==9
xtset  Prov año
spxtregress I_y ln_y_1, fe dvarlag(W) ivarlag(W:ln_y_1)
spxtregress I_y ln_y_1, fe dvarlag(A) ivarlag(A:ln_y_1)

spxtregress I_p ln_p_1, fe dvarlag(W) ivarlag(W:ln_p_1)
spxtregress I_p ln_p_1, fe dvarlag(A) ivarlag(A:ln_p_1)
restore

*_Cross section 
reg I_y_22 O_y if Sect ==9 , robust   // PIB
reg I_p_22 O_p if Sect ==9 , robust   // productividada

* Modelo alternitivo
reg lny_yT lnyT_y_1 i.Prov if Sect ==9 , robust    // PIB
reg lnp_pT lnpT_p_1 i.Prov if Sect ==9  , robust   // PIB

reg I_yyT lnyT_y_1 i.Prov if Sect ==9 , robust nocons   // PIB
reg I_ppT lnpT_p_1 i.Prov if Sect ==9  , robust   // PIB




*Sectores
forvalue s= 1/7 {
* 1: Industrias Manufactureras
* 2: Servicios financieros, inmobiliarios, profesionales, admin
* 3: Servicios: Admin Pub, Sanidad, educacion ,artistas y entretenimiento y otros
* 4: Sector Primario
* 5: Servicios: omercio; rep. de vehículos, transporte, hostelería, Comunicacion
* 6: COnstrucción
* 7: Indusstria
display `s'
*_ Básico: INCONDICIONADA
reg I_p ln_p_1 if Sect == `s' , robust   // productividada

*_ Básico: CONDICIONADA
reg I_p ln_p_1  i.Prov if Sect ==`s' , robust   // productividada                             
reg lnp_pT lnpT_p_1 i.Prov if Sect == `s' , robust 
}


*MAPAS y DESCRIPTIVOS
use "Principal.dta", clear

*El desempleo es a partir del 2002
bysort Provincia Sect: egen u= total(U) if año<=2002
replace U= u if año==2000     //usamos los valores del 2002 en el 2000 meramente por facilidad en el código

*El capital Humano es solo del 2002
bysort Provincia Sect: egen Hh= total(hh) if año<=2002
replace hh= Hh if año==2000     //usamos los valores del 2002 en el 2000 meramente por facilidad en el código

keep if año== 2000 | año== 2010 | año>= 2021   //2000, 2010, 2021*, 2022
keep if Sect==9
drop ln* I_* O_* u Hh

*Peso de las Provincia en GDP total
gen W_VAB = VAB/VAB_T

*Peso de las Provincia en Población total
gen W_Pop = Poblacion/ Poblacion_T

*Variación PIB 
sort Provincia año
by Provincia: gen I_Y= (VAB - VAB[_n-1])/VAB[_n-1]  if año!=2022
by Provincia: gen I_Y2=(VAB - VAB[_n-2])/VAB[_n-2]
by Provincia: gen I_Y3=(VAB - VAB[_n-3])/VAB[_n-3]

*Variación PIB per capita
sort Provincia año
by Provincia: gen I_y= (y - y[_n-1])/y[_n-1]  if año!=2022
by Provincia: gen I_y2= (y - y[_n-2])/y[_n-2]
by Provincia: gen I_y3= (y - y[_n-3])/y[_n-3]

*Variación Población
sort Provincia año
by Provincia: gen I_Pop= (Poblacion - Poblacion[_n-1])/Poblacion[_n-1]  if año!=2022
by Provincia: gen I_Pop2=(Poblacion - Poblacion[_n-2])/Poblacion[_n-2]
by Provincia: gen I_Pop3=(Poblacion - Poblacion[_n-3])/Poblacion[_n-3]

*Variación Empleo
sort Provincia año
by Provincia: gen I_E= (Empl - Empl[_n-1])/Empl[_n-1]  if año!=2022
by Provincia: gen I_E2=(Empl - Empl[_n-2])/Empl[_n-2]
by Provincia: gen I_E3=(Empl - Empl[_n-3])/Empl[_n-3]

*Variación Productividad
sort Provincia año
by Provincia: gen I_p= (p - p[_n-1])/p[_n-1]  if año!=2022
by Provincia: gen I_p2=(p - p[_n-2])/p[_n-2]
by Provincia: gen I_p3=(p - p[_n-3])/p[_n-3]

*Variación Kapital y Kapital intensity 
sort Provincia año
by Provincia: gen I_K= (K - K[_n-1])/K[_n-1] 
by Provincia: gen I_K2=(K - K[_n-2])/K[_n-2]

by Provincia: gen I_k= (k - k[_n-1])/k[_n-1]  
by Provincia: gen I_k2=(k - k[_n-2])/k[_n-2]

*Variación Infraestructuras
sort Provincia año
forvalues i= 1/7 { 
by Provincia: gen I_F`i'= (Inf`i' - Inf`i'[_n-1])/Inf`i'[_n-1] 
by Provincia: gen I_F`i'2=(Inf`i' - Inf`i'[_n-2])/Inf`i'[_n-2]

by Provincia: gen I_f`i'= (f_`i' - f_`i'[_n-1])/f_`i'[_n-1] 
by Provincia: gen I_f`i'2=(f_`i' - f_`i'[_n-2])/f_`i'[_n-2]

by Provincia: gen I_fy`i'= (fy_`i' - fy_`i'[_n-1])/fy_`i'[_n-1]  
by Provincia: gen I_fy`i'2=(fy_`i'- fy_`i'[_n-2])/fy_`i'[_n-2]
}

*Variación TASA de desempleo
sort Provincia año
by Provincia: gen I_U = (U - U[_n-1])
by Provincia: gen I_U2=(U - U[_n-2])
by Provincia: gen I_U3=(U - U[_n-3])

*Variación RATIO Capital Humano
sort Provincia año
by Provincia: gen I_hh2=(hh - hh[_n-2])
by Provincia: gen I_hh3=(hh - hh[_n-3])

*Variación TFP
sort Provincia año
by Provincia: gen I_A=(lnA - lnA[_n-1])
by Provincia: gen I_A2=(lnA - lnA[_n-2])
by Provincia: gen I_A3=(lnA - lnA[_n-3])

drop C sector Sect Prov    h_lowm h_mhigh Empl_H 
global X K K_T k Inf* f* lnA Poblacion VAB Empl hh y p Empl_PT Empl_T Empl_TT UT UTT VAB_T Poblacion_T W_VAB W_Pop I_y* I_Pop* I_E* I_p* I_Y* I_K* I_k* I_f* I_F* I_U* I_h* I_A* act ocup U Superf H_lmT H_mhT
keep $X año
reshape wide $X , i(Provincia) j( año)
drop Empl_TT* Empl_PT* UT* // si solo usamos un sector
drop I_p2000 I_y2000 I_E2000 I_U2000 I_Pop2000  *22000 *32000 *22010 *32010 *32021 I_p2022 I_y2022 I_E2022 I_Pop2022 
drop I_K2000 I_k2000 I_K2022 I_k2022 I_K22022 I_k22022 k2022 

export excel Provincia VAB* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("PIB") sheetmodify firstrow(variables)
export excel Provincia Pobla* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("Población") sheetmodify firstrow(variables)
export excel Provincia Empl*  using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("Empleo") sheetmodify firstrow(variables)
export excel Provincia U*  using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("Paro") sheetmodify firstrow(variables)

export excel Provincia y* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("Pib per capita") sheetmodify firstrow(variables)
export excel Provincia p* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("Productividad") sheetmodify firstrow(variables)
export excel Provincia W* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("Pesos") sheetmodify firstrow(variables)
export excel Provincia k* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("KapInte") sheetmodify firstrow(variables)
export excel Provincia f* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("Infraestructuras") sheetmodify firstrow(variables)
export excel Provincia hh* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("Capital humano") sheetmodify firstrow(variables)
export excel Provincia lnA*  using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("TFP") sheetmodify firstrow(variables)

export excel Provincia I_p* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_p") sheetmodify firstrow(variables)
export excel Provincia I_y* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_y") sheetmodify firstrow(variables)
export excel Provincia I_Pop* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_Pop") sheetmodify firstrow(variables)
export excel Provincia I_E* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_E") sheetmodify firstrow(variables)
export excel Provincia I_K* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_K") sheetmodify firstrow(variables)
export excel Provincia I_k* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_ki") sheetmodify firstrow(variables)
export excel Provincia I_f* I_F* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_f") sheetmodify firstrow(variables)
export excel Provincia I_U* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_U") sheetmodify firstrow(variables)
export excel Provincia I_hh* using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_H") sheetmodify firstrow(variables)
export excel Provincia I_A*  using "D:\Doctorado\Convergencia\DATOS\Resultados\Descriptivos.xlsx", sheet("I_TFP") sheetmodify firstrow(variables)


twoway (scatter I_y32022 y2000, mlabel(Provincia)) (qfit I_y32022 y2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\ploty.png", as(png) replace
twoway (scatter I_p32022 p2000, mlabel(Provincia)) (qfit I_p32022 p2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotp.png", as(png) replace

twoway (scatter I_y2010 y2000, mlabel(Provincia)) (qfit I_y2010 y2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\ploty0010.png", as(png) replace
twoway (scatter I_p2010 p2000, mlabel(Provincia)) (qfit I_p2010 p2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotp0010.png", as(png) replace

twoway (scatter I_y22022 y2010, mlabel(Provincia)) (qfit I_y22022 y2010)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\ploty1022.png", as(png) replace
twoway (scatter I_p22022 p2010, mlabel(Provincia)) (qfit I_p22022 p2010)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotp1022.png", as(png) replace

*Kapìtal
twoway (scatter I_k22021 k2000, mlabel(Provincia)) (qfit I_k22021 k2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotk.png", as(png) replace
twoway (scatter I_k2010 k2000, mlabel(Provincia)) (qfit I_k2010 k2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotk0010.png", as(png) replace
twoway (scatter I_k2021 k2010, mlabel(Provincia)) (qfit I_k2021 k2010)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotk1022.png", as(png) replace

*Infraestructuras
forvalues i= 1/7 {
twoway (scatter I_f`i'22021 f_`i'2000, mlabel(Provincia)) (qfit I_f`i'22021 f_`i'2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotf`i'.png", as(png) replace
twoway (scatter I_fy`i'22021 fy_`i'2000, mlabel(Provincia)) (qfit I_fy`i'22021 fy_`i'2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotfy`i'.png", as(png) replace
twoway (scatter I_F`i'22021 fs_`i'2000, mlabel(Provincia)) (qfit I_F`i'22021 fs_`i'2000)

twoway (scatter I_f`i'2010 f_`i'2000, mlabel(Provincia)) (qfit I_f`i'2010 f_`i'2000)
twoway (scatter I_f`i'2021 f_`i'2010, mlabel(Provincia)) (qfit I_f`i'2021 f_`i'2010)
}

twoway (scatter I_y22021 f_12000, mlabel(Provincia)) (qfit I_y22021 f_12000)
twoway (scatter I_p22021 f_12000, mlabel(Provincia)) (qfit I_p22021 f_12000)

twoway (scatter I_y22021 fy_12000, mlabel(Provincia)) (qfit I_y22021 fy_12000)
twoway (scatter I_p22021 fy_12000, mlabel(Provincia)) (qfit I_p22021 fy_12000)

twoway (scatter I_y22021 fs_12000, mlabel(Provincia)) (qfit I_y22021 fs_12000)
twoway (scatter I_p22021 fs_12000, mlabel(Provincia)) (qfit I_p22021 fs_12000)


twoway (scatter I_f22021 y2000, mlabel(Provincia)) (qfit I_f22021 y2000)      
twoway (scatter I_f22021 p2000, mlabel(Provincia)) (qfit I_f22021 p2000)

*Desempleo
twoway (scatter I_U32022 U2000, mlabel(Provincia)) (qfit I_U32022 U2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotU.png", as(png) replace

twoway (scatter I_U2010 U2000, mlabel(Provincia)) (qfit I_U2010 U2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotU0010.png", as(png) replace

twoway (scatter I_U22022 U2010, mlabel(Provincia)) (qfit I_U22022 U2010)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotU1022.png", as(png) replace

twoway (scatter  I_Pop32022 I_U32022 , mlabel(Provincia)) (qfit  I_Pop32022 I_U32022)


*Capital humano
twoway (scatter I_hh32022 hh2000, mlabel(Provincia)) (qfit I_hh32022 hh2000)
graph export "D:\Doctorado\Convergencia\DATOS\Resultados\plotH.png", as(png) replace

twoway (scatter I_p32022 hh2000, mlabel(Provincia)) (qfit I_p32022 hh2000)
twoway (scatter p2000 hh2000, mlabel(Provincia)) (qfit p2000 hh2000)
twoway (scatter p2022 hh2022, mlabel(Provincia)) (qfit p2022 hh2022)

*TFP



*===============================================================================

*===============================================================================

*===============================================================================


*Histograma tasa crecimiento del VAB
by sector Provincia:  gen Iy= (y - y[_n-22])/y[_n-22]
histogram Iy Prov if Sect==9
*Histograma tasa crecimiento del Empleo
*Histograma tasa crecimiento de la Población

 *ssc install spmap
 *ssc install shp2dta
 *ssc install mif2dta
 cd "D:\Doctorado\Convergencia\DATOS\Mapa"
 
spshape2dta recintos_provinciales_inspire_canarias_regcan95
spshape2dta  recintos_provinciales_inspire_peninbal_etrs89


use recintos_provinciales_inspire_canarias_regcan95, clear
 spset
 spmap using  recintos_provinciales_inspire_canarias_regcan95_shp, id( _ID )
 
  use recintos_provinciales_inspire_peninbal_etrs89, clear
 spset
 spmap using  recintos_provinciales_inspire_peninbal_etrs89_shp, id( _ID )

 
 
use recintos_provinciales_inspire_canarias_regcan95, clear
replace _ID= _ID + 51
append using recintos_provinciales_inspire_peninbal_etrs89
spset
spmatrix create contiguity W
spmatrix create idistance A


save Provincias_map , replace

use recintos_provinciales_inspire_canarias_regcan95_shp, clear
replace _ID= _ID + 51
append using recintos_provinciales_inspire_peninbal_etrs89_shp
save Provincias_map_shp , replace
 spmap using  recintos_provinciales_inspire_peninbal_etrs89_shp, id( _ID )


*--------------------Modelo 2---------------------------------------------------
//Modelo estructural 
 

*-----------------------------

Agglomeration effects KM2 
"labour productiviy disparities in Europea regions 




spset

spmatrix

spregress

spxtregress

estat moran




constraint 1 lnK + lnL + lnM + lnH = 1
cnsreg lnY lnK lnL lnM lnH, constraints(1)
