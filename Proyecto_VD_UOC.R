# ==============================================================================
# Proyecto: Visualización de la Gentrificación y Presión Turística
# Autor: Joaquín Paredes Ribera
# Asignatura: Proyecto de Visualización de Datos (Máster en Ingeniería Informática)
# Centro: Universitat Oberta de Catalunya (UOC)
# Fecha: Junio 2026
# Licencia: MIT License (Garantiza código abierto y reproducibilidad científica)
# ==============================================================================

# ==============================================================================
# 0. CARGA E INSTALACIÓN DE LIBRERÍAS
# ==============================================================================

# Se descarga e instala los paquetes desde el repositorio oficial CRAN si no están presentes,
# para garantizar la portabilidad del script; cualquier máquina que ejecute el archivo 
# instalará automáticamente las dependencias necesarias sin intervención manual.
install.packages(c("readxl", "writexl", "dplyr", "stringr", "purrr", "readr", "tidyr"))

# Carga las librerías en el entorno activo de memoria de R.
library(readxl)   # Permite la lectura eficiente de hojas de cálculo .xlsx multipestaña.
library(writexl)  # Permite exportar los dataframes finales estructurados de vuelta a Excel.
library(dplyr)    # Proporciona la gramática de manipulación de datos (verbos mutate, filter, etc.).
library(stringr)  # Ofrece funciones para la limpieza y formateo de cadenas de texto.
library(purrr)    # Para programación funcional (mapeo iterativo seguro celda por celda).
library(readr)    # Para la importación rápida de archivos de texto plano y CSV.
library(tidyr)    # Para pivotar, reestructurar y asegurar la limpieza formal del dataset.

# ==============================================================================
# 1. VECTOR MAESTRO DE MAPEO Y FUNCIONES DE INGENIERÍA DE DATOS
# ==============================================================================

# Vector asociativo (diccionario) que empareja códigos catastrales/INE con nombres geográficos.
# Las fuentes oficiales (INE, MIVAU) usan códigos numéricos estrictos, mientras que Airbnb 
# registra nombres de texto. Este vector actúa como puente de traducción.
mapeo_completo <- c(
  "0801902" = "Eixample", "0801903" = "Sants-Montjuïc", "0801906" = "Gràcia",
  "0801910" = "Sant Martí", "0801901" = "Ciutat Vella", "0801907" = "Horta-Guinardó",
  "0801908" = "Nou Barris", "0801904" = "Les Corts", "0801905" = "Sarrià-Sant Gervasi",
  "0801909" = "Sant Andreu", "0801907" = "Horta", "0801905" = "el Putxet i el Farró",
  "20069" = "GuipÃºzcoa", "48201" = "Vizcaya", "01059" = "Ã lava", "48201" = "Bilbao",
  "01033" = "LantarÃ³n", "20069" = "Donostia-San SebastiÃ¡n", "48068" = "Mundaka",
  "20001" = "Abaltzisketa", "48026" = "Dima", "48018" = "Bermeo", "48073" = "Ondarroa",
  "48013" = "Barakaldo", "20079" = "Zarautz", "17079" = "Girona", "17013" = "Begur",
  "17095" = "Lloret de Mar", "17089" = "Llagostera", "17199" = "Torroella de Montgrí",
  "17148" = "Riudarenes", "17187" = "Saus", "17153" = "Rupià", "17065" = "Esponellà",
  "17210" = "Ventalló", "17117" = "Palafrugell", "17042" = "Capmany", "17118" = "Palamós",
  "17202" = "Tossa de Mar", "17032" = "Cadaqués", "17124" = "Pals", "17062" = "L'Escala",
  "17115" = "Ordis", "17152" = "Roses", "17023" = "Blanes", "17048" = "Castell-Platja d'Aro",
  "17034" = "Calonge", "17047" = "Castelló d'Empúries", "17204" = "Ullà", "17098" = "Maià de Montcal",
  "17114" = "Olot", "17181" = "Santa Cristina d'Aro", "17020" = "Bescanó", "17022" = "La Bisbal d'Emporda",
  "17207" = "La Vall d'en Bas", "17208" = "La Vall de Bianya", "17188" = "La Selva de Mar",
  "17109" = "Montagut i Oix", "17164" = "Sant Hilari Sacalm", "17021" = "Beuda", "17092" = "Llançà",
  "17094" = "Llívia", "17090" = "Llambilles", "17142" = "Quart", "17057" = "Corçà",
  "17212" = "Vidrà", "17902" = "Forallac", "17111" = "Navata", "17184" = "Santa Pau",
  "17151" = "Riumors", "17141" = "Puigcerdà", "17193" = "Sils", "17015" = "Banyoles",
  "17218" = "Vilademuls", "17120" = "Palau-saverdera", "17088" = "Lladó", "17197" = "Torrent",
  "17206" = "Urús", "17082" = "Guils de Cerdanya", "17173" = "Sant Martí Vell", "17039" = "Camprodon",
  "17024" = "Bolvir", "17147" = "Ripoll", "17012" = "Avinyonet de Puigventós", "17016" = "Bàscara",
  "17198" = "Torroella de Fluvià", "17140" = "El Port de la Selva", "17167" = "Sant Joan de les Abadesses",
  "17145" = "Ribes de Freser", "17185" = "Sant Joan les Fonts", "17166" = "Sant Jordi Desvalls",
  "17046" = "Castellfollit de la Roca", "17116" = "Osor", "17121" = "Palau-sator", "17201" = "Toses",
  "17080" = "Gombrèn", "17160" = "Sant Feliu de Guíxols", "17069" = "Fontanals de Cerdanya",
  "17143" = "Rabós", "17901" = "Cruïlles, Monells i Sant Sadurní de l'Heura", "17102" = "Maçanet de Cabrenys",
  "17165" = "Sant Jaume de Llierca", "17190" = "Serinyà", "17058" = "Crespià", "17132" = "Peralada",
  "17056" = "Cornellà del Terri", "17158" = "Sant Climent Sescebes", "17189" = "La Cellera de Ter",
  "17019" = "Besalú", "17004" = "Albons", "17191" = "Serra de Daró", "17035" = "Camós",
  "17162" = "Sant Ferriol", "17172" = "Sant Martí de Llémena", "17009" = "Arbúcies", "17007" = "Amer",
  "17180" = "Santa Coloma de Farners", "17083" = "Hostalric", "17096" = "Les Llosses", "17051" = "Cistella",
  "17136" = "Pontós", "17228" = "Vilanant", "17170" = "Vallfogona de Ripollès", "17154" = "Sales de Llierca",
  "17157" = "Sant Andreu Salou", "17041" = "Cantallops", "17128" = "Pau", "17093" = "Llers",
  "17097" = "Madremanya", "17085" = "Jafre", "17064" = "Espolla", "17183" = "Sant Aniol de Finestres",
  "17029" = "Boadella d'Empordà", "17135" = "Pont de Molins", "17222" = "Vilaür", "17227" = "Vilamaniscle",
  "17137" = "Porqueres", "17107" = "Molló", "17192" = "Setcases", "17061" = "Das", "17105" = "Mieres",
  "17002" = "Aiguaviva", "17146" = "Riells i Viabrea", "17070" = "Fontanilles", "17144" = "Regencós",
  "17087" = "Juià", "17073" = "Fornells de la Selva", "17077" = "Garriguella", "17054" = "Colera",
  "17037" = "Campelles", "17003" = "Albanyà", "17161" = "Sant Feliu de Pallerols",
  "2807910" = "Latina", "2807901" = "Centro", "2807904" = "Salamanca", "2807908" = "Fuencarral - El Pardo",
  "2807915" = "Ciudad Lineal", "2807917" = "Villaverde", "2807907" = "Chamberí", "2807916" = "Hortaleza",
  "2807902" = "Arganzuela", "2807906" = "Tetuán", "2807903" = "Retiro", "2807911" = "Carabanchel",
  "2807920" = "San Blas - Canillejas", "2807921" = "Barajas", "2807905" = "Chamartín", "2807912" = "Usera",
  "2807918" = "Villa de Vallecas", "2807913" = "Puente de Vallecas", "2807909" = "Moncloa - Aravaca",
  "2807914" = "Moratalaz", "2807919" = "Vicálvaro", "2906702" = "Este", "2906701" = "Centro",
  "2906708" = "Churriana", "2906707" = "Carretera de Cadiz", "2906706" = "Cruz De Humilladero",
  "2906711" = "Teatinos-Universidad", "2906710" = "Puerto de la Torre", "2906703" = "Ciudad Jardin",
  "2906704" = "Bailen-Miraflores", "2906705" = "Palma-Palmilla", "2906709" = "Campanillas",
  "07040" = "Palma de Mallorca", "07051" = "Sant Llorenç des Cardassar", "07058" = "Selva",
  "07007" = "Banyalbufar", "07003" = "Alcúdia", "07057" = "Santanyí", "07025" = "Fornalutx",
  "07013" = "Campos", "07031" = "Llucmajor", "07055" = "Santa Margalida", "07022" = "Felanitx",
  "07030" = "Llubí", "07020" = "Esporles", "07039" = "Muro", "07028" = "Lloret de Vistalegre",
  "07042" = "Pollença", "07061" = "Sóller", "07011" = "Calvià", "07033" = "Manacor",
  "07062" = "Son Servera", "07005" = "Andratx", "07059" = "Ses Salines", "07014" = "Capdepera",
  "07027" = "Inca", "07010" = "Bunyola", "07018" = "Deyá", "07044" = "Sa Pobla",
  "07006" = "Artà", "07056" = "Santa María del Camí", "07043" = "Porreres", "07038" = "Montuïri",
  "07001" = "Alaró", "07019" = "Escorca", "07034" = "Mancor de la Vall", "07029" = "Lloseta",
  "07036" = "Marratxí", "07041" = "Petra", "07004" = "Algaida", "07012" = "Campanet",
  "07016" = "Consell", "07045" = "Puigpunyent", "07008" = "Binissalem", "07047" = "Sencelles",
  "07017" = "Costitx", "07060" = "Sineu", "07065" = "Vilafranc de Bonany", "07009" = "Búger",
  "07053" = "Santa Eugènia", "07901" = "Ariany", "07021" = "Estellencs", "07063" = "Valldemossa",
  "07049" = "Sant Joan", "07035" = "Maria de la Salut", "07024" = "Es Mercadal",
  "07015" = "Ciutadella de Menorca", "07002" = "Alaior", "07032" = "Mahón", "07064" = "Es Castell",
  "07052" = "Sant Lluís", "07023" = "Ferreries", "07902" = "Es Migjorn Gran",
  "4109101" = "Casco Antiguo", "4109102" = "Macarena", "4109103" = "Nervión", "4109104" = "Cerro - Amate",
  "4109107" = "Triana", "4109105" = "Sur", "4109109" = "Palmera - Bellavista", "4109108" = "San Pablo - Santa Justa",
  "4109111" = "Macarena - Norte", "4109110" = "Este - Alcosa - Torreblanca", "4109106" = "Los Remedios",
  "4625005" = "LA SAIDIA", "4625019" = "POBLATS DEL SUD", "4625011" = "POBLATS MARITIMS",
  "4625003" = "EXTRAMURS", "4625012" = "CAMINS AL GRAU", "4625001" = "CIUTAT VELLA",
  "4625006" = "EL PLA DEL REAL", "4625013" = "ALGIROS", "4625002" = "L'EIXAMPLE",
  "4625010" = "QUATRE CARRERES", "4625016" = "BENICALAP", "4625004" = "CAMPANAR",
  "4625007" = "L'OLIVERETA", "4625018" = "POBLATS DE L'OEST", "4625009" = "JESUS",
  "4625014" = "BENIMACLET", "4625015" = "RASCANYA", "4625008" = "PATRAIX",
  "4625017" = "POBLATS DEL NORD"
)

# Convierte el vector asociativo en un dataframe de dos columnas y construye la zona geográfica.
df_map_ciudades <- data.frame(
  CLAVE_GEO = names(mapeo_completo), # Extrae las claves numéricas como primera columna.
  Zona_Txt = mapeo_completo,         # Asigna los nombres textuales como segunda columna.
  stringsAsFactors = FALSE           # Evita la conversión obsoleta de textos a factores categóricos.
) %>%
  # Evalúa los prefijos de los códigos INE para agrupar los barrios en su provincia/ciudad correspondiente.
  # Es indispensable tener una etiqueta de ciudad explícita para poder realizar los "group_by(Ciudad)" locales en la Etapa 4.
  mutate(Ciudad = case_when(
    str_starts(CLAVE_GEO, "08019") ~ "Barcelona", # CUMUN/CUDIS específico de Barcelona capital.
    str_starts(CLAVE_GEO, "28079") ~ "Madrid",    # CUMUN/CUDIS específico de Madrid capital.
    str_starts(CLAVE_GEO, "46250") ~ "Valencia",  # CUMUN/CUDIS específico de Valencia capital.
    str_starts(CLAVE_GEO, "41091") ~ "Sevilla",   # CUMUN/CUDIS específico de Sevilla capital.
    str_starts(CLAVE_GEO, "29067") ~ "Malaga",    # CUMUN/CUDIS específico de Málaga capital.
    str_starts(CLAVE_GEO, "17")    ~ "Girona",    # Códigos que inician por 17 corresponden a la provincia de Girona.
    str_starts(CLAVE_GEO, "01") | str_starts(CLAVE_GEO, "48") | str_starts(CLAVE_GEO, "20") ~ "Euskadi", # Provincias vascas (Álava, Vizcaya, Guipúzcoa).
    CLAVE_GEO %in% c("07024", "07015", "07002", "07032", "07064", "07052", "07023", "07902") ~ "Menorca", # Segmentación explícita para municipios de Menorca.
    str_starts(CLAVE_GEO, "07") ~ "Mallorca",   # El resto de códigos de Baleares se consolidan en Mallorca.
    TRUE ~ NA_character_                        # Cláusula de seguridad para capturar registros fuera de rango.
  ))

# Función constructora para extraer, limpiar y parsear datos estructurados de hojas del INE.
# Automatiza la limpieza repetitiva de las tablas del INE (filas de metadatos, cabeceras, espacios).
extraer_datos_ine <- function(ruta_archivo, col_valor, longitud_codigo, nombre_metrica) {
  df_raw <- read_excel(ruta_archivo, col_names = FALSE) # Lee el archivo.
  
  df_raw %>%
    mutate(Texto_Celda = str_trim(as.character(...1))) %>% # Convierte a texto la primera columna y elimina espacios en blanco.
    filter(str_detect(Texto_Celda, "^\\d+")) %>%  # Conserva únicamente las filas cuyo primer elemento empiece por números (registros reales).
    mutate(RAW_CLAVE = str_extract(Texto_Celda, "^\\d+")) %>% # Extrae de forma aislada el código numérico inicial del municipio/distrito.
    filter(str_length(RAW_CLAVE) == longitud_codigo) %>% # Validación estricta de longitud (7 para distritos, 5 para municipios).
    transmute(
      CLAVE_GEO = RAW_CLAVE, # Almacena la clave estandarizada.
      !!nombre_metrica := as.numeric(.data[[names(df_raw)[col_valor]]]) # Evalúa e inyecta dinámicamente la columna numérica objetivo si es necesario (!!).
    ) %>%
    filter(!is.na(!!sym(nombre_metrica))) # Purga registros que no contengan datos válidos en la métrica extraída.
}

# Extracción directa de las variables demográficas base para los municipios.
# Proporciona la población y total de hogares monoparentales para calcular los denominadores de densidad y tasas de vulnerabilidad.
df_pob_base_mun <- extraer_datos_ine("población.xlsx", 2, 5, "Pob_Residente") %>% distinct(CLAVE_GEO, .keep_all = TRUE)
df_hog_base_mun <- extraer_datos_ine("hogares1persona.xlsx", 3, 5, "Hogares_1Pers") %>% distinct(CLAVE_GEO, .keep_all = TRUE)

# ==============================================================================
# ETAPA 1: PROCESAMIENTO SOCIOECONÓMICO A NIVEL DE DISTRITOS (7 DÍGITOS)
# ==============================================================================

# Extrae la renta media por hogar de las grandes capitales usando códigos de 7 dígitos.
df_rentas_dist <- extraer_datos_ine("rentas.xlsx", 11, 7, "Renta_Hogar") %>% distinct(CLAVE_GEO, .keep_all = TRUE)

# Importa los datos oficiales de precios de alquiler del Ministerio de Vivienda (MIVAU).
df_mivau_dist <- read_excel("precio-alquiler.xlsx", sheet = "Distritos") %>%
  select(CUDIS, ALQM2_LV_M_VC_19, ALQM2_LV_M_VC_24) %>% # Selecciona el código del distrito y los precios medios de 2019 y 2024.
  mutate(CLAVE_GEO = str_pad(as.character(CUDIS), width = 7, side = "left", pad = "0")) %>% # Rellena con ceros a la izquierda hasta asegurar las 7 cifras.
  select(-CUDIS) # Pasa a usar la clave homogénea y elimina la columna original redundante.

# Importa y parsea la población residente oficial registrada en los distritos.
df_pob_real_dist <- read_excel("poblacionDistritos.xlsx") %>%
  mutate(CLAVE_GEO = str_pad(as.character(CUDIS), width = 7, side = "left", pad = "0")) %>% # Relleno homogéneo de 7 cifras.
  transmute(CLAVE_GEO, Pob_Residente = as.numeric(Total)) %>% # Fuerza conversión a numérico para operaciones matemáticas.
  distinct(CLAVE_GEO, .keep_all = TRUE) # Elimina duplicaciones accidentales en el censo.

# Importa y aísla la cantidad de hogares de un solo habitante a nivel de distrito.
df_hog_real_dist <- read_excel("hogaresDistritos.xlsx") %>%
  mutate(CLAVE_GEO = str_pad(as.character(CUDIS), width = 7, side = "left", pad = "0")) %>%
  transmute(CLAVE_GEO, Hogares_1Pers = as.numeric(`1 Persona`)) %>% 
  distinct(CLAVE_GEO, .keep_all = TRUE)

# Unificación estructural para la capa de distritos urbanos.
# Junta todas las fuentes dispersas (Renta, Alquiler MIVAU, Censos) mediante un identificador espacial común.
df_socio_dist_total <- df_rentas_dist %>%
  inner_join(df_mivau_dist, by = "CLAVE_GEO") %>% # Cruza datos económicos del alquiler.
  left_join(df_pob_real_dist, by = "CLAVE_GEO") %>% # Añade población manteniendo registros base.
  left_join(df_hog_real_dist, by = "CLAVE_GEO") %>% # Añade estructuras de hogar monoparental.
  inner_join(df_map_ciudades, by = "CLAVE_GEO") # Incorpora los nombres de zonas textuales y etiquetas de Ciudad.

# ==============================================================================
# ETAPA 2: PROCESAMIENTO SOCIOECONÓMICO A NIVEL DE MUNICIPIOS (5 DÍGITOS)
# ==============================================================================

# Extrae los indicadores de renta media por hogar para los municipios completos (código de 5 cifras).
df_rentas_mun <- extraer_datos_ine("rentas.xlsx", 11, 5, "Renta_Hogar") %>% distinct(CLAVE_GEO, .keep_all = TRUE)

# Importa los precios oficiales de alquiler de la pestaña municipal del Ministerio.
df_mivau_mun <- read_excel("precio-alquiler.xlsx", sheet = "Municipios") %>%
  select(CUMUN, ALQM2_LV_M_VC_19, ALQM2_LV_M_VC_24) %>%
  mutate(CLAVE_GEO = str_pad(as.character(CUMUN), width = 5, side = "left", pad = "0")) %>% # Relleno homogéneo a 5 dígitos para pueblos y municipios.
  select(-CUMUN)

# Unificación para la capa municipal.
# Al igual que en la etapa anterior, junta las fuentes pero a una escala territorial más agregada.
df_socio_mun_total <- df_pob_base_mun %>%
  inner_join(df_rentas_mun, by = "CLAVE_GEO") %>%
  inner_join(df_hog_base_mun, by = "CLAVE_GEO") %>%
  inner_join(df_mivau_mun, by = "CLAVE_GEO") %>%
  inner_join(df_map_ciudades, by = "CLAVE_GEO")

# Unificación vertical (concatenación de filas) de ambas capas territoriales.
# Consolida en una única matriz maestra de datos socioeconómicos tanto las grandes ciudades segmentadas como los municipios autónomos pequeños.
df_socio_total_limpio <- bind_rows(df_socio_dist_total, df_socio_mun_total)

# ==============================================================================
# ETAPA 3: NORMALIZACIÓN MULTIFORMATO DE COORDENADAS AIRBNB
# ==============================================================================

# Función para corregir la latitud celda por celda mediante reducción iterativa.
# Corrige de forma adaptativa anomalías de padding o ausencia de punto decimal (ej: convierte 4041556 en 40.41556).
clean_lat_cell <- function(lat) {
  lat_num <- as.numeric(lat)
  if (is.na(lat_num)) return(NA_real_) # Control de nulos.
  
  # La latitud geográfica de España se sitúa en un rango estricto entre 35 y 44 grados.
  # El bucle divide el número entre 10 sucesivamente hasta forzar que caiga en este rango real.
  while (abs(lat_num) > 45 && lat_num != 0) {
    lat_num <- lat_num / 10
  }
  return(lat_num)
}

# Función para corregir la longitud celda por celda mediante reducción iterativa.
clean_lon_cell <- function(lon) {
  lon_num <- as.numeric(lon)
  if (is.na(lon_num)) return(NA_real_)
  
  # La longitud en el España se situa dentro del rango de -10 a 5 grados decimales.
  # El bucle divide el número entre 10 sucesivamente hasta forzar que caiga en este rango real.
  while (abs(lon_num) > 10 && lon_num != 0) {
    lon_num <- lon_num / 10
  }
  return(lon_num)
}

# Importación cartográfica de Airbnb.
hojas_airbnb <- excel_sheets("airbnb.xlsx") # Extrae los nombres de las pestañas reales del archivo.
todas_las_ciudades <- c('Madrid', 'Barcelona', 'Valencia', 'Sevilla', 'Malaga', 'Girona', 'Euskadi', 'Mallorca', 'Menorca')
airbnb_consolidado_list <- list() # Estructura de almacenamiento temporal tipo lista.

# Bucle sobre cada mercado turístico registrado en las pestañas del Excel.
for (ciudad in todas_las_ciudades) {
  if (ciudad %in% hojas_airbnb) {
    
    # Carga los alojamientos y aplica un filtro restrictivo sobre tipología de oferta.
    # El estudio se centra exclusivamente en viviendas completas ("Entire home/apt") por ser las que ejercen 
    # una presión directa de sustitución habitacional sobre la población residencial, descartando habitaciones sueltas.
    df_air <- read_excel("airbnb.xlsx", sheet = "Madrid") %>%
      filter(room_type == "Entire home/apt")
    
    # Invoca el limpiador funcional fila por fila sobre las coordenadas originales.
    # Al usar 'map_dbl' de purrr, R evalúa celda por celda de forma aislada, para asegurar que cada cifra tenga el formato correcto.
    df_air <- df_air %>%
      mutate(
        latitude_clean = map_dbl(latitude, clean_lat_cell),
        longitude_clean = map_dbl(longitude, clean_lon_cell)
      )
    
    # Ajuste de escala explícito exclusivo para la longitud de Valencia.
    # Tras eliminar los dígitos masivos, la longitud de Valencia quedaba desplazada un decimal (ej: -3.46 en vez de -0.34), al quedar muy próximo a 0.
    if (ciudad == "Valencia") {
      df_air <- df_air %>%
        mutate(longitude_clean = if_else(abs(longitude_clean) > 1, longitude_clean / 10, longitude_clean))
    }
    
    # Estructura condicional para seleccionar la columna de agrupación según el tipo de territorio.
    # Las ciudades grandes se agrupan por distritos municipales ("neighbourhood_group"), mientras que las provincias o 
    # islas se analizan directamente agregando a nivel de municipio o barrio local ("neighbourhood").
    col_agrupacion <- ifelse(ciudad %in% c("Malaga", "Girona", "Euskadi", "Mallorca", "Menorca"), "neighbourhood", "neighbourhood_group")
    
    # Agrupación espacial y cálculo de estadísticas descriptivas de la oferta turística.
    df_air_agrupado <- df_air %>%
      group_by(Zona_Txt = .data[[col_agrupacion]]) %>% # Agrupa dinámicamente según la regla territorial previa.
      summarise(
        Viviendas_Tur_Total = n(), # Conteo absoluto de ofertas turísticas vigentes en el entorno urbano.
        
        # Sumatorio condicional basado en el volumen de anuncios por anfitrión.
        # Se considera que un anfitrión con más de 3 anuncios operativos corresponde 
        # a un perfil de gestión profesional.
        Viviendas_Prof_Total = sum(calculated_host_listings_count > 3, na.rm = TRUE),
        
        # Cálculo del centroide medio geográfico de los puntos limpios normalizados.
        # Genera las coordenadas únicas promedio requeridas por Flourish para situar con precisión el marcador físico en el mapa.
        Latitud_Media = mean(latitude_clean, na.rm = TRUE),
        Longitud_Media = mean(longitude_clean, na.rm = TRUE),
        .groups = "drop" # Deshace la agrupación interna para liberar memoria de procesamiento.
      ) %>% 
      mutate(Ciudad = "Madrid")
    
    airbnb_consolidado_list[[ciudad]] <- df_air_agrupado # Guarda los resultados limpios del territorio actual en la lista maestra.
  }
}

# Consolidación unificada de toda la oferta turística nacional de Airbnb en una única matriz plana.
df_airbnb_total <- bind_rows(airbnb_consolidado_list)

# ==============================================================================
# ETAPA 4: CRUCE FINAL, NORMALIZACIÓN LOCAL (NUMÉRICA Y TEXTO) Y EXPORTACIÓN
# ==============================================================================

message(">>> Ejecutando cruce final y generando métricas relativas locales...")

df_resultado_final <- df_airbnb_total %>%
  # Realiza un cruce relacional estricto (INNER JOIN) entre Airbnb y la matriz socioeconómica oficial (INE/MIVAU).
  # Descarta automáticamente cualquier zona que no tenga correspondencia estadística oficial censada.
  inner_join(df_socio_total_limpio, by = c("Ciudad", "Zona_Txt")) %>%
  
  # Bloque para la construcción de los indicadores definitivos del modelo matemático.
  mutate(
    # ID único en mayúsculas para indexación limpia.
    ID_UNICO = str_c(str_to_upper(Ciudad), "_", str_to_upper(Zona_Txt)),
    
    # Cálculo de la variación porcentual acumulada del precio del alquiler por metro cuadrado entre 2019 y 2024.
    # Captura la velocidad y la tasa de encarecimiento del mercado de arrendamiento residencial a medio plazo.
    Variacion_Alquiler_Pct = round(((ALQM2_LV_M_VC_24 - ALQM2_LV_M_VC_19) / ALQM2_LV_M_VC_19) * 100, 2),
    
    # Tasa de densidad turística: volumen de alojamientos de Airbnb por cada 1.000 ciudadanos residentes.
    # Relativiza la presencia turística en función de la población fija, reflejando el grado de saturación espacial real.
    Densidad_Turistica = round((Viviendas_Tur_Total / Pob_Residente) * 1000, 2),
    
    # Cálculo del grado de comercialización del mercado turístico de alquiler de la zona.
    Pct_Profesional = round((Viviendas_Prof_Total / Viviendas_Tur_Total) * 100, 2),
    
    # CÁLCULO DEL IPG GLOBAL ABSOLUTO.
    # Es la métrica principal. Multiplica la densidad por la variación de precios para capturar el 
    # efecto multiplicador: la presión es máxima donde hay saturación de viviendas vacacionales Y encarecimiento desmedido.
    IPG_Indice = round(Densidad_Turistica * Variacion_Alquiler_Pct, 2),
    
    # Tasa de vulnerabilidad sociodemográfica (porcentaje de hogares habitados por una sola persona).
    # Los hogares unipersonales representan estructuras demográficas con menor adaptabilidad financiera ante subidas del alquiler.
    Pct_Vulnerabilidad_G = round((Hogares_1Pers / Pob_Residente) * 100, 2)
  ) %>%
  
  # Aislamiento operativo por ciudad.
  # Segmenta los cálculos posteriores para que las funciones estadísticas operen de forma local e independiente en cada ciudad.
  group_by(Ciudad) %>%
  mutate(
    # FÓRMULA DE ESCALADO MIN-MAX REGIONAL (IPG LOCAL NUMÉRICO).
    # Transforma el IPG absoluto a una escala uniforme de 0 a 100 exclusiva para cada entorno urbano. El distrito más 
    # tensionado de CADA ciudad obtendrá siempre un 100 (máximo tamaño del punto en el mapa de Flourish) y el menor un 0, para equilibraro la representación gráfica.
    IPG_Relativo_Raw = ((IPG_Indice - min(IPG_Indice, na.rm = TRUE)) / 
                          (max(IPG_Indice, na.rm = TRUE) - min(IPG_Indice, na.rm = TRUE))) * 100,
    
    # Tratamiento de seguridad contra la división por cero (NaN).
    # Si una ciudad registra un único barrio o valores idénticos, la resta del denominador daría 0. Con esto lo fijamos a 100 de forma segura.
    IPG_Relativo_Raw = if_else(is.nan(IPG_Relativo_Raw), 100, IPG_Relativo_Raw),
    
    # Redondeo del indicador numérico de tamaño.
    IPG_Local_Num = round(IPG_Relativo_Raw, 2)
  ) %>%
  ungroup() %>% # Disuelve la agrupación local para permitir la reestructuración global del dataset final.
  
  # Conversión de la escala numérica continua en categorías discretas y jerarquizadas de TEXTO.
  # Mapea los rangos del IPG local en etiquetas semánticas inteligibles que Flourish interpretará en su canal de "Color".
  # Se antepone un prefijo numérico ("1.", "2."...) para forzar al motor del mapa a ordenar la leyenda de forma lógica y no alfabética.
  mutate(
    Tension_Turistica_Local = case_when(
      IPG_Local_Num >= 80 ~ "1. Critico (Top local)",
      IPG_Local_Num >= 50 & IPG_Local_Num < 80 ~ "2. Alto",
      IPG_Local_Num >= 20 & IPG_Local_Num < 50 ~ "3. Moderado",
      TRUE ~ "4. Bajo"
    )
  ) %>%
  
  # Selección, formateo y ordenación estructural de la matriz de salida.
  # Deja los campos del dataset final limpios, renombrados y ordenados.
  select(
    ID_UNICO, CLAVE_GEO, Ciudad, Barrio_Distrito_Municipio = Zona_Txt, 
    Latitud_Media, Longitud_Media, Pob_Residente, Renta_Hogar, 
    Pct_Vulnerabilidad_G, Viviendas_Tur_Total, Pct_Profesional, 
    Densidad_Turistica, Variacion_Alquiler_Pct, 
    IPG_Indice, # Métrica absoluta útil para incorporar en textos dinámicos o tooltips.
    IPG_Local_Num, # Métrica numérica normalizada asignable al campo VALUE (SIZE) de Flourish.
    Tension_Turistica_Local # Métrica de texto categórica asignable al campo COLOR de Flourish.
  )

# Escritura física y persistencia del dataframe unificado estructurado en el disco local.
# Genera el archivo definitivo .xlsx limpio y compatible, listo para ser consumido directamente por Flourish sin necesidad de recálculos.
write_xlsx(df_resultado_final, "dataset_visualizacion_final.xlsx")

message(">>> ¡PROCESO COMPLETADO! Archivo compatible con Flourish generado con éxito.")
