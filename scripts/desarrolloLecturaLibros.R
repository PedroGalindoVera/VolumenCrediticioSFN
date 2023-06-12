requerirPaquetes("dplyr")

rutas_libros_excel <-
  list.files("data/Fuentes/SEPS/Reportes/Volumen de Credito Mensual",
             recursive = TRUE, full.names = TRUE)

nombres_hojas_existentes <-
  lapply(rutas_libros_excel, readxl::excel_sheets) %>%
  unlist() %>%
  unique() %>%
  sort()

hojas_por_libro <- data.frame()
for (ruta in rutas_libros_excel) {
  nueva_fila <-
    data.frame(
      Libro = basename(ruta),
      `Año` = gsub(".*(\\d{4}).*", "\\1", basename(ruta)),
      Hojas = paste(sapply(ruta, readxl::excel_sheets), collapse = ", ")
    )
  hojas_por_libro <- hojas_por_libro %>% dplyr::bind_rows(nueva_fila)
}

## OBSERVACIÓN
# Al revisar "hojas_por_libro", corroboramos que las hojas con nombre "VOLUMEN
# DE CRÉDITO HISTORICO", es una hoja que contiene una tabla dinámica que
# resulta a partir de la hoja "Base_vcredito". A su vez, por inspección de
# "nombres_hojas_existentes" elegimos las 7 primeras componentes del vector.

hojas_volumen_crediticio <- head(nombres_hojas_existentes,7)

# Para busqueda ----
# hojas_por_libro <- lapply(rutas_libros_excel, readxl::excel_sheets)
# names(hojas_por_libro) <- basename(rutas_libros_excel)
# 
# hojas_por_libro <- unlist(hojas_por_libro)
# names(hojas_por_libro)[hojas_por_libro == "BASE VOL.CRÉDITO 2013"]
# ----

biblioteca_libros <- lapply(rutas_libros_excel, function(ruta) {
  nombres_hojas <- readxl::excel_sheets(ruta)
  hoja_seleccionada <- intersect(nombres_hojas, hojas_volumen_crediticio)
  libro <- readxl::read_excel(ruta, sheet = hoja_seleccionada)
  attr(libro, "hoja_seleccionada") <- hoja_seleccionada
  return(libro)
  })
names(biblioteca_libros) <- gsub("\\.[^.]*$", "", basename(rutas_libros_excel))

nombres_columnas_en_hojas_volumen_crediticio <-
  lapply(biblioteca_libros, names) %>% unlist() %>% unique() %>% sort()

# Idea para funcion resumenBibliotecaLibros ----
# resumen_biblioteca_libros <- data.frame(DataFrame = names(biblioteca_libros))
# for (nombre_columna in nombres_columnas_en_hojas_volumen_crediticio) {
#   resumen_biblioteca_libros[[nombre_columna]] <-
#     sapply(biblioteca_libros, function(df) {
#       ifelse(test = nombre_columna %in% names(df),
#              yes = class(df[[nombre_columna]]),
#              no = "")
#   })
# }
# ----

resumenBibliotecaLibros <- function(lista_data_frames) {
  library(dplyr)
  biblioteca_libros <- lista_data_frames
  nombres_columnas_en_hojas_volumen_crediticio <-
    lapply(biblioteca_libros, names) %>% unlist() %>% unique() %>% sort()
  resumen_biblioteca_libros <- data.frame(
    Libro = names(biblioteca_libros),
    Hoja = unname(sapply(
      biblioteca_libros, function(df) attr(df, "hoja_seleccionada"))),
    `Año` = gsub(".*(\\d{4}).*", "\\1", names(biblioteca_libros))
    )
  for (nombre_columna in nombres_columnas_en_hojas_volumen_crediticio) {
    resumen_biblioteca_libros[[nombre_columna]] <-
      sapply(biblioteca_libros, function(df) {
        ifelse(test = nombre_columna %in% names(df),
               yes = class(df[[nombre_columna]]),
               no = "")
      })
  }
  return(resumen_biblioteca_libros)
}

resumen_biblioteca_libros <- resumenBibliotecaLibros(biblioteca_libros)

# SOLO HACER DESDE 2016
# LAS SIGUIENTES COLUMNAS YA NO ESTARIAN: "Nivel 1", "Nivel 2", "No,Op", "No.Op"

desde <- 2016
hasta <- as.integer(format(Sys.Date(),"%Y"))
expresion_regular_anios_selecionados <- 
  paste0(seq(desde, hasta), collapse = "|")
indice_data_frame_selecionados <-
  grep(expresion_regular_anios_selecionados, names(biblioteca_libros))

resumen_biblioteca_libros_seleccion <-
  resumenBibliotecaLibros(biblioteca_libros[indice_data_frame_selecionados])

biblioteca_libros_estandarizada <- lapply(
  biblioteca_libros[indice_data_frame_selecionados], function(df) {
  nombres_columnas <- names(df)
  nombres_columnas[
    nombres_columnas %in% 
      c("Actividad Económica", "Actividad Económica ( productivas)", 
        "ACTIVIDAD_ECONOMICA", "Actividades no productivas")] <- 
    "Actividad"
  # Agrega aquí más reglas de estandarización si es necesario
  nombres_columnas[
    nombres_columnas %in% c("CANTON", "Cantón")] <-
    "Cantón"
  nombres_columnas[
    nombres_columnas %in% c("DESTINO_FINANCIERO")] <-
    "Destino Financiero"
  nombres_columnas[
    nombres_columnas %in%c("Estado de Operación","ESTADO_OPERACION")] <-
    "Estado de Operación"
  nombres_columnas[
    nombres_columnas %in% 
      c("Fecha de corte", "Fecha de Corte", "FECHA_CORTE")] <-
    "Fecha"
  nombres_columnas[
    nombres_columnas %in% c("Institución", "RAZON_SOCIAL")] <- 
    "Entidad"#"Razón Social"
  nombres_columnas[
    nombres_columnas %in% c("Provincia", "PROVINCIA")] <- 
    "Provincia"
  nombres_columnas[
    nombres_columnas %in% c("REGION", "Región")] <- 
    "Región"
  nombres_columnas[
    nombres_columnas %in% c("OPERACIONES")] <- 
    "Número de Operaciones"
  nombres_columnas[
    nombres_columnas %in% c("NUM_RUC")] <-
    "RUC"
  nombres_columnas[
    nombres_columnas %in% c("SEGMENTO")] <-
    "Segmento"
  nombres_columnas[
    nombres_columnas %in% c("SUJETOS DE CREDITO", "SUJETOS DE CREDITOS")] <-
    "Sujetos de Crédito"
  nombres_columnas[
    nombres_columnas %in%
      c("Tipo de Crédito", "TIPO DE CRÉDITO GENERAL", "TIPO_CREDITO")] <-
    "Tipo de Crédito"
  nombres_columnas[
    nombres_columnas %in%
      c("TIPO DE CRÉDITO ESPECÍFICO", "TIPO_CREDITO_nuevo")] <-
    "Tipo de Crédito Específico"
  nombres_columnas[
    nombres_columnas %in% c("Monto", "VAL_OPERACION")] <-
    "Valor"
  names(df) <- nombres_columnas
  df
})

resumen_biblioteca_libros_estandarizada <-
  resumenBibliotecaLibros(biblioteca_libros_estandarizada)

biblioteca_libros_corregida <- lapply(
  biblioteca_libros_estandarizada, function(df) {
    if ("Fecha de Corte" %in% names(df)) {
      df$`Fecha de Corte` <- as.Date(df$`Fecha de Corte`, origin = "1899-12-30")
    }
    df$Columna1 <- NULL
    df
})

resumen_biblioteca_libros_corregida <-
  resumenBibliotecaLibros(biblioteca_libros_corregida)

tabla_concatenada_SEPS <- bind_rows(biblioteca_libros_corregida)
