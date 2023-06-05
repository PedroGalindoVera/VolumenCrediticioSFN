# General----

requerirPaquetes <- function(...) {
  lista_paquetes_empleados <-
    list(
      "beepr",
      "data.table", "dplyr",
      "httr",
      "lubridate",
      "openxlsx",
      "parallel", "parsedate", "purrr",
      "readr", "readxl", "reshape2", "rlang", "rvest",
      "stats", "stringdist", "stringr",  
      "tools",
      "utils"
    )
  lista_paquetes <-
    if ( length(list(...)) == 0 ) {
      lista_paquetes_empleados
    } else {
      list(...)
    }
  paquetes <- unlist(lista_paquetes)
  for (paquete in paquetes) {
    if ( !require(paquete, character.only = TRUE) ) {
      install.packages(paquete)
      library(paquete, character.only = TRUE)
    }
  }
}

formatoTiempoHMS <- function(tiempo) {
  tiempo <- as.numeric(tiempo)
  tiempo <- as.POSIXct(tiempo, origin = "1970-01-01", tz = "UTC")
  tiempo <- format(tiempo,"%H:%M:%S")
  return(tiempo)
}

barraProgreso <- function(conjunto) {
  barra_progreso <- txtProgressBar(min = 0, max = length(conjunto), style = 3)
  numero_elementos <- length(conjunto)
  if ( exists("contador_progreso") ) {
    setTxtProgressBar(barra_progreso, contador_progreso)
    marcador_progreso_cronometro <- Sys.time()
    tiempo_transcurrido <-
      difftime(
        marcador_progreso_cronometro,
        marcador_inicio_cronometro,
        units = "sec")
    estimador_tiempo_proceso <-
      numero_elementos*(tiempo_transcurrido/(contador_progreso))
    cat("\n  \033[34mTiempo transcurrido:\033[0m",
        "\033[1;34m",formatoTiempoHMS(tiempo_transcurrido),"\033[0m", " de ",
        "\033[1;34m", formatoTiempoHMS(estimador_tiempo_proceso),"\033[0m")
    contador_progreso <<- contador_progreso + 1
  } else {
    marcador_inicio_cronometro <<- Sys.time()
    contador_progreso <<- 1
    setTxtProgressBar(barra_progreso, contador_progreso)
  }
  cat(paste0("\n[",contador_progreso,"] "))
  if ( contador_progreso == length(conjunto) ) {
    close(barra_progreso)
    rm(contador_progreso, envir = .GlobalEnv)
    rm(marcador_inicio_cronometro, envir = .GlobalEnv)
  }
}

barraProgresoReinicio <- function() {
  if (exists("contador_progreso"))
    rm(contador_progreso, envir = .GlobalEnv)
  if (exists("marcador_inicio_cronometro"))
    rm(marcador_inicio_cronometro, envir = .GlobalEnv)
}

# Descarga----

analisisVinculosPaginaWebSEPS <- function() {
  
  requerirPaquetes("dplyr","rvest")
  
  # protocolo_dominio <- "https://estadisticas.seps.gob.ec"
  # ruta <- "/estadisticas-sfps/"
  # link <- paste0(protocolo_dominio, ruta)
  link <- "https://estadisticas.seps.gob.ec/index.php/estadisticas-sfps/"
  selectorCSS_volumen_credito_mensual <- "#collapse_9" # verificado 2023/05/06
  pagina <- rvest::read_html(link)
  entorno_enlaces_descarga <-
    pagina %>%
    rvest::html_nodes(selectorCSS_volumen_credito_mensual)
  enlaces_descarga <-
    rvest::html_nodes(entorno_enlaces_descarga,"a") %>%
    rvest::html_attr("href")
  
  cat("\n\033[1mEnlaces de descarga encontrados en la página:\033[0m [",link,"]\n")
  print(enlaces_descarga)
  
  return(enlaces_descarga)
}

obtenerEnlacesDescarga <- function(enlaces_descarga, identificador) {
  
  requerirPaquetes("httr")
  
  enlaces <- enlaces_descarga
  
  barraProgresoReinicio()
  
  informacion <- data.frame()
  
  for (enlace in enlaces) {
    indice <- match(enlace, enlaces)
    
    head <- httr::HEAD(enlace)
    
    
    nueva_fila <-
      data.frame(
        time = Sys.time(),
        link = enlace,
        url = head$url,
        status_code = head$status_code,
        content_type = head$headers$`content-type`,
        last_modified =
          ifelse(!is.null(head$headers$`last-modified`),
            head$headers$`last-modified`, NA),
        content_length =
          round(as.numeric(head$headers$`content-length`) / 2^20, 2)
      )
    
    informacion <- informacion %>% bind_rows(nueva_fila)
    
    if ( indice == 1 ) { cat("\nRutas de descarga:") }
    barraProgreso(enlaces)
    cat("\033[1;32mObteniendo ruta de descarga...\033[0m\n")
    cat("Del vínculo:\n\t[", enlace, "]",
        "\nse ha capturado la ruta de descarga:\n\t[", head$url, "].\n")
  }
  
  cat("\n\n\033[1mResumen:\033[0m\n")
  print(informacion)
  
  exportarReporteTabla(dataFrame =  informacion, nombre_archivo = paste("Reporte Enlaces de Descarga", identificador))
  
  #actualizarReporte(dataFrame = informacion, nombre_archivo = paste("Reporte Enlaces de Descarga", identificador))
  
  return(informacion)
}

descargarArchivosEnlacesAnalizados <- function(enlaces, informacion, ruta_destino) {
  
  if ( !dir.exists(ruta_destino) ) dir.create(ruta_destino, recursive = TRUE)
  url <- informacion$url
  status <- informacion$status_code
  barraProgresoReinicio()
  cat("\n\n")
  for ( k in seq_along(url) ) {
    archivo_destino <- file.path(ruta_destino, basename(url[k]))
    if ( (!file.exists(archivo_destino) || 
          grepl(format(Sys.Date(), "%Y"),url[k])
          ) && status[k] == 200 ) {
      # El argumento `timeout = 300` indica que R esperará hasta 300 segundos antes de cancelar la descarga si no recibe una respuesta del servidor.
      download.file(url[k], archivo_destino, timeout = 300)
    } else if ( file.exists(archivo_destino) ) {
      cat("\nEl archivo: [", basename(url[k]),"]",
          "ya existe en el directorio: [", archivo_destino, "].\n")
    } else if ( status[k] == 404 ) {
      cat("\nEl archivo: [", basename(url[k]),"]"
          ,"NO está disponible en la dirección: [", url[k], "].\n")
    }
    barraProgreso(seq_along(url))
    cat("\033[1;32mAnalizando descarga:\033[0m ")
  }
  cat("\n\n")
}

descargarArchivosEnlacesAnalizados(enlaces_descarga, informacion, "data")

# SEPS----

gestorDescargasDescompresionSEPS <- function() {
  
  enlaces_SEPS <- analisisVinculosPaginaWebSEPS()
  
  info_enlaces_SEPS <-
    obtenerEnlacesDescarga(
      enlaces_descarga = enlaces_SEPS, identificador = "SEPS")
  
  ruta_descargas_SEPS <-
    "data/Descargas/SEPS/Bases de Datos/Estados Financieros"
  
  descargarArchivosEnlacesAnalizados(
    enlaces_SEPS, info_enlaces_SEPS, ruta_descargas_SEPS)
  
  ruta_fuentes_SEPS <- "data/Fuente/SEPS/Bases de Datos/Estados Financieros"
  descomprimirArchivosDirectorioZip(ruta_descargas_SEPS, ruta_fuentes_SEPS) # Verificado en prueba individual 2023/05/09
  
}
