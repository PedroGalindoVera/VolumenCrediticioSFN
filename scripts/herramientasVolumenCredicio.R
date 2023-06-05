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

crearDirectorio <- function(ruta_directorio) {
  
  # Esta función permite crear cualesquier ruta especificada, dentro del directorio del proyecto.
  
  # EJEMPLO: crearDirectorio("data/Fuente/SB/PRIVADA")
  
  existe_ruta <- dir.exists(ruta_directorio)
  
  if ( !existe_ruta  ) {
    dir.create(ruta_directorio, recursive = TRUE)
    cat("\n\033[1mSe creo la carpeta:\033[0m [",basename(ruta_directorio),"]",
        "con la ruta: [", normalizePath(ruta_directorio),"].\n")
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

exportarReporteTabla <- function(dataFrame, nombre_archivo) {
  requerirPaquetes("openxlsx")
  creear_libro_trabajo <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(creear_libro_trabajo, "Reporte")
  openxlsx::writeData(creear_libro_trabajo, "Reporte", dataFrame) #, autoWidth = TRUE)
  directorio_reportes <- "data/Reportes"
  crearDirectorio(directorio_reportes)
  nombre_archivo <-
    paste0(nombre_archivo, format(Sys.time(), " %Y-%m-%d_%HH%M.xlsx"))
  ruta_archivo <- file.path(directorio_reportes, nombre_archivo)
  openxlsx::saveWorkbook(creear_libro_trabajo, ruta_archivo, overwrite = TRUE)
  #openxlsx::write.xlsx(informacion, file.path(directorio_reportes,paste("Reporte Enlaces de Descarga",format(Sys.Date(), "%Y-%m-%d.xlsx"))), rowNames = FALSE)
  cat("\nSe ha creado el archivo con la ruta: [", normalizePath(ruta_archivo), "]\n")
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
  
  requerirPaquetes("httr","dplyr")
  
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
    
    informacion <- informacion %>% dplyr::bind_rows(nueva_fila)
    
    if ( indice == 1 ) { cat("\nRutas de descarga:") }
    barraProgreso(enlaces)
    cat("\033[1;32mObteniendo ruta de descarga...\033[0m\n")
    cat("Del vínculo:\n\t[", enlace, "]",
        "\nse ha capturado la ruta de descarga:\n\t[", head$url, "].\n")
  }
  
  cat("\n\n\033[1mResumen:\033[0m\n")
  print(informacion)
  
  exportarReporteTabla(
    dataFrame =  informacion,
    nombre_archivo = paste("Reporte Enlaces de Descarga", identificador))
  
  #actualizarReporte(dataFrame = informacion, nombre_archivo = paste("Reporte Enlaces de Descarga", identificador))
  
  return(informacion)
}

descargarArchivosEnlacesAnalizados <- function(enlaces, informacion, ruta_destino) {
  
  crearDirectorio(ruta_destino)
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

descomprimirArchivosDirectorioZip <- function(origen, destino) {
  
  # EJEMPLO:
  # origen <- "data/Descargas/SEPS/Bases de Datos"
  # destino <- "data/Fuente/SEPS/Bases de Datos"
  # descomprimirArchivosDirectorioZip(origen, destino)
  
  requerirPaquetes("utils")
  descompresionZip <- function(ruta_origen, directorio_destino) {
    #if ( !dir.exists(directorio_destino) ) 
    crearDirectorio(directorio_destino)
    tryCatch(
      {
        utils::unzip(ruta_origen, exdir = directorio_destino)
      },
      error = function(e) {
        message("Ocurrió un error al descomprimir el archivo zip: ", e$message,
                "\nEmpleando 7-Zip para completar la descompresión...")
        # Código para manejar el error, utilizando una herramienta externa para descomprimir el archivo zip
        ruta_origen_normalizado <- normalizePath(ruta_origen)
        ruta_destino_normalizado <- normalizePath(directorio_destino)
        # Descompresión externa de archivos
        descompresion7zip(ruta_origen_normalizado, ruta_destino_normalizado)
      }
    )
  }
  archivoZip <- function() {
    #archivos_contenidos <- utils::unzip(ruta_origen, list = TRUE)$Name
    hay_contenidos_zip <- any(grepl("\\.zip$", archivos_contenidos))
    if ( hay_contenidos_zip ) {
      directorio_destino_temporal <-
        dirname(gsub("Descargas","Temporal",ruta_origen))
      #if ( !dir.exists(directorio_destino_temporal) )
      crearDirectorio(directorio_destino_temporal)
      descompresionZip(ruta_origen, directorio_destino_temporal)
      ruta_comprimido_temporal <-
        file.path(directorio_destino_temporal, archivos_contenidos)
      directorio_origen_temporal <- directorio_destino_temporal
      descomprimirArchivosDirectorioZip(origen = directorio_origen_temporal,
                                        destino = directorio_destino)
      unlink(ruta_comprimido_temporal, recursive = TRUE)
      #file.remove(ruta_comprimido_temporal)
    } else {
      descompresionZip(ruta_origen, directorio_destino)
    }
  }
  copiarArchivo <- function() {
    ruta_verificacion <- file.path(destino, archivo)
    if ( !file.exists(ruta_verificacion) ) {
      cat("\nCopiando el archivo: [", normalizePath(ruta_origen),"] ...\n")
      file.copy( ruta_origen, ruta_verificacion )
    }
  }
  decidirAccion <- function() {
    tiene_extension_zip <- grepl("\\.zip$", ruta_origen)
    if ( tiene_extension_zip ) {
      archivoZip()
    } else {
      copiarArchivo()
    }
  }
  
  archivos_origen <- list.files(origen, recursive = TRUE)
  # Elegimos únicamente los archivos con extensión zip
  archivos <- grep("\\.zip$", archivos_origen, value = TRUE)
  for ( archivo in archivos ) {
    ruta_origen <- file.path( origen, archivo )
    # Establecer el directorio de destino para los archivos Descomprimidos
    va_a_data_fuente_SB <-
      grepl("(?=.*data)(?=.*Fuente)(?=.*SB)", destino, perl = TRUE)
    if ( va_a_data_fuente_SB ) {
      directorio_destino <-
        dirname(gsub("Descargas|Temporal","Fuente",ruta_origen))
    } else {
      nombre_archivo_zip <- gsub("\\.zip$","",basename(ruta_origen))
      directorio_destino <- file.path(destino, nombre_archivo_zip)
    }
    #if ( !dir.exists(directorio_destino) )
    crearDirectorio(directorio_destino)
    archivos_contenidos <-
      utils::unzip(ruta_origen, list = TRUE)$Name
    alguno_de_los_caracteres_no_es_utf8 <-
      any(is.na(unlist(sapply(archivos_contenidos, utf8ToInt))))
    if ( !alguno_de_los_caracteres_no_es_utf8 ) {
      ruta_archivos_descomprimidos <-
        file.path(directorio_destino, archivos_contenidos)
      existe_archivo_descomprimido <-
        all(file.exists(ruta_archivos_descomprimidos))
    } else {
      existe_archivo_descomprimido <- FALSE
    }
    if ( !existe_archivo_descomprimido ) {
      decidirAccion()
      barraProgreso(archivos)
      cat("\033[1;32mDescomprimiendo el archivo:\033[0m [", normalizePath(ruta_origen), "]\n")
    }
  }
  cat("\n")
}

# SEPS----

gestorDescargasDescompresionSEPS <- function() {# Verificado en prueba individual 2023/06/05
  
  enlaces_SEPS <- analisisVinculosPaginaWebSEPS()
  
  info_enlaces_SEPS <-
    obtenerEnlacesDescarga(
      enlaces_descarga = enlaces_SEPS,
      identificador = "SEPS Volumen Crediticio")
  
  ruta_descargas_SEPS <-
    "data/Descargas/SEPS/Reportes/Volumen de Credito Mensual"
  
  descargarArchivosEnlacesAnalizados(
    enlaces_SEPS, info_enlaces_SEPS, ruta_descargas_SEPS)
  
  ruta_fuentes_SEPS <- "data/Fuentes/SEPS/Reportes/Volumen de Credito Mensual"
  
  descomprimirArchivosDirectorioZip(ruta_descargas_SEPS, ruta_fuentes_SEPS)
  
}
