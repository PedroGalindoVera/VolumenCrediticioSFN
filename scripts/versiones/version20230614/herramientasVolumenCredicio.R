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

exportarResultadosCSV <- function(tabla, nombre_archivo, ruta_directorio = NULL) {
  requerirPaquetes("beepr")
  #cat("\n\nExportando resultados...\n")
  cat("\n\n\033[1;32mExportando resultados...\033[0m\n")
  if ( is.null(ruta_directorio) ) {
    ruta_directorio <- "data/Base de Datos"
  }
  indice_columna_fecha <-
    head(grep("fecha", names(tabla), ignore.case = TRUE), 1)
  fecha_maxima_tabla <-
    max(as.Date(tabla[[indice_columna_fecha]]), na.rm = TRUE)
  fecha <-
    if (length(fecha_maxima_tabla) > 0) {
      fecha_maxima_tabla
    } else {
      as.character(Sys.Date())
    }
  nombre_archivo <- paste0(nombre_archivo, " ", fecha, ".csv")
  ruta_directorio_normalizada <- normalizePath(ruta_directorio)
  dir.create(ruta_directorio_normalizada,recursive = TRUE,showWarnings = FALSE)
  if ( dir.exists(ruta_directorio_normalizada) ) {
    ruta_archivo_normalizada <-
      paste0(ruta_directorio_normalizada, "\\", nombre_archivo)
    data.table::fwrite(tabla, ruta_archivo_normalizada)
    beepr::beep(8)
    cat("\nSe ha creado el archivo con la ruta: [", ruta_archivo_normalizada, "]\n")
  }
}

verificarInstalacion <- function(ruta_intalacion) {
  
  # Esta función ejecuta un comando de PowerShell para buscar la ubicación del ejecutable de Excel en el Registro de Windows
  
  # EJEMPLO:
  # ruta_intalacion <- "'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe'"
  # verificarInstalacion(ruta_intalacion)
  # verificarInstalacion("'C:\\Program Files\\7-Zip\\7zFM.exe'")
  
  # Script de Power Shell para verificar instalción en determinada ruta
  script <- paste0("Test-Path ", ruta_intalacion)
  # Ejecutar el script de PowerShell y capturar el resultado
  is_installed <- system2("powershell", script, stdout = TRUE)
  
  return(as.logical(is_installed))
}

cerrarLibroExcel<- function(ruta_libro) {
  
  # EJMPLO:
  # ruta_libro <- "D:\\INNOVACION\\PASANTE\\DESARROLLO\\BalanceFinacieroSFN\\data\\mtcars.csv"
  # cerrarLibroExcel(ruta_libro)
  
  # Cerrarmos toda la aplicación
  # Get-Process excel | Foreach-Object { $_.CloseMainWindow() }
  
  ruta_libro_normalizada <- normalizePath(ruta_libro)
  
  if ( length(ruta_libro_normalizada) < 1  ) {
    cat("\nIngrese una ruta válida")
  } else if ( length(ruta_libro_normalizada) == 1 ) {
    rutas_cerrar <- paste0("'",ruta_libro_normalizada,"'")
  } else {
    rutas_cerrar <- paste(paste0("'",ruta_libro_normalizada,"'"), collapse = ",")
  }
  # Script para cerrar el libro de Excel especificado empleando Powershell
  script <- paste0("$filesToClose = @(", rutas_cerrar, ");",
                   " ([Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')).Workbooks",
                   " | Where-Object { $filesToClose -contains $_.FullName }",
                   " | ForEach-Object { $_.Close() }")
  # Ejecucion del script con Powershell
  system2("powershell", script)
}

xlsb2xlsx <- function(ruta_archivo_xlsb) {
  
  # Esta función permite transformar el formato de un archivo de Excel con extención ".xlsb" a ".xlsx" y reemplazarlo empleando Windows PowerShell
  
  # EJEMPLO:
  # ruta_archivo_xlsb <- "data/Fuente/Casos Particulares/BOL_FIN_PUB_SEPT_20.xlsb"
  # xlsb2xlsx(ruta_archivo_xlsb)
  
  requerirPaquetes("tools")
  
  verificadorExcel <- function() {
    # Ruta genérica para que PowerShell encuentre a Excel
    ruta_excel <- "'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe'"
    return ( verificarInstalacion(ruta_excel) )
  }
  verificadorArchivo <- function(ruta_archivo_xlsb) {
    if (file.exists(ruta_archivo_xlsb)) {
      return(TRUE)
    } else {
      cat(paste("\nLa ruta no existe\n"))
      return(FALSE)
    }
  }
  verificadorFormato_xlsb <- function(ruta_archivo_xlsb) {
    if (tools::file_ext(ruta_archivo_xlsb) == "xlsb") {
      return(TRUE)
    } else {
      cat(paste("\nEl archivo con ruta: [", ruta_archivo_xlsb,"],",
                "no es de formato \".xlsb\" por lo que no se realizaron cambios\n"))
      return(FALSE)
    }
    
  }
  trasnformador2xlsb <- function(ruta_archivo_xlsb) {
    # Ruta del archivo xlsb
    xlsbFile <- normalizePath(ruta_archivo_xlsb)
    
    # Ruta del archivo xlsx
    xlsxFile <- normalizePath(gsub(".xlsb", ".xlsx",ruta_archivo_xlsb))
    
    # Crear el script de PowerShell
    script <- paste(
      # Ruta del archivo xlsb
      paste0("$xlsbFile = ", '"', xlsbFile, '"'),
      # Ruta del archivo xlsx
      paste0("$xlsxFile = ", '"', xlsxFile, '"'),
      # Crear un objeto COM de Excel
      "$excel = New-Object -ComObject Excel.Application",
      # Deshabilitar las alertas
      "$excel.DisplayAlerts = $false",
      # Abrir el archivo xlsb
      "$workbook = $excel.Workbooks.Open($xlsbFile)",
      # Guardar como archivo xlsx, donde 1 corresponde al formato xls y 51 a xlsx
      "$workbook.SaveAs($xlsxFile, 51)",
      # Cerrar el libro y salir de Excel
      "$workbook.Close()",
      "$excel.Quit()",
      sep = "\n"
    )
    
    # Guardar el script en un archivo en el directorio principal
    writeLines(script, "convert.ps1")
    
    # Mensaje
    cat(paste("\nSe remplazo el archivo \".xlsb\" de ruta: [", xlsbFile,"],",
              "con el archivo \".xlsx\" de ruta [", xlsxFile,"]\n"))
    
    # Ejecutar el script de PowerShell
    shell("powershell -File convert.ps1", wait = TRUE)
    
    # Eliminar el archivo de script y el .xlsb original
    file.remove("convert.ps1",xlsbFile)
  }
  
  # Condiciones no admisibles
  if ( verificadorArchivo(ruta_archivo_xlsb) && verificadorFormato_xlsb(ruta_archivo_xlsb)) {
    if ( verificadorExcel() ) {
      trasnformador2xlsb(ruta_archivo_xlsb)
    } else {
      cat("\nRequiere instalación de Microsoft Excel")
    }
  }
  
}

# Descarga----

obtenerEnlacesDescarga <- function(enlaces_descarga, identificador) {
  
  requerirPaquetes("httr","dplyr")
  
  enlaces <- enlaces_descarga
  
  barraProgresoReinicio()
  
  informacion <- data.frame()
  
  for (enlace in enlaces) {
    indice <- match(enlace, enlaces)
    
    response <- httr::HEAD(enlace)
    
    filename_temporal <-
      response$headers$`content-disposition` %>%
      sub(".*filename=\"([^\"]+)\".*", "\\1", .)
    
    nueva_fila <-
      data.frame(
        time = Sys.time(),
        link = enlace,
        url = response$url,
        filename =
          ifelse(length(filename_temporal) > 0,
                 filename_temporal, basename(response$url)),
        status_code = response$status_code,
        content_type = response$headers$`content-type`,
        last_modified =
          ifelse(!is.null(response$headers$`last-modified`),
                 response$headers$`last-modified`, NA),
        content_length =
          round(as.numeric(response$headers$`content-length`) / 2^20, 2)
      )
    
    informacion <- informacion %>% dplyr::bind_rows(nueva_fila)
    
    if ( indice == 1 ) { cat("\nRutas de descarga:") }
    barraProgreso(enlaces)
    cat("\033[1;32mObteniendo ruta de descarga...\033[0m\n")
    cat("Del vínculo:\n\t[", enlace, "]",
        "\nse ha capturado la ruta de descarga:\n\t[", response$url, "].\n")
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
  # destino <- "data/Fuentes/SEPS/Bases de Datos"
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
      grepl("(?=.*data)(?=.*Fuentes)(?=.*SB)", destino, perl = TRUE)
    if ( va_a_data_fuente_SB ) {
      directorio_destino <-
        dirname(gsub("Descargas|Temporal","Fuentes",ruta_origen))
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

generarListaHojasVolumenCreditoMensualSEPS <- function(){
  
  requerirPaquetes("readxl", "dplyr")
  
  ruta_directorio <- "data/Fuentes/SEPS/Reportes/Volumen de Credito Mensual"
  
  rutas_libros_excel <-
    list.files(ruta_directorio, recursive = TRUE, full.names = TRUE)
  
  nombres_libros_excel <- gsub("\\.[^.]*$", "", basename(rutas_libros_excel))
  
  nombres_hojas_existentes <-
    lapply(rutas_libros_excel, readxl::excel_sheets) %>%
    unlist() %>% unique() %>% sort()
  
  expresion_buscada <- "base"
  
  hojas_volumen_credito_mensual <-
    grep(expresion_buscada, nombres_hojas_existentes,
         ignore.case = TRUE, value = TRUE)
  
  biblioteca_libros <-
    lapply(rutas_libros_excel, function(ruta) {
      nombres_hojas <- readxl::excel_sheets(ruta)
      hoja_seleccionada <-
        intersect(nombres_hojas, hojas_volumen_credito_mensual)
      libro <- readxl::read_excel(ruta, sheet = hoja_seleccionada)
      attr(libro, "nombre_hoja") <- hoja_seleccionada
      return(libro)
    })
  
  names(biblioteca_libros) <- nombres_libros_excel
  
  return(biblioteca_libros)
}

resumenBibliotecaLibros <- function(lista_data_frames) {
  requerirPaquetes("dplyr")
  biblioteca_libros <- lista_data_frames
  nombres_libros <- names(biblioteca_libros)
  nombres_hojas <-
    sapply(biblioteca_libros, function(df) attr(df, "nombre_hoja")) %>%
    unname()
  nombres_columnas_en_hojas <-
    lapply(biblioteca_libros, names) %>% unlist() %>% unique() %>% sort()
  
  resumen_biblioteca_libros <- data.frame(Libro = nombres_libros)
  if ( length(nombres_hojas) == nrow(resumen_biblioteca_libros) ) {
    resumen_biblioteca_libros$Hoja = nombres_hojas
  }
  resumen_biblioteca_libros$`Año` = gsub(".*(\\d{4}).*", "\\1", nombres_libros)
  
  for (nombre_columna in nombres_columnas_en_hojas) {
    resumen_biblioteca_libros[[nombre_columna]] <-
      sapply(biblioteca_libros, function(df) {
        ifelse(test = nombre_columna %in% names(df),
               yes = class(df[[nombre_columna]]),
               no = "")
      })
  }
  return(resumen_biblioteca_libros)
}

crearEstadosFinancierosSEPS <- function() {
  requerirPaquetes("dplyr")
  
  estandarizarNombresColumnasVolumenCreditoMensualSEPS <- function(data_frame) {
    requerirPaquetes("dplyr")
    nuevos_nombres_columnas <-
      dplyr::case_when(
        names(data_frame) %in% c("Actividad Económica", "Actividad Económica ( productivas)", 
                                 "ACTIVIDAD_ECONOMICA", "Actividades no productivas") ~ "Actividad Económica",
        names(data_frame) %in% c("CANTON", "Cantón") ~ "Cantón",
        names(data_frame) %in% c("DESTINO_FINANCIERO") ~ "Destino Financiero",
        names(data_frame) %in% c("Estado de Operación","ESTADO_OPERACION") ~ "Estado de Operación",
        names(data_frame) %in% c("Fecha de corte", "Fecha de Corte", "FECHA_CORTE") ~ "Fecha",
        names(data_frame) %in% c("Institución", "RAZON_SOCIAL") ~ "Entidad Financiera",
        names(data_frame) %in% c("Provincia", "PROVINCIA") ~ "Provincia",
        names(data_frame) %in% c("OPERACIONES") ~ "Número de Operaciones",
        names(data_frame) %in% c("REGION", "Región") ~ "Región",
        names(data_frame) %in% c("NUM_RUC") ~ "RUC",
        names(data_frame) %in% c("SEGMENTO") ~ "Segmento",
        names(data_frame) %in% c("SUJETOS DE CREDITO", "SUJETOS DE CREDITOS") ~ "Sujetos de Crédito",
        names(data_frame) %in% c("Tipo de Crédito", "TIPO DE CRÉDITO GENERAL", 
                                 "TIPO_CREDITO") ~ "Tipo de Crédito",
        names(data_frame) %in% c("TIPO DE CRÉDITO ESPECÍFICO", 
                                 "TIPO_CREDITO_nuevo") ~ "Tipo de Crédito Específico",
        names(data_frame) %in% c("Monto", "VAL_OPERACION") ~ "Valor",
        TRUE ~ names(data_frame)
      )
    names(data_frame) <- nuevos_nombres_columnas
    return(data_frame)
  }
  
  tic_general <- Sys.time()
  
  desde <- 2016
  hasta <- as.integer(format(Sys.Date(),"%Y"))
  expresion_regular_anios_selecionados <- 
    paste0(seq(desde, hasta), collapse = "|")
  
  biblioteca_libros <- generarListaHojasVolumenCreditoMensualSEPS()
  
  indice_data_frame_selecionados <-
    grep(expresion_regular_anios_selecionados, names(biblioteca_libros))
  
  tabla_concatenada <-
    biblioteca_libros[indice_data_frame_selecionados] %>%
    lapply(., estandarizarNombresColumnasVolumenCreditoMensualSEPS) %>%
    lapply(., function(data_frame) {
      if ("Fecha de Corte" %in% names(data_frame)) {
        data_frame$`Fecha de Corte` <-
          as.Date(data_frame$`Fecha de Corte`, origin = "1899-12-30")
      }
      if ("Número de Operaciones" %in% names(data_frame)) {
        data_frame$`Número de Operaciones` <-
          as.integer(data_frame$`Número de Operaciones`)
      }
      if ("Sujetos de Crédito" %in% names(data_frame)) {
        data_frame$`Sujetos de Crédito` <-
          as.integer(data_frame$`Sujetos de Crédito`)
      }
      if ("Valor" %in% names(data_frame)) {
        data_frame$`Valor` <-
          as.numeric(data_frame$`Valor`)
      }
      data_frame$Columna1 <- NULL
      return(data_frame)
    }) %>%
    dplyr::bind_rows()
  
  exportarResultadosCSV(tabla_concatenada,"SEPS Volumen de Credito")
  cat("\n\n  \033[1;34mDuración total del proceso \"Volumen de Credito Mensual SEPS\":",
      formatoTiempoHMS(difftime(Sys.time(), tic_general, units = "secs")), "\033[0m\n")
  
  return(tabla_concatenada)
}

# SB ----

decargaDesdePortalEstudiosSB <- function(ruta_archivo_html) {
  
  requerirPaquetes("rvest","dplyr","stringr","readr")
  
  leer_pagina_html <- function(ruta_archivo_html, codificacion = NA) {
    if ( is.na(codificacion) ) {
      codificacion <-
        readr::guess_encoding(ruta_archivo_html) %>%
        slice(which.max(confidence)) %>%
        pull(encoding)
    }
    pagina_html <- rvest::read_html(ruta_archivo_html, encoding = codificacion)
    return(pagina_html)
  }
  scraping_descarga_entry_file <- function(pagina_html, informacion_enlaces_data_frame) {
    nodos_descarga_entry_file <- pagina_html %>% html_nodes(".entry.file")
    if ( length(nodos_descarga_entry_file) > 0 ) {
      informacion_enlaces <- data.frame(
        boletin =
          pagina_html %>% html_nodes(".entry.file") %>% html_attr("data-name"),
        nombre_archivo =
          pagina_html %>% html_nodes(".entry.file") %>%
          html_nodes(".entry-info-name span") %>% html_text(),
        enlace_descarga =
          pagina_html %>% html_nodes(".entry_link.entry_action_download") %>% #no se puede separar el selector
          html_attr("href"),
        ids_archivo =
          pagina_html %>% html_nodes(".entry.file") %>% html_attr("data-id"),
        fecha_modificacion =
          pagina_html %>% html_nodes(".entry-info-modified-date") %>% html_text(),
        fecha_descripcion =
          pagina_html %>% html_node(".description-file-info") %>% html_text() %>%
          stringr::str_extract("\\d+\\s[a-zA-Z]+,\\s\\d{4}\\s\\d+:\\d+\\s[ap]m"),
        tamanio_archivo =
          pagina_html %>% html_nodes(".entry-info-size") %>% html_text() %>%
          readr::parse_number()
      )
    } else {
      informacion_enlaces <- informacion_enlaces_data_frame
    }
    return(informacion_enlaces)
  }
  scraping_descarga_a <- function(pagina_html, informacion_enlaces_data_frame) {
    nodos_descarga_entry_file <- pagina_html %>% html_nodes(".entry.file")
    nodos_descarga_a <- pagina_html %>% html_nodes("a")
    if ( length(nodos_descarga_entry_file) == 0 &
         length(nodos_descarga_a) > 0
    ) {
      # informacion_enlaces <-
      #   data.frame(enlaces_descarga = nodos_descarga_a %>% html_attr("href")) %>%
      #   mutate(nombre_archivo = basename(enlaces_descarga))
      enlaces_descarga <- nodos_descarga_a %>% html_attr("href")
      informacion_enlaces <-
        obtenerEnlacesDescarga(enlaces_descarga,"SB") %>%
        rename(
          fecha_acceso = "time",
          enlace_descarga = "url",
          nombre_archivo = "filename",
          fecha_modificacion = "last_modified",
          tamanio_archivo = "content_length"
        )
    } else {
      informacion_enlaces <- informacion_enlaces_data_frame
    }
    return(informacion_enlaces)
  }
  scraping_descarga_error <- function(informacion_enlaces_data_frame) {
    if ( nrow(informacion_enlaces_data_frame) == 0 ) {
      stop(paste("\nEl proceso se ha interrumpido,",
                 "el Web Scraping no pudo realizarse,",
                 "la función 'decargaDesdePortalEstudiosSB'",
                 "requiere mantenimiento."))
    }
  }
  directorioCarpetaDescarga <- function(ruta_archivo_html, nombre_archivo) {
    directorio_carpeta <-
      ruta_archivo_html %>%
      gsub(".html", "",.) %>%
      gsub("html/", "data/Descargas/",.)
    anio_actual <- as.numeric(format(Sys.Date(), "%Y"))
    expresion_regular_anios <- paste(seq(1990,anio_actual), collapse = "|")
    prueba_anio <- grepl(expresion_regular_anios, directorio_carpeta)
    if ( !prueba_anio ) {
      nombre_carpeta <- gsub(".zip", "", nombre_archivo)
      coincidencias_anio <- gregexpr(expresion_regular_anios, nombre_carpeta)
      nombre_carpeta <- unlist(regmatches(nombre_carpeta, coincidencias_anio))
      directorio_carpeta <- 
        gsub(basename(directorio_carpeta), nombre_carpeta, directorio_carpeta)
    }
    crearDirectorio(directorio_carpeta)
    return(directorio_carpeta)
  }
  
  pagina_html <- leer_pagina_html(ruta_archivo_html)
  
  informacion_enlases_selecionados <- data.frame()
  informacion_enlases_selecionados <- scraping_descarga_entry_file(pagina_html, informacion_enlases_selecionados)
  informacion_enlases_selecionados <- scraping_descarga_a(pagina_html, informacion_enlases_selecionados)
  scraping_descarga_error(informacion_enlases_selecionados)
  
  exportarReporteTabla(
    dataFrame =  informacion_enlases_selecionados,
    nombre_archivo =
      paste("Reporte Enlaces de Descarga SB",
            tools::file_path_sans_ext(basename(ruta_archivo_html))))
  
  #barraProgresoReinicio()
  
  for (k in 1:nrow(informacion_enlases_selecionados) ) {
    link <- informacion_enlases_selecionados$enlace_descarga[k]
    nombre_archivo <- informacion_enlases_selecionados$nombre_archivo[k]
    directorio_descarga <-
      directorioCarpetaDescarga(ruta_archivo_html, nombre_archivo)
    ruta_archivo <- file.path(directorio_descarga, nombre_archivo)
    #tamanio_archivo <- informacion_enlases_selecionados$tamanos_archivo[k]
    if ( 
      !file.exists(ruta_archivo) #| file.size(ruta_archivo) < tamanio_archivo
    ) {
      download.file(link, ruta_archivo, mode = "wb", timeout = 300)
      cat("\033[1;32mSe descargó el archivo en la ruta:\033[0m",
          "[", normalizePath(ruta_archivo),"]\n\n")
      #barraProgreso(seq_along(download_links_elegidos))
      #cat("Descargando... ")
    }
  }
}

ejecutarDecargaDesdePortalEstudiosSB <- function() {
  ruta_directorio_html_SB <- "html/SB/Volumen de Credito"
  rutas_archivos_html_SB <-
    list.files(ruta_directorio_html_SB, recursive = TRUE, full.names = TRUE)
  
  # ruta_directorio_html_SB_Privados <-
  #   paste0(ruta_directorio_html_SB,"/Bancos Privados/",
  #          format(Sys.Date(),"%Y"),".html")
  # ruta_directorio_html_SB_Publicas <-
  #   paste0(ruta_directorio_html_SB,"/Instituciones Publicas/",
  #          format(Sys.Date(),"%Y"),".html")
  # prueba_ruta_anio_actual <-
  #   all(c(ruta_directorio_html_SB_Privados,
  #         ruta_directorio_html_SB_Publicas) %in% rutas_archivos_html_SB)
  # if ( ! prueba_ruta_anio_actual  ) {
  #   file.create(ruta_directorio_html_SB_Privados,
  #               ruta_directorio_html_SB_Publicas)
  # }
  
  barraProgresoReinicio()
  for (ruta in rutas_archivos_html_SB) {
    decargaDesdePortalEstudiosSB(ruta)
    barraProgreso(rutas_archivos_html_SB)
    cat("\033[1;32mAnalizando archivo html en la ruta:\033[0m [", ruta, "]\n\n")
  }
}

hojaToTablaBoletinesFinancierosSB <- function(ruta_libro, nombre_hoja, fecha_corte = NULL) {
  
  # Esta función permite extraer la tabla de datos contenida en un hoja de cálculo correspondiente a los "Boletines Financieros mensuales" de la SB
  
  # ARGUMENTOS:
  # ruta_libro <- "data/Fuente/SB/PRIVADA/2023/FINANCIERO MENSUAL BANCA PRIVADA 2023_02.xlsx"
  # nombre_hoja <- "BALANCE"
  # fecha_corte <- "2023-02-29"
  # EJEMPLO: tabla <- hojaToTablaBoletinesFinancierosSB(ruta_libro, nombre_hoja, fecha_corte)
  
  requerirPaquetes("dplyr","readxl")
  
  nombreHojaSimilar <- function(ruta_libro, nombre_hoja_buscado) {
    
    # Determina la hoja en un libro de excel que tiene mayor similitud a la hoja buscada
    
    requerirPaquetes("readxl","stringdist")
    
    nombres_hojas <- readxl::excel_sheets(ruta_libro)
    distancia <- stringdist::stringsimmatrix(nombre_hoja_buscado, nombres_hojas, method = "jw")
    indice_hoja_similar <- which.max(distancia)
    nombre_hoja_similar <- nombres_hojas[indice_hoja_similar]
    return(nombre_hoja_similar)
  }
  analisisDifusoNLPFechaCorte <- function(tabla) {
    
    # Esta función procesa un texto relacionado a un la fecha de corte de los "Balances Financieros" de SB y devuelve el date más cercano a fecha de corte.
    
    requerirPaquetes("lubridate","parsedate","stringdist")
    
    traductor_mes <- function(texto) {
      
      # Esta función modifica con la traducción al ingles correspondiente sean los nombres completos o las abreviaciones de los meses, para un posterior reconocimiento optimo de fecha
      
      texto_original <- tolower(texto)
      # Creamos un diccionario para traducción y posterior reconocimiento optimo de fechas
      meses <-
        data.frame(
          es = c("ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"),
          #en = substr(strsplit(tolower(month.name), " "), 1, 3)
          en = tolower(month.name)
        )
      # Definimos el patrón buscado en el texto
      patron <- meses$es
      # Definimos el texto de reemplazo
      reemplazo <- meses$en
      # Definimos los separadores admisibles para las palabras
      separadores <- "[-,/, ]"
      # Separamos cada palabra en sus letras componentes
      palabras <- strsplit(texto_original, separadores)[[1]]
      # palabras <- unlist(strsplit(texto_original, separadores))
      # Elegimos únicamente las 3 primeros caracteres de cada palabra para obtener expresiones como: "ene"
      palabras_abreviadas <- substr(palabras, 1, 3)
      # Calculamos las similitudes entres las palabras abreviadas y el patrón de busqueda
      similitudes <- stringdist::stringsimmatrix(palabras_abreviadas, patron, method = "jw")
      # Se emplea una probabilidad de similitud del 90% para compensar el error por identidad con el máximo
      #posiciones_max <- as.data.frame(which(similitudes >= 0.8*max(similitudes), arr.ind = TRUE))
      # Buscamos los índices con las mayores coincidencias
      posiciones_max <- as.data.frame(which(similitudes == max(similitudes), arr.ind = TRUE))
      # Determinamos las abreviaciones similares
      palabra_similar <- palabras[posiciones_max$row]
      # Reemplazamos con la palabra completa las abreviaciones similares
      reemplazo_similar <- reemplazo[posiciones_max$col]
      # Modificamos uno a uno los nombres de los mes traducidos
      texto_modificado <- texto_original
      for ( k in seq_along(palabra_similar) ) {
        texto_modificado <- gsub(palabra_similar[k], reemplazo_similar[k], texto_modificado)
      }
      return(texto_modificado)
    }
    prueba_anio <- function(texto) {
      # Año actual a texto, para generar expresión regular de año, usan Sys.Date() y descomponiéndolo
      anio_num <- year(Sys.Date())
      anio_text <- strsplit(as.character(anio_num), split = "")[[1]]
      # Establecemos una expresión regular que acepta 2000 hasta el año actual
      expresion_regular_anio <- paste0("\\b(",anio_text[1],"[",0,"-",anio_text[2],"][",0,"-",anio_text[3],"][0-9])\\b")
      # expresion_regular_anio <- "\\b(20[0-3][0-9])\\b$" # acepta desde 2000 hasta 2039
      return(grepl(expresion_regular_anio, texto, ignore.case = TRUE))
    }
    prueba_mes <- function(texto) {
      # Establecemos una expresión regular
      expresion_regular_mes <- paste0(c("ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"), collapse = "|")
      return(grepl(expresion_regular_mes, texto, ignore.case = TRUE))
    }
    prueba_dia <- function(texto) {
      # Establecemos una expresión regular del día del mes
      expresion_regular_dia <- "\\b([1-2]?[0-9]|3[0-1])\\b"
      return(grepl(expresion_regular_dia, texto, ignore.case = TRUE))
    }
    formato_numerico_excel <- function(fecha) {
      # Función para transformar a formato numérico de Excel una fecha date
      # Fecha base de Excel
      fecha_base_excel <- as.Date("1899-12-30")
      return(as.numeric(difftime(as.Date(fecha), as.Date("1899-12-30"))))
    }
    prueba_fecha_excel <- function(texto) {
      # Fecha de inicio de busqueda en formato numérico de Excel
      fecha_num_excel_inicio <- formato_numerico_excel("2000-01-01")
      # Fecha de inicio descompuesta en caracteres para formar expresión regular
      fechaI <- strsplit(as.character(fecha_num_excel_inicio), split = "")[[1]]
      # Fecha de actual en formato numérico de Excel para busqueda
      fecha_num_excel_fin <- formato_numerico_excel(Sys.Date())
      # Fecha de fin descompuesta en caracteres para formar expresión regular
      fechaF <- strsplit(as.character(fecha_num_excel_fin), split = "")[[1]]
      # Establecemos una expresión regular que acepte los formatos numéricos para fecha de Excel
      expresion_regular_fecha_num_excel <-
        paste0(
          "^(",fechaI[1],"[",fechaI[2],"-9]","[",fechaI[3],"-9]","[",fechaI[4],"-9]","[",fechaI[5],"-9]|",
          fechaF[1],"[0-9]{", length(fechaF)-1, "})"
        )
      return(grepl(expresion_regular_fecha_num_excel, texto))
    }
    prueba_fecha_date <- function(texto) {
      # Establecemos una expresión regular que acepte variantes de formato fecha
      expresion_regular_fecha_date <-
        paste(c(
          "\\b(20[0-9]{2}[-/][0-1][0-9][-/][0-3]?[0-9])\\b",
          #"\\b(20[0-9]{2}[-/][[:alpha:]]{1,10}[-/][0-3]?[0-9])\\b",# NO USAR ALTERA EN EL CONDICIONAL
          "\\b([0-3]?[0-9][-/][0-1][0-9][-/]20[0-9]{2})\\b" #,
          #"\\b([0-3]?[0-9][-/][[:alpha:]]{1,10}[-/]20[0-9]{2})\\b"
        ), collapse = "|"
        )
      return(grepl(expresion_regular_fecha_date, texto))
    }
    
    # Empleamos la función indicePrimeraFilDecimalTabla() para identificar la primera fila decimal
    indice_fila_nombres <- indicePrimeraFilDecimalTabla(tabla)
    # Subtabla previa a los valores decimales, y a la fila de nombres de columnas, por eso se resta 2
    subtabla <- tabla[1:(indice_fila_nombres-2),]
    # Determinamos las coincidencias en la subtabla
    coincidencias <-
      apply(
        subtabla, 2,
        function(fila) {
          prueba_fecha_date(fila) | prueba_fecha_excel(fila) | (prueba_anio(fila) & prueba_mes(fila)) 
        })
    # Identificamos los indices de las entradas con coincidencias
    indices_celda <- data.frame(which(coincidencias, arr.ind = TRUE))
    # Exigimos que haya al menos un resultado
    if ( length(indices_celda) > 0 ) {
      # Especificamos la primera coincidencia
      contenido_celda <- as.character(subtabla[indices_celda$row[1], indices_celda$col[1]])
    } else {
      cat("\nNo se pudo encontrar una fecha.\n")
      break
    } 
    
    # Establecemos el proceso directo para formatos de fecha
    
    if ( prueba_fecha_date(contenido_celda) ) {
      
      fecha_identificada <- parsedate::parse_date(contenido_celda)
      
      # Establecemos la condición para cuando el texto leído corresponde a fecha en formato numérico de Excel
      
    } else if ( prueba_fecha_excel(contenido_celda) ) {
      
      # Determinamos el valor de la celda buscada con la fecha de corte
      num_fecha_corte <- as.numeric(contenido_celda)
      # Determinamos la fecha de corte
      fecha_identificada <- as.Date( num_fecha_corte, origin = "1899-12-30")
      
      # Establecemos el procedimiento para el caso de tener un mes y un año reconocibles
      
    } else if ( prueba_anio(contenido_celda) & prueba_mes(contenido_celda) ) {
      
      # Dividimos el texto original en sus componentes por si hubiera más de una fecha
      texto_dividido <- unlist(strsplit(contenido_celda, " "))
      # Establecemos el proceso cuando haya solo un año, solo un mes, y no más de un día del mes
      if ( sum(prueba_anio(texto_dividido)) == 1 & 
           sum(prueba_mes(texto_dividido)) == 1 & 
           sum(prueba_dia(texto_dividido)) <= 1 ) {
        fechas_reconocidas <- traductor_mes(contenido_celda)
        fecha_identificada <- parsedate::parse_date(fechas_reconocidas)
        # Establecemos el proceso cuando hay más de una fecha en la celda elegida
      } else {
        # Traducimos el contenido de la celda elegida
        fechas_reconocidas <- traductor_mes(contenido_celda)
        # Empleamos un selector para el separador de frases, según formato
        separadores_fechas <-
          if ( grepl("-",contenido_celda, ignore.case = TRUE) ) {
            " "
          } else if ( grepl("de",contenido_celda, ignore.case = TRUE) ) {
            c(" al "," hasta ")
          }
        # Separamos las diferentes frases relacionadas a fechas
        fechas_reconocidas <- strsplit(fechas_reconocidas, separadores_fechas)[[1]]
        # Agregamos un filtro para evitar frases sin el año
        fechas_reconocidas <- fechas_reconocidas[prueba_anio(fechas_reconocidas)]
        fechas_reconocidas <- parsedate::parse_date(fechas_reconocidas)
        # Agregamos un filtro para elegir siempre la mayor de las fechas
        fecha_identificada <- fechas_reconocidas[which.max(fechas_reconocidas)]
      }
    }
    
    # Determinamos el año
    anio <- format(fecha_identificada, "%Y")
    # Determinamos el mes
    mes <- format(fecha_identificada, "%m")
    # Determinamos una fecha preliminar
    fecha_corte_preliminar <- paste(anio,mes,"01",sep = "-")
    # Determinamos el último día del respectivo mes
    fecha_corte <- as.Date(fecha_corte_preliminar) + months(1) - days(1)
    
    return(fecha_corte)
  }
  
  nombre_hoja <- nombreHojaSimilar(ruta_libro, nombre_hoja)
  hoja <- suppressMessages(readxl::read_excel(ruta_libro, sheet = nombre_hoja, col_names = FALSE, n_max = 30))
  fecha_corte <-
    if ( is.null(fecha_corte) ) {
      analisisDifusoNLPFechaCorte(hoja)
    } else {
      fecha_corte
    }
  # Determinamos la fila más probable con los nombres de las columnas
  indice_fila_nombres_columnas <- indicePrimeraFilDecimalTabla(hoja) - 1
  # Almacenamos la fila con los nombres de las columnas
  nombres_columnas <- unname(unlist(hoja[indice_fila_nombres_columnas,]))
  # Importamos una tabla de prueba para verificar la correcta asignación de los nombres de las columnas en sus 20 primeras filas
  tabla_prueba <- suppressMessages(readxl::read_excel(ruta_libro, sheet = nombre_hoja, col_names = TRUE, skip = indice_fila_nombres_columnas, n_max = 20))
  # Verificamos si coinciden adecuadamente los nombres de las columnas
  if ( mean(nombres_columnas == names(tabla_prueba), na.rm = TRUE) < 0.8 ) {
    # Retrocedemos un índice en las filas previo a iterear para incluir cualquier caso exepcional
    indice_fila_nombres_columnas <- indice_fila_nombres_columnas - 2
    # Iteramos hasta que hayan coincidencias en al menos el 80%
    while ( mean(nombres_columnas == names(tabla_prueba), na.rm = TRUE) < 0.8 & indice_fila_nombres_columnas <= 20 ) {
      # Incrementamos el índice de la fila para continuar la prueba
      indice_fila_nombres_columnas <- indice_fila_nombres_columnas + 1
      # Reimportamos la tabla de prueba para verificar la correcta asignación de los nombres de las columnas en sus 20 primeras filas
      tabla_prueba <- suppressMessages(readxl::read_excel(ruta_libro, sheet = nombre_hoja, col_names = TRUE, skip = indice_fila_nombres_columnas, n_max = 20))
    }
  }
  # Inicializamos la variable para almacenar la advertencias
  advertencias <- NULL
  # Volvemos a importar la hoja de cálculo pero especificando la fija de inicio, para que se reconozca el tipo de dato y nombre de cada columna
  tabla <-
    # Usamos withCallingHandlers() para capturar las advertencias generadas durante la ejecución del código y almacenarlas en una variable
    withCallingHandlers(
      # Importamos únicamente la tabla de datos contenida en la hoja especificada, saltando las primeras filas
      suppressMessages(
        readxl::read_excel(ruta_libro,
                           sheet = nombre_hoja,
                           col_names = TRUE,
                           skip = indice_fila_nombres_columnas)),
      # Empleamos una función como manejador de advertencias
      warning = function(w) {
        # La función toma un argumento w, que es un objeto de advertencia que contiene información sobre la advertencia generada
        advertencias <<- c(advertencias, w$message)
        # Suprimimos la advertencia y evitamos que la advertencia se muestre en la consola y permite que el código continúe ejecutándose normalmente
        invokeRestart("muffleWarning")
      }
    )
  # Agregamos las advertencias como un atributo de la tabla
  attr(tabla, "advertencias") <- advertencias
  # Agregamos la columna con la fecha del "Boletín Financiero mensual"
  tabla_modificada <-
    tabla %>%
    # Eliminamos las columnas que no contengan caracteres alfabéticos
    select( -matches("^[^[:alpha:]]+$", .) ) %>%
    # Eliminamos las filas que contienen únicamente valores NA
    filter( !if_all(everything(), is.na) ) %>%
    # Empleamos la función creada para modificar los nombres de las columnas según un catálogo por defecto
    modificarNombreColumnaSB(tabla = ., precision = 0.8) %>%
    # Modificamos la columna CODIGO a texto
    mutate(CODIGO = as.character(CODIGO)) %>%
    # Modificamos la columna CUENTA a texto
    mutate(CUENTA = as.character(CUENTA)) %>%
    # Modificamos el resto de columnas a numéricas
    mutate_at(vars(-CODIGO, -CUENTA), as.numeric) %>%
    # Eliminamos todas las filas donde el valor en las columnas "CODIGO" y "CUENTA" es NA
    filter( !(is.na(CODIGO) & is.na(CUENTA)) ) %>%
    # Eliminamos las filas donde todas las columnas son NA excepto CUENTA
    filter( !if_all(-CUENTA, is.na) ) %>%
    # Eliminamos las filas donde la columna CODIGO tenga letras mientras todas las las demás columnas son NA
    filter( !(grepl("[[:alpha:]]+",CODIGO) & if_all(-CODIGO, is.na)) ) %>%
    # Agregamos la columna con la fechas de corte
    mutate(`FECHA` = rep(fecha_corte)) %>%
    # Movemos la columna FECHA al inicio de la tabla
    select(`FECHA`, everything())
  # Agregamos metadatos como atributo de la tabla
  #attr(tabla, "fecha_creacion") <- Sys.Date()
  
  return(tabla_modificada)
}

compilarHojasBalanceFinancieroSB <- function(ruta_directorio = NULL) {
  
  # Esta función realiza todo el proceso necesario para crear la base de datos de los Balances Financieros mensuales de la SB
  
  requerirPaquetes("dplyr","purrr","readxl","reshape2","tools")
  
  # # Cerramos todos los libros de Excel abiertos
  # system2("powershell", "Get-Process excel | Foreach-Object { $_.CloseMainWindow() }")
  if ( is.null(ruta_directorio) ) {
    ruta_directorio <- "data/Fuentes/SB/Volumen de Credito"
  }
  archivos_directorio <- list.files(ruta_directorio, recursive = TRUE)
  #tiene_extension_zip <- tools::file_ext(archivos_directorio) == "zip"
  tiene_extension_zip <- grepl("\\.zip$", archivos_directorio)
  archivos_directorio <- archivos_directorio[!tiene_extension_zip]
  rutas_libros <- file.path(ruta_directorio, archivos_directorio)
  rutas_transformar <- rutas_libros[tools::file_ext(rutas_libros) == "xlsb"]
  if ( length(rutas_transformar) > 0 ) {
    purrr::map(rutas_transformar, xlsb2xlsx)
    archivos_directorio <- list.files(ruta_directorio, recursive = TRUE)
    rutas_libros <- file.path(ruta_directorio, archivos_directorio)
  }
  # prueba_anio <- grepl("(201[3-9])|(202[0-9])",rutas_libros)
  anio_inicio <- 2005
  anio_actual <- as.numeric(format(Sys.Date(), "%Y"))
  expresion_regular_anios <-
    paste(seq(anio_inicio, anio_actual), collapse = "|")
  prueba_anio <- grepl(expresion_regular_anios, rutas_libros)
  rutas_libros_seleccionados <- rutas_libros[prueba_anio]
  cat("\n\nCerrando los los libros de Excel realacionados...\n")
  cerrarLibroExcel(rutas_libros_seleccionados)
  barraProgresoReinicio()
  lista_tablas_BAL_PYG_concatenadas <- list()
  for ( ruta_libro in rutas_libros_seleccionados ) {
    # Importamos las 20 primeras filas de la hoja BALANCE para identificar la fecha de corte
    hoja <-
      suppressMessages(
        readxl::read_excel(ruta_libro, sheet = "BALANCE", n_max = 20))
    # Identificamos la fecha de corte
    fecha_corte <- analisisDifusoNLPFechaCorte(hoja)
    # Extraemos la tabla de BALANCE
    tabla_BAL <-
      hojaToTablaBoletinesFinancierosSB(ruta_libro, "BALANCE", fecha_corte)
    # Extraemos la tabla de PYG
    tabla_PYG <-
      hojaToTablaBoletinesFinancierosSB(ruta_libro, "PYG", fecha_corte)
    # Definimos el nombre de para cada tabla
    nombre_tabla <- basename(ruta_libro)
    # Asignamos la tabla concatenada de BALANCE y PYG a un elemento de la lista de tablas
    lista_tablas_BAL_PYG_concatenadas[[nombre_tabla]] <-
      dplyr::bind_rows(tabla_BAL,tabla_PYG)
    # Ejecutamos el código para la barra de progreso
    barraProgreso(rutas_libros_seleccionados)
    # Mostramos la ruta del archivo en proceso
    cat("\033[1;32mImportando y procesando el archivo:\033[0m",
        "[", normalizePath(ruta_libro), "]\n")
  }
  # Concatenamos todas las tablas de la lista generada
  tabla_BAL_PYG <- dplyr::bind_rows(lista_tablas_BAL_PYG_concatenadas)
  # Asignamos el registro completo de advertencias (warnings) generadas al convertir a tabla las hojas de cálculo
  registro_advertencias <-
    sapply(seq_along(lista_tablas_BAL_PYG_concatenadas),
           function(k) attr(lista_tablas_BAL_PYG_concatenadas[[k]],"advertencias"))
  # Recuperamos los nombres de cada archivo para el registro de advertencias
  names(registro_advertencias) <- names(lista_tablas_BAL_PYG_concatenadas)
  # Asignamos la información de las advertencias a un data frame
  reporte_consolidacion_BAL_PYG <-
    data.frame(
      Archivo = names(unlist(registro_advertencias)),
      Advertencia = unname(unlist(registro_advertencias)))
  # Exportamos el reporte con el registro de las advertencias
  exportarReporteTabla(
    reporte_consolidacion_BAL_PYG,
    paste("Reporte Advertencias en Consolidación Balances Financieros SB",
          basename(ruta_directorio)))
  # Fundimos (melting) las tablas
  tabla_BAL_PYG_fundida <-
    reshape2::melt(tabla_BAL_PYG,
                   id.vars = colnames(tabla_BAL_PYG)[1:3],
                   variable.name = "RAZON_SOCIAL",
                   value.name = "VALOR")
  
  return(tabla_BAL_PYG_fundida)
}

crearBalancesFinancierosSB <- function() {
  
  requerirPaquetes("dplyr")
  
  tic_general <- Sys.time()
  
  ejecutarDecargaDesdePortalEstudiosSB() # Verificado en prueba 2023/06/14
  
  origen <- "data/Descargas/SB/Volumen de Credito"
  destino <- "data/Fuentes/SB/Volumen de Credito"
  descomprimirArchivosDirectorioZip(origen, destino) # Verificado en prueba 2023/06/14
  
  privada <- compilarHojasBalanceFinancieroSB(destino) %>%
    mutate(SEGMENTO = "PRIVADA") # Verificado en prueba individual 2023/05/11
  
  # ETAPA 4: Concatenación todas de tablas consolidadas ----
  cat("\n\nConcatenando tablas y agregando RUC...\n")
  consolidada <-
    dplyr::bind_rows(privada,publica) %>%
    agregarRUCenSB() %>%
    dplyr::select(FECHA, SEGMENTO, RUC, RAZON_SOCIAL, CODIGO, CUENTA, VALOR)
  
  # ETAPA 5: Exportación de base de datos generada ----
  exportarResultadosCSV(consolidada,"SB Balances Financieros")
  cat("\n\n  \033[1;34mDuración total del proceso \"Balances Financieros SB\":",
      formatoTiempoHMS(difftime(Sys.time(), tic_general, units = "secs")), "\033[0m\n")
  
  return(consolidada)
}


