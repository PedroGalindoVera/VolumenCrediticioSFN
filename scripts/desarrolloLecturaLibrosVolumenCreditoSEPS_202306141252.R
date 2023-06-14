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
