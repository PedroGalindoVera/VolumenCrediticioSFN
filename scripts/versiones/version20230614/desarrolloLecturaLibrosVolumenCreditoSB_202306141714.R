

rutas_libros <- rutas_libros_seleccionados

ruta_directorio <- "data/Fuentes/SB/Volumen de Credito"

resumenLibrosExcelDirectorio <- function(ruta_directorio) {
  requerirPaquetes("dplyr")
  
  rutas_libros_excel <-
    list.files(ruta_directorio, recursive = TRUE, full.names = TRUE)
  nombres_libros_excel <- gsub("\\.[^.]*$", "", basename(rutas_libros_excel))
  
  nombres_hojas_existentes <-
    lapply(rutas_libros_excel, readxl::excel_sheets) %>%
    unlist() %>% unique() %>% sort()
  
  resumen_libros_directorio <- data.frame(Libro = nombres_libros_excel)

  resumen_libros_directorio$`Año` <-
    gsub(".*(\\d{4}).*", "\\1", nombres_libros_excel)
  
  for (nombre_hoja in nombres_hojas_existentes) {
    resumen_libros_directorio[[nombre_hoja]] <-
      sapply(rutas_libros_excel, function(ruta) {
        ifelse(test = nombre_hoja %in% readxl::excel_sheets(ruta),
               yes = TRUE,
               no = "")
      })
  }
  return(resumen_libros_directorio)
}