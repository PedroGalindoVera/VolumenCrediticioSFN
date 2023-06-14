nombres_hojas_existentes <-
  lapply(rutas_libros_seleccionados, readxl::excel_sheets) %>%
  unlist() %>% unique() %>% sort()


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