

n <- sapply(seq_along(lista_tablas), function(k) names(lista_tablas[[k]])) %>% unlist() %>% unique() %>% sort()

columnas_identificadas <- c(
  "ACTIVIDAD", "CANTON", "ENTIDAD","ESTADO DE LA OPERACION",
  "ESTADO DE LA OPERACIÓN", "FECHA", "MONTO OTORGADO", "NUMERO DE OPERACIONES",
  "NÚMERO DE OPERACIONES", "PROVINCIA", "REGION", "SECTOR", "SUBSECTOR",
  "SUBSISTEMA", "TIPO DE CREDITO", "TIPO DE OPERACION", "TIPO DE OPERACIÓN")

lista_data_frames <- lista_tablas

resumenBibliotecaLibrosSB <- function(lista_data_frames) {
  requerirPaquetes("dplyr")
  biblioteca_libros <- lista_data_frames
  nombres_libros <- names(biblioteca_libros)
  nombres_hojas <-
    sapply(biblioteca_libros, function(df) attr(df, "nombre_hoja")) %>%
    unlist() %>% unname()
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