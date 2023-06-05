requerirPaquetes("dplyr")

rutas_libros_excel <-
  list.files("data/Fuentes/SEPS/Reportes/Volumen de Credito Mensual",
             recursive = TRUE, full.names = TRUE)

nombres_hojas_existentes <-
  lapply(rutas_libros_excel, readxl::excel_sheets) %>%
  unlist() %>%
  unique() %>%
  sort()

hojas_volumen_crediticio <- head(nombres_hojas_existentes,7)

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

hojas_por_libro <- lapply(rutas_libros_excel, readxl::excel_sheets)
names(hojas_por_libro) <- basename(rutas_libros_excel)

hojas_por_libro <- unlist(hojas_por_libro)
names(hojas_por_libro)[hojas_por_libro == "BASE VOL.CRÉDITO 2013"]

biblioteca_libros <- lapply(rutas_libros_excel, function(ruta) {
  nombres_hojas <- readxl::excel_sheets(ruta)
  hoja_seleccionada <- intersect(nombres_hojas, hojas_volumen_crediticio)
  readxl::read_excel(ruta, sheet = hoja_seleccionada)
  })

lapply(biblioteca_libros, names) %>% unlist() %>% unique() %>% sort()
