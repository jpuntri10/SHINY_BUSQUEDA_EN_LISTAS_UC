
library(shiny)
library(readxl)
library(dplyr)
library(DT)
library(writexl)
library(stringr)
library(purrr)
library(tidyr)

# Para PDF sin LaTeX (opcional)
has_pkg <- function(pkg) requireNamespace(pkg, quietly = TRUE)

# =============================
#  Zona horaria: America/Lima
# =============================
TZ_APP <- "America/Lima"
Sys.setenv(TZ = TZ_APP)
options(tz = TZ_APP)

fecha_local <- function(fmt = "%Y-%m-%d %H:%M:%S") {
  format(Sys.time(), tz = TZ_APP, fmt)
}

# Logs
message("== Zona horaria de la app ==")
message("TZ env: '", Sys.getenv("TZ"), "'")
message("Sys.timezone(): ", Sys.timezone())
message("as.POSIXlt(Sys.time())$zone: ",
        paste(unique(as.POSIXlt(Sys.time())$zone), collapse = ", "))

ui <- fluidPage(
  titlePanel("Cruce de Documentos con listas UC (Excel base automático: más reciente)"),
  
  # ---------- CSS ----------
  tags$style(HTML("
    .sidebarPanel small { display: block; line-height: 1.2; margin-bottom: 6px; }
    .path-trunc { max-width: 100%; display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    code { font-family: Consolas, 'Courier New', monospace; }
    @media (min-width: 768px) { .col-sm-4 { min-width: 320px; } }
  ")),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("archivo_consulta", "Sube tu Excel de consulta (columna COD_DOCUM o COD_ID)",
                accept = c(".xlsx", ".xls")),
      textInput("doc_manual", "O ingresa un documento manualmente"),
      actionButton("procesar", "Procesar"),
      br(), br(),
      uiOutput("estado_base"),
      uiOutput("detalle_base"),
      uiOutput("hojas_base")
    ),
    mainPanel(
      DTOutput("tabla_resultado"),
      br(),
      fluidRow(
        column(6, downloadButton("descargar_excel", "Descargar Resultado en Excel")),
        column(6, downloadButton("descargar_pdf", "Descargar Resultado en PDF"))
      )
    )
  )
)

server <- function(input, output, session) {
  
  # === Columnas visibles/exportables ===
  columnas_mostrar <- c(
    "TIP_DOCUM",
    "COD_DOCUM",
    "FUENTE_HOJA",
    "LISTAS",
    "TIPO_ENTIDAD",
    "NOMBRE",
    "FECHA_BUSQUEDA",
    "ESTADO"
  )
  
  cols_presentes <- function(df, deseadas) intersect(deseadas, names(df))
  
  base_listas      <- reactiveVal(NULL)
  info_carga       <- reactiveVal(list(path = NA_character_, hojas = character(0)))
  resultado_cruce  <- reactiveVal(data.frame())
  
  # ---------- Normalización de encabezados ----------
  normalizar_y_mapear <- function(df) {
    names(df) <- names(df) |> stringr::str_trim() |> toupper()
    # Mapeos de compatibilidad (por si algún archivo trae variantes)
    mapa <- c(
      "COD_ID"        = "COD_DOCUM",
      "COD_DOC"       = "COD_DOCUM",
      "TIPO_ID"       = "TIP_DOCUM",
      "TIP_DOCU"      = "TIP_DOCUM",
      "TIPO_ENT"      = "TIPO_ENTIDAD",
      "NOM_COMPLETO"  = "NOMBRES",
      "OBSERVACIONES" = "DETALLE"
    )
    for (origen in names(mapa)) {
      destino <- mapa[[origen]]
      if (origen %in% names(df) && !(destino %in% names(df))) {
        df <- dplyr::rename(df, !!destino := dplyr::all_of(origen))
      }
    }
    df
  }
  
  # Superset por compatibilidad
  columnas_superset <- c(
    "TIP_DOCUM", "COD_DOCUM", "TARGEN", "NOMBRE", "MCA_INH", "FEC_ACTU", "DETALLE",
    "TIPO_ENTIDAD", "NOMBRES", "LISTAS", "NOMBRE_O_RAZON_SOCIAL"
  )
  
  estandarizar_columnas <- function(df, hoja) {
    df <- normalizar_y_mapear(df)
    
    if (!("COD_DOCUM" %in% names(df))) {
      stop(paste0("La hoja '", hoja, "' no contiene la columna COD_DOCUM (ni mapeable)."))
    }
    
    # Crear faltantes del superset
    faltantes <- setdiff(columnas_superset, names(df))
    if (length(faltantes) > 0) for (col in faltantes) df[[col]] <- NA
    
    # Normalizar LISTAS (solo por robustez)
    if ("LISTAS" %in% names(df)) {
      df$LISTAS <- df$LISTAS |> as.character() |> stringr::str_trim() |> toupper()
      df$LISTAS <- dplyr::case_when(
        df$LISTAS %in% c("PEP", "PEPS") ~ "PEP",
        df$LISTAS %in% c("OBSERVADOS", "OBSERVADO") ~ "OBSERVADOS",
        TRUE ~ df$LISTAS
      )
    }
    
    # Orden y tipos (vectorizado y seguro)
    df <- df[, unique(c(columnas_superset, names(df))), drop = FALSE]
    df <- df |>
      dplyr::mutate(
        COD_DOCUM = as.character(COD_DOCUM),
        COD_DOCUM = stringr::str_trim(COD_DOCUM)
      )
    df
  }
  
  # ---------- Cargar Excel base: SOLO UNA HOJA ----------
  # Prioriza hoja 'datos' (case-insensitive); si no está, lee la PRIMERA hoja.
  # Lee TODAS las columnas como texto para preservar ceros a la izquierda.
  cargar_excel_base <- function(path) {
    if (!file.exists(path)) stop(paste0("No se encontró el archivo base en: ", path))
    
    hojas <- readxl::excel_sheets(path)
    hojas_lower <- tolower(hojas)
    idx_datos <- which(hojas_lower == "datos")
    hoja_sel <- if (length(idx_datos) == 1) hojas[idx_datos] else hojas[1]
    message("Leyendo hoja: ", hoja_sel, " del archivo base: ", basename(path))
    
    # Leer una vez para conocer el número de columnas
    tmp <- readxl::read_excel(path, sheet = hoja_sel, col_names = TRUE)
    n <- ncol(tmp)
    # Leer nuevamente forzando todo como texto
    df <- readxl::read_excel(path, sheet = hoja_sel, col_names = TRUE,
                             col_types = rep("text", n))
    
    df <- estandarizar_columnas(df, hoja_sel)
    df$FUENTE_HOJA <- hoja_sel
    
    # Filtrar si venía MCA_INH
    if ("MCA_INH" %in% names(df)) {
      df <- df |> dplyr::filter(is.na(MCA_INH) | MCA_INH != "S")
    }
    df <- df |> dplyr::distinct()
    
    list(base = df, hojas = hoja_sel)
  }
  
  # =============================
  # Opción A: seleccionar el Excel más reciente por fecha de modificación
  # =============================
  
  # Patrón de nombres (ajústalo si usas otro prefijo)
  PATRON_ARCHIVO <- "^reporte_.*\\.xlsx$"
  
  CARPETAS_BUSQUEDA <- function() {
    app_dir <- normalizePath(".", mustWork = TRUE)
    unique(c(app_dir, file.path(app_dir, "data")))
  }
  
  listar_candidatos <- function(pattern = PATRON_ARCHIVO) {
    dirs <- CARPETAS_BUSQUEDA()
    files <- unlist(lapply(dirs, function(d) {
      if (!dir.exists(d)) return(character(0))
      list.files(d, pattern = pattern, full.names = TRUE, ignore.case = TRUE)
    }))
    unique(files)
  }
  
  elegir_por_mtime <- function(files) {
    if (length(files) == 0) return(NA_character_)
    info <- file.info(files)
    files[order(info$mtime, decreasing = TRUE)][1]
  }
  
  resolver_archivo_base_mas_reciente <- function() {
    cand <- listar_candidatos(PATRON_ARCHIVO)
    message("Candidatos detectados (reporte_*.xlsx): ", paste(basename(cand), collapse = ", "))
    path <- elegir_por_mtime(cand)
    if (is.na(path)) NA_character_ else normalizePath(path, mustWork = TRUE)
  }
  
  # ---------- Resolver automáticamente el Excel BASE MÁS RECIENTE ----------
  observe({
    # 1) Intentar encontrar el más reciente por patrón 'reporte_*.xlsx'
    path_ok <- resolver_archivo_base_mas_reciente()
    
    # 2) Si no hay candidatos 'reporte_*.xlsx', usar fallback: carga_prueba.xlsx o peps_diciembre.xlsx
    if (is.na(path_ok)) {
      app_dir <- normalizePath(".", mustWork = TRUE)
      candidatos_fallback <- c(
        file.path(app_dir, "data", "carga_prueba.xlsx"),
        file.path(app_dir, "carga_prueba.xlsx"),
        file.path(app_dir, "data", "peps_diciembre.xlsx"),
        file.path(app_dir, "peps_diciembre.xlsx")
      )
      existe <- file.exists(candidatos_fallback)
      message("Fallback candidatos: ", paste(candidatos_fallback, collapse = " | "))
      message("Fallback existe: ", paste(existe, collapse = ", "))
      if (any(existe)) {
        path_ok <- normalizePath(candidatos_fallback[which(existe)[1]], mustWork = TRUE)
        message("Usando fallback: ", path_ok)
      }
    }
    
    # 3) Si aún no tenemos path, avisar y salir
    if (is.na(path_ok)) {
      showNotification(
        "No se encontró un archivo 'reporte_*.xlsx' ni los fallbacks (carga_prueba.xlsx/peps_diciembre.xlsx).",
        type = "error", duration = 10
      )
      base_listas(NULL)
      info_carga(list(path = NA_character_, hojas = character(0)))
      return(NULL)
    }
    
    # 4) Cargar el Excel (lee una sola hoja y todo como texto)
    tryCatch({
      res <- cargar_excel_base(path_ok)
      base_listas(res$base)
      info_carga(list(path = path_ok, hojas = res$hojas))
      showNotification(
        paste0("Excel base cargado (más reciente): ", basename(path_ok)),
        type = "message"
      )
    }, error = function(e) {
      message("ERROR al leer el Excel base reciente: ", e$message)
      base_listas(NULL)
      info_carga(list(path = path_ok, hojas = character(0)))
      showNotification(paste("Error leyendo el Excel base:", e$message),
                       type = "error", duration = 12)
    })
  })
  
  # ---------- UI: estado y detalles ----------
  output$estado_base <- renderUI({
    if (is.null(base_listas())) {
      tags$span(style = "color:#a94442;", "Excel base: NO cargado")
    } else {
      tags$span(style = "color:#3c763d;", "Excel base: cargado correctamente")
    }
  })
  
  output$detalle_base <- renderUI({
    inf <- info_carga()
    if (!is.null(inf$path) && !is.na(inf$path) && !is.null(base_listas())) {
      base <- base_listas()
      tags$small(
        tags$span(class = "path-trunc", tags$code(title = inf$path, inf$path)),
        sprintf("Filas: %d | Columnas: %d", nrow(base), ncol(base))
      )
    }
  })
  
  # === NUEVO: Mostrar hoja leída, archivo base (nombre) y última modificación ===
  output$hojas_base <- renderUI({
    inf <- info_carga()
    if (!is.null(inf$path) && !is.na(inf$path) && length(inf$hojas) > 0) {
      # Obtener fecha de modificación del archivo y formatearla en TZ Lima
      mtime <- tryCatch(file.info(inf$path)$mtime, error = function(e) NA)
      mtime_fmt <- if (!is.na(mtime)) format(mtime, tz = TZ_APP, "%d-%m-%Y %H:%M:%S") else "N/A"
      
      tags$small(
        tags$span(style = "color:#3c763d;", paste0("Hoja leída: ", paste(inf$hojas, collapse = ", "))),
        tags$br(),
        tags$span(
          class = "path-trunc",
          tags$code(
            title = inf$path,                                 # tooltip: ruta completa
            paste0("Archivo base: ", basename(inf$path))      # visible: nombre del archivo
          )
        ),
        tags$br(),
        paste0("Última modificación: ", mtime_fmt)
      )
    }
  })
  
  # === Procesar cruce (incluye SIN COINCIDENCIA para búsqueda masiva) ===
  observeEvent(input$procesar, {
    base <- base_listas()
    req(base)
    
    fecha_busqueda <- fecha_local()
    
    # 1) Construir tabla de consulta (manual o Excel)
    codigos_df <- NULL
    if (nzchar(input$doc_manual)) {
      codigos_df <- data.frame(
        COD_DOCUM = as.character(stringr::str_trim(input$doc_manual)),
        TIP_DOCUM = NA_character_,
        stringsAsFactors = FALSE
      )
    } else {
      req(input$archivo_consulta)
      consulta <- readxl::read_excel(input$archivo_consulta$datapath)
      names(consulta) <- names(consulta) |> stringr::str_trim() |> toupper()
      
      if (!("COD_DOCUM" %in% names(consulta)) && "COD_ID" %in% names(consulta)) {
        consulta <- consulta |> dplyr::rename(COD_DOCUM = COD_ID)
      }
      if (!("COD_DOCUM" %in% names(consulta))) {
        showNotification("El archivo de consulta debe tener la columna COD_DOCUM (o COD_ID).",
                         type = "error")
        return(NULL)
      }
      
      cols_consulta <- intersect(c("COD_DOCUM", "TIP_DOCUM"), names(consulta))
      codigos_df <- consulta[, cols_consulta, drop = FALSE]
    }
    
    # Normalizar y quitar duplicados de consulta
    codigos_df <- codigos_df |>
      dplyr::mutate(
        COD_DOCUM = as.character(COD_DOCUM),
        COD_DOCUM = stringr::str_trim(COD_DOCUM)
      ) |>
      dplyr::distinct()
    
    # 2) Left join desde la consulta hacia la base (para incluir SIN COINCIDENCIA)
    base_limpia <- base |>
      dplyr::mutate(
        COD_DOCUM = as.character(COD_DOCUM),
        COD_DOCUM = stringr::str_trim(COD_DOCUM)
      )
    
    cruce_full <- codigos_df |> dplyr::left_join(base_limpia, by = "COD_DOCUM")
    
    # --- Unificar TIP_DOCUM (evita TIP_DOCUM.x / TIP_DOCUM.y) ---
    if ("TIP_DOCUM.x" %in% names(cruce_full) || "TIP_DOCUM.y" %in% names(cruce_full)) {
      cruce_full <- cruce_full |>
        dplyr::mutate(
          TIP_DOCUM = dplyr::coalesce(.data[["TIP_DOCUM.x"]], .data[["TIP_DOCUM.y"]])
        ) |>
        dplyr::select(-dplyr::any_of(c("TIP_DOCUM.x", "TIP_DOCUM.y")))
    }
    
    # --- Unificar NOMBRE ---
    cruce_full <- cruce_full |>
      dplyr::mutate(
        NOMBRE = dplyr::coalesce(NOMBRE, NOMBRES, NOMBRE_O_RAZON_SOCIAL)
      ) |>
      dplyr::select(-dplyr::any_of(c("NOMBRES", "NOMBRE_O_RAZON_SOCIAL")))
    
    # 3) Completar campos para los que no tuvieron match
    sin_match <- is.na(cruce_full$FUENTE_HOJA)
    cruce_full <- cruce_full |>
      dplyr::mutate(
        FUENTE_HOJA = dplyr::if_else(sin_match, "SIN COINCIDENCIA", FUENTE_HOJA),
        TIPO_ENTIDAD = dplyr::if_else(sin_match, NA_character_, TIPO_ENTIDAD),
        LISTAS = dplyr::if_else(sin_match, NA_character_, LISTAS),
        NOMBRE = dplyr::if_else(sin_match, NA_character_, NOMBRE),
        FECHA_BUSQUEDA = fecha_busqueda,
        ESTADO = dplyr::if_else(sin_match, "NO ENCONTRADO", "ENCONTRADO")
      )
    
    # 4) Seleccionar solo columnas visibles (sin romper si faltan)
    cols <- cols_presentes(cruce_full, columnas_mostrar)
    cruce_final <- cruce_full |>
      dplyr::select(dplyr::all_of(cols)) |>
      dplyr::distinct() |>
      dplyr::mutate(.coincide = ESTADO == "ENCONTRADO", .es_pep = LISTAS == "PEP") |>
      dplyr::arrange(dplyr::desc(.coincide), dplyr::desc(.es_pep), TIP_DOCUM, COD_DOCUM) |>
      dplyr::select(-.coincide, -.es_pep)
    
    resultado_cruce(cruce_final)
    
    output$tabla_resultado <- renderDT({
      datatable(
        cruce_final,
        options = list(
          pageLength = 10,
          scrollX = TRUE,
          dom = 'Brtip',        # sin 'f' para ocultar buscador global; usa 'Bfrtip' si lo quieres
          buttons = c('copy', 'csv', 'excel')
        ),
        rownames = FALSE,
        filter = "none"         # quita fila de filtros ("All")
      )
    })
  })
  
  # Excel
  output$descargar_excel <- downloadHandler(
    filename = function() paste0(
      "resultado_cruce_", fecha_local("%Y%m%d_%H%M%S"), ".xlsx"
    ),
    content = function(file) {
      df <- resultado_cruce()
      cols <- cols_presentes(df, columnas_mostrar)
      df <- dplyr::select(df, dplyr::all_of(cols))
      write_xlsx(df, path = file)
    }
  )
  
  # PDF horizontal (landscape) robusto (sin LaTeX)
  output$descargar_pdf <- downloadHandler(
    filename = function() paste0(
      "resultado_cruce_", fecha_local("%Y%m%d_%H%M%S"), ".pdf"
    ),
    contentType = "application/pdf",
    content = function(file) {
      df <- resultado_cruce()
      cols <- cols_presentes(df, columnas_mostrar)
      df <- dplyr::select(df, dplyr::all_of(cols))
      
      # Método A: pagedown + Chrome
      if (has_pkg("pagedown") && !is.null(pagedown::find_chrome())) {
        html_path <- tempfile(fileext = ".html")
        html_header <- '
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Resultado de búsqueda</title>
<style>
@page { size: letter landscape; margin: 24mm; }
body { font-family: Arial, sans-serif; margin: 24px; }
h1 { margin-bottom: 4px; font-size: 18px; }
p  { margin: 0 0 8px 0; font-size: 11px; }
table { border-collapse: collapse; width: 100%; font-size: 10px; table-layout: fixed; }
th, td { border: 1px solid #777; padding: 5px; text-align: left; vertical-align: top; word-break: break-word; hyphens: auto; }
thead { background: #f0f0f0; }
</style>
</head>
<body>
'
        html_footer <- '
</body>
</html>
'
        tbl_html <- knitr::kable(df, format = "html", table.attr = 'class="table"')
        
        html_content <- paste0(
          html_header,
          sprintf("<h1>Resultado de búsqueda en listas</h1>"),
          sprintf("<p><strong>Generado:</strong> %s</p>", fecha_local()),
          sprintf("<p><strong>Registros:</strong> %d</p>", nrow(df)),
          as.character(tbl_html),
          html_footer
        )
        writeLines(html_content, con = html_path)
        pagedown::chrome_print(input = html_path, output = file)
        return(invisible(NULL))
      }
      
      # Método B: gridExtra::tableGrob
      if (has_pkg("gridExtra")) {
        pdf(file, width = 11, height = 8.5)
        grid::grid.newpage()
        grid::grid.text("Resultado de búsqueda en listas", x = 0.5, y = 0.95,
                        gp = grid::gpar(fontsize = 14, fontface = "bold"))
        grid::grid.text(sprintf("Generado: %s", fecha_local()),
                        x = 0.5, y = 0.92, gp = grid::gpar(fontsize = 10))
        grid::grid.text(sprintf("Registros: %d", nrow(df)),
                        x = 0.5, y = 0.89, gp = grid::gpar(fontsize = 10))
        tg <- gridExtra::tableGrob(df, rows = NULL, theme = gridExtra::ttheme_minimal(base_size = 9))
        tg$widths <- rep(grid::unit(1, "null"), ncol(df))
        grid::grid.draw(tg)
        dev.off()
        return(invisible(NULL))
      }
      
      # Método C: base R
      pdf(file, width = 11, height = 8.5)
      op <- par(mar = c(1,1,1,1))
      plot.new()
      mtext("Resultado de búsqueda en listas", side = 3, line = -2, cex = 1.2, font = 2)
      mtext(sprintf("Generado: %s", fecha_local()), side = 3, line = -1, cex = 0.9)
      mtext(sprintf("Registros: %d", nrow(df)), side = 3, line = 0, cex = 0.9)
      N <- min(nrow(df), 60); y <- 0.8; step <- 0.015
      headers <- paste(names(df), collapse = " | ")
      text(x = 0.02, y = y, labels = headers, adj = c(0,1), cex = 0.7); y <- y - step
      for (i in seq_len(N)) {
        row_str <- paste(df[i, ], collapse = " | ")
        text(x = 0.02, y = y, labels = row_str, adj = c(0,1), cex = 0.7)
        y <- y - step; if (y < 0.05) break
      }
      par(op); dev.off()
    }
  )
}

shinyApp(ui = ui, server = server)

