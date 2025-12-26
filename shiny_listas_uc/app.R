


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

# Helper para obtener fecha/hora formateada siempre en Lima
fecha_local <- function(fmt = "%Y-%m-%d %H:%M:%S") {
  format(Sys.time(), tz = TZ_APP, fmt)
}

# Logs para diagnosticar TZ en los registros del servidor
message("== Zona horaria de la app ==")
message("TZ env: '", Sys.getenv("TZ"), "'")
message("Sys.timezone(): ", Sys.timezone())
message("as.POSIXlt(Sys.time())$zone: ",
        paste(unique(as.POSIXlt(Sys.time())$zone), collapse = ", "))

ui <- fluidPage(
  titlePanel("Cruce de Documentos con listas UC (Excel fijo)"),
  
  # ---------- CSS para truncar ruta con ellipsis y mostrar tooltip ----------
  tags$style(HTML("
    /* Estiliza textos del sidebar */
    .sidebarPanel small {
      display: block;
      line-height: 1.2;
      margin-bottom: 6px;
    }
    /* Clase para truncar en una sola línea con puntos suspensivos */
    .path-trunc {
      max-width: 100%;
      display: block;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    /* Estilo monoespaciado y buen contraste para rutas */
    code {
      font-family: Consolas, 'Courier New', monospace;
    }
    /* (Opcional) Evitar que el sidebar se haga demasiado angosto */
    @media (min-width: 768px) {
      .col-sm-4 { min-width: 320px; }
    }
  ")),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("archivo_consulta", "Sube tu Excel y que el nombre de la columna sea COD_DOCUM",
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
  # TIP_DOCUM primero, luego COD_DOCUM. LISTAS removida del resultado.
  columnas_mostrar <- c(
    "TIP_DOCUM",             # primero
    "COD_DOCUM",             # luego
    "FUENTE_HOJA",
    "TIPO_ENTIDAD",
    "NOMBRES",
    # "LISTAS",              # <- oculta del resultado final
    "NOMBRE_O_RAZON_SOCIAL",
    "FECHA_BUSQUEDA",
    "ESTADO" # opcional; se ve si hubo match o no
  )
  
  # Helper: seleccionar solo columnas presentes (evita errores)
  cols_presentes <- function(df, deseadas) intersect(deseadas, names(df))
  
  # Reactivos
  base_listas      <- reactiveVal(NULL)
  info_carga       <- reactiveVal(list(path = NA_character_, hojas = character(0)))
  resultado_cruce  <- reactiveVal(data.frame())
  
  # Hojas objetivo
  hojas_objetivo <- c("peps", "obserbados", "observados")
  
  # Normaliza encabezados y mapea nombres comunes
  normalizar_y_mapear <- function(df) {
    names(df) <- names(df) |>
      str_trim() |>
      toupper()
    
    mapa <- c(
      "COD_ID"        = "COD_DOCUM",
      "TIPO_ID"       = "TIP_DOCUM",
      "NOM_COMPLETO"  = "NOMBRES",   # si prefieres mapear a NOMBRE_O_RAZON_SOCIAL, cambia aquí
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
  
  # Superset para estandarizar
  columnas_superset <- c(
    "TIP_DOCUM", "COD_DOCUM", "TARGEN", "NOMBRE", "MCA_INH", "FEC_ACTU", "DETALLE",
    "TIPO_ENTIDAD", "NOMBRES", "LISTAS", "NOMBRE_O_RAZON_SOCIAL"
  )
  
  # Estandariza columnas al superset
  estandarizar_columnas <- function(df, hoja) {
    df <- normalizar_y_mapear(df)
    
    if (!("COD_DOCUM" %in% names(df))) {
      stop(paste0("La hoja '", hoja, "' no contiene la columna COD_DOCUM (ni mapeable)."))
    }
    
    faltantes <- setdiff(columnas_superset, names(df))
    if (length(faltantes) > 0) for (col in faltantes) df[[col]] <- NA
    
    df <- df[, unique(c(columnas_superset, names(df))), drop = FALSE]
    
    df <- df |> mutate(COD_DOCUM = as.character(COD_DOCUM) |> str_trim())
    df
  }
  
  # Cargar Excel y consolidar
  cargar_excel_base <- function(path) {
    if (!file.exists(path)) stop(paste0("No se encontró el archivo base en: ", path))
    
    hojas <- readxl::excel_sheets(path)
    hojas_lower <- tolower(hojas)
    idx_obj <- which(hojas_lower %in% hojas_objetivo)
    idx_lectura <- if (length(idx_obj) == 0) seq_along(hojas) else idx_obj
    
    lista_hojas <- purrr::map(idx_lectura, function(i) {
      hoja <- hojas[i]
      df <- readxl::read_excel(path, sheet = hoja)
      df <- estandarizar_columnas(df, hoja)
      df$FUENTE_HOJA <- hoja
      df
    })
    
    base <- dplyr::bind_rows(lista_hojas)
    
    if ("MCA_INH" %in% names(base)) {
      base <- base |> dplyr::filter(is.na(MCA_INH) | MCA_INH != "S")
    }
    
    base <- base |> dplyr::distinct()
    
    list(base = base, hojas = hojas[idx_lectura])
  }
  
  # Resolver ruta del Excel (data/ primero, luego raíz)
  observe({
    app_dir <- normalizePath(".", mustWork = TRUE)
    
    candidatos <- c(
      file.path(app_dir, "data", "peps_diciembre.xlsx"),
      file.path(app_dir, "peps_diciembre.xlsx")
    )
    existe <- file.exists(candidatos)
    
    # Logs
    message("WD: ", getwd())
    message("app_dir: ", app_dir)
    message("Candidatos: ", paste(candidatos, collapse = " | "))
    message("Existe: ", paste(existe, collapse = ", "))
    if (dir.exists(file.path(app_dir, "data"))) {
      message("list.files(data): ", paste(list.files(file.path(app_dir, "data")), collapse = ", "))
    } else {
      message("No existe carpeta data/")
    }
    message("list.files(app_dir): ", paste(list.files(app_dir), collapse = ", "))
    
    if (!any(existe)) {
      showNotification(
        paste0("No se encontró el archivo base. Revise:\n- ", candidatos[1], "\n- ", candidatos[2]),
        type = "error", duration = 12
      )
      base_listas(NULL)
      info_carga(list(path = NA_character_, hojas = character(0)))
      return(NULL)
    }
    
    path_ok <- normalizePath(candidatos[which(existe)[1]], mustWork = TRUE)
    
    tryCatch({
      res <- cargar_excel_base(path_ok)
      base_listas(res$base)
      info_carga(list(path = path_ok, hojas = res$hojas))
      showNotification("Excel base cargado correctamente.", type = "message")
    }, error = function(e) {
      base_listas(NULL)
      info_carga(list(path = path_ok, hojas = character(0)))
      showNotification(paste("Error leyendo el Excel base:", e$message),
                       type = "error", duration = 12)
    })
  })
  
  # UI: estado y detalles
  output$estado_base <- renderUI({
    if (is.null(base_listas())) {
      tags$span(style = "color:#a94442;", "Excel base: NO cargado")
    } else {
      tags$span(style = "color:#3c763d;", "Excel base: cargado correctamente")
    }
  })
  
  # ---------- Opción A: ruta truncada con ellipsis y tooltip ----------
  output$detalle_base <- renderUI({
    inf <- info_carga()
    if (!is.null(inf$path) && !is.na(inf$path) && !is.null(base_listas())) {
      base <- base_listas()
      tags$small(
        # Línea de la ruta (truncada, con tooltip). Se usa <code> para monoespaciado.
        tags$span(
          class = "path-trunc",
          tags$code(title = inf$path, inf$path)
        ),
        # Línea de métricas
        sprintf("Filas: %d | Columnas: %d", nrow(base), ncol(base))
      )
    }
  })
  
  output$hojas_base <- renderUI({
    inf <- info_carga()
    if (length(inf$hojas) > 0) {
      tags$small(paste0("Hojas cargadas: ", paste(inf$hojas, collapse = ", ")))
    }
  })
  
  # === Procesar cruce (incluye SIN COINCIDENCIA para búsqueda masiva) ===
  observeEvent(input$procesar, {
    base <- base_listas()
    req(base)
    
    # Fecha/Hora siempre en America/Lima
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
      
      # Mapear COD_ID -> COD_DOCUM si fuese necesario
      if (!("COD_DOCUM" %in% names(consulta)) && "COD_ID" %in% names(consulta)) {
        consulta <- consulta |> dplyr::rename(COD_DOCUM = COD_ID)
      }
      if (!("COD_DOCUM" %in% names(consulta))) {
        showNotification("El archivo de consulta debe tener la columna COD_DOCUM (o COD_ID).",
                         type = "error")
        return(NULL)
      }
      
      # Nos quedamos con COD_DOCUM y TIP_DOCUM si viene
      cols_consulta <- intersect(c("COD_DOCUM", "TIP_DOCUM"), names(consulta))
      codigos_df <- consulta[, cols_consulta, drop = FALSE]
    }
    
    # Normalizar y quitar duplicados de consulta
    codigos_df <- codigos_df |>
      dplyr::mutate(COD_DOCUM = as.character(COD_DOCUM) |> stringr::str_trim()) |>
      dplyr::distinct()
    
    # 2) Left join desde la consulta hacia la base (para incluir SIN COINCIDENCIA)
    base_limpia <- base |>
      dplyr::mutate(COD_DOCUM = as.character(COD_DOCUM) |> stringr::str_trim())
    
    cruce_full <- codigos_df |>
      dplyr::left_join(base_limpia, by = "COD_DOCUM")
    
    # --- Unificar TIP_DOCUM (evita TIP_DOCUM.x / TIP_DOCUM.y) ---
    if ("TIP_DOCUM.x" %in% names(cruce_full) || "TIP_DOCUM.y" %in% names(cruce_full)) {
      cruce_full <- cruce_full |>
        dplyr::mutate(
          TIP_DOCUM = dplyr::coalesce(.data[["TIP_DOCUM.x"]], .data[["TIP_DOCUM.y"]])
        ) |>
        dplyr::select(-dplyr::any_of(c("TIP_DOCUM.x", "TIP_DOCUM.y")))
    }
    # --- Fin unificación ---
    
    # 3) Completar campos para los que no tuvieron match
    sin_match <- is.na(cruce_full$FUENTE_HOJA)
    cruce_full <- cruce_full |>
      dplyr::mutate(
        FUENTE_HOJA = dplyr::if_else(sin_match, "SIN COINCIDENCIA", FUENTE_HOJA),
        TIPO_ENTIDAD = dplyr::if_else(sin_match, NA_character_, TIPO_ENTIDAD),
        NOMBRES = dplyr::if_else(sin_match, NA_character_, NOMBRES),
        LISTAS = dplyr::if_else(sin_match, NA_character_, LISTAS),  # se mantiene internamente
        NOMBRE_O_RAZON_SOCIAL = dplyr::if_else(sin_match, NA_character_, NOMBRE_O_RAZON_SOCIAL),
        FECHA_BUSQUEDA = fecha_busqueda,
        ESTADO = dplyr::if_else(sin_match, "NO ENCONTRADO", "ENCONTRADO")
      )
    
    # 4) Seleccionar solo columnas visibles (sin romper si faltan)
    cols <- cols_presentes(cruce_full, columnas_mostrar)
    cruce_final <- cruce_full |>
      dplyr::select(dplyr::all_of(cols)) |>
      dplyr::distinct()
    
    # (Opcional) Orden: primero encontrados, luego por TIP_DOCUM y COD_DOCUM
    cruce_final <- cruce_final |>
      dplyr::mutate(.coincide = ESTADO == "ENCONTRADO") |>
      dplyr::arrange(dplyr::desc(.coincide), TIP_DOCUM, COD_DOCUM) |>
      dplyr::select(-.coincide)
    
    # 5) Guardar y mostrar
    resultado_cruce(cruce_final)
    
    output$tabla_resultado <- renderDT({
      datatable(
        cruce_final,
        options = list(pageLength = 10, scrollX = TRUE),
        rownames = FALSE
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
      
      # Método A: pagedown + Chrome (horizontal vía CSS @page)
      if (has_pkg("pagedown") && !is.null(pagedown::find_chrome())) {
        html_path <- tempfile(fileext = ".html")
        html_header <- '
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Resultado de búsqueda</title>
<style>
@page { size: letter landscape; margin: 24mm; }  /* <-- Carta horizontal */
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
      
      # Método B: gridExtra::tableGrob (horizontal por dimensiones del dispositivo)
      if (has_pkg("gridExtra")) {
        pdf(file, width = 11, height = 8.5)  # <-- Carta horizontal
        grid::grid.newpage()
        grid::grid.text("Resultado de búsqueda en listas", x = 0.5, y = 0.95,
                        gp = grid::gpar(fontsize = 14, fontface = "bold"))
        grid::grid.text(sprintf("Generado: %s", fecha_local()),
                        x = 0.5, y = 0.92, gp = grid::gpar(fontsize = 10))
        grid::grid.text(sprintf("Registros: %d", nrow(df)),
                        x = 0.5, y = 0.89, gp = grid::gpar(fontsize = 10))
        tg <- gridExtra::tableGrob(
          df, rows = NULL, theme = gridExtra::ttheme_minimal(base_size = 9)
        )
        tg$widths <- rep(grid::unit(1, "null"), ncol(df))
        grid::grid.draw(tg)
        dev.off()
        return(invisible(NULL))
      }
      
      # Método C: PDF básico con base R (horizontal por width/height)
      pdf(file, width = 11, height = 8.5)  # <-- Carta horizontal
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
