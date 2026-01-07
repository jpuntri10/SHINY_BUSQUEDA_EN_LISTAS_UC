


library(shiny)
library(readxl)
library(dplyr)
library(DT)
library(writexl)
library(stringr)
library(purrr)
library(tidyr)
library(gridExtra)   # <- asegúrate que se instale en shinyapps.io

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

ui <- fluidPage(
  titlePanel("Cruce de Documentos con listas UC (Excel base automático: más reciente)"),
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
  
  # === Columnas visibles/exportables (sin FUENTE_HOJA; NOMBRE 3ra) ===
  columnas_mostrar <- c(
    "TIP_DOCUM",
    "COD_DOCUM",
    "NOMBRE",          # 3ra columna (más ancha)
    "LISTAS",
    "TIPO_ENTIDAD",
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
  
  columnas_superset <- c(
    "TIP_DOCUM", "COD_DOCUM", "TARGEN", "NOMBRE", "MCA_INH", "FEC_ACTU", "DETALLE",
    "TIPO_ENTIDAD", "NOMBRES", "LISTAS", "NOMBRE_O_RAZON_SOCIAL", "FUENTE_HOJA"
  )
  
  estandarizar_columnas <- function(df, hoja) {
    df <- normalizar_y_mapear(df)
    if (!("COD_DOCUM" %in% names(df))) {
      stop(paste0("La hoja '", hoja, "' no contiene la columna COD_DOCUM (ni mapeable)."))
    }
    faltantes <- setdiff(columnas_superset, names(df))
    if (length(faltantes) > 0) for (col in faltantes) df[[col]] <- NA
    
    if ("LISTAS" %in% names(df)) {
      df$LISTAS <- df$LISTAS |> as.character() |> stringr::str_trim() |> toupper()
      df$LISTAS <- dplyr::case_when(
        df$LISTAS %in% c("PEP", "PEPS") ~ "PEP",
        df$LISTAS %in% c("OBSERVADOS", "OBSERVADO") ~ "OBSERVADOS",
        TRUE ~ df$LISTAS
      )
    }
    
    df <- df[, unique(c(columnas_superset, names(df))), drop = FALSE]
    df <- df |>
      dplyr::mutate(
        COD_DOCUM = as.character(COD_DOCUM),
        COD_DOCUM = stringr::str_trim(COD_DOCUM)
      )
    df
  }
  
  # ---------- Cargar Excel base: SOLO UNA HOJA ----------
  cargar_excel_base <- function(path) {
    if (!file.exists(path)) stop(paste0("No se encontró el archivo base en: ", path))
    
    hojas <- readxl::excel_sheets(path)
    hojas_lower <- tolower(hojas)
    idx_datos <- which(hojas_lower == "datos")
    hoja_sel <- if (length(idx_datos) == 1) hojas[idx_datos] else hojas[1]
    message("Leyendo hoja: ", hoja_sel, " del archivo base: ", basename(path))
    
    tmp <- readxl::read_excel(path, sheet = hoja_sel, col_names = TRUE)
    n <- ncol(tmp)
    df <- readxl::read_excel(path, sheet = hoja_sel, col_names = TRUE,
                             col_types = rep("text", n))
    
    df <- estandarizar_columnas(df, hoja_sel)
    df$FUENTE_HOJA <- hoja_sel
    
    if ("MCA_INH" %in% names(df)) {
      df <- df |> dplyr::filter(is.na(MCA_INH) | MCA_INH != "S")
    }
    df <- df |> dplyr::distinct()
    
    list(base = df, hojas = hoja_sel)
  }
  
  # =============================
  # Opción A: seleccionar el Excel más reciente por fecha de modificación
  # =============================
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
    path_ok <- resolver_archivo_base_mas_reciente()
    
    if (is.na(path_ok)) {
      app_dir <- normalizePath(".", mustWork = TRUE)
      candidatos_fallback <- c(
        file.path(app_dir, "data", "carga_prueba.xlsx"),
        file.path(app_dir, "carga_prueba.xlsx"),
        file.path(app_dir, "data", "peps_diciembre.xlsx"),
        file.path(app_dir, "peps_diciembre.xlsx")
      )
      existe <- file.exists(candidatos_fallback)
      if (any(existe)) {
        path_ok <- normalizePath(candidatos_fallback[which(existe)[1]], mustWork = TRUE)
      }
    }
    
    if (is.na(path_ok)) {
      showNotification(
        "No se encontró un archivo 'reporte_*.xlsx' ni los fallbacks (carga_prueba.xlsx/peps_diciembre.xlsx).",
        type = "error", duration = 10
      )
      base_listas(NULL)
      info_carga(list(path = NA_character_, hojas = character(0)))
      return(NULL)
    }
    
    tryCatch({
      res <- cargar_excel_base(path_ok)
      base_listas(res$base)
      info_carga(list(path = path_ok, hojas = res$hojas))
      showNotification(
        paste0("Excel base cargado (más reciente): ", basename(path_ok)),
        type = "message"
      )
    }, error = function(e) {
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
  
  output$hojas_base <- renderUI({
    inf <- info_carga()
    if (!is.null(inf$path) && !is.na(inf$path) && length(inf$hojas) > 0) {
      mtime <- tryCatch(file.info(inf$path)$mtime, error = function(e) NA)
      mtime_fmt <- if (!is.na(mtime)) format(mtime, tz = TZ_APP, "%d-%m-%Y %H:%M:%S") else "N/A"
      tags$small(
        tags$span(style = "color:#3c763d;", paste0("Hoja leída: ", paste(inf$hojas, collapse = ", "))),
        tags$br(),
        tags$span(
          class = "path-trunc",
          tags$code(title = inf$path, paste0("Archivo base: ", basename(inf$path)))
        ),
        tags$br(),
        paste0("Última modificación: ", mtime_fmt)
      )
    }
  })
  
  # === Procesar cruce (incluye SIN COINCIDENCIA) ===
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
    
    # 2) Left join desde la consulta hacia la base
    base_limpia <- base |>
      dplyr::mutate(
        COD_DOCUM = as.character(COD_DOCUM),
        COD_DOCUM = stringr::str_trim(COD_DOCUM)
      )
    
    cruce_full <- codigos_df |> dplyr::left_join(base_limpia, by = "COD_DOCUM")
    
    # --- Unificar TIP_DOCUM ---
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
        FUENTE_HOJA    = dplyr::if_else(sin_match, "SIN COINCIDENCIA", FUENTE_HOJA),
        TIPO_ENTIDAD   = dplyr::if_else(sin_match, NA_character_, TIPO_ENTIDAD),
        LISTAS         = dplyr::if_else(sin_match, NA_character_, LISTAS),
        NOMBRE         = dplyr::if_else(sin_match, NA_character_, NOMBRE),
        FECHA_BUSQUEDA = fecha_busqueda,
        ESTADO         = dplyr::if_else(sin_match, "NO ENCONTRADO", "ENCONTRADO")
      )
    
    # 4) Salida ordenada con NOMBRE 3ra
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
          dom = 'Brtip',
          buttons = c('copy', 'csv', 'excel')
        ),
        rownames = FALSE,
        filter = "none"
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
      writexl::write_xlsx(df, path = file)
    }
  )
  
  # ============================
  # PDF (solo R puro: gridExtra o Base R)
  # ============================
  output$descargar_pdf <- downloadHandler(
    filename = function() paste0(
      "resultado_cruce_", fecha_local("%Y%m%d_%H%M%S"), ".pdf"
    ),
    contentType = "application/pdf",
    content = function(file) {
      df <- resultado_cruce()
      cols <- cols_presentes(df, columnas_mostrar)
      df <- dplyr::select(df, dplyr::all_of(cols))
      
      # ---- Sanitización segura ----
      clean_text <- function(x) {
        x <- as.character(x)
        x <- iconv(x, from = "", to = "UTF-8", sub = " ")
        x <- gsub("[[:cntrl:]]", " ", x, perl = TRUE)
        x <- gsub("\r\n|\r", "\n", x, perl = TRUE)
        trimws(x)
      }
      chr_cols <- vapply(df, is.character, logical(1))
      if (any(chr_cols)) df[chr_cols] <- lapply(df[chr_cols], clean_text)
      
      # ---- SIEMPRE escribir un PDF (tryCatch) ----
      tryCatch({
        pdf(file, width = 11, height = 8.5)  # Carta apaisado
        on.exit(dev.off(), add = TRUE)
        
        chunk_size <- 24  # menos filas por página para evitar cortes
        n_pages <- max(1L, ceiling(max(1L, nrow(df)) / chunk_size))
        
        for (i in seq_len(n_pages)) {
          idx_ini <- ((i - 1) * chunk_size) + 1
          idx_fin <- min(i * chunk_size, nrow(df))
          df_i <- if (nrow(df) == 0) df[0, , drop = FALSE] else df[idx_ini:idx_fin, , drop = FALSE]
          
          title_g <- grid::textGrob("Resultado de búsqueda en listas UC",
                                    gp = grid::gpar(fontsize = 16, fontface = "bold"),
                                    x = 0.5, y = 0.5, just = "center")
          subtitle_g <- grid::textGrob(
            sprintf("Generado: %s    |    Registros totales: %d    |    Página %d de %d",
                    fecha_local(), nrow(df), i, n_pages),
            gp = grid::gpar(fontsize = 11),
            x = 0.5, y = 0.5, just = "center"
          )
          spacer_top <- grid::nullGrob()
          spacer_bottom <- grid::nullGrob()
          
          if (nrow(df_i) == 0) {
            msg_g <- grid::textGrob("No hay registros para mostrar.",
                                    gp = grid::gpar(fontsize = 11),
                                    x = 0.5, y = 0.5, just = "center")
            page_g <- gridExtra::arrangeGrob(
              title_g, subtitle_g, spacer_top, msg_g, spacer_bottom, ncol = 1,
              heights = grid::unit(c(0.18, 0.12, 0.06, 0.60, 0.04), "npc")
            )
            grid::grid.newpage(); grid::grid.draw(page_g)
          } else {
            tg <- gridExtra::tableGrob(
              df_i, rows = NULL,
              theme = gridExtra::ttheme_minimal(
                base_size = 9,
                core = list(fg_params = list(hjust = 0.5, x = grid::unit(0.5, "npc"))),
                colhead = list(fg_params = list(hjust = 0.5, x = grid::unit(0.5, "npc")))
              )
            )
            # Ensanchar columna 3 (NOMBRE) y alinear a la izquierda
            tg$widths <- rep(grid::unit(1, "null"), ncol(df_i))
            if (ncol(df_i) >= 3) tg$widths[3] <- grid::unit(3, "null")
            lay <- tg$layout
            idx_head_col3 <- which(lay$name == "colhead" & lay$l == 3)
            for (j in idx_head_col3) { tg$grobs[[j]]$x <- grid::unit(0, "npc"); tg$grobs[[j]]$hjust <- 0 }
            idx_core_col3 <- which(lay$name == "core" & lay$l == 3)
            for (j in idx_core_col3) { tg$grobs[[j]]$x <- grid::unit(0, "npc"); tg$grobs[[j]]$hjust <- 0 }
            
            page_g <- gridExtra::arrangeGrob(
              title_g, subtitle_g, spacer_top, tg, spacer_bottom,
              ncol = 1,
              heights = grid::unit(c(0.18, 0.12, 0.06, 0.60, 0.04), "npc")
            )
            grid::grid.newpage(); grid::grid.draw(page_g)
          }
        }
      }, error = function(e) {
        # Fallback: PDF de 1 página con el mensaje de error (nunca 500)
        pdf(file, width = 11, height = 8.5)
        on.exit(dev.off(), add = TRUE)
        par(mar = c(1,1,1,1)); plot.new()
        mtext("Error generando PDF", side = 3, line = -1.5, cex = 1.3, font = 2)
        text(0.02, 0.92, labels = fecha_local(), adj = c(0,1), cex = 0.9)
        text(0.02, 0.86, labels = paste("Mensaje del servidor:", e$message), adj = c(0,1), cex = 0.85)
        text(0.02, 0.80, labels = "Intenta nuevamente o contacta al administrador.", adj = c(0,1), cex = 0.85)
      })
    }
  )
}

shinyApp(ui = ui, server = server)
