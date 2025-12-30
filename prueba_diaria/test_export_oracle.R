

suppressPackageStartupMessages({
  library(RODBC)
  library(openxlsx)
})

DSN         <- "TRON_BI"
ORACLE_USER <- Sys.getenv("ORACLE_USER", "jpuntri")
ORACLE_PASS <- Sys.getenv("ORACLE_PASS", "Peru.500600")  # ideal mover a variable de entorno
OUTPUT_DIR  <- "D:/Descargas/Pedidos Hugo/Shiny_busqueda_listas_UC/SHINY_BUSQUEDA_EN_LISTAS_UC/ver_log"
FILE_PREFIX <- "reporte"
SHEET_NAME  <- "Datos"
ADD_TIME    <- TRUE

QUERY <- "
SELECT 
    t73.TIPO_ENTIDAD,t73.TIP_DOCUM,t73.COD_DOCUM,
    (t73.APELLIDO_PATERNO || ' ' || t73.APELLIDO_MATERNO || ' ' ||
            t73.NOMBRE1 || ' ' || t73.NOMBRE2 || ' ' || t73.NOMBRE3) AS NOMBRE,'PEP' AS LISTAS FROM targen73 t73
 --WHERE ROWNUM < 2  -- solo para pruebas
UNION ALL
SELECT
    /* targen66 no tiene TIPO_ENTIDAD */
    CAST(NULL AS VARCHAR2(150)) AS TIPO_ENTIDAD,t66.TIP_DOCUM,t66.COD_DOCUM,
    (t66.NOMBRE_O_RAZON_SOCIAL) AS NOMBRE,'OBSERVADOS' AS LISTAS FROM targen66 t66
 --WHERE ROWNUM < 2  -- solo para pruebas
"

if (!dir.exists(OUTPUT_DIR)) dir.create(OUTPUT_DIR, recursive = TRUE, showWarnings = FALSE)
date_str  <- format(Sys.Date(), "%Y-%m-%d")
time_str  <- format(Sys.time(), "%H%M")
base_name <- if (ADD_TIME) paste0(FILE_PREFIX, "_", date_str, "_", time_str) else paste0(FILE_PREFIX, "_", date_str)
out_path  <- file.path(OUTPUT_DIR, paste0(base_name, ".xlsx"))

excel_limit <- 1048576L

# Conexión y tuning
ch <- odbcConnect(dsn = DSN, uid = ORACLE_USER, pwd = ORACLE_PASS, believeNRows = FALSE)
on.exit(try(odbcClose(ch), silent = TRUE), add = TRUE)

attr(ch, "rows_at_time") <- 5000L  # prueba con 5000, luego 10000 si tu red/cliente soporta

cat("Ejecutando consulta completa...\n")
res <- try(sqlQuery(ch, QUERY, as.is = TRUE, errors = TRUE), silent = TRUE)

if (inherits(res, "try-error")) {
  # Capturamos el mensaje real del error
  err <- attr(res, "condition")
  cat("Error en sqlQuery: ", conditionMessage(err), "\n")
  stop(conditionMessage(err))
}
if (!is.data.frame(res)) stop("La consulta no retornó un data.frame.")
n <- nrow(res)
cat(sprintf("Filas recibidas: %d\n", n))

if (n > excel_limit) stop(sprintf("Resultado (%d) excede el límite de Excel (%d).", n, excel_limit))

wb <- createWorkbook()
addWorksheet(wb, SHEET_NAME)
writeData(wb, SHEET_NAME, res, withFilter = FALSE)

tryCatch({
  saveWorkbook(wb, out_path, overwrite = TRUE)
  cat(sprintf("Excel generado: %s | Filas: %d\n", out_path, n))
}, error = function(e) {
  cat("ERROR en saveWorkbook: ", conditionMessage(e), "\n")
  stop(e)
})

