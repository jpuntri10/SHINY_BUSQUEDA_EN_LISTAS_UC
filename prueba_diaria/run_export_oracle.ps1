


# D:\Descargas\prueba_diaria\run_export_oracle.ps1
# --- rutas ---
$Rscript = 'D:\Program Files\R\R-4.5.1\bin\x64\Rscript.exe'   # ajusta si es distinto
$Rfile   = 'D:\Descargas\Pedidos Hugo\Shiny_busqueda_listas_UC\SHINY_BUSQUEDA_EN_LISTAS_UC\prueba_diaria\test_export_oracle.R'
$logDir  = 'D:\Descargas\Pedidos Hugo\Shiny_busqueda_listas_UC\SHINY_BUSQUEDA_EN_LISTAS_UC\Ver_log'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

# --- log ---
$ts      = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$logMain = Join-Path $logDir "export_prueba_$ts.log"

try {
    # Ejecuta Rscript y captura stdout + stderr juntos
    $rOut = & "$Rscript" "--vanilla" "$Rfile" 2>&1

    # Mantén SOLO estas líneas en el log:
    # - Filas recibidas: N
    # - OK: Excel generado: ...
    # - ERROR: ...
    $filtered = $rOut | Where-Object {
        $_ -match '^Filas recibidas:' -or $_ -match '^OK:' -or $_ -match '^ERROR'
    }

    if (-not $filtered -or $filtered.Count -eq 0) {
        # Por si algún día no se imprimen las líneas clave:
        $filtered = @("Sin líneas clave. Últimas líneas:",
                      ($rOut | Select-Object -Last 10))
    }

    # Escribe el log minimalista (solo esas líneas)
    $filtered | Out-File -FilePath $logMain -Encoding UTF8

    # Revisa el código de salida de R
    if ($LASTEXITCODE -ne 0) {
        throw "Rscript terminó con código $LASTEXITCODE"
    }

    # (Opcional) Validar que el Excel existe
    $lastXlsx = Get-ChildItem -Path $logDir -Filter 'reporte_*.xlsx' |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $lastXlsx) {
        "ERROR: No se encontró el Excel generado." | Out-File -FilePath $logMain -Encoding UTF8 -Append
        exit 1
    }

    # Rotación: conserva solo los últimos 50 logs
    Get-ChildItem -Path $logDir -Filter "export_prueba_*.log" |
      Sort-Object LastWriteTime -Descending | Select-Object -Skip 50 |
      Remove-Item -Force -ErrorAction SilentlyContinue
}
catch {
    # Log minimal de error
    "ERROR: $($_.Exception.Message)" | Out-File -FilePath $logMain -Encoding UTF8 -Append
    exit 1
}
