# Crea EnvioWhatsApp.xlsm con las hojas, botones y la macro de envío.
# Lo ejecuta INSTALAR.bat; no hace falta abrirlo directamente.

$ErrorActionPreference = 'Stop'
$carpeta = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$modulo  = Join-Path $PSScriptRoot 'ModuloWhatsApp.bas'
$destino = Join-Path $carpeta 'EnvioWhatsApp.xlsm'

if (Test-Path $destino) {
    $respaldo = Join-Path $carpeta ("EnvioWhatsApp_respaldo_{0:yyyyMMdd_HHmmss}.xlsm" -f (Get-Date))
    Rename-Item $destino $respaldo
    Write-Host "Ya existía un EnvioWhatsApp.xlsm: se guardó como respaldo en $respaldo"
}

Write-Host 'Abriendo Excel...'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$moduloImportado = $false

try {
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 4) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $config   = $wb.Worksheets.Item(1); $config.Name   = 'Config'
    $mensajes = $wb.Worksheets.Item(2); $mensajes.Name = 'Mensajes'
    $base     = $wb.Worksheets.Item(3); $base.Name     = 'Base'
    $bajas    = $wb.Worksheets.Item(4); $bajas.Name    = 'Bajas'

    # ---------------- Config ----------------
    $config.Range('A1').Value2 = 'ENVÍO DE MENSAJES POR WHATSAPP'
    $config.Range('A1').Font.Size = 16; $config.Range('A1').Font.Bold = $true
    $config.Range('A3').Value2 = 'Ajuste'; $config.Range('B3').Value2 = 'Valor'; $config.Range('C3').Value2 = 'Qué significa'
    $config.Range('A3:C3').Font.Bold = $true
    $filas = @(
        @('Segundos entre mensajes', 20, 'Tiempo de espera antes de abrir el siguiente chat.', 'cfgEspera'),
        @('Mensaje a usar (N°)', 1, 'Número del mensaje de la hoja Mensajes.', 'cfgMensaje'),
        @('Máximo de chats por sesión', 50, 'Cuántos chats abre cada vez que presionas Enviar.', 'cfgMaximo'),
        @('Vendedor', '', 'Vacío = todos. O escribe el nombre exacto, ej: VENDEDOR 1', 'cfgVendedor'),
        @('Canal', 'ESCRITORIO', 'ESCRITORIO = app WhatsApp de Windows (recomendado). WEB = navegador.', 'cfgCanal'),
        @('Confirmar cada envío', 'SI', 'SI = después de cada chat preguntas si se envió. NO = solo espera los segundos.', 'cfgConfirmar')
    )
    $r = 4
    foreach ($f in $filas) {
        $config.Cells.Item($r, 1).Value2 = $f[0]
        $config.Cells.Item($r, 2).Value2 = $f[1]
        $config.Cells.Item($r, 3).Value2 = $f[2]
        [void]$wb.Names.Add($f[3], "=Config!`$B`$$r")
        $config.Cells.Item($r, 2).Interior.Color = 0xCCF2FF   # amarillo claro = editable
        $r++
    }

    $config.Range('A11').Value2 = 'Resumen'; $config.Range('A11').Font.Bold = $true
    $resumen = @('ENVIADO', 'ABIERTO', 'OMITIDO', 'BAJA', 'DUPLICADO', 'INVALIDO')
    $r = 12
    foreach ($e in $resumen) {
        $config.Cells.Item($r, 1).Value2 = $e
        $config.Cells.Item($r, 2).Formula = "=COUNTIF(Base!`$A:`$Z,""$e"")"
        $r++
    }
    $config.Cells.Item($r, 1).Value2 = 'PENDIENTES'
    $config.Cells.Item($r, 2).Formula = "=MAX(0,COUNTA(Base!`$A:`$A)-1-SUM(B12:B17))"
    $config.Range("A$($r):B$($r)").Font.Bold = $true

    $config.Columns.Item(1).ColumnWidth = 30
    $config.Columns.Item(2).ColumnWidth = 16
    $config.Columns.Item(3).ColumnWidth = 75
    $config.Range('B4:B20').HorizontalAlignment = -4108   # centrado

    # Botones
    $botones = @(
        @('1. Importar base', 'ImportarBase'),
        @('2. Probar mensaje', 'ProbarMensaje'),
        @('3. ENVIAR MENSAJES', 'EnviarMensajes'),
        @('Marcar baja (fila de Base)', 'MarcarBaja'),
        @('Reiniciar estados', 'ReiniciarEstados')
    )
    $izq = $config.Range('E3').Left + 10
    $arriba = $config.Range('E3').Top
    foreach ($b in $botones) {
        try {
            $boton = $config.Buttons().Add($izq, $arriba, 190, 30)
            $boton.Caption = $b[0]
            $boton.OnAction = $b[1]
        } catch {
            $forma = $config.Shapes.AddFormControl(0, $izq, $arriba, 190, 30)   # 0 = botón
            $forma.OnAction = $b[1]
            $forma.TextFrame.Characters().Text = $b[0]
        }
        $arriba += 38
    }

    # ---------------- Mensajes ----------------
    $mensajes.Range('A1').Value2 = 'N°'; $mensajes.Range('B1').Value2 = 'Nombre'; $mensajes.Range('C1').Value2 = 'Texto del mensaje'
    $mensajes.Range('A1:C1').Font.Bold = $true
    $mensajes.Range('A2').Value2 = 1
    $mensajes.Range('B2').Value2 = 'Oferta plan'
    $mensajes.Range('C2').Value2 = "Hola! Te escribimos de [TU COMPAÑÍA]. Vimos que recargas `${MONTO_RECARGA} en tu prepago.`nPor un valor parecido puedes tener un plan con más gigas y minutos ilimitados.`n¿Te cuento los detalles?`n`nSi no quieres recibir más mensajes, responde BAJA."
    $mensajes.Range('A3').Value2 = 2
    $mensajes.Range('B3').Value2 = 'Portabilidad'
    $mensajes.Range('C3').Value2 = "Hola! Tenemos una oferta especial para clientes prepago: cámbiate a plan y conserva tu mismo número.`n¿Te interesa que te llame un ejecutivo?`n`nSi no quieres recibir más mensajes, responde BAJA."
    $mensajes.Range('E1').Value2 = 'Cómo escribir mensajes'
    $mensajes.Range('E1').Font.Bold = $true
    $mensajes.Range('E2').Value2 = 'Agrega o cambia filas con un N° distinto. Para un salto de línea dentro de la celda: Alt + Enter.'
    $mensajes.Range('E3').Value2 = 'Datos del cliente: escribe el nombre de una columna de la Base entre llaves, ej: {MONTO_RECARGA}, {DESC_PRODUCTO}, {SUBSEGMENTO}, {VENDEDOR}.'
    $mensajes.Range('E4').Value2 = 'Termina siempre con la opción de darse de baja.'
    $mensajes.Columns.Item(1).ColumnWidth = 6
    $mensajes.Columns.Item(2).ColumnWidth = 18
    $mensajes.Columns.Item(3).ColumnWidth = 80
    $mensajes.Columns.Item(3).WrapText = $true
    $mensajes.Range('A:B').VerticalAlignment = -4160  # arriba

    # ---------------- Bajas ----------------
    $bajas.Range('A1').Value2 = 'NUMERO'; $bajas.Range('B1').Value2 = 'FECHA'; $bajas.Range('C1').Value2 = 'MOTIVO'
    $bajas.Range('A1:C1').Font.Bold = $true
    $bajas.Columns.Item(1).NumberFormat = '@'
    $bajas.Columns.Item(1).ColumnWidth = 16
    $bajas.Columns.Item(2).ColumnWidth = 18
    $bajas.Columns.Item(3).ColumnWidth = 40

    # ---------------- Macro ----------------
    try {
        [void]$wb.VBProject.VBComponents.Import($modulo)
        $moduloImportado = $true
    } catch {
        $moduloImportado = $false
    }

    $config.Activate()
    $wb.SaveAs($destino, 52)   # 52 = libro con macros (.xlsm)
    $wb.Close($false)
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

Write-Host ''
Write-Host "Archivo creado: $destino"
if (-not $moduloImportado) {
    Write-Host ''
    Write-Host 'ATENCIÓN: Excel no permitió agregar la macro automáticamente.' -ForegroundColor Yellow
    Write-Host 'Opción A (recomendada): en Excel ve a Archivo > Opciones > Centro de confianza >'
    Write-Host '  Configuración del Centro de confianza > Configuración de macros, marca'
    Write-Host '  "Confiar en el acceso al modelo de objetos de proyectos de VBA" y vuelve a ejecutar INSTALAR.bat.'
    Write-Host 'Opción B: abre EnvioWhatsApp.xlsm, presiona Alt+F11, menú Archivo > Importar archivo,'
    Write-Host "  elige $modulo y guarda."
    exit 2
}
exit 0
