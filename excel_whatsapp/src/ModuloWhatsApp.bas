Attribute VB_Name = "ModuloWhatsApp"
'==============================================================================
' Envío de mensajes de WhatsApp desde Excel
'
' Hojas:
'   Base      -> la BBDD (se carga con el botón "Importar base")
'   Config    -> tiempos, mensaje a usar, límites y canal
'   Mensajes  -> los textos (se pueden cambiar cuando quieras)
'   Bajas     -> números que NUNCA deben recibir mensajes
'
' Cada mensaje se abre en WhatsApp ya escrito y la persona lo envía con Enter.
' Para detener el envío en cualquier momento: tecla Esc.
'==============================================================================
Option Explicit

Private Const HOJA_BASE As String = "Base"
Private Const HOJA_MENSAJES As String = "Mensajes"
Private Const HOJA_BAJAS As String = "Bajas"

Private Const COL_ESTADO As String = "ESTADO_ENVIO"
Private Const COL_FECHA As String = "FECHA_ENVIO"
Private Const COL_MENSAJE As String = "MENSAJE_N"

'------------------------------------------------------------------------------
' Botón: Importar base
'------------------------------------------------------------------------------
Public Sub ImportarBase()
    Dim archivo As Variant, origen As Workbook, ws As Worksheet

    archivo = Application.GetOpenFilename( _
        "Excel o CSV (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", , _
        "Elige la BBDD a importar")
    If VarType(archivo) = vbBoolean Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(HOJA_BASE)
    If Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
        If MsgBox("La hoja Base ya tiene datos y se reemplazarán " & _
                  "(se pierden los estados de envío)." & vbLf & vbLf & "¿Continuar?", _
                  vbExclamation + vbYesNo, "Importar base") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Importando base..."
    Set origen = Workbooks.Open(CStr(archivo), ReadOnly:=True)
    ws.Cells.Clear
    origen.Worksheets(1).UsedRange.Copy Destination:=ws.Range("A1")
    origen.Close SaveChanges:=False

    ws.Columns.AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True
    PrepararColumnas ws
    ThisWorkbook.Save

    MsgBox "Base importada: " & Format(UltimaFila(ws) - 1, "#,##0") & " filas.", _
           vbInformation, "Importar base"
End Sub

'------------------------------------------------------------------------------
' Botón: Enviar mensajes
'------------------------------------------------------------------------------
Public Sub EnviarMensajes()
    Dim ws As Worksheet, datos As Variant, enc As Variant
    Dim colTel As Long, colVend As Long, colEst As Long, colFec As Long, colMsj As Long
    Dim espera As Long, numMensaje As Long, maximo As Long
    Dim vendedor As String, canal As String, confirmar As Boolean
    Dim plantilla As String, texto As String, tel As String
    Dim bajas As Object, yaEnviados As Object
    Dim fila As Long, ultima As Long, abiertos As Long, resp As VbMsgBoxResult
    Dim estado As String

    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Manejo

    Set ws = ThisWorkbook.Worksheets(HOJA_BASE)
    ultima = UltimaFila(ws)
    If ultima < 2 Then
        MsgBox "La hoja Base está vacía. Usa primero el botón Importar base.", vbExclamation
        Exit Sub
    End If
    PrepararColumnas ws

    ' --- Configuración ---
    espera = CLng(Val(Cfg("cfgEspera")))
    numMensaje = CLng(Val(Cfg("cfgMensaje")))
    maximo = CLng(Val(Cfg("cfgMaximo")))
    vendedor = UCase$(Trim$(CStr(Cfg("cfgVendedor"))))
    canal = UCase$(Trim$(CStr(Cfg("cfgCanal"))))
    confirmar = (UCase$(Left$(Trim$(CStr(Cfg("cfgConfirmar"))), 1)) = "S")
    If espera < 1 Then espera = 1
    If maximo < 1 Then maximo = 1

    plantilla = ObtenerPlantilla(numMensaje)
    If Len(plantilla) = 0 Then
        MsgBox "No existe el mensaje N° " & numMensaje & " en la hoja Mensajes.", vbExclamation
        Exit Sub
    End If

    ' --- Columnas ---
    enc = ws.Range(ws.Cells(1, 1), ws.Cells(1, UltimaColumna(ws))).Value
    colTel = ColumnaTelefono(enc)
    colVend = BuscarColumna(enc, "VENDEDOR")
    colEst = BuscarColumna(enc, COL_ESTADO)
    colFec = BuscarColumna(enc, COL_FECHA)
    colMsj = BuscarColumna(enc, COL_MENSAJE)
    If colTel = 0 Then
        MsgBox "No encontré la columna del número en la hoja Base.", vbCritical
        Exit Sub
    End If

    datos = ws.Range(ws.Cells(1, 1), ws.Cells(ultima, UBound(enc, 2))).Value
    Set bajas = CargarBajas()
    Set yaEnviados = CreateObject("Scripting.Dictionary")
    For fila = 2 To ultima
        estado = CStr(datos(fila, colEst))
        If estado = "ENVIADO" Or estado = "ABIERTO" Then
            tel = NormalizarTelefono(datos(fila, colTel))
            If Len(tel) > 0 Then yaEnviados(tel) = True
        End If
    Next fila

    If MsgBox("Se abrirán hasta " & maximo & " chats con el mensaje N° " & numMensaje & _
              ", esperando " & espera & " segundos entre cada uno." & vbLf & vbLf & _
              "En cada chat presiona Enter en WhatsApp para enviar." & vbLf & _
              "Para detener: tecla Esc." & vbLf & vbLf & "¿Comenzar?", _
              vbQuestion + vbYesNo, "Enviar mensajes") = vbNo Then Exit Sub

    ' --- Envío ---
    Application.Calculation = xlCalculationManual   ' más rápido con bases grandes
    For fila = 2 To ultima
        If abiertos >= maximo Then Exit For
        If Len(Trim$(CStr(datos(fila, colEst)))) = 0 Then
            If Len(vendedor) = 0 Or colVend = 0 Then
                estado = "?"
            ElseIf UCase$(Trim$(CStr(datos(fila, colVend)))) = vendedor Then
                estado = "?"
            Else
                estado = ""
            End If

            If estado = "?" Then
                tel = NormalizarTelefono(datos(fila, colTel))
                If Len(tel) = 0 Then
                    MarcarFila ws, fila, colEst, colFec, colMsj, "INVALIDO", 0
                ElseIf bajas.Exists(tel) Then
                    MarcarFila ws, fila, colEst, colFec, colMsj, "BAJA", 0
                ElseIf yaEnviados.Exists(tel) Then
                    MarcarFila ws, fila, colEst, colFec, colMsj, "DUPLICADO", 0
                Else
                    texto = ArmarMensaje(plantilla, enc, datos, fila)
                    Application.Goto ws.Cells(fila, 1), True
                    Application.StatusBar = "Chat " & (abiertos + 1) & " de " & maximo & ": +" & tel
                    AbrirChat tel, texto, canal

                    If confirmar Then
                        resp = MsgBox("Chat abierto con +" & tel & "." & vbLf & vbLf & _
                            "1) Ve a WhatsApp y presiona Enter para enviar." & vbLf & _
                            "2) Vuelve aquí y responde:" & vbLf & vbLf & _
                            "Sí = se envió    No = omitir    Cancelar = detener", _
                            vbQuestion + vbYesNoCancel + vbDefaultButton2, "¿Se envió?")
                        If resp = vbCancel Then Exit For
                        If resp = vbYes Then
                            MarcarFila ws, fila, colEst, colFec, colMsj, "ENVIADO", numMensaje
                            yaEnviados(tel) = True
                        Else
                            MarcarFila ws, fila, colEst, colFec, colMsj, "OMITIDO", numMensaje
                        End If
                    Else
                        MarcarFila ws, fila, colEst, colFec, colMsj, "ABIERTO", numMensaje
                        yaEnviados(tel) = True
                    End If

                    abiertos = abiertos + 1
                    If abiertos Mod 10 = 0 Then ThisWorkbook.Save
                    If abiertos < maximo Then Esperar espera, abiertos, maximo
                End If
            End If
        End If
    Next fila

Fin:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.Save
    MsgBox "Proceso terminado. Chats abiertos en esta sesión: " & abiertos, vbInformation
    Exit Sub

Manejo:
    Application.Calculation = xlCalculationAutomatic
    If Err.Number = 18 Then
        Application.StatusBar = False
        ThisWorkbook.Save
        MsgBox "Envío detenido. Chats abiertos en esta sesión: " & abiertos, vbInformation
    Else
        Application.StatusBar = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    End If
End Sub

'------------------------------------------------------------------------------
' Botón: Probar mensaje (muestra cómo queda con la primera fila pendiente)
'------------------------------------------------------------------------------
Public Sub ProbarMensaje()
    Dim ws As Worksheet, enc As Variant, datos As Variant, plantilla As String
    Dim fila As Long, colEst As Long, numMensaje As Long

    Set ws = ThisWorkbook.Worksheets(HOJA_BASE)
    If UltimaFila(ws) < 2 Then
        MsgBox "La hoja Base está vacía. Usa primero el botón Importar base.", vbExclamation
        Exit Sub
    End If
    PrepararColumnas ws
    numMensaje = CLng(Val(Cfg("cfgMensaje")))
    plantilla = ObtenerPlantilla(numMensaje)
    If Len(plantilla) = 0 Then
        MsgBox "No existe el mensaje N° " & numMensaje & " en la hoja Mensajes.", vbExclamation
        Exit Sub
    End If

    enc = ws.Range(ws.Cells(1, 1), ws.Cells(1, UltimaColumna(ws))).Value
    colEst = BuscarColumna(enc, COL_ESTADO)
    datos = ws.Range(ws.Cells(1, 1), ws.Cells(UltimaFila(ws), UBound(enc, 2))).Value
    For fila = 2 To UBound(datos, 1)
        If Len(Trim$(CStr(datos(fila, colEst)))) = 0 Then Exit For
    Next fila
    If fila > UBound(datos, 1) Then fila = 2

    MsgBox ArmarMensaje(plantilla, enc, datos, fila), vbInformation, _
           "Vista previa del mensaje N° " & numMensaje & " (fila " & fila & ")"
End Sub

'------------------------------------------------------------------------------
' Botón: Marcar baja (fila seleccionada en la hoja Base)
'------------------------------------------------------------------------------
Public Sub MarcarBaja()
    Dim ws As Worksheet, enc As Variant, colTel As Long, tel As String, fila As Long

    Set ws = ThisWorkbook.Worksheets(HOJA_BASE)
    ' Toma la fila que quedó seleccionada en la hoja Base.
    If Not ActiveSheet Is ws Then ws.Activate
    fila = ActiveCell.Row
    If fila < 2 Then
        MsgBox "En la hoja Base, selecciona una celda de la fila del cliente y presiona otra vez " & _
               "Marcar baja." & vbLf & "(También puedes escribir el número en la hoja Bajas.)", _
               vbInformation
        Exit Sub
    End If

    PrepararColumnas ws
    enc = ws.Range(ws.Cells(1, 1), ws.Cells(1, UltimaColumna(ws))).Value
    colTel = ColumnaTelefono(enc)
    tel = NormalizarTelefono(ws.Cells(fila, colTel).Value)
    If Len(tel) = 0 Then
        MsgBox "El número de la fila " & fila & " no es válido.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Fila " & fila & ": ¿marcar +" & tel & " como BAJA?" & vbLf & _
              "No volverá a recibir mensajes.", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub

    AgregarBaja tel
    ws.Cells(fila, BuscarColumna(enc, COL_ESTADO)).Value = "BAJA"
    ws.Cells(fila, BuscarColumna(enc, COL_FECHA)).Value = Now
    ThisWorkbook.Save
End Sub

'------------------------------------------------------------------------------
' Botón: Reiniciar estados (para una nueva campaña)
'------------------------------------------------------------------------------
Public Sub ReiniciarEstados()
    Dim ws As Worksheet, enc As Variant, ultima As Long, c As Variant

    If MsgBox("Se borrarán los estados ENVIADO, ABIERTO y OMITIDO para hacer una nueva campaña." & _
              vbLf & "Las BAJAS se mantienen siempre." & vbLf & vbLf & "¿Continuar?", _
              vbExclamation + vbYesNo, "Reiniciar estados") = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(HOJA_BASE)
    ultima = UltimaFila(ws)
    If ultima < 2 Then Exit Sub
    PrepararColumnas ws
    enc = ws.Range(ws.Cells(1, 1), ws.Cells(1, UltimaColumna(ws))).Value

    Application.ScreenUpdating = False
    Dim colEst As Long, colFec As Long, colMsj As Long, fila As Long, est As Variant, fec As Variant, msj As Variant
    colEst = BuscarColumna(enc, COL_ESTADO)
    colFec = BuscarColumna(enc, COL_FECHA)
    colMsj = BuscarColumna(enc, COL_MENSAJE)
    est = ws.Range(ws.Cells(1, colEst), ws.Cells(ultima, colEst)).Value
    fec = ws.Range(ws.Cells(1, colFec), ws.Cells(ultima, colFec)).Value
    msj = ws.Range(ws.Cells(1, colMsj), ws.Cells(ultima, colMsj)).Value
    For fila = 2 To ultima
        c = est(fila, 1)
        If c = "ENVIADO" Or c = "ABIERTO" Or c = "OMITIDO" Or c = "DUPLICADO" Then
            est(fila, 1) = Empty: fec(fila, 1) = Empty: msj(fila, 1) = Empty
        End If
    Next fila
    ws.Range(ws.Cells(1, colEst), ws.Cells(ultima, colEst)).Value = est
    ws.Range(ws.Cells(1, colFec), ws.Cells(ultima, colFec)).Value = fec
    ws.Range(ws.Cells(1, colMsj), ws.Cells(ultima, colMsj)).Value = msj
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox "Estados reiniciados.", vbInformation
End Sub

'==============================================================================
' Funciones de apoyo
'==============================================================================
Private Sub AbrirChat(ByVal tel As String, ByVal texto As String, ByVal canal As String)
    Dim url As String, enc As String
    enc = Application.WorksheetFunction.EncodeURL(texto)
    If canal = "WEB" Then
        url = "https://web.whatsapp.com/send?phone=" & tel & "&text=" & enc
    Else
        url = "whatsapp://send?phone=" & tel & "&text=" & enc
    End If
    CreateObject("Shell.Application").ShellExecute url
End Sub

Private Sub Esperar(ByVal segundos As Long, ByVal hechos As Long, ByVal maximo As Long)
    Dim i As Long
    For i = segundos To 1 Step -1
        Application.StatusBar = "Enviados/abiertos: " & hechos & " de " & maximo & _
                                "  |  Siguiente chat en " & i & " s  (Esc para detener)"
        Application.Wait Now + TimeSerial(0, 0, 1)
        DoEvents
    Next i
End Sub

Private Sub MarcarFila(ws As Worksheet, ByVal fila As Long, ByVal colEst As Long, _
                       ByVal colFec As Long, ByVal colMsj As Long, ByVal estado As String, _
                       ByVal numMensaje As Long)
    ws.Cells(fila, colEst).Value = estado
    ws.Cells(fila, colFec).Value = Now
    If numMensaje > 0 Then ws.Cells(fila, colMsj).Value = numMensaje
End Sub

Private Function ArmarMensaje(ByVal plantilla As String, enc As Variant, datos As Variant, _
                              ByVal fila As Long) As String
    ' Reemplaza {ENCABEZADO} por el valor de esa columna. Ej: {MONTO_RECARGA}
    Dim c As Long, texto As String, valor As String
    texto = plantilla
    For c = 1 To UBound(enc, 2)
        If Len(CStr(enc(1, c))) > 0 Then
            valor = CStr(datos(fila, c))
            If IsNumeric(datos(fila, c)) And Len(valor) < 9 And Len(valor) > 3 Then
                valor = Format(CDbl(datos(fila, c)), "#,##0")
            End If
            texto = Replace(texto, "{" & UCase$(Trim$(CStr(enc(1, c)))) & "}", valor, , , vbTextCompare)
        End If
    Next c
    ArmarMensaje = texto
End Function

Private Function ObtenerPlantilla(ByVal numMensaje As Long) As String
    Dim ws As Worksheet, fila As Long
    Set ws = ThisWorkbook.Worksheets(HOJA_MENSAJES)
    For fila = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Val(ws.Cells(fila, 1).Value) = numMensaje Then
            ObtenerPlantilla = Trim$(CStr(ws.Cells(fila, 3).Value))
            Exit Function
        End If
    Next fila
End Function

Private Function CargarBajas() As Object
    Dim ws As Worksheet, fila As Long, tel As String, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(HOJA_BAJAS)
    For fila = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        tel = NormalizarTelefono(ws.Cells(fila, 1).Value)
        If Len(tel) > 0 Then d(tel) = True
    Next fila
    Set CargarBajas = d
End Function

Private Sub AgregarBaja(ByVal tel As String)
    Dim ws As Worksheet, fila As Long
    If CargarBajas().Exists(tel) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(HOJA_BAJAS)
    fila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(fila, 1).NumberFormat = "@"
    ws.Cells(fila, 1).Value = tel
    ws.Cells(fila, 2).Value = Now
    ws.Cells(fila, 3).Value = "Pidió no recibir mensajes"
End Sub

Public Function NormalizarTelefono(ByVal v As Variant) As String
    ' Devuelve el móvil chileno como 569XXXXXXXX, o "" si no es válido.
    Dim s As String, d As String, i As Long, ch As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then s = Format(CDbl(v), "0") Else s = CStr(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then d = d & ch
    Next i
    If Len(d) = 8 Then d = "569" & d
    If Len(d) = 9 And Left$(d, 1) = "9" Then d = "56" & d
    If Len(d) = 11 And Left$(d, 3) = "569" Then NormalizarTelefono = d
End Function

Private Sub PrepararColumnas(ws As Worksheet)
    ' Agrega al final las columnas de control si todavía no existen.
    Dim enc As Variant, nombres As Variant, i As Long, ultimaCol As Long
    nombres = Array(COL_ESTADO, COL_FECHA, COL_MENSAJE)
    For i = 0 To 2
        ultimaCol = UltimaColumna(ws)
        enc = ws.Range(ws.Cells(1, 1), ws.Cells(1, ultimaCol + 1)).Value
        If BuscarColumna(enc, CStr(nombres(i))) = 0 Then
            ws.Cells(1, ultimaCol + 1).Value = nombres(i)
            ws.Cells(1, ultimaCol + 1).Font.Bold = True
            If nombres(i) = COL_FECHA Then
                ws.Columns(ultimaCol + 1).NumberFormat = "dd-mm-yyyy hh:mm"
            End If
        End If
    Next i
End Sub

Private Function BuscarColumna(enc As Variant, ByVal nombre As String) As Long
    Dim c As Long
    For c = 1 To UBound(enc, 2)
        If UCase$(Trim$(CStr(enc(1, c)))) = UCase$(nombre) Then
            BuscarColumna = c
            Exit Function
        End If
    Next c
End Function

Private Function ColumnaTelefono(enc As Variant) As Long
    Dim c As Long, h As String, claves As Variant, k As Variant
    claves = Array("CELULAR", "TELEFONO", "TELÉFONO", "FONO", "MOVIL", "MÓVIL", "MSISDN", "NUMERO", "NÚMERO", "COMPA")
    For Each k In claves
        For c = 1 To UBound(enc, 2)
            h = UCase$(CStr(enc(1, c)))
            If InStr(1, h, CStr(k), vbTextCompare) > 0 Then
                ColumnaTelefono = c
                Exit Function
            End If
        Next c
    Next k
End Function

Private Function UltimaFila(ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find("*", , xlValues, , xlByRows, xlPrevious)
    If r Is Nothing Then UltimaFila = 1 Else UltimaFila = r.Row
End Function

Private Function UltimaColumna(ws As Worksheet) As Long
    Dim c As Long
    c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If c = 1 And Len(CStr(ws.Cells(1, 1).Value)) = 0 Then c = 0
    UltimaColumna = c
End Function

Private Function Cfg(ByVal nombre As String) As Variant
    ' Lee un valor de la hoja Config usando su nombre definido (cfgEspera, cfgMensaje, ...).
    Cfg = ThisWorkbook.Names(nombre).RefersToRange.Value
End Function
