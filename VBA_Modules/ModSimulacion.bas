Attribute VB_Name = "ModSimulacion"
'==============================================================================
' MÓDULO: ModSimulacion
' DESCRIPCIÓN: Simulación y análisis de aspirantes vs. matriculados
'              Aspirantes = total de estudiantes que aplicaron
'              Matriculados = confirmados (Estado = "Activo")
' PROYECTO: Automatizador_Matricula.xlsm
' VERSIÓN: 1.0
' FECHA: Febrero 2026
'==============================================================================
Option Explicit

' ============================================================
' MACRO PÚBLICA PARA EL PANEL
' ============================================================

Public Sub Btn_GenerarSimulacion()
    ActualizarEstadoPanelSim "Generando simulación..."
    GenerarSimulacionAspirantes
    ActualizarEstadoPanelSim "Listo"
End Sub

' ============================================================
' PROCEDIMIENTO PRINCIPAL DE SIMULACIÓN
' ============================================================

Public Sub GenerarSimulacionAspirantes()
    '
    ' Lee los datos importados de la hoja Datos_SIGU y genera un informe
    ' de simulación en la hoja Simulacion_Aspirantes, mostrando:
    '   - Total de aspirantes (todos los registros)
    '   - Matriculados confirmados (Estado = "Activo")
    '   - Pendientes, Inactivos y Retirados
    '   - Tasas de confirmación por carrera y por sede
    '
    Dim wsSIGU As Worksheet
    Dim wsSim  As Worksheet
    Dim ultimaFila As Long
    
    On Error GoTo ErrorHandler
    
    ' ---- Verificar que existan datos en la hoja SIGU ----
    On Error Resume Next
    Set wsSIGU = ThisWorkbook.Sheets(HOJA_SIGU)
    On Error GoTo ErrorHandler
    
    If wsSIGU Is Nothing Then
        MsgBox "La hoja '" & HOJA_SIGU & "' no existe." & vbCrLf & _
               "Inicialice el sistema e importe datos SIGU primero.", _
               vbExclamation, "Simulación"
        Exit Sub
    End If
    
    ultimaFila = ObtenerUltimaFila(wsSIGU)
    
    If ultimaFila < 2 Then
        MsgBox "No hay datos en la hoja '" & HOJA_SIGU & "'." & vbCrLf & _
               "Importe los datos SIGU antes de ejecutar la simulación.", _
               vbExclamation, "Simulación"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' ---- Preparar hoja de simulación ----
    Set wsSim = PrepararHojaSimulacion()
    
    ' ---- Recopilar estadísticas desde los datos SIGU ----
    Dim dicCarreras  As Object
    Dim dicSedes     As Object
    Dim totalAsp     As Long
    Dim totalMat     As Long
    Dim totalPen     As Long
    Dim totalIna     As Long
    
    Set dicCarreras = CreateObject("Scripting.Dictionary")
    Set dicSedes    = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To ultimaFila
        Dim cedula  As String
        Dim carrera As String
        Dim sede    As String
        Dim estado  As String
        
        cedula  = Trim(CStr(wsSIGU.Cells(i, COL_SIGU_CEDULA).Value))
        carrera = Trim(CStr(wsSIGU.Cells(i, COL_SIGU_CARRERA).Value))
        sede    = Trim(CStr(wsSIGU.Cells(i, COL_SIGU_SEDE).Value))
        estado  = Trim(CStr(wsSIGU.Cells(i, COL_SIGU_ESTADO).Value))
        
        If cedula = "" Then GoTo SiguienteRegistro
        
        totalAsp = totalAsp + 1
        
        ' --- Acumuladores por carrera ---
        If Not dicCarreras.Exists(carrera) Then
            dicCarreras(carrera) = Array(0, 0, 0, 0)  ' Asp, Mat, Pen, Ina
        End If
        Dim cArr As Variant
        cArr = dicCarreras(carrera)
        cArr(0) = cArr(0) + 1
        
        ' --- Acumuladores por sede ---
        If Not dicSedes.Exists(sede) Then
            dicSedes(sede) = Array(0, 0, 0, 0)
        End If
        Dim sArr As Variant
        sArr = dicSedes(sede)
        sArr(0) = sArr(0) + 1
        
        ' --- Clasificar por estado ---
        Select Case UCase(estado)
            Case UCase(ESTADO_ACTIVO)
                totalMat = totalMat + 1
                cArr(1) = cArr(1) + 1
                sArr(1) = sArr(1) + 1
            Case UCase(ESTADO_PENDIENTE)
                totalPen = totalPen + 1
                cArr(2) = cArr(2) + 1
                sArr(2) = sArr(2) + 1
            Case Else
                totalIna = totalIna + 1
                cArr(3) = cArr(3) + 1
                sArr(3) = sArr(3) + 1
        End Select
        
        dicCarreras(carrera) = cArr
        dicSedes(sede)       = sArr
        
SiguienteRegistro:
    Next i
    
    ' ---- Escribir resumen general ----
    EscribirResumenGeneral wsSim, totalAsp, totalMat, totalPen, totalIna
    
    ' ---- Escribir tabla por carrera ----
    Dim filaInicioCarreras As Long
    filaInicioCarreras = 15
    EscribirTablaCarreras wsSim, dicCarreras, filaInicioCarreras
    
    ' ---- Escribir tabla por sede ----
    Dim filaInicioSedes As Long
    filaInicioSedes = filaInicioCarreras + dicCarreras.Count + 4
    EscribirTablaSedes wsSim, dicSedes, filaInicioSedes
    
    ' ---- Registrar en historial ----
    RegistrarHistorial "SIMULACION", "", _
        "Simulación generada: " & totalAsp & " aspirantes, " & totalMat & " matriculados", "ÉXITO"
    
    ' ---- Mostrar resultado ----
    wsSim.Activate
    Application.ScreenUpdating = True
    
    MsgBox "SIMULACIÓN COMPLETADA" & vbCrLf & vbCrLf & _
           "Total aspirantes  : " & totalAsp & vbCrLf & _
           "Matriculados      : " & totalMat & " (" & PorcentajeStr(totalMat, totalAsp) & ")" & vbCrLf & _
           "Pendientes        : " & totalPen & " (" & PorcentajeStr(totalPen, totalAsp) & ")" & vbCrLf & _
           "Inact./Retirados  : " & totalIna & " (" & PorcentajeStr(totalIna, totalAsp) & ")" & vbCrLf & vbCrLf & _
           "Revise la hoja '" & HOJA_SIMULACION & "'.", _
           vbInformation, "Simulación Aspirantes vs. Matriculados"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error al generar la simulación: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================================
' PREPARAR / LIMPIAR HOJA DE SIMULACIÓN
' ============================================================

Private Function PrepararHojaSimulacion() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HOJA_SIMULACION)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = HOJA_SIMULACION
    Else
        ws.Cells.Clear
    End If
    
    ' Título principal
    ws.Range("A1:H1").Merge
    ws.Range("A1").Value = "SIMULACIÓN: ASPIRANTES VS. MATRICULADOS - PERÍODO " & ObtenerPeriodoActual()
    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Range("A2").Value = "Generado el: " & Format(Now, "dd/mm/yyyy hh:mm")
    ws.Range("A2").Font.Italic = True
    
    ws.Range("A3").Value = "Aspirantes = total de estudiantes que aplicaron  |  " & _
                           "Matriculados = confirmados (Estado Activo)"
    ws.Range("A3").Font.Color = RGB(128, 0, 0)
    
    Set PrepararHojaSimulacion = ws
End Function

' ============================================================
' RESUMEN GENERAL
' ============================================================

Private Sub EscribirResumenGeneral(ws As Worksheet, _
                                   totalAsp As Long, totalMat As Long, _
                                   totalPen As Long, totalIna As Long)
    Dim fila As Long
    fila = 5
    
    ws.Range("A" & fila).Value = "RESUMEN GENERAL"
    ws.Range("A" & fila).Font.Bold = True
    ws.Range("A" & fila).Font.Size = 12
    ws.Range("A" & fila).Font.Color = RGB(0, 51, 102)
    fila = fila + 1
    
    ' Encabezados
    Dim headers As Variant
    headers = Array("Categoría", "Cantidad", "Porcentaje", "Observación")
    Dim h As Long
    For h = 0 To UBound(headers)
        ws.Cells(fila, h + 1).Value = headers(h)
        ws.Cells(fila, h + 1).Font.Bold = True
        ws.Cells(fila, h + 1).Interior.Color = RGB(0, 102, 153)
        ws.Cells(fila, h + 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(fila, h + 1).Borders.LineStyle = xlContinuous
    Next h
    fila = fila + 1
    
    ' Filas de datos
    Dim datos(3, 3) As Variant
    datos(0, 0) = "Total Aspirantes"
    datos(0, 1) = totalAsp
    datos(0, 2) = "100%"
    datos(0, 3) = "Todos los que aplicaron al período"
    
    datos(1, 0) = "Matriculados (confirmados)"
    datos(1, 1) = totalMat
    datos(1, 2) = PorcentajeStr(totalMat, totalAsp)
    datos(1, 3) = "Estado: Activo"
    
    datos(2, 0) = "Pendientes"
    datos(2, 1) = totalPen
    datos(2, 2) = PorcentajeStr(totalPen, totalAsp)
    datos(2, 3) = "Estado: Pendiente"
    
    datos(3, 0) = "Inactivos / Retirados"
    datos(3, 1) = totalIna
    datos(3, 2) = PorcentajeStr(totalIna, totalAsp)
    datos(3, 3) = "Estado: Inactivo o Retirado"
    
    Dim colores(3) As Long
    colores(0) = RGB(204, 229, 255)
    colores(1) = RGB(198, 239, 206)
    colores(2) = RGB(255, 235, 156)
    colores(3) = RGB(255, 199, 206)
    
    Dim r As Long
    For r = 0 To 3
        Dim c As Long
        For c = 0 To 3
            ws.Cells(fila, c + 1).Value = datos(r, c)
            ws.Cells(fila, c + 1).Interior.Color = colores(r)
            ws.Cells(fila, c + 1).Borders.LineStyle = xlContinuous
        Next c
        ws.Cells(fila, 1).Font.Bold = (r = 0)
        fila = fila + 1
    Next r
    
    ws.Columns("A:D").AutoFit
End Sub

' ============================================================
' TABLA POR CARRERA
' ============================================================

Private Sub EscribirTablaCarreras(ws As Worksheet, dic As Object, filaInicio As Long)
    Dim fila As Long
    fila = filaInicio
    
    ws.Range("A" & fila).Value = "DETALLE POR CARRERA"
    ws.Range("A" & fila).Font.Bold = True
    ws.Range("A" & fila).Font.Size = 12
    ws.Range("A" & fila).Font.Color = RGB(0, 51, 102)
    fila = fila + 1
    
    Dim cols As Variant
    cols = Array("Carrera", "Aspirantes", "Matriculados", "Pendientes", "Inact./Retirados", "% Confirmación")
    Dim h As Long
    For h = 0 To UBound(cols)
        ws.Cells(fila, h + 1).Value = cols(h)
        ws.Cells(fila, h + 1).Font.Bold = True
        ws.Cells(fila, h + 1).Interior.Color = RGB(0, 102, 153)
        ws.Cells(fila, h + 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(fila, h + 1).Borders.LineStyle = xlContinuous
    Next h
    fila = fila + 1
    
    Dim totalAsp As Long
    Dim totalMat As Long
    Dim totalPen As Long
    Dim totalIna As Long
    
    Dim key As Variant
    For Each key In dic.Keys
        Dim arr As Variant
        arr = dic(key)
        Dim asp As Long: asp = arr(0)
        Dim mat As Long: mat = arr(1)
        Dim pen As Long: pen = arr(2)
        Dim ina As Long: ina = arr(3)
        Dim pct As Double
        If asp > 0 Then pct = mat / asp * 100 Else pct = 0
        
        ws.Cells(fila, 1).Value = key
        ws.Cells(fila, 2).Value = asp
        ws.Cells(fila, 3).Value = mat
        ws.Cells(fila, 4).Value = pen
        ws.Cells(fila, 5).Value = ina
        ws.Cells(fila, 6).Value = Format(pct, "0.0") & "%"
        
        ' Colorear % Confirmación
        Dim colorPct As Long
        If pct >= 80 Then
            colorPct = RGB(198, 239, 206)
        ElseIf pct >= 65 Then
            colorPct = RGB(255, 235, 156)
        Else
            colorPct = RGB(255, 199, 206)
        End If
        ws.Cells(fila, 6).Interior.Color = colorPct
        
        Dim c As Long
        For c = 1 To 5
            ws.Cells(fila, c).Borders.LineStyle = xlContinuous
        Next c
        ws.Cells(fila, 6).Borders.LineStyle = xlContinuous
        
        totalAsp = totalAsp + asp
        totalMat = totalMat + mat
        totalPen = totalPen + pen
        totalIna = totalIna + ina
        
        fila = fila + 1
    Next key
    
    ' Fila totales
    ws.Cells(fila, 1).Value = "TOTAL"
    ws.Cells(fila, 2).Value = totalAsp
    ws.Cells(fila, 3).Value = totalMat
    ws.Cells(fila, 4).Value = totalPen
    ws.Cells(fila, 5).Value = totalIna
    Dim pctTotal As Double
    If totalAsp > 0 Then pctTotal = totalMat / totalAsp * 100 Else pctTotal = 0
    ws.Cells(fila, 6).Value = Format(pctTotal, "0.0") & "%"
    
    For c = 1 To 6
        ws.Cells(fila, c).Font.Bold = True
        ws.Cells(fila, c).Interior.Color = RGB(191, 191, 191)
        ws.Cells(fila, c).Borders.LineStyle = xlContinuous
    Next c
    
    ws.Columns("A:F").AutoFit
End Sub

' ============================================================
' TABLA POR SEDE
' ============================================================

Private Sub EscribirTablaSedes(ws As Worksheet, dic As Object, filaInicio As Long)
    Dim fila As Long
    fila = filaInicio
    
    ws.Range("A" & fila).Value = "DETALLE POR SEDE"
    ws.Range("A" & fila).Font.Bold = True
    ws.Range("A" & fila).Font.Size = 12
    ws.Range("A" & fila).Font.Color = RGB(0, 51, 102)
    fila = fila + 1
    
    Dim cols As Variant
    cols = Array("Sede", "Aspirantes", "Matriculados", "Pendientes", "Inact./Retirados", "% Confirmación")
    Dim h As Long
    For h = 0 To UBound(cols)
        ws.Cells(fila, h + 1).Value = cols(h)
        ws.Cells(fila, h + 1).Font.Bold = True
        ws.Cells(fila, h + 1).Interior.Color = RGB(0, 102, 153)
        ws.Cells(fila, h + 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(fila, h + 1).Borders.LineStyle = xlContinuous
    Next h
    fila = fila + 1
    
    Dim totalAsp As Long
    Dim totalMat As Long
    Dim totalPen As Long
    Dim totalIna As Long
    
    Dim key As Variant
    For Each key In dic.Keys
        Dim arr As Variant
        arr = dic(key)
        Dim asp As Long: asp = arr(0)
        Dim mat As Long: mat = arr(1)
        Dim pen As Long: pen = arr(2)
        Dim ina As Long: ina = arr(3)
        Dim pct As Double
        If asp > 0 Then pct = mat / asp * 100 Else pct = 0
        
        ws.Cells(fila, 1).Value = key
        ws.Cells(fila, 2).Value = asp
        ws.Cells(fila, 3).Value = mat
        ws.Cells(fila, 4).Value = pen
        ws.Cells(fila, 5).Value = ina
        ws.Cells(fila, 6).Value = Format(pct, "0.0") & "%"
        
        Dim colorPct As Long
        If pct >= 80 Then
            colorPct = RGB(198, 239, 206)
        ElseIf pct >= 65 Then
            colorPct = RGB(255, 235, 156)
        Else
            colorPct = RGB(255, 199, 206)
        End If
        ws.Cells(fila, 6).Interior.Color = colorPct
        
        Dim c As Long
        For c = 1 To 6
            ws.Cells(fila, c).Borders.LineStyle = xlContinuous
        Next c
        
        totalAsp = totalAsp + asp
        totalMat = totalMat + mat
        totalPen = totalPen + pen
        totalIna = totalIna + ina
        
        fila = fila + 1
    Next key
    
    ' Fila totales
    ws.Cells(fila, 1).Value = "TOTAL"
    ws.Cells(fila, 2).Value = totalAsp
    ws.Cells(fila, 3).Value = totalMat
    ws.Cells(fila, 4).Value = totalPen
    ws.Cells(fila, 5).Value = totalIna
    Dim pctTotal As Double
    If totalAsp > 0 Then pctTotal = totalMat / totalAsp * 100 Else pctTotal = 0
    ws.Cells(fila, 6).Value = Format(pctTotal, "0.0") & "%"
    
    For c = 1 To 6
        ws.Cells(fila, c).Font.Bold = True
        ws.Cells(fila, c).Interior.Color = RGB(191, 191, 191)
        ws.Cells(fila, c).Borders.LineStyle = xlContinuous
    Next c
    
    ws.Columns("A:F").AutoFit
End Sub

' ============================================================
' FUNCIONES AUXILIARES
' ============================================================

Private Function PorcentajeStr(parte As Long, total As Long) As String
    If total = 0 Then
        PorcentajeStr = "0.0%"
    Else
        PorcentajeStr = Format(parte / total * 100, "0.0") & "%"
    End If
End Function

Private Function ObtenerPeriodoActual() As String
    ' Lee el período desde los datos SIGU; si no hay datos, retorna el año actual
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_SIGU)
    If Not ws Is Nothing Then
        Dim valor As String
        valor = Trim(CStr(ws.Cells(2, COL_SIGU_PERIODO).Value))
        If valor <> "" Then
            ObtenerPeriodoActual = valor
            Exit Function
        End If
    End If
    On Error GoTo 0
    ObtenerPeriodoActual = "I-" & Year(Now)
End Function

Private Sub ActualizarEstadoPanelSim(estado As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_PANEL)
    If Not ws Is Nothing Then
        ws.Range("B5").Value = estado
        If UCase(estado) = "LISTO" Then
            ws.Range("B5").Font.Color = RGB(0, 128, 0)
        Else
            ws.Range("B5").Font.Color = RGB(0, 0, 200)
        End If
    End If
    On Error GoTo 0
    DoEvents
End Sub
