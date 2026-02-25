Attribute VB_Name = "ModSimulacion"
'==============================================================================
' MÓDULO: ModSimulacion
' DESCRIPCIÓN: Simulación del proceso de matrícula basada en el listado
'              "Listado de personas matriculadas 2026".
'              Distingue entre ASPIRANTES (quienes aplicaron) y
'              MATRICULADOS (quienes confirmaron la matrícula).
' PROYECTO: Automatizador_Matricula.xlsm
' VERSIÓN: 1.0
' FECHA: Febrero 2026
'==============================================================================
Option Explicit

' ============================================================
' CONSTANTES INTERNAS DE SIMULACIÓN
' ============================================================

Private Const HOJA_SIMULACION As String = "Simulacion"
Private Const HOJA_ASPIRANTES As String = "Aspirantes"

' Columnas de la hoja Simulacion / resumen
Private Const COL_SIM_CATEGORIA As Long = 1    ' A - Categoría
Private Const COL_SIM_CANTIDAD  As Long = 2    ' B - Cantidad
Private Const COL_SIM_PORCENTAJE As Long = 3   ' C - Porcentaje

' ============================================================
' DATOS DE MUESTRA — nombres y apellidos costarricenses
' ============================================================

Private Function ObtenerNombres() As Variant
    ObtenerNombres = Array( _
        "Andrés", "María", "Carlos", "Laura", "José", _
        "Andrea", "Luis", "Sofía", "Diego", "Valeria", _
        "Miguel", "Daniela", "Jorge", "Alejandra", "Pablo", _
        "Gabriela", "Ricardo", "Melissa", "Eduardo", "Natalia", _
        "Fernando", "Lucía", "Óscar", "Carmen", "Sergio", _
        "Paola", "Roberto", "Jimena", "Adrián", "Fabiola")
End Function

Private Function ObtenerApellidos() As Variant
    ObtenerApellidos = Array( _
        "Rodríguez", "González", "Vargas", "Mora", "Jiménez", _
        "Solano", "Arias", "Quesada", "Herrera", "Rojas", _
        "Chaves", "Castro", "Pérez", "Zamora", "Brenes", _
        "Acuña", "Vindas", "Salas", "Méndez", "Alpízar", _
        "Fallas", "Ugalde", "Picado", "Badilla", "Sánchez", _
        "López", "Cordero", "Mora", "Blanco", "Vega")
End Function

Private Function ObtenerCarreras() As Variant
    ObtenerCarreras = Array( _
        "Administración de Empresas", _
        "Ingeniería en Computación", _
        "Contaduría Pública", _
        "Gestión del Turismo", _
        "Ingeniería en Producción Industrial", _
        "Educación Técnica")
End Function

Private Function ObtenerSedes() As Variant
    ObtenerSedes = Array( _
        "Central", "San Carlos", "Alajuela", "Limon")
End Function

Private Function ObtenerDocumentosLista() As Variant
    ObtenerDocumentosLista = Array( _
        "Cédula", "Foto", "Título Secundaria", _
        "Notas Secundaria", "Constancia CCSS", "Declaración Jurada")
End Function

' ============================================================
' MACRO PRINCIPAL: EJECUTAR SIMULACIÓN
' ============================================================

Public Sub EjecutarSimulacion()
    '
    ' Genera datos de muestra para ASPIRANTES y MATRICULADOS,
    ' los carga en las hojas del sistema y ejecuta el cruce
    ' para producir un informe de simulación.
    '
    Dim totalAspirantes As Long
    Dim porcentajeConversion As Double
    
    ' Pedir parámetros al usuario
    Dim inputAsp As String
    inputAsp = InputBox( _
        "SIMULACIÓN DE MATRÍCULA 2026" & vbCrLf & vbCrLf & _
        "Ingrese la cantidad de ASPIRANTES (estudiantes que aplicaron):", _
        "Parámetros de Simulación", "50")
    
    If inputAsp = "" Then Exit Sub
    If Not IsNumeric(inputAsp) Then
        MsgBox "Debe ingresar un número válido.", vbExclamation, "Error"
        Exit Sub
    End If
    totalAspirantes = CLng(inputAsp)
    If totalAspirantes < 1 Or totalAspirantes > 500 Then
        MsgBox "Ingrese un número entre 1 y 500.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Dim inputPct As String
    inputPct = InputBox( _
        "¿Qué porcentaje de aspirantes confirmaron matrícula?" & vbCrLf & _
        "(ingrese un número entre 1 y 100):", _
        "Parámetros de Simulación", "70")
    
    If inputPct = "" Then Exit Sub
    If Not IsNumeric(inputPct) Then
        MsgBox "Debe ingresar un número válido.", vbExclamation, "Error"
        Exit Sub
    End If
    porcentajeConversion = CDbl(inputPct)
    If porcentajeConversion < 1 Or porcentajeConversion > 100 Then
        MsgBox "Ingrese un porcentaje entre 1 y 100.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Dim totalMatriculados As Long
    totalMatriculados = CLng(totalAspirantes * (porcentajeConversion / 100))
    
    If MsgBox("RESUMEN DE SIMULACIÓN" & vbCrLf & vbCrLf & _
              "Aspirantes a generar:   " & totalAspirantes & vbCrLf & _
              "Matriculados esperados: " & totalMatriculados & _
              " (" & Format(porcentajeConversion, "0.0") & "%)" & vbCrLf & _
              "Sin confirmar:          " & (totalAspirantes - totalMatriculados) & vbCrLf & vbCrLf & _
              "Se sobrescribirán los datos de las hojas SIGU y Manual." & vbCrLf & _
              "¿Desea continuar?", _
              vbYesNo + vbQuestion, "Confirmar Simulación") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Asegurarse de que las hojas base existen
    InicializarHojasBase
    
    MostrarProgreso "Generando datos de aspirantes...", 10
    GenerarDatosAspirantes totalAspirantes
    
    MostrarProgreso "Cargando datos de aspirantes en hoja SIGU...", 30
    CargarAspirantesEnSIGU totalAspirantes
    
    MostrarProgreso "Cargando matriculados en hoja Manual...", 50
    CargarMatriculadosEnManual totalAspirantes, totalMatriculados
    
    MostrarProgreso "Ejecutando cruce de datos...", 70
    EjecutarCruceDeDatos
    
    MostrarProgreso "Generando hoja de resumen de simulación...", 85
    GenerarResumenSimulacion totalAspirantes, totalMatriculados, porcentajeConversion
    
    MostrarProgreso "Simulación completada.", 100
    
    RegistrarHistorial "SIMULACION", "", _
        "Simulación: " & totalAspirantes & " aspirantes, " & _
        totalMatriculados & " matriculados (" & _
        Format(porcentajeConversion, "0.0") & "%)", "ÉXITO"
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    LimpiarProgreso
    
    ' Mostrar la hoja de resumen
    On Error Resume Next
    ThisWorkbook.Sheets(HOJA_SIMULACION).Activate
    On Error GoTo 0
    
    MsgBox "SIMULACIÓN COMPLETADA" & vbCrLf & vbCrLf & _
           "Aspirantes generados:   " & totalAspirantes & vbCrLf & _
           "Matriculados (SIGU):    " & totalMatriculados & vbCrLf & _
           "Sin confirmar:          " & (totalAspirantes - totalMatriculados) & vbCrLf & _
           "Tasa de conversión:     " & Format(porcentajeConversion, "0.0") & "%" & vbCrLf & vbCrLf & _
           "Revise la hoja '" & HOJA_SIMULACION & "' para el detalle.", _
           vbInformation, "Simulación Finalizada"
    
    Exit Sub

ErrorHandler:
    MsgBox "Error en la simulación: " & Err.Description, vbCritical, "Error"
    RegistrarHistorial "SIMULACION", "", "Error: " & Err.Description, "ERROR"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    LimpiarProgreso
End Sub

' ============================================================
' GENERADOR DE DATOS DE MUESTRA
' ============================================================

' Array modular para guardar los registros generados en memoria
Private gAspirantes() As String   ' cedula|carnet|nombre|ap1|ap2|correo|tel|carrera|sede|docs
Private gTotalGenerados As Long

Private Sub GenerarDatosAspirantes(totalAspirantes As Long)
    Dim nombres As Variant
    Dim apellidos As Variant
    Dim carreras As Variant
    Dim sedes As Variant
    Dim docs As Variant
    
    nombres = ObtenerNombres()
    apellidos = ObtenerApellidos()
    carreras = ObtenerCarreras()
    sedes = ObtenerSedes()
    docs = ObtenerDocumentosLista()
    
    ReDim gAspirantes(1 To totalAspirantes, 1 To 10)
    gTotalGenerados = totalAspirantes
    
    Randomize
    
    Dim i As Long
    For i = 1 To totalAspirantes
        ' Cédula ficticia (9 dígitos)
        Dim cedula As String
        cedula = Format(100000000 + i, "000000000")
        
        ' Carnet (periodo 2026 + correlativo)
        Dim carnet As String
        carnet = "2026" & Format(i, "0000")
        
        ' Nombre y apellidos aleatorios
        Dim idxNom As Long
        Dim idxAp1 As Long
        Dim idxAp2 As Long
        idxNom = Int(Rnd() * (UBound(nombres) + 1))
        idxAp1 = Int(Rnd() * (UBound(apellidos) + 1))
        idxAp2 = Int(Rnd() * (UBound(apellidos) + 1))
        
        Dim nombre As String: nombre = CStr(nombres(idxNom))
        Dim ap1 As String:    ap1 = CStr(apellidos(idxAp1))
        Dim ap2 As String:    ap2 = CStr(apellidos(idxAp2))
        
        ' Correo
        Dim correo As String
        correo = LCase(nombre) & "." & LCase(ap1) & i & "@est.utn.ac.cr"
        correo = Replace(correo, " ", "")
        correo = Replace(correo, "á", "a")
        correo = Replace(correo, "é", "e")
        correo = Replace(correo, "í", "i")
        correo = Replace(correo, "ó", "o")
        correo = Replace(correo, "ú", "u")
        correo = Replace(correo, "ó", "o")
        correo = Replace(correo, "ñ", "n")
        
        ' Teléfono ficticio (8 dígitos)
        Dim tel As String
        tel = Format(60000000 + Int(Rnd() * 39999999), "00000000")
        tel = Left(tel, 4) & "-" & Right(tel, 4)
        
        ' Carrera y sede
        Dim idxCarr As Long: idxCarr = Int(Rnd() * (UBound(carreras) + 1))
        Dim idxSede As Long: idxSede = Int(Rnd() * (UBound(sedes) + 1))
        
        ' Documentos entregados (algunos completos, algunos parciales)
        Dim docsEntregados As String
        Dim numDocs As Long
        numDocs = Int(Rnd() * (UBound(docs) + 2))  ' 0 .. totalDocs
        Dim docList As String: docList = ""
        Dim d As Long
        For d = 0 To numDocs - 1
            If d <= UBound(docs) Then
                If docList <> "" Then docList = docList & ","
                docList = docList & CStr(docs(d))
            End If
        Next d
        
        ' Almacenar
        gAspirantes(i, 1) = cedula
        gAspirantes(i, 2) = carnet
        gAspirantes(i, 3) = nombre
        gAspirantes(i, 4) = ap1
        gAspirantes(i, 5) = ap2
        gAspirantes(i, 6) = correo
        gAspirantes(i, 7) = tel
        gAspirantes(i, 8) = CStr(carreras(idxCarr))
        gAspirantes(i, 9) = CStr(sedes(idxSede))
        gAspirantes(i, 10) = docList
    Next i
End Sub

' ============================================================
' CARGAR DATOS EN HOJA SIGU (todos los aspirantes)
' ============================================================

Private Sub CargarAspirantesEnSIGU(totalAspirantes As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_SIGU)
    
    LimpiarHoja ws, True
    
    Dim i As Long
    For i = 1 To totalAspirantes
        Dim fila As Long: fila = i + 1
        
        ws.Cells(fila, COL_SIGU_CEDULA).Value = gAspirantes(i, 1)
        ws.Cells(fila, COL_SIGU_CARNET).Value = gAspirantes(i, 2)
        ws.Cells(fila, COL_SIGU_NOMBRE).Value = gAspirantes(i, 3)
        ws.Cells(fila, COL_SIGU_APELLIDO1).Value = gAspirantes(i, 4)
        ws.Cells(fila, COL_SIGU_APELLIDO2).Value = gAspirantes(i, 5)
        ws.Cells(fila, COL_SIGU_CORREO).Value = gAspirantes(i, 6)
        ws.Cells(fila, COL_SIGU_TELEFONO).Value = gAspirantes(i, 7)
        ws.Cells(fila, COL_SIGU_CARRERA).Value = gAspirantes(i, 8)
        ws.Cells(fila, COL_SIGU_SEDE).Value = gAspirantes(i, 9)
        ws.Cells(fila, COL_SIGU_ESTADO).Value = "Aspirante"
        ws.Cells(fila, COL_SIGU_PERIODO).Value = "2026-I"
        ws.Cells(fila, COL_SIGU_FECHA_MATRICULA).Value = ""
        ws.Cells(fila, COL_SIGU_DOCUMENTOS).Value = gAspirantes(i, 10)
        ws.Cells(fila, COL_SIGU_OBSERVACIONES).Value = ""
    Next i
    
    ' Actualizar contador en configuración
    On Error Resume Next
    ThisWorkbook.Sheets(HOJA_CONFIG).Range("B13").Value = totalAspirantes
    On Error GoTo 0
End Sub

' ============================================================
' CARGAR DATOS EN HOJA MANUAL (solo matriculados = confirmados)
' ============================================================

Private Sub CargarMatriculadosEnManual(totalAspirantes As Long, totalMatriculados As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_MANUAL)
    
    LimpiarHoja ws, True
    
    ' Los primeros totalMatriculados aspirantes se consideran confirmados
    Dim i As Long
    For i = 1 To totalMatriculados
        Dim fila As Long: fila = i + 1
        
        ws.Cells(fila, COL_MAN_CEDULA).Value = gAspirantes(i, 1)
        ws.Cells(fila, COL_MAN_CARNET).Value = gAspirantes(i, 2)
        ws.Cells(fila, COL_MAN_NOMBRE).Value = gAspirantes(i, 3)
        ws.Cells(fila, COL_MAN_APELLIDO1).Value = gAspirantes(i, 4)
        ws.Cells(fila, COL_MAN_APELLIDO2).Value = gAspirantes(i, 5)
        ws.Cells(fila, COL_MAN_CORREO).Value = gAspirantes(i, 6)
        ws.Cells(fila, COL_MAN_TELEFONO).Value = gAspirantes(i, 7)
        ws.Cells(fila, COL_MAN_CARRERA).Value = gAspirantes(i, 8)
        ws.Cells(fila, COL_MAN_SEDE).Value = gAspirantes(i, 9)
        ws.Cells(fila, COL_MAN_ESTADO).Value = "Matriculado"
        ws.Cells(fila, COL_MAN_PERIODO).Value = "2026-I"
        ws.Cells(fila, COL_MAN_FECHA_MATRICULA).Value = Format(Date, "dd/mm/yyyy")
        ws.Cells(fila, COL_MAN_DOCUMENTOS).Value = gAspirantes(i, 10)
        ws.Cells(fila, COL_MAN_OBSERVACIONES).Value = ""
        ws.Cells(fila, COL_MAN_ULTIMA_SYNC).Value = Format(Now, "dd/mm/yyyy hh:mm")
        ws.Cells(fila, COL_MAN_CORREO_ENVIADO).Value = "No"
        ws.Cells(fila, COL_MAN_FECHA_CORREO).Value = ""
    Next i
    
    ' Actualizar contador en configuración
    On Error Resume Next
    ThisWorkbook.Sheets(HOJA_CONFIG).Range("B14").Value = totalMatriculados
    On Error GoTo 0
End Sub

' ============================================================
' GENERAR HOJA DE RESUMEN DE SIMULACIÓN
' ============================================================

Private Sub GenerarResumenSimulacion(totalAspirantes As Long, totalMatriculados As Long, _
                                     porcentajeConversion As Double)
    ' Crear o limpiar hoja de Simulación
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
    
    Dim sinConfirmar As Long
    sinConfirmar = totalAspirantes - totalMatriculados
    
    ' --- Título ---
    With ws.Range("A1:E1")
        .Merge
        .Value = "SIMULACIÓN DE MATRÍCULA 2026 — ASPIRANTES VS MATRICULADOS"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 36
    
    ' --- Fecha ---
    ws.Range("A2").Value = "Generado el:"
    ws.Range("B2").Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
    ws.Range("A2").Font.Bold = True
    
    ' --- Tabla de resumen ---
    Dim encFila As Long: encFila = 4
    
    ws.Cells(encFila, COL_SIM_CATEGORIA).Value = "Categoría"
    ws.Cells(encFila, COL_SIM_CANTIDAD).Value = "Cantidad"
    ws.Cells(encFila, COL_SIM_PORCENTAJE).Value = "% del Total"
    
    With ws.Range(ws.Cells(encFila, 1), ws.Cells(encFila, 3))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 153)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Fila 1: Aspirantes
    ws.Cells(encFila + 1, COL_SIM_CATEGORIA).Value = "Aspirantes (total aplicaron)"
    ws.Cells(encFila + 1, COL_SIM_CANTIDAD).Value = totalAspirantes
    ws.Cells(encFila + 1, COL_SIM_PORCENTAJE).Value = 1
    ws.Cells(encFila + 1, COL_SIM_PORCENTAJE).NumberFormat = "0.0%"
    ws.Range(ws.Cells(encFila + 1, 1), ws.Cells(encFila + 1, 3)).Interior.Color = RGB(189, 215, 238)
    
    ' Fila 2: Matriculados
    ws.Cells(encFila + 2, COL_SIM_CATEGORIA).Value = "Matriculados (confirmaron)"
    ws.Cells(encFila + 2, COL_SIM_CANTIDAD).Value = totalMatriculados
    If totalAspirantes > 0 Then
        ws.Cells(encFila + 2, COL_SIM_PORCENTAJE).Value = totalMatriculados / totalAspirantes
    Else
        ws.Cells(encFila + 2, COL_SIM_PORCENTAJE).Value = 0
    End If
    ws.Cells(encFila + 2, COL_SIM_PORCENTAJE).NumberFormat = "0.0%"
    ws.Range(ws.Cells(encFila + 2, 1), ws.Cells(encFila + 2, 3)).Interior.Color = RGB(198, 239, 206)
    
    ' Fila 3: Sin confirmar
    ws.Cells(encFila + 3, COL_SIM_CATEGORIA).Value = "Sin confirmar (no matricularon)"
    ws.Cells(encFila + 3, COL_SIM_CANTIDAD).Value = sinConfirmar
    If totalAspirantes > 0 Then
        ws.Cells(encFila + 3, COL_SIM_PORCENTAJE).Value = sinConfirmar / totalAspirantes
    Else
        ws.Cells(encFila + 3, COL_SIM_PORCENTAJE).Value = 0
    End If
    ws.Cells(encFila + 3, COL_SIM_PORCENTAJE).NumberFormat = "0.0%"
    ws.Range(ws.Cells(encFila + 3, 1), ws.Cells(encFila + 3, 3)).Interior.Color = RGB(255, 199, 206)
    
    ' Bordes
    With ws.Range(ws.Cells(encFila, 1), ws.Cells(encFila + 3, 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' --- Estadísticas del cruce ---
    Dim wsCruce As Worksheet
    Set wsCruce = ThisWorkbook.Sheets(HOJA_CRUCE)
    Dim ultimaFilaCruce As Long
    ultimaFilaCruce = ObtenerUltimaFila(wsCruce)
    
    Dim detFila As Long: detFila = encFila + 6
    
    ws.Cells(detFila, 1).Value = "RESULTADO DEL CRUCE DE DATOS"
    ws.Cells(detFila, 1).Font.Bold = True
    ws.Cells(detFila, 1).Font.Size = 12
    ws.Range(ws.Cells(detFila, 1), ws.Cells(detFila, 3)).Interior.Color = RGB(220, 220, 220)
    
    detFila = detFila + 1
    ws.Cells(detFila, 1).Value = "Tipo de Cambio"
    ws.Cells(detFila, 2).Value = "Cantidad"
    With ws.Range(ws.Cells(detFila, 1), ws.Cells(detFila, 2))
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 153)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    Dim tipos As Variant
    Dim descs As Variant
    Dim colores As Variant
    
    tipos = Array(TIPO_NUEVO, TIPO_CAMBIO_ESTADO, TIPO_CAMBIO_SEDE, _
                  TIPO_CAMBIO_CARRERA, TIPO_INCONSISTENTE, TIPO_DOC_PENDIENTE, TIPO_DATO_ACTUALIZADO)
    descs = Array("Nuevos en SIGU (sin confirmar)", "Cambio de estado", "Cambio de sede", _
                  "Cambio de carrera", "Inconsistentes", "Documentos pendientes", "Datos actualizados")
    colores = Array(RGB(198, 239, 206), RGB(255, 235, 156), RGB(189, 215, 238), _
                    RGB(189, 215, 238), RGB(255, 199, 206), RGB(255, 220, 180), RGB(230, 230, 250))
    
    Dim j As Long
    For j = LBound(tipos) To UBound(tipos)
        detFila = detFila + 1
        Dim cnt As Long
        cnt = ContarPorTipoCambio(wsCruce, CStr(tipos(j)))
        ws.Cells(detFila, 1).Value = CStr(descs(j))
        ws.Cells(detFila, 2).Value = cnt
        ws.Range(ws.Cells(detFila, 1), ws.Cells(detFila, 2)).Interior.Color = CLng(colores(j))
    Next j
    
    With ws.Range(ws.Cells(encFila + 7, 1), ws.Cells(detFila, 2)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' --- Pendientes de documentos ---
    Dim pendFila As Long: pendFila = detFila + 3
    ws.Cells(pendFila, 1).Value = "MATRICULADOS CON DOCUMENTOS PENDIENTES"
    ws.Cells(pendFila, 1).Font.Bold = True
    ws.Cells(pendFila, 1).Font.Size = 12
    ws.Range(ws.Cells(pendFila, 1), ws.Cells(pendFila, 5)).Interior.Color = RGB(255, 220, 180)
    
    pendFila = pendFila + 1
    ws.Cells(pendFila, 1).Value = "Cédula"
    ws.Cells(pendFila, 2).Value = "Nombre Completo"
    ws.Cells(pendFila, 3).Value = "Sede"
    ws.Cells(pendFila, 4).Value = "Carrera"
    ws.Cells(pendFila, 5).Value = "Documentos Pendientes"
    With ws.Range(ws.Cells(pendFila, 1), ws.Cells(pendFila, 5))
        .Font.Bold = True
        .Interior.Color = RGB(204, 102, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Listar matriculados con docs pendientes desde la hoja SIGU
    Dim wsSIGU As Worksheet
    Set wsSIGU = ThisWorkbook.Sheets(HOJA_SIGU)
    Dim wsManual As Worksheet
    Set wsManual = ThisWorkbook.Sheets(HOJA_MANUAL)
    
    Dim ultimaFilaMan As Long
    ultimaFilaMan = ObtenerUltimaFila(wsManual)
    
    Dim pendCount As Long: pendCount = 0
    Dim m As Long
    For m = 2 To ultimaFilaMan
        Dim cedMan As String
        cedMan = Trim(CStr(wsManual.Cells(m, COL_MAN_CEDULA).Value))
        Dim docsEnt As String
        docsEnt = Trim(CStr(wsManual.Cells(m, COL_MAN_DOCUMENTOS).Value))
        Dim pendientes As String
        pendientes = ObtenerDocumentosPendientes(docsEnt)
        
        If pendientes <> "" Then
            pendFila = pendFila + 1
            pendCount = pendCount + 1
            ws.Cells(pendFila, 1).Value = cedMan
            ws.Cells(pendFila, 2).Value = Trim(CStr(wsManual.Cells(m, COL_MAN_NOMBRE).Value)) & " " & _
                                           Trim(CStr(wsManual.Cells(m, COL_MAN_APELLIDO1).Value)) & " " & _
                                           Trim(CStr(wsManual.Cells(m, COL_MAN_APELLIDO2).Value))
            ws.Cells(pendFila, 3).Value = wsManual.Cells(m, COL_MAN_SEDE).Value
            ws.Cells(pendFila, 4).Value = wsManual.Cells(m, COL_MAN_CARRERA).Value
            ws.Cells(pendFila, 5).Value = pendientes
            
            If pendCount Mod 2 = 0 Then
                ws.Range(ws.Cells(pendFila, 1), ws.Cells(pendFila, 5)).Interior.Color = RGB(255, 243, 220)
            End If
        End If
    Next m
    
    If pendCount = 0 Then
        pendFila = pendFila + 1
        ws.Cells(pendFila, 1).Value = "Todos los matriculados tienen documentos completos."
        ws.Cells(pendFila, 1).Font.Italic = True
    End If
    
    ' Totalizador de pendientes
    pendFila = pendFila + 1
    ws.Cells(pendFila, 1).Value = "Total con documentos pendientes:"
    ws.Cells(pendFila, 1).Font.Bold = True
    ws.Cells(pendFila, 2).Value = pendCount
    ws.Cells(pendFila, 2).Font.Bold = True
    
    ' Ajustar columnas
    ws.Columns("A:E").AutoFit
    ws.Columns("A").ColumnWidth = 35
    ws.Columns("E").ColumnWidth = 50
End Sub

' ============================================================
' MACRO DE ACCESO RÁPIDO — llamada desde el panel
' ============================================================

Public Sub Btn_Simulacion()
    EjecutarSimulacion
End Sub
