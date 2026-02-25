# ==============================================================================
# Script: generar_simulacion.ps1
# Descripcion: Crea archivos Excel simulados de matricula UTN
# Aspirantes = total que aplicaron / Matriculados = confirmados
# Universidad Tecnica Nacional - Costa Rica
# Fecha: Febrero 2026
# ==============================================================================

$carpeta = "c:\Users\LINC\Desktop\cda"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

# ==============================================================================
# DATOS BASE
# ==============================================================================

$carreras = @(
    @{ nombre = "Ingenieria en Tecnologias de Informacion"; codigo = "ITI"; area = "Ingenieria" },
    @{ nombre = "Administracion de Empresas";               codigo = "ADE"; area = "Ciencias Economicas" },
    @{ nombre = "Contaduria Publica";                       codigo = "CPU"; area = "Ciencias Economicas" },
    @{ nombre = "Gestion del Turismo Sostenible";           codigo = "GTS"; area = "Turismo" },
    @{ nombre = "Ingenieria Electromecanica Industrial";    codigo = "IEI"; area = "Ingenieria" },
    @{ nombre = "Produccion Industrial";                    codigo = "PIN"; area = "Ingenieria" },
    @{ nombre = "Salud Ocupacional";                        codigo = "SOC"; area = "Salud" },
    @{ nombre = "Asistencia Administrativa";                codigo = "AAD"; area = "Ciencias Economicas" },
    @{ nombre = "Ingles como Lengua Extranjera";            codigo = "ILE"; area = "Educacion" },
    @{ nombre = "Agronomia";                                codigo = "AGR"; area = "Ciencias Agrarias" }
)

$sedes = @("Sede Central Alajuela", "Sede Pacifico", "Sede Guanacaste", "Sede Atenas", "Sede San Carlos")

# aspirantesPorSede[carrera][sede]
$aspirantesPorCarrera = @(
    @(85, 42, 38, 20, 30),   # ITI
    @(70, 55, 45, 35, 40),   # ADE
    @(40, 30, 25, 18, 22),   # CPU
    @(25, 45, 35, 10, 15),   # GTS
    @(50, 20, 15, 25, 18),   # IEI
    @(35, 15, 12, 20, 10),   # PIN
    @(30, 25, 20, 12, 15),   # SOC
    @(45, 35, 30, 25, 28),   # AAD
    @(20, 15, 18,  8, 10),   # ILE
    @(15, 10, 12, 30, 35)    # AGR
)

# Tasa de confirmacion por carrera (matriculados/aspirantes)
$tasaConfirmacion = @(0.72, 0.78, 0.80, 0.65, 0.70, 0.75, 0.82, 0.77, 0.60, 0.68)

$nombresM = @("Jose","Carlos","Luis","Juan","Miguel","Daniel","Andres","Diego","David","Kevin",
              "Bryan","Esteban","Fabian","Gabriel","Alejandro","Roberto","Francisco","Fernando",
              "Marco","Javier","Oscar","Adrian","Cristian","Eduardo","Sebastian","Mauricio",
              "Rodrigo","Alvaro","Pablo","Hector","Ricardo","Gerardo","Randall","Jeffry","Josue")

$nombresF = @("Maria","Ana","Sofia","Laura","Andrea","Carolina","Daniela","Valeria","Natalia",
              "Gabriela","Paola","Katherine","Stephanie","Melissa","Monica","Rebeca","Adriana",
              "Lucia","Elena","Tatiana","Priscilla","Diana","Wendy","Karla","Fernanda","Isabel",
              "Marcela","Yuliana","Kimberly","Genesis","Alison","Pamela","Viviana","Nicole")

$apellidos = @("Rodriguez","Jimenez","Mora","Hernandez","Vargas","Solis","Chaves","Castro",
               "Arias","Rojas","Cordero","Monge","Gonzalez","Lopez","Ramirez","Perez",
               "Retana","Urena","Villalobos","Brenes","Ugalde","Sanchez","Vega","Salazar",
               "Quiros","Calderon","Araya","Campos","Valverde","Gutierrez","Barrantes",
               "Segura","Nunez","Madrigal","Camacho","Fallas","Bonilla","Zamora",
               "Alvarado","Quesada","Vindas","Aguilar","Montero","Picado","Cespedes",
               "Chavarria","Murillo","Gamboa","Leon","Acuna")

$documentosRequeridos = @("Cedula de identidad","Foto tamano pasaporte","Titulo de secundaria",
                          "Notas de secundaria","Constancia CCSS","Declaracion jurada")

$periodoActual = "I-2026"

# ==============================================================================
# FUNCIONES
# ==============================================================================

function Get-CedulaCR {
    $prov  = Get-Random -Minimum 1 -Maximum 8
    $tomo  = Get-Random -Minimum 100 -Maximum 2000
    $asie  = Get-Random -Minimum 100 -Maximum 1000
    return "$prov-$('{0:D4}' -f $tomo)-$('{0:D4}' -f $asie)"
}

function Get-CarnetUTN { param([int]$idx)
    $anio = Get-Random -Minimum 2022 -Maximum 2027
    return "UTN-$anio-$('{0:D5}' -f ($idx + 10000))"
}

function Get-Persona {
    $mujer = ((Get-Random -Minimum 0 -Maximum 2) -eq 1)
    if ($mujer) { $nom = $nombresF[(Get-Random -Minimum 0 -Maximum $nombresF.Length)] }
    else        { $nom = $nombresM[(Get-Random -Minimum 0 -Maximum $nombresM.Length)] }
    $ap1 = $apellidos[(Get-Random -Minimum 0 -Maximum $apellidos.Length)]
    $ap2 = $apellidos[(Get-Random -Minimum 0 -Maximum $apellidos.Length)]
    while ($ap2 -eq $ap1) { $ap2 = $apellidos[(Get-Random -Minimum 0 -Maximum $apellidos.Length)] }
    return @{ nombre=$nom; ap1=$ap1; ap2=$ap2 }
}

function Get-Correo { param([string]$nom, [string]$ap1)
    $n = $nom.ToLower().Substring(0,1)
    $a = $ap1.ToLower()
    $num = Get-Random -Minimum 1 -Maximum 99
    return "$n$a$num@est.utn.ac.cr"
}

function Get-Telefono {
    $pref = @("6","7","8")[(Get-Random -Minimum 0 -Maximum 3)]
    $num  = Get-Random -Minimum 1000000 -Maximum 9999999
    $s    = $num.ToString()
    return "+506 $pref$($s.Substring(0,3))-$($s.Substring(3,4))"
}

function Get-Documentos { param([string]$estado)
    if ($estado -eq "Activo") {
        if ((Get-Random -Minimum 0 -Maximum 100) -lt 75) {
            return ($documentosRequeridos -join ", ")
        } else {
            $n = $documentosRequeridos.Length - (Get-Random -Minimum 1 -Maximum 3)
            $sel = $documentosRequeridos | Get-Random -Count $n
            return ($sel -join ", ")
        }
    } elseif ($estado -eq "Pendiente") {
        $n = Get-Random -Minimum 1 -Maximum 4
        $sel = $documentosRequeridos | Get-Random -Count $n
        return ($sel -join ", ")
    } else {
        $n = Get-Random -Minimum 0 -Maximum 3
        if ($n -eq 0) { return "Ninguno" }
        $sel = $documentosRequeridos | Get-Random -Count $n
        return ($sel -join ", ")
    }
}

function Get-Faltantes { param([string]$entregados)
    if ($entregados -eq ($documentosRequeridos -join ", ")) { return "Ninguno" }
    if ($entregados -eq "Ninguno") { return ($documentosRequeridos -join ", ") }
    $lista = $entregados -split ", "
    $falt  = $documentosRequeridos | Where-Object { $_ -notin $lista }
    if ($falt.Count -eq 0) { return "Ninguno" }
    return ($falt -join ", ")
}

# ==============================================================================
# GENERAR LISTA DE ESTUDIANTES
# ==============================================================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  GENERADOR DE DATOS SIMULADOS UTN" -ForegroundColor Cyan
Write-Host "  Aspirantes vs Matriculados - $periodoActual" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

$estudiantes = [System.Collections.ArrayList]@()
$idx = 0

for ($c = 0; $c -lt $carreras.Length; $c++) {
    $carrera = $carreras[$c]
    $tasa    = $tasaConfirmacion[$c]
    $aspPorSede = $aspirantesPorCarrera[$c]

    for ($s = 0; $s -lt $sedes.Length; $s++) {
        $totalAsp = $aspPorSede[$s]
        $totalMat = [math]::Round($totalAsp * $tasa)
        $totalPen = [math]::Round(($totalAsp - $totalMat) * 0.6)

        for ($i = 0; $i -lt $totalAsp; $i++) {
            $idx++
            $p = Get-Persona

            if ($i -lt $totalMat) {
                $estado = "Activo"
                $fechaMat = (Get-Date).AddDays(-(Get-Random -Minimum 5 -Maximum 45)).ToString("dd/MM/yyyy")
            } elseif ($i -lt ($totalMat + $totalPen)) {
                $estado = "Pendiente"
                $fechaMat = ""
            } else {
                if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) { $estado = "Inactivo" } else { $estado = "Retirado" }
                $fechaMat = ""
            }

            $docs  = Get-Documentos $estado
            $falt  = Get-Faltantes $docs

            if ($estado -eq "Activo") {
                if ($falt -ne "Ninguno") { $obs = "Matriculado - documentos pendientes" }
                else                     { $obs = "Matricula completa" }
            } elseif ($estado -eq "Pendiente") {
                $obs = "Aspirante pendiente. Faltan: $falt"
            } elseif ($estado -eq "Inactivo") {
                $obs = "No completo proceso de matricula"
            } else {
                $obs = "Desistio del proceso"
            }

            $null = $estudiantes.Add(@{
                idx      = $idx
                cedula   = Get-CedulaCR
                carnet   = Get-CarnetUTN $idx
                nombre   = $p.nombre
                ap1      = $p.ap1
                ap2      = $p.ap2
                correo   = Get-Correo $p.nombre $p.ap1
                tel      = Get-Telefono
                carrera  = $carrera.nombre
                cod      = $carrera.codigo
                area     = $carrera.area
                sede     = $sedes[$s]
                estado   = $estado
                periodo  = $periodoActual
                fechaMat = $fechaMat
                docs     = $docs
                falt     = $falt
                obs      = $obs
            })
        }
    }
}

$totalEst = $estudiantes.Count
$totalMat = ($estudiantes | Where-Object { $_.estado -eq "Activo" }).Count
$totalPen = ($estudiantes | Where-Object { $_.estado -eq "Pendiente" }).Count
$totalIna = ($estudiantes | Where-Object { $_.estado -eq "Inactivo" -or $_.estado -eq "Retirado" }).Count

Write-Host ""
Write-Host "Total aspirantes: $totalEst" -ForegroundColor White
Write-Host "  Matriculados  : $totalMat" -ForegroundColor Green
Write-Host "  Pendientes    : $totalPen" -ForegroundColor Yellow
Write-Host "  Inact/Retirado: $totalIna" -ForegroundColor Red
Write-Host ""

# ==============================================================================
# LIBRO 1: EXPORTACION SIGU
# ==============================================================================

Write-Host "Creando archivo SIGU..." -ForegroundColor Cyan
$libroSIGU = $excel.Workbooks.Add()

# ---- Hoja 1: Resumen Aspirantes vs Matriculados ----
$hResumen = $libroSIGU.Sheets.Item(1)
$hResumen.Name = "Resumen_Aspirantes"

$hResumen.Range("A1:H1").Merge()
$hResumen.Cells.Item(1,1) = "UTN - Resumen Aspirantes y Matriculados - $periodoActual"
$hResumen.Cells.Item(1,1).Font.Bold = $true
$hResumen.Cells.Item(1,1).Font.Size = 14
$hResumen.Cells.Item(1,1).Font.Color = 16777215
$hResumen.Cells.Item(1,1).Interior.Color = 1644820   # Azul UTN

$hResumen.Cells.Item(2,1) = "Generado el: $(Get-Date -Format 'dd/MM/yyyy HH:mm')"

$hRes = @("Carrera","Codigo","Area","Sede","Aspirantes","Matriculados","Pendientes","% Confirmacion")
for ($h = 0; $h -lt $hRes.Length; $h++) {
    $hResumen.Cells.Item(4, $h+1) = $hRes[$h]
    $hResumen.Cells.Item(4, $h+1).Font.Bold = $true
    $hResumen.Cells.Item(4, $h+1).Interior.Color = 5287936
    $hResumen.Cells.Item(4, $h+1).Font.Color = 16777215
    $hResumen.Cells.Item(4, $h+1).Borders.LineStyle = 1
}

$fila = 5
$totAsp = 0; $totMat = 0; $totPen = 0

for ($c = 0; $c -lt $carreras.Length; $c++) {
    $carrera = $carreras[$c]
    for ($s = 0; $s -lt $sedes.Length; $s++) {
        $asp = $aspirantesPorCarrera[$c][$s]
        $mat = [math]::Round($asp * $tasaConfirmacion[$c])
        $pen = $asp - $mat
        $pct = [math]::Round(($mat / $asp) * 100, 1)
        $totAsp += $asp; $totMat += $mat; $totPen += $pen

        $hResumen.Cells.Item($fila,1) = $carrera.nombre
        $hResumen.Cells.Item($fila,2) = $carrera.codigo
        $hResumen.Cells.Item($fila,3) = $carrera.area
        $hResumen.Cells.Item($fila,4) = $sedes[$s]
        $hResumen.Cells.Item($fila,5) = $asp
        $hResumen.Cells.Item($fila,6) = $mat
        $hResumen.Cells.Item($fila,7) = $pen
        $hResumen.Cells.Item($fila,8) = "$pct%"

        if ($pct -ge 80)      { $hResumen.Cells.Item($fila,8).Interior.Color = 5287936;  $hResumen.Cells.Item($fila,8).Font.Color = 16777215 }
        elseif ($pct -ge 65)  { $hResumen.Cells.Item($fila,8).Interior.Color = 65535 }
        else                  { $hResumen.Cells.Item($fila,8).Interior.Color = 255;       $hResumen.Cells.Item($fila,8).Font.Color = 16777215 }

        for ($col = 1; $col -le 8; $col++) { $hResumen.Cells.Item($fila,$col).Borders.LineStyle = 1 }
        $fila++
    }
}

# Fila TOTAL
$hResumen.Cells.Item($fila,1) = "TOTAL GENERAL"
$hResumen.Cells.Item($fila,5) = $totAsp
$hResumen.Cells.Item($fila,6) = $totMat
$hResumen.Cells.Item($fila,7) = $totPen
$pctGen = [math]::Round(($totMat / $totAsp) * 100, 1)
$hResumen.Cells.Item($fila,8) = "$pctGen%"
for ($col = 1; $col -le 8; $col++) {
    $hResumen.Cells.Item($fila,$col).Font.Bold = $true
    $hResumen.Cells.Item($fila,$col).Interior.Color = 12632256
    $hResumen.Cells.Item($fila,$col).Borders.LineStyle = 1
}
for ($col = 1; $col -le 8; $col++) { $hResumen.Columns.Item($col).AutoFit() | Out-Null }

# ---- Hoja 2: Detalle SIGU ----
$hDetalle = $libroSIGU.Sheets.Add([System.Reflection.Missing]::Value, $hResumen)
$hDetalle.Name = "Datos_SIGU"

$hSIGU = @("Cedula","Carnet","Nombre","Primer Apellido","Segundo Apellido",
           "Correo Electronico","Telefono","Carrera","Sede","Estado",
           "Periodo","Fecha Matricula","Documentos Entregados","Observaciones")

for ($h = 0; $h -lt $hSIGU.Length; $h++) {
    $hDetalle.Cells.Item(1,$h+1) = $hSIGU[$h]
    $hDetalle.Cells.Item(1,$h+1).Font.Bold = $true
    $hDetalle.Cells.Item(1,$h+1).Interior.Color = 5287936
    $hDetalle.Cells.Item(1,$h+1).Font.Color = 16777215
    $hDetalle.Cells.Item(1,$h+1).Borders.LineStyle = 1
}

Write-Host "Escribiendo $totalEst registros en SIGU..." -ForegroundColor Yellow
$fila = 2; $cont = 0
foreach ($e in $estudiantes) {
    $cont++
    if ($cont % 200 -eq 0) { Write-Host "  $cont / $totalEst" -ForegroundColor Gray }

    $hDetalle.Cells.Item($fila,1)  = $e.cedula
    $hDetalle.Cells.Item($fila,2)  = $e.carnet
    $hDetalle.Cells.Item($fila,3)  = $e.nombre
    $hDetalle.Cells.Item($fila,4)  = $e.ap1
    $hDetalle.Cells.Item($fila,5)  = $e.ap2
    $hDetalle.Cells.Item($fila,6)  = $e.correo
    $hDetalle.Cells.Item($fila,7)  = $e.tel
    $hDetalle.Cells.Item($fila,8)  = $e.carrera
    $hDetalle.Cells.Item($fila,9)  = $e.sede
    $hDetalle.Cells.Item($fila,10) = $e.estado
    $hDetalle.Cells.Item($fila,11) = $e.periodo
    $hDetalle.Cells.Item($fila,12) = $e.fechaMat
    $hDetalle.Cells.Item($fila,13) = $e.docs
    $hDetalle.Cells.Item($fila,14) = $e.obs

    if    ($e.estado -eq "Activo")   { $col = 13434828 }
    elseif($e.estado -eq "Pendiente"){ $col = 10092543 }
    elseif($e.estado -eq "Inactivo") { $col = 13421823 }
    else                             { $col = 12632256 }

    for ($c = 1; $c -le 14; $c++) {
        $hDetalle.Cells.Item($fila,$c).Interior.Color = $col
        $hDetalle.Cells.Item($fila,$c).Borders.LineStyle = 1
    }
    $fila++
}
for ($col = 1; $col -le 14; $col++) { $hDetalle.Columns.Item($col).AutoFit() | Out-Null }
$hDetalle.Range("A1:N$($fila-1)").AutoFilter() | Out-Null

# ---- Hoja 3: Resumen por Sede ----
$hSede = $libroSIGU.Sheets.Add([System.Reflection.Missing]::Value, $hDetalle)
$hSede.Name = "Resumen_Sedes"
$hSede.Cells.Item(1,1) = "Resumen por Sede - $periodoActual"
$hSede.Cells.Item(1,1).Font.Bold = $true

$hS = @("Sede","Total Aspirantes","Matriculados","% Confirmacion")
for ($h = 0; $h -lt $hS.Length; $h++) {
    $hSede.Cells.Item(3,$h+1) = $hS[$h]
    $hSede.Cells.Item(3,$h+1).Font.Bold = $true
    $hSede.Cells.Item(3,$h+1).Interior.Color = 5287936
    $hSede.Cells.Item(3,$h+1).Font.Color = 16777215
}
$fila = 4
foreach ($sede in $sedes) {
    $aS = ($estudiantes | Where-Object { $_.sede -eq $sede }).Count
    $mS = ($estudiantes | Where-Object { $_.sede -eq $sede -and $_.estado -eq "Activo" }).Count
    $pS = if ($aS -gt 0) { [math]::Round(($mS/$aS)*100,1) } else { 0 }
    $hSede.Cells.Item($fila,1) = $sede
    $hSede.Cells.Item($fila,2) = $aS
    $hSede.Cells.Item($fila,3) = $mS
    $hSede.Cells.Item($fila,4) = "$pS%"
    for ($c = 1; $c -le 4; $c++) { $hSede.Cells.Item($fila,$c).Borders.LineStyle = 1 }
    $fila++
}
for ($col = 1; $col -le 4; $col++) { $hSede.Columns.Item($col).AutoFit() | Out-Null }

# Guardar SIGU
$rutaSIGU = "$carpeta\Datos_SIGU\SIGU_Exportacion_$(Get-Date -Format 'yyyyMMdd').xlsx"
$libroSIGU.SaveAs($rutaSIGU)
Write-Host "SIGU guardado: $rutaSIGU" -ForegroundColor Green

# ==============================================================================
# LIBRO 2: CONTROL MANUAL ADMINISTRATIVO
# ==============================================================================

Write-Host ""
Write-Host "Creando archivo Manual..." -ForegroundColor Cyan
$libroMan = $excel.Workbooks.Add()
$hMan = $libroMan.Sheets.Item(1)
$hMan.Name = "Control_Manual"

$hMANh = @("Cedula","Carnet","Nombre","Primer Apellido","Segundo Apellido",
           "Correo Electronico","Telefono","Carrera","Sede","Estado",
           "Periodo","Fecha Matricula","Documentos Entregados","Observaciones",
           "Ultima Sincronizacion","Correo Enviado","Fecha Correo")

for ($h = 0; $h -lt $hMANh.Length; $h++) {
    $hMan.Cells.Item(1,$h+1) = $hMANh[$h]
    $hMan.Cells.Item(1,$h+1).Font.Bold = $true
    $hMan.Cells.Item(1,$h+1).Interior.Color = 1644820   # Azul UTN
    $hMan.Cells.Item(1,$h+1).Font.Color = 16777215
    $hMan.Cells.Item(1,$h+1).Borders.LineStyle = 1
}

# Tomar 85% de registros (el otro 15% seran "nuevos" al cruzar)
$totalParaMan = [math]::Round($estudiantes.Count * 0.85)
$todosIdx = 0..($estudiantes.Count - 1)
$seleccionados = $todosIdx | Get-Random -Count $totalParaMan | Sort-Object

Write-Host "Escribiendo $totalParaMan registros en Manual..." -ForegroundColor Yellow
Write-Host "  ($($totalEst - $totalParaMan) seran NUEVOS en el cruce)" -ForegroundColor Yellow

$fila = 2; $modificados = 0

foreach ($i in $seleccionados) {
    $e = $estudiantes[$i]

    $estMan  = $e.estado
    $sedeMan = $e.sede
    $corrMan = $e.correo
    $obsMan  = $e.obs
    $cEnv    = ""
    $fEnv    = ""

    # 10% con discrepancias para el cruce
    if ((Get-Random -Minimum 0 -Maximum 100) -lt 10) {
        $modificados++
        $tipo = Get-Random -Minimum 0 -Maximum 3
        if ($tipo -eq 0) {
            if    ($estMan -eq "Activo")   { $estMan = "Pendiente" }
            elseif($estMan -eq "Pendiente"){ $estMan = "Activo" }
            $obsMan = "Estado actualizado manualmente"
        } elseif ($tipo -eq 1) {
            $sedesAlt = $sedes | Where-Object { $_ -ne $e.sede }
            $sedeMan  = $sedesAlt[(Get-Random -Minimum 0 -Maximum $sedesAlt.Count)]
            $obsMan   = "Cambio de sede registrado"
        } else {
            $corrMan = $e.nombre.ToLower() + "@gmail.com"
            $obsMan  = "Correo personal registrado"
        }
    }

    # Pendientes: 40% ya recibieron recordatorio
    if ($e.estado -eq "Pendiente" -and (Get-Random -Minimum 0 -Maximum 100) -lt 40) {
        $cEnv = "Si"
        $fEnv = (Get-Date).AddDays(-(Get-Random -Minimum 3 -Maximum 20)).ToString("dd/MM/yyyy")
    }

    $hMan.Cells.Item($fila,1)  = $e.cedula
    $hMan.Cells.Item($fila,2)  = $e.carnet
    $hMan.Cells.Item($fila,3)  = $e.nombre
    $hMan.Cells.Item($fila,4)  = $e.ap1
    $hMan.Cells.Item($fila,5)  = $e.ap2
    $hMan.Cells.Item($fila,6)  = $corrMan
    $hMan.Cells.Item($fila,7)  = $e.tel
    $hMan.Cells.Item($fila,8)  = $e.carrera
    $hMan.Cells.Item($fila,9)  = $sedeMan
    $hMan.Cells.Item($fila,10) = $estMan
    $hMan.Cells.Item($fila,11) = $e.periodo
    $hMan.Cells.Item($fila,12) = $e.fechaMat
    $hMan.Cells.Item($fila,13) = $e.docs
    $hMan.Cells.Item($fila,14) = $obsMan
    $hMan.Cells.Item($fila,15) = (Get-Date).AddDays(-(Get-Random -Minimum 10 -Maximum 30)).ToString("dd/MM/yyyy")
    $hMan.Cells.Item($fila,16) = $cEnv
    $hMan.Cells.Item($fila,17) = $fEnv

    if    ($estMan -eq "Activo")   { $col = 13434828 }
    elseif($estMan -eq "Pendiente"){ $col = 10092543 }
    elseif($estMan -eq "Inactivo") { $col = 13421823 }
    else                           { $col = 12632256 }

    for ($c = 1; $c -le 17; $c++) {
        $hMan.Cells.Item($fila,$c).Interior.Color = $col
        $hMan.Cells.Item($fila,$c).Borders.LineStyle = 1
    }
    $fila++
}

for ($col = 1; $col -le 17; $col++) { $hMan.Columns.Item($col).AutoFit() | Out-Null }
$hMan.Range("A1:Q$($fila-1)").AutoFilter() | Out-Null

$rutaMan = "$carpeta\Datos_Manual\Control_Administrativo_Manual.xlsx"
$libroMan.SaveAs($rutaMan)
Write-Host "Manual guardado: $rutaMan" -ForegroundColor Green

# ==============================================================================
# RESUMEN FINAL
# ==============================================================================

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  SIMULACION COMPLETADA" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "ARCHIVOS:" -ForegroundColor White
Write-Host "  SIGU  : $rutaSIGU" -ForegroundColor Green
Write-Host "  Manual: $rutaMan"  -ForegroundColor Green
Write-Host ""
Write-Host "ESTADISTICAS GENERALES:" -ForegroundColor White
Write-Host "  Total aspirantes    : $totalEst" -ForegroundColor Cyan
Write-Host "  Matriculados        : $totalMat  ($([math]::Round($totalMat/$totalEst*100,1))%)" -ForegroundColor Green
Write-Host "  Pendientes          : $totalPen  ($([math]::Round($totalPen/$totalEst*100,1))%)" -ForegroundColor Yellow
Write-Host "  Inactivos/Retirados : $totalIna  ($([math]::Round($totalIna/$totalEst*100,1))%)" -ForegroundColor Red
Write-Host ""
Write-Host "  En archivo Manual   : $totalParaMan registros" -ForegroundColor White
Write-Host "  Solo en SIGU (nuevos): $($totalEst - $totalParaMan) registros" -ForegroundColor White
Write-Host "  Con discrepancias   : $modificados registros" -ForegroundColor White
Write-Host ""
Write-Host "POR CARRERA:" -ForegroundColor White
foreach ($carrera in $carreras) {
    $aCarr = ($estudiantes | Where-Object { $_.carrera -eq $carrera.nombre }).Count
    $mCarr = ($estudiantes | Where-Object { $_.carrera -eq $carrera.nombre -and $_.estado -eq "Activo" }).Count
    Write-Host ("  {0,-5} {1,-45} Asp:{2,4}  Mat:{3,4}" -f $carrera.codigo, $carrera.nombre, $aCarr, $mCarr) -ForegroundColor Gray
}
Write-Host ""
Write-Host "POR SEDE:" -ForegroundColor White
foreach ($sede in $sedes) {
    $aSede = ($estudiantes | Where-Object { $_.sede -eq $sede }).Count
    $mSede = ($estudiantes | Where-Object { $_.sede -eq $sede -and $_.estado -eq "Activo" }).Count
    Write-Host ("  {0,-30} Asp:{1,4}  Mat:{2,4}" -f $sede, $aSede, $mSede) -ForegroundColor Gray
}
Write-Host ""
Write-Host "Los archivos estan abiertos en Excel." -ForegroundColor Yellow
Write-Host "Presione Enter para finalizar..." -ForegroundColor Yellow
Read-Host
