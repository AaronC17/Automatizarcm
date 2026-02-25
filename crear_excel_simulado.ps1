# ==============================================================================
# Script para crear archivos Excel simulados de matrícula UTN
# Basado en estructura: Aspirantes (total estudiantes) vs Matriculados (confirmados)
# Universidad Técnica Nacional - Costa Rica
# Fecha: Febrero 2026
# ==============================================================================

$carpeta = "c:\Users\LINC\Desktop\cda"

# Crear objeto de Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

# ==============================================================================
# DATOS BASE DE LA UTN
# ==============================================================================

# Carreras ofertadas por la UTN
$carreras = @(
    @{ nombre = "Ingeniería en Tecnologías de Información"; codigo = "ITI"; area = "Ingeniería" },
    @{ nombre = "Administración de Empresas"; codigo = "ADE"; area = "Ciencias Económicas" },
    @{ nombre = "Contaduría Pública"; codigo = "CPU"; area = "Ciencias Económicas" },
    @{ nombre = "Gestión del Turismo Sostenible"; codigo = "GTS"; area = "Turismo" },
    @{ nombre = "Ingeniería Electromecánica Industrial"; codigo = "IEI"; area = "Ingeniería" },
    @{ nombre = "Producción Industrial"; codigo = "PIN"; area = "Ingeniería" },
    @{ nombre = "Salud Ocupacional"; codigo = "SOC"; area = "Salud" },
    @{ nombre = "Asistencia Administrativa"; codigo = "AAD"; area = "Ciencias Económicas" },
    @{ nombre = "Inglés como Lengua Extranjera"; codigo = "ILE"; area = "Educación" },
    @{ nombre = "Agronomía"; codigo = "AGR"; area = "Ciencias Agrarias" }
)

# Sedes de la UTN
$sedes = @("Sede Central Alajuela", "Sede Pacífico", "Sede Guanacaste", "Sede Atenas", "Sede San Carlos")

# Distribución de aspirantes por carrera (simulación realista)
# aspirantes = total que aplicaron, matriculados = los confirmados
$distribucionCarreras = @(
    @{ carrera = 0; aspirantesPorSede = @(85, 42, 38, 20, 30) },   # ITI - muy popular
    @{ carrera = 1; aspirantesPorSede = @(70, 55, 45, 35, 40) },   # ADE
    @{ carrera = 2; aspirantesPorSede = @(40, 30, 25, 18, 22) },   # CPU
    @{ carrera = 3; aspirantesPorSede = @(25, 45, 35, 10, 15) },   # GTS - más en Pacífico
    @{ carrera = 4; aspirantesPorSede = @(50, 20, 15, 25, 18) },   # IEI
    @{ carrera = 5; aspirantesPorSede = @(35, 15, 12, 20, 10) },   # PIN
    @{ carrera = 6; aspirantesPorSede = @(30, 25, 20, 12, 15) },   # SOC
    @{ carrera = 7; aspirantesPorSede = @(45, 35, 30, 25, 28) },   # AAD
    @{ carrera = 8; aspirantesPorSede = @(20, 15, 18, 8, 10) },    # ILE
    @{ carrera = 9; aspirantesPorSede = @(15, 10, 12, 30, 35) }    # AGR - más en Atenas/San Carlos
)

# Porcentaje de confirmación (matriculados) por carrera
$tasaConfirmacion = @(0.72, 0.78, 0.80, 0.65, 0.70, 0.75, 0.82, 0.77, 0.60, 0.68)

# Nombres costarricenses
$nombresM = @("Jose", "Carlos", "Luis", "Juan", "Miguel", "Daniel", "Andres", "Diego", "David", "Kevin",
              "Bryan", "Esteban", "Fabian", "Gabriel", "Alejandro", "Roberto", "Francisco", "Fernando",
              "Marco", "Javier", "Oscar", "Adrian", "Cristian", "Eduardo", "Sebastian", "Mauricio",
              "Rodrigo", "Alvaro", "Pablo", "Hector", "Ricardo", "Gerardo", "Randall", "Jeffry", "Josue")

$nombresF = @("Maria", "Ana", "Sofia", "Laura", "Andrea", "Carolina", "Daniela", "Valeria", "Natalia",
              "Gabriela", "Paola", "Katherine", "Stephanie", "Melissa", "Monica", "Rebeca", "Adriana",
              "Lucia", "Elena", "Tatiana", "Priscilla", "Diana", "Wendy", "Karla", "Fernanda", "Isabel",
              "Marcela", "Yuliana", "Kimberly", "Genesis", "Alison", "Pamela", "Viviana", "Nicole")

$apellidos = @("Rodriguez", "Jimenez", "Mora", "Hernandez", "Vargas", "Solis", "Chaves", "Castro",
               "Arias", "Rojas", "Cordero", "Monge", "Gonzalez", "Lopez", "Ramirez", "Perez",
               "Retana", "Urena", "Villalobos", "Brenes", "Ugalde", "Sanchez", "Vega", "Salazar",
               "Quiros", "Calderon", "Araya", "Campos", "Valverde", "Gutierrez", "Barrantes",
               "Segura", "Nunez", "Madrigal", "Camacho", "Fallas", "Bonilla", "Zamora",
               "Alvarado", "Quesada", "Vindas", "Aguilar", "Montero", "Picado", "Cespedes",
               "Chavarria", "Murillo", "Gamboa", "Leon", "Acuna")

# Documentos requeridos para matricula
$documentosRequeridos = @("Cedula de identidad", "Foto tamano pasaporte", "Titulo de secundaria",
                          "Notas de secundaria", "Constancia CCSS", "Declaracion jurada")

# Periodos
$periodoActual = "I-2026"

# ==============================================================================
# FUNCIONES UTILITARIAS
# ==============================================================================

function Get-CedulaCR {
    # Genera una cédula costarricense simulada (formato: X-XXXX-XXXX)
    $provincia = Get-Random -Minimum 1 -Maximum 8
    $tomo = Get-Random -Minimum 100 -Maximum 2000
    $asiento = Get-Random -Minimum 100 -Maximum 1000
    return "$provincia-$('{0:D4}' -f $tomo)-$('{0:D4}' -f $asiento)"
}

function Get-CarnetUTN {
    param([int]$indice)
    $anio = Get-Random -Minimum 2022 -Maximum 2027
    return "UTN-$anio-$('{0:D5}' -f ($indice + 10000))"
}

function Get-NombreCompleto {
    $esMujer = (Get-Random -Minimum 0 -Maximum 2) -eq 1
    if ($esMujer) {
        $nombre = $nombresF[(Get-Random -Minimum 0 -Maximum $nombresF.Length)]
    } else {
        $nombre = $nombresM[(Get-Random -Minimum 0 -Maximum $nombresM.Length)]
    }
    $ap1 = $apellidos[(Get-Random -Minimum 0 -Maximum $apellidos.Length)]
    $ap2 = $apellidos[(Get-Random -Minimum 0 -Maximum $apellidos.Length)]
    while ($ap2 -eq $ap1) {
        $ap2 = $apellidos[(Get-Random -Minimum 0 -Maximum $apellidos.Length)]
    }
    return @{ nombre = $nombre; apellido1 = $ap1; apellido2 = $ap2; esMujer = $esMujer }
}

function Get-CorreoUTN {
    param([string]$nombre, [string]$apellido1)
    $n = $nombre.ToLower()
    $a = $apellido1.ToLower()
    $num = Get-Random -Minimum 1 -Maximum 99
    return "$($n.Substring(0,1))$($a)$num@est.utn.ac.cr"
}

function Get-TelefonoCR {
    $prefijos = @("6", "7", "8")
    $prefijo = $prefijos[(Get-Random -Minimum 0 -Maximum $prefijos.Length)]
    $num = Get-Random -Minimum 1000000 -Maximum 9999999
    return "+506 $prefijo$($num.ToString().Substring(0,3))-$($num.ToString().Substring(3,4))"
}

function Get-DocumentosEntregados {
    param([string]$estado)
    
    if ($estado -eq "Activo") {
        # Matriculados: la mayoría tiene todos los documentos
        $probabilidadCompleto = 0.75
        if ((Get-Random -Minimum 0 -Maximum 100) -lt ($probabilidadCompleto * 100)) {
            return ($documentosRequeridos -join ", ")
        } else {
            $cantFaltantes = Get-Random -Minimum 1 -Maximum 3
            $entregados = $documentosRequeridos | Get-Random -Count ($documentosRequeridos.Length - $cantFaltantes)
            return ($entregados -join ", ")
        }
    }
    elseif ($estado -eq "Pendiente") {
        # Pendientes: faltan documentos
        $cantEntregados = Get-Random -Minimum 1 -Maximum 4
        $entregados = $documentosRequeridos | Get-Random -Count $cantEntregados
        return ($entregados -join ", ")
    }
    else {
        # Inactivos/retirados: pocos documentos
        $cantEntregados = Get-Random -Minimum 0 -Maximum 3
        if ($cantEntregados -eq 0) { return "Ninguno" }
        $entregados = $documentosRequeridos | Get-Random -Count $cantEntregados
        return ($entregados -join ", ")
    }
}

function Get-DocumentosFaltantes {
    param([string]$entregados)
    if ($entregados -eq ($documentosRequeridos -join ", ")) { return "Ninguno" }
    if ($entregados -eq "Ninguno") { return ($documentosRequeridos -join ", ") }
    
    $listaEntregados = $entregados -split ", "
    $faltantes = $documentosRequeridos | Where-Object { $_ -notin $listaEntregados }
    if ($faltantes.Count -eq 0) { return "Ninguno" }
    return ($faltantes -join ", ")
}

# ==============================================================================
# GENERAR LISTA COMPLETA DE ESTUDIANTES
# ==============================================================================

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " GENERADOR DE DATOS SIMULADOS UTN" -ForegroundColor Cyan
Write-Host " Aspirantes vs Matriculados" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$estudiantes = @()
$indiceGlobal = 0

foreach ($dist in $distribucionCarreras) {
    $carrera = $carreras[$dist.carrera]
    $tasa = $tasaConfirmacion[$dist.carrera]
    
    for ($s = 0; $s -lt $sedes.Length; $s++) {
        $totalAspirantesSede = $dist.aspirantesPorSede[$s]
        $totalMatriculados = [math]::Round($totalAspirantesSede * $tasa)
        $totalPendientes = [math]::Round(($totalAspirantesSede - $totalMatriculados) * 0.6)
        $totalInactivos = $totalAspirantesSede - $totalMatriculados - $totalPendientes
        
        for ($i = 0; $i -lt $totalAspirantesSede; $i++) {
            $indiceGlobal++
            $persona = Get-NombreCompleto
            $cedula = Get-CedulaCR
            $carnet = Get-CarnetUTN $indiceGlobal
            
            # Determinar estado
            if ($i -lt $totalMatriculados) {
                $estado = "Activo"
                $fechaMatricula = (Get-Date).AddDays(-(Get-Random -Minimum 5 -Maximum 45)).ToString("dd/MM/yyyy")
            }
            elseif ($i -lt ($totalMatriculados + $totalPendientes)) {
                $estado = "Pendiente"
                $fechaMatricula = ""
            }
            else {
                $estadosInactivos = @("Inactivo", "Retirado")
                $estado = $estadosInactivos[(Get-Random -Minimum 0 -Maximum 2)]
                $fechaMatricula = ""
            }
            
            $docsEntregados = Get-DocumentosEntregados $estado
            $docsFaltantes = Get-DocumentosFaltantes $docsEntregados
            
            if ($estado -eq "Activo") {
                if ($docsFaltantes -ne "Ninguno") { $observacion = "Matriculado - documentos pendientes" } else { $observacion = "Matricula completa" }
            } elseif ($estado -eq "Pendiente") {
                $observacion = "Aspirante pendiente de completar proceso. Falta: $docsFaltantes"
            } elseif ($estado -eq "Inactivo") {
                $observacion = "No completo proceso de matricula"
            } else {
                $observacion = "Desistio del proceso de admision"
            }
            
            $estudiantes += @{
                indice       = $indiceGlobal
                cedula       = $cedula
                carnet       = $carnet
                nombre       = $persona.nombre
                apellido1    = $persona.apellido1
                apellido2    = $persona.apellido2
                correo       = Get-CorreoUTN $persona.nombre $persona.apellido1
                telefono     = Get-TelefonoCR
                carrera      = $carrera.nombre
                codigoCarrera = $carrera.codigo
                area         = $carrera.area
                sede         = $sedes[$s]
                estado       = $estado
                periodo      = $periodoActual
                fechaMatricula = $fechaMatricula
                documentos   = $docsEntregados
                docsFaltantes = $docsFaltantes
                observaciones = $observacion
            }
        }
    }
}

$totalEstudiantes = $estudiantes.Count
$totalMatriculados = ($estudiantes | Where-Object { $_.estado -eq "Activo" }).Count
$totalPendientes = ($estudiantes | Where-Object { $_.estado -eq "Pendiente" }).Count
$totalInactivos = ($estudiantes | Where-Object { $_.estado -eq "Inactivo" -or $_.estado -eq "Retirado" }).Count

Write-Host "Total aspirantes generados: $totalEstudiantes" -ForegroundColor Green
Write-Host "  Matriculados (confirmados): $totalMatriculados" -ForegroundColor Green
Write-Host "  Pendientes: $totalPendientes" -ForegroundColor Yellow
Write-Host "  Inactivos/Retirados: $totalInactivos" -ForegroundColor Red
Write-Host ""

# ==============================================================================
# CREAR LIBRO 1: ARCHIVO SIGU (Datos_SIGU)
# ==============================================================================

Write-Host "Creando archivo SIGU..." -ForegroundColor Cyan

$libroSIGU = $excel.Workbooks.Add()

# --- HOJA 1: Resumen por Carrera (Aspirantes vs Matriculados) ---
$hojaResumen = $libroSIGU.Sheets.Item(1)
$hojaResumen.Name = "Resumen_Aspirantes"

# Título
$hojaResumen.Range("A1:H1").Merge()
$hojaResumen.Cells.Item(1, 1) = "UNIVERSIDAD TECNICA NACIONAL - Resumen de Aspirantes y Matriculados - $periodoActual"
$hojaResumen.Cells.Item(1, 1).Font.Bold = $true
$hojaResumen.Cells.Item(1, 1).Font.Size = 14
$hojaResumen.Cells.Item(1, 1).Font.Color = 16777215
$hojaResumen.Cells.Item(1, 1).Interior.Color = 8388608  # Azul oscuro

$hojaResumen.Range("A2:H2").Merge()
$hojaResumen.Cells.Item(2, 2) = "Fecha de generación: $(Get-Date -Format 'dd/MM/yyyy HH:mm')"

# Headers
$headersResumen = @("Carrera", "Código", "Área", "Sede", "Aspirantes", "Matriculados", "Pendientes", "% Confirmación")
$fila = 4

for ($h = 0; $h -lt $headersResumen.Length; $h++) {
    $hojaResumen.Cells.Item($fila, $h + 1) = $headersResumen[$h]
    $hojaResumen.Cells.Item($fila, $h + 1).Font.Bold = $true
    $hojaResumen.Cells.Item($fila, $h + 1).Interior.Color = 5287936  # Verde oscuro
    $hojaResumen.Cells.Item($fila, $h + 1).Font.Color = 16777215
    $hojaResumen.Cells.Item($fila, $h + 1).Borders.LineStyle = 1
}

$fila = 5
$totalGeneralAsp = 0
$totalGeneralMat = 0
$totalGeneralPen = 0

foreach ($dist in $distribucionCarreras) {
    $carrera = $carreras[$dist.carrera]
    
    for ($s = 0; $s -lt $sedes.Length; $s++) {
        $asp = $dist.aspirantesPorSede[$s]
        $mat = [math]::Round($asp * $tasaConfirmacion[$dist.carrera])
        $pen = $asp - $mat
        $pct = [math]::Round(($mat / $asp) * 100, 1)
        
        $totalGeneralAsp += $asp
        $totalGeneralMat += $mat
        $totalGeneralPen += $pen
        
        $hojaResumen.Cells.Item($fila, 1) = $carrera.nombre
        $hojaResumen.Cells.Item($fila, 2) = $carrera.codigo
        $hojaResumen.Cells.Item($fila, 3) = $carrera.area
        $hojaResumen.Cells.Item($fila, 4) = $sedes[$s]
        $hojaResumen.Cells.Item($fila, 5) = $asp
        $hojaResumen.Cells.Item($fila, 6) = $mat
        $hojaResumen.Cells.Item($fila, 7) = $pen
        $hojaResumen.Cells.Item($fila, 8) = "$pct%"
        
        # Colorear fila según porcentaje
        if ($pct -ge 80) {
            $hojaResumen.Cells.Item($fila, 8).Interior.Color = 5287936  # Verde
            $hojaResumen.Cells.Item($fila, 8).Font.Color = 16777215
        } elseif ($pct -ge 65) {
            $hojaResumen.Cells.Item($fila, 8).Interior.Color = 65535    # Amarillo
        } else {
            $hojaResumen.Cells.Item($fila, 8).Interior.Color = 255      # Rojo
            $hojaResumen.Cells.Item($fila, 8).Font.Color = 16777215
        }
        
        # Bordes
        for ($c = 1; $c -le 8; $c++) {
            $hojaResumen.Cells.Item($fila, $c).Borders.LineStyle = 1
        }
        
        $fila++
    }
}

# Fila total
$hojaResumen.Cells.Item($fila, 1) = "TOTAL GENERAL"
$hojaResumen.Cells.Item($fila, 1).Font.Bold = $true
$hojaResumen.Cells.Item($fila, 5) = $totalGeneralAsp
$hojaResumen.Cells.Item($fila, 5).Font.Bold = $true
$hojaResumen.Cells.Item($fila, 6) = $totalGeneralMat
$hojaResumen.Cells.Item($fila, 6).Font.Bold = $true
$hojaResumen.Cells.Item($fila, 7) = $totalGeneralPen
$hojaResumen.Cells.Item($fila, 7).Font.Bold = $true
$pctGeneral = [math]::Round(($totalGeneralMat / $totalGeneralAsp) * 100, 1)
$hojaResumen.Cells.Item($fila, 8) = "$pctGeneral%"
$hojaResumen.Cells.Item($fila, 8).Font.Bold = $true
for ($c = 1; $c -le 8; $c++) {
    $hojaResumen.Cells.Item($fila, $c).Interior.Color = 12632256
    $hojaResumen.Cells.Item($fila, $c).Borders.LineStyle = 1
}

# Ajustar ancho
for ($c = 1; $c -le 8; $c++) {
    $hojaResumen.Columns.Item($c).AutoFit() | Out-Null
}

# --- HOJA 2: Detalle de Estudiantes (formato SIGU) ---
$hojaDetalle = $libroSIGU.Sheets.Add([System.Reflection.Missing]::Value, $hojaResumen)
$hojaDetalle.Name = "Datos_SIGU"

$headersSIGU = @("Cédula", "Carnet", "Nombre", "Primer Apellido", "Segundo Apellido", 
                  "Correo Electrónico", "Teléfono", "Carrera", "Sede", "Estado", 
                  "Período", "Fecha Matrícula", "Documentos Entregados", "Observaciones")

# Encabezados
for ($h = 0; $h -lt $headersSIGU.Length; $h++) {
    $hojaDetalle.Cells.Item(1, $h + 1) = $headersSIGU[$h]
    $hojaDetalle.Cells.Item(1, $h + 1).Font.Bold = $true
    $hojaDetalle.Cells.Item(1, $h + 1).Interior.Color = 5287936
    $hojaDetalle.Cells.Item(1, $h + 1).Font.Color = 16777215
    $hojaDetalle.Cells.Item(1, $h + 1).Borders.LineStyle = 1
}

# Datos de todos los estudiantes (aspirantes)
Write-Host "Escribiendo $totalEstudiantes registros en hoja SIGU..." -ForegroundColor Yellow
$fila = 2
$contador = 0
foreach ($est in $estudiantes) {
    $contador++
    if ($contador % 100 -eq 0) {
        Write-Host "  Progreso: $contador / $totalEstudiantes" -ForegroundColor Gray
    }
    
    $hojaDetalle.Cells.Item($fila, 1) = $est.cedula
    $hojaDetalle.Cells.Item($fila, 2) = $est.carnet
    $hojaDetalle.Cells.Item($fila, 3) = $est.nombre
    $hojaDetalle.Cells.Item($fila, 4) = $est.apellido1
    $hojaDetalle.Cells.Item($fila, 5) = $est.apellido2
    $hojaDetalle.Cells.Item($fila, 6) = $est.correo
    $hojaDetalle.Cells.Item($fila, 7) = $est.telefono
    $hojaDetalle.Cells.Item($fila, 8) = $est.carrera
    $hojaDetalle.Cells.Item($fila, 9) = $est.sede
    $hojaDetalle.Cells.Item($fila, 10) = $est.estado
    $hojaDetalle.Cells.Item($fila, 11) = $est.periodo
    $hojaDetalle.Cells.Item($fila, 12) = $est.fechaMatricula
    $hojaDetalle.Cells.Item($fila, 13) = $est.documentos
    $hojaDetalle.Cells.Item($fila, 14) = $est.observaciones
    
    # Colorear segun estado
    if ($est.estado -eq "Activo")         { $colorFila = 13434828 }  # Verde claro
    elseif ($est.estado -eq "Pendiente")  { $colorFila = 10092543 }  # Amarillo claro
    elseif ($est.estado -eq "Inactivo")   { $colorFila = 13421823 }  # Rojo claro
    else                                  { $colorFila = 12632256 }  # Gris claro
    for ($c = 1; $c -le 14; $c++) {
        $hojaDetalle.Cells.Item($fila, $c).Interior.Color = $colorFila
        $hojaDetalle.Cells.Item($fila, $c).Borders.LineStyle = 1
    }
    
    $fila++
}

# Ajustar ancho
for ($c = 1; $c -le 14; $c++) {
    $hojaDetalle.Columns.Item($c).AutoFit() | Out-Null
}

# Filtro automático
$rangoFiltro = $hojaDetalle.Range("A1:N$($fila - 1)")
$rangoFiltro.AutoFilter() | Out-Null

# --- HOJA 3: Resumen por Sede ---
$hojaSede = $libroSIGU.Sheets.Add([System.Reflection.Missing]::Value, $hojaDetalle)
$hojaSede.Name = "Resumen_Sedes"

$hojaSede.Range("A1:D1").Merge()
$hojaSede.Cells.Item(1, 1) = "Resumen por Sede - $periodoActual"
$hojaSede.Cells.Item(1, 1).Font.Bold = $true
$hojaSede.Cells.Item(1, 1).Font.Size = 12

$headersSede = @("Sede", "Total Aspirantes", "Matriculados", "% Confirmación")
for ($h = 0; $h -lt $headersSede.Length; $h++) {
    $hojaSede.Cells.Item(3, $h + 1) = $headersSede[$h]
    $hojaSede.Cells.Item(3, $h + 1).Font.Bold = $true
    $hojaSede.Cells.Item(3, $h + 1).Interior.Color = 5287936
    $hojaSede.Cells.Item(3, $h + 1).Font.Color = 16777215
}

$fila = 4
foreach ($sede in $sedes) {
    $aspSede = ($estudiantes | Where-Object { $_.sede -eq $sede }).Count
    $matSede = ($estudiantes | Where-Object { $_.sede -eq $sede -and $_.estado -eq "Activo" }).Count
    $pctSede = if ($aspSede -gt 0) { [math]::Round(($matSede / $aspSede) * 100, 1) } else { 0 }
    
    $hojaSede.Cells.Item($fila, 1) = $sede
    $hojaSede.Cells.Item($fila, 2) = $aspSede
    $hojaSede.Cells.Item($fila, 3) = $matSede
    $hojaSede.Cells.Item($fila, 4) = "$pctSede%"
    
    for ($c = 1; $c -le 4; $c++) {
        $hojaSede.Cells.Item($fila, $c).Borders.LineStyle = 1
    }
    $fila++
}

for ($c = 1; $c -le 4; $c++) {
    $hojaSede.Columns.Item($c).AutoFit() | Out-Null
}

# Guardar archivo SIGU
$archivoSIGU = "$carpeta\Datos_SIGU\SIGU_Exportacion_$(Get-Date -Format 'yyyyMMdd').xlsx"
$libroSIGU.SaveAs($archivoSIGU)
Write-Host "Archivo SIGU guardado: $archivoSIGU" -ForegroundColor Green

# ==============================================================================
# CREAR LIBRO 2: ARCHIVO MANUAL ADMINISTRATIVO (Datos_Manual)
# ==============================================================================

Write-Host ""
Write-Host "Creando archivo Manual Administrativo..." -ForegroundColor Cyan

$libroManual = $excel.Workbooks.Add()
$hojaManual = $libroManual.Sheets.Item(1)
$hojaManual.Name = "Control_Manual"

$headersManual = @("Cédula", "Carnet", "Nombre", "Primer Apellido", "Segundo Apellido",
                   "Correo Electrónico", "Teléfono", "Carrera", "Sede", "Estado",
                   "Período", "Fecha Matrícula", "Documentos Entregados", "Observaciones",
                   "Última Sincronización", "Correo Enviado", "Fecha Correo")

# Encabezados
for ($h = 0; $h -lt $headersManual.Length; $h++) {
    $hojaManual.Cells.Item(1, $h + 1) = $headersManual[$h]
    $hojaManual.Cells.Item(1, $h + 1).Font.Bold = $true
    $hojaManual.Cells.Item(1, $h + 1).Interior.Color = 16711680  # Azul
    $hojaManual.Cells.Item(1, $h + 1).Font.Color = 16777215
    $hojaManual.Cells.Item(1, $h + 1).Borders.LineStyle = 1
}

# Para el archivo manual, simular que está ligeramente desactualizado:
# - Incluir solo ~85% de los estudiantes (algunos son "nuevos" en SIGU)
# - Algunos con estado diferente (para detectar cambios)
# - Algunos con sede diferente

$totalParaManual = [math]::Round($estudiantes.Count * 0.85)

# Tomar un subconjunto aleatorio
$todosIndices = 0..($estudiantes.Count - 1)
$indicesSeleccionados = $todosIndices | Get-Random -Count $totalParaManual | Sort-Object

Write-Host "Escribiendo $totalParaManual registros en archivo Manual (de $totalEstudiantes totales)..." -ForegroundColor Yellow
Write-Host "  ($($totalEstudiantes - $totalParaManual) estudiantes seran detectados como NUEVOS en el cruce)" -ForegroundColor Yellow

$fila = 2
$modificados = 0

foreach ($idx in $indicesSeleccionados) {
    $est = $estudiantes[$idx]
    
    # Simular desactualización (10% de registros con datos diferentes)
    $estadoManual = $est.estado
    $sedeManual = $est.sede
    $correoManual = $est.correo
    $observManual = $est.observaciones
    $correoEnviado = ""
    $fechaCorreo = ""
    
    if ((Get-Random -Minimum 0 -Maximum 100) -lt 10) {
        $modificados++
        $cambio = Get-Random -Minimum 0 -Maximum 3
        switch ($cambio) {
            0 { 
                # Cambio de estado
                if ($estadoManual -eq "Activo") { $estadoManual = "Pendiente" }
                elseif ($estadoManual -eq "Pendiente") { $estadoManual = "Activo" }
                $observManual = "Estado actualizado manualmente"
            }
            1 {
                # Cambio de sede
                $sedesDisponibles = $sedes | Where-Object { $_ -ne $est.sede }
                $sedeManual = $sedesDisponibles[(Get-Random -Minimum 0 -Maximum $sedesDisponibles.Count)]
                $observManual = "Cambio de sede registrado"
            }
            2 {
                # Correo diferente
                $correoManual = ($est.nombre.ToLower().Replace("á","a").Replace("é","e").Replace("í","i").Replace("ó","o").Replace("ú","u")) + "@gmail.com"
                $observManual = "Correo personal registrado"
            }
        }
    }
    
    # Simular que algunos ya recibieron correo de recordatorio
    if ($est.estado -eq "Pendiente" -and (Get-Random -Minimum 0 -Maximum 100) -lt 40) {
        $correoEnviado = "Sí"
        $fechaCorreo = (Get-Date).AddDays(-(Get-Random -Minimum 3 -Maximum 20)).ToString("dd/MM/yyyy")
    }
    
    $hojaManual.Cells.Item($fila, 1) = $est.cedula
    $hojaManual.Cells.Item($fila, 2) = $est.carnet
    $hojaManual.Cells.Item($fila, 3) = $est.nombre
    $hojaManual.Cells.Item($fila, 4) = $est.apellido1
    $hojaManual.Cells.Item($fila, 5) = $est.apellido2
    $hojaManual.Cells.Item($fila, 6) = $correoManual
    $hojaManual.Cells.Item($fila, 7) = $est.telefono
    $hojaManual.Cells.Item($fila, 8) = $est.carrera
    $hojaManual.Cells.Item($fila, 9) = $sedeManual
    $hojaManual.Cells.Item($fila, 10) = $estadoManual
    $hojaManual.Cells.Item($fila, 11) = $est.periodo
    $hojaManual.Cells.Item($fila, 12) = $est.fechaMatricula
    $hojaManual.Cells.Item($fila, 13) = $est.documentos
    $hojaManual.Cells.Item($fila, 14) = $observManual
    $hojaManual.Cells.Item($fila, 15) = (Get-Date).AddDays(-(Get-Random -Minimum 10 -Maximum 30)).ToString("dd/MM/yyyy")
    $hojaManual.Cells.Item($fila, 16) = $correoEnviado
    $hojaManual.Cells.Item($fila, 17) = $fechaCorreo
    
    # Colorear segun estado
    if ($estadoManual -eq "Activo")         { $colorFila = 13434828 }
    elseif ($estadoManual -eq "Pendiente")  { $colorFila = 10092543 }
    elseif ($estadoManual -eq "Inactivo")   { $colorFila = 13421823 }
    else                                    { $colorFila = 12632256 }
    for ($c = 1; $c -le 17; $c++) {
        $hojaManual.Cells.Item($fila, $c).Interior.Color = $colorFila
        $hojaManual.Cells.Item($fila, $c).Borders.LineStyle = 1
    }
    
    $fila++
}

# Ajustar ancho
for ($c = 1; $c -le 17; $c++) {
    $hojaManual.Columns.Item($c).AutoFit() | Out-Null
}

# Filtro automático
$rangoFiltroManual = $hojaManual.Range("A1:Q$($fila - 1)")
$rangoFiltroManual.AutoFilter() | Out-Null

Write-Host "  Registros con datos modificados (para cruce): $modificados" -ForegroundColor Yellow

# Guardar archivo Manual
$archivoManual = "$carpeta\Datos_Manual\Control_Administrativo_Manual.xlsx"
$libroManual.SaveAs($archivoManual)
Write-Host "Archivo Manual guardado: $archivoManual" -ForegroundColor Green

# ==============================================================================
# RESUMEN FINAL
# ==============================================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " SIMULACIÓN COMPLETADA" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "ARCHIVOS GENERADOS:" -ForegroundColor White
Write-Host "  1. SIGU: $archivoSIGU" -ForegroundColor Green
Write-Host "  2. Manual: $archivoManual" -ForegroundColor Green
Write-Host ""
Write-Host "ESTADÍSTICAS:" -ForegroundColor White
Write-Host "  Total Aspirantes: $totalEstudiantes" -ForegroundColor Cyan
    Write-Host "  Matriculados (confirmados): $totalMatriculados" -ForegroundColor Green
    Write-Host "  Pendientes: $totalPendientes" -ForegroundColor Yellow
    Write-Host "  Inactivos/Retirados: $totalInactivos" -ForegroundColor Red
Write-Host ""
Write-Host "  Registros en archivo manual: $totalParaManual" -ForegroundColor White
Write-Host "  Estudiantes 'nuevos' (solo en SIGU): $($totalEstudiantes - $totalParaManual)" -ForegroundColor White
Write-Host "  Registros modificados (discrepancias): $modificados" -ForegroundColor White
Write-Host ""
Write-Host "CARRERAS:" -ForegroundColor White
foreach ($carrera in $carreras) {
    $aspCarrera = ($estudiantes | Where-Object { $_.carrera -eq $carrera.nombre }).Count
    $matCarrera = ($estudiantes | Where-Object { $_.carrera -eq $carrera.nombre -and $_.estado -eq "Activo" }).Count
    Write-Host "  $($carrera.codigo) - $($carrera.nombre): $aspCarrera aspirantes, $matCarrera matriculados" -ForegroundColor Gray
}
Write-Host ""
Write-Host "SEDES:" -ForegroundColor White
foreach ($sede in $sedes) {
    $aspSede = ($estudiantes | Where-Object { $_.sede -eq $sede }).Count
    $matSede = ($estudiantes | Where-Object { $_.sede -eq $sede -and $_.estado -eq "Activo" }).Count
    Write-Host "  $sede : $aspSede aspirantes, $matSede matriculados" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Los archivos ya están abiertos en Excel." -ForegroundColor Yellow
Write-Host "Presione Enter para finalizar..." -ForegroundColor Yellow
Read-Host
