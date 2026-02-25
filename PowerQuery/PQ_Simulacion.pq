// ==============================================================================
// POWER QUERY: Simulación Aspirantes vs Matriculados 2026
// ARCHIVO: PQ_Simulacion.pq
// DESCRIPCIÓN: Cruza los datos de aspirantes (hoja Datos_SIGU, estado="Aspirante")
//              con los datos de matriculados confirmados (hoja Datos_Manual,
//              estado="Matriculado") para calcular tasas de conversión,
//              documentos pendientes y distribución por sede/carrera.
//
// INSTRUCCIONES:
// 1. En Excel, vaya a Datos > Obtener Datos > De otras fuentes > Consulta en blanco
// 2. Abra el Editor Avanzado y pegue este código
// 3. Asegúrese de tener las tablas "Datos_SIGU" y "Datos_Manual" definidas
// ==============================================================================

let
    // ============================
    // PASO 1: Cargar fuentes
    // ============================
    DatosSIGU   = Excel.CurrentWorkbook(){[Name="Datos_SIGU"]}[Content],
    DatosManual = Excel.CurrentWorkbook(){[Name="Datos_Manual"]}[Content],

    // Asegurar tipo texto en cédula
    SIGU   = Table.TransformColumnTypes(DatosSIGU,   {{"Cédula", type text}, {"Estado", type text}}),
    Manual = Table.TransformColumnTypes(DatosManual, {{"Cédula", type text}, {"Estado", type text}}),

    // ============================
    // PASO 2: Filtrar aspirantes y matriculados
    // ============================
    Aspirantes   = Table.SelectRows(SIGU,   each Text.Upper(Text.Trim([Estado] ?? "")) = "ASPIRANTE"),
    Matriculados = Table.SelectRows(Manual, each Text.Upper(Text.Trim([Estado] ?? "")) = "MATRICULADO"),

    // ============================
    // PASO 3: LEFT JOIN aspirantes -> matriculados
    // ============================
    Cruce = Table.NestedJoin(
        Aspirantes,   {"Cédula"},
        Matriculados, {"Cédula"},
        "DatosMatriculado",
        JoinKind.LeftOuter
    ),

    // Extraer solo si existe en manual (confirmado = true/false)
    AgregarConfirmado = Table.AddColumn(Cruce, "Confirmado", each
        let
            sub = [DatosMatriculado],
            cnt = if sub = null then 0 else Table.RowCount(sub)
        in
            if cnt > 0 then "Sí" else "No",
        type text
    ),

    // ============================
    // PASO 4: Calcular documentos pendientes
    // ============================
    DocsRequeridos = {"Cédula", "Foto", "Título Secundaria",
                      "Notas Secundaria", "Constancia CCSS", "Declaración Jurada"},

    AgregarDocsPendientes = Table.AddColumn(AgregarConfirmado, "Documentos Pendientes", each
        let
            entregados = if [Documentos Entregados] = null or [Documentos Entregados] = ""
                         then {}
                         else List.Transform(
                                Text.Split([Documentos Entregados], ","),
                                each Text.Trim(_)
                              ),
            faltantes = List.RemoveMatchingItems(
                            DocsRequeridos, entregados, Comparer.OrdinalIgnoreCase
                        )
        in
            Text.Combine(faltantes, ", "),
        type text
    ),

    AgregarDocsCompletos = Table.AddColumn(AgregarDocsPendientes, "Docs Completos", each
        if [Documentos Pendientes] = "" or [Documentos Pendientes] = null then "Sí" else "No",
        type text
    ),

    // ============================
    // PASO 5: Seleccionar columnas relevantes
    // ============================
    Resultado = Table.SelectColumns(AgregarDocsCompletos, {
        "Cédula", "Carnet", "Nombre", "Primer Apellido", "Segundo Apellido",
        "Sede", "Carrera", "Estado", "Confirmado",
        "Documentos Entregados", "Documentos Pendientes", "Docs Completos"
    }),

    // ============================
    // PASO 6: Ordenar — matriculados primero, luego por apellido
    // ============================
    Ordenado = Table.Sort(Resultado, {
        {"Confirmado", Order.Descending},
        {"Primer Apellido", Order.Ascending},
        {"Nombre", Order.Ascending}
    })

in
    Ordenado
