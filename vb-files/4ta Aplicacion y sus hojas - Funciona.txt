'Esto es para automatizar todos los preservables

Sub DCuartaAplicacion()
    Dim respuesta As VbMsgBoxResult
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim rng As Range
    Dim Criterios As Variant
    Dim i As Variant ' Cambiado a Variant por si el usuario cancela el InputBox

    Set ws1 = Sheets("Columnas Ordenadas")

    Do
MuestreoStock:
        i = InputBox("Indicar que almacenes mostrar en el reporte: " & vbCrLf & "(1) Servibles" & vbCrLf & "(2) Reparables" & vbCrLf & "(3) Cuarentena", "Variable i")
        
        ' Validar si el usuario presionó Cancelar
        If i = "" Then Exit Sub
        
        If Not IsNumeric(i) Then
            MsgBox "Por favor, ingresa un número.", vbCritical
            GoTo MuestreoStock
        ElseIf i > 3 Or i < 1 Then
            MsgBox "¿Valor incorrecto? Por favor, intenta de nuevo.", vbQuestion, "Error"
            GoTo MuestreoStock
        End If
        
        ' Definir criterios y hoja de destino
        Select Case CInt(i)
            Case 1
                Criterios = Array("(002)CALLAO - SERVICIABLES CON CTO", "(009)KITENI SERVICIABLES CON CTO", "(013)MALVINAS SERVICIABLES CON CTO", "(015)NUEVO MUNDO SERVICIABLES CON CTO", "(031)AYACUCHO SERVICIABLES CON CTO", "(101)CALLAO - SERVICIABLES SIN CTO", "(105)KITENI SERVICIABLES SIN CTO", "(107)MALVINAS SERVICIABLES SIN CTO", "(108)NUEVO MUNDO SERVICIABLES SIN CTO", "(124)AYACUCHO SERVICIABLES SIN CTO", "(138)LURIN - SERVICIABLE CON CTO", "(139)LURIN - SERVICIABLE SIN CTO", "(153)CALLAO - BK SERVIBLE SIN COSTO", "(157)NUEVO MUNDO BK - SERVICIABLES SIN CTO", "(164)CALLAO - BELL SERVIBLE CON COSTO", "(169)MALVINAS SERVICIABLES BK - SIN CTO", "(172)MALVINAS SERVICIABLES BELL - CON CTO", "(173)MALVINAS SERVICIABLES BELL - SIN CTO", "(184)CEMAE - SERVICIABLES CON CTO", "(185)CEMAE - SERVICIABLES SIN CTO", "(200)AYACUCHO SERVICIABLES BELL SIN CTO")
                Set ws2 = Sheets("Servibles")
            Case 2
                Criterios = Array("(003)CALLAO - REPARABLES SIN CTO", "(036)AYACUCHO REPARABLE CON CTO", "(119)CALLAO - REPARABLES CON CTO", "(134)LURIN - REPARABLE SIN RECURSO CON COSTO", "(135)LURIN - REPARABLE SIN RECURSO SIN COSTO", "(140)LURIN - REPARABLE CON CTO", "(141)LURIN - REPARABLE SIN CTO", "(160)CALLAO - BK REPARABLE CON COSTO", "(161)CALLAO - BK REPARABLE SIN COSTO", "(176)LURIN - REPARABLE NO APLICABLE CON CTO", "(177)LURIN - REPARABLE NO APLICABLE SIN CTO", "(186)CEMAE - REPARABLES SIN CTO", "(187)CEMAE - REPARABLES CON CTO", "(196)AYACUCHO - REPARABLES SIN CTO", "(208)CALLAO - BELL REPARABLE CON COSTO", "(209)CALLAO - BELL REPARABLE SIN COSTO")
                Set ws2 = Sheets("Reparables")
            Case 3
                Criterios = Array("(094)CALLAO - CUARENTENA CALIDAD C/COSTO", "(095)CALLAO - CUARENTENA CALIDAD S/COSTO", "(214)KITENI CUARENTENA CON CTO", "(215)KITENI CUARENTENA SIN CTO", "(216)AYACUCHO CUARENTENA CON CTO", "(217)AYACUCHO CUARENTENA SIN CTO", "(218)CEMAE - CUARENTENA CON CTO", "(219)CEMAE - CUARENTENA SIN CTO", "(220)NUEVO MUNDO CUARENTENA CON CTO", "(221)NUEVO MUNDO CUARENTENA SIN CTO")
                Set ws2 = Sheets("Cuarentena")
        End Select

        ' Limpiar filtros y definir rango
        If ws1.AutoFilterMode Then ws1.AutoFilterMode = False
        Set rng = ws1.Range("A1").CurrentRegion
        
        ' Aplicar el filtro
        rng.AutoFilter Field:=3, Criteria1:=Criterios, Operator:=xlFilterValues
        
        ' --- SOLUCIÓN AL ERROR 1004 ---
        ' Primero limpiamos el destino
        ws2.Range("A2:M" & ws2.Cells(Rows.Count, 1).End(xlUp).Row + 1).ClearContents
        
        ' Ahora copiamos y pegamos inmediatamente
        ' Solo copiamos las filas visibles (sin el encabezado para no borrarlo luego)
        On Error Resume Next ' Por si no hay datos filtrados
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy
        ws2.Cells(2, 1).PasteSpecial Paste:=xlPasteValues
        On Error GoTo 0
        
        Application.CutCopyMode = False
        ws1.AutoFilterMode = False

        ' Preguntar al usuario si desea repetir
        respuesta = MsgBox("¿Deseas repetir el proceso de filtrado?", _
                           vbQuestion + vbYesNo + vbDefaultButton2, _
                           "Repetir Macro")

    Loop While respuesta = vbYes

    MsgBox "Proceso finalizado.", vbInformation
    
End Sub


Sub CuartaAplicacionServibles()

Dim wbOrigen, wbDestino As Workbook
Dim wsOrigen, wsDestino As Worksheet
Dim rng As Range
Dim respuesta As String

    ' 1. Configuración de nombres de libros y hojas
    
    Set wbOrigen = Workbooks("069. Reporte de vencimiento preservado con flota - Mejorado con Macros.xlsm")
    Set wbDestino = Workbooks("a) Reporte de control de articulos preservados (serviciables).xlsx")
    
    ' 2. Validar que los libros estén abiertos
    If wbOrigen Is Nothing Or wbDestino Is Nothing Then
        MsgBox "Asegúrate de que ambos libros estén abiertos antes de ejecutar la macro.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set wsOrigen = wbOrigen.Sheets("Servibles")
    
respuesta = MsgBox("Iniciar procedimiento?", vbYesNo + vbQuestion, "Inicio")

Select Case respuesta

    Case vbYes
        
        '0) General

        Set rng = wsOrigen.Range("A1").CurrentRegion
        Set wsDestino = wbDestino.Sheets("GENERAL - Servible")
        
        If wsOrigen.FilterMode Then wsOrigen.ShowAllData
        
        wsDestino.Range("TablaGeneralServible").ClearContents
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(4, 1)
        
        wsDestino.Cells(2, 3).Value = Date
        
        '1) Lima
        
        Set rng = wsOrigen.Range("A1").CurrentRegion
        Set wsDestino = wbDestino.Sheets("Lima - Servible")
        
        wsDestino.Range("TablaLimaServible").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Lima"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
                               
        
        '2) Malvinas
        
        Set wsDestino = wbDestino.Sheets("Malvinas - Servible")
        
        wsDestino.Range("TablaMalvinasServible").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Malvinas"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
        
        '3) Ayacucho
        
        Set wsDestino = wbDestino.Sheets("Ayacucho - Servible")
        
        wsDestino.Range("TablaAyacuchoServible").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Ayacucho"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)

        '4) Kiteni
        
        Set wsDestino = wbDestino.Sheets("Kiteni - Servible")
        
        wsDestino.Range("TablaKiteniServible").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Kiteni"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)

        '5) NuevoMundo
        
        Set wsDestino = wbDestino.Sheets("NuevoMundo - Servible")

        wsDestino.Range("TablaNuevoMundoServible").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Nuevo Mundo"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)

        '6) Arequipa
        
        Set wsDestino = wbDestino.Sheets("Arequipa - Servible")
        
        wsDestino.Range("TablaArequipaServible").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("AREQUIPA"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
        
        wsOrigen.ShowAllData
        
        MsgBox "Procedimiento terminado", vbCritical, "Final"
        
    Case vbNo
        
        MsgBox "Procedimiento cancelado", vbCritical, "Final"
        
        Exit Sub
        
    End Select
    
End Sub

Sub CuartaAplicacionReparables()

Dim wbOrigen, wbDestino As Workbook
Dim wsOrigen, wsDestino As Worksheet
Dim rng As Range
Dim respuesta As String

    ' 1. Configuración de nombres de libros y hojas
    
    Set wbOrigen = Workbooks("069. Reporte de vencimiento preservado con flota - Mejorado con Macros.xlsm")
    Set wbDestino = Workbooks("b) Reporte de control de articulos preservados (reparables).xlsx")
    
    ' 2. Validar que los libros estén abiertos
    If wbOrigen Is Nothing Or wbDestino Is Nothing Then
        MsgBox "Asegúrate de que ambos libros estén abiertos antes de ejecutar la macro.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set wsOrigen = wbOrigen.Sheets("Reparables")
    
respuesta = MsgBox("Iniciar procedimiento?", vbYesNo + vbQuestion, "Inicio")

Select Case respuesta

    Case vbYes
        
        '0) General

        Set rng = wsOrigen.Range("A1").CurrentRegion
        Set wsDestino = wbDestino.Sheets("GENERAL - Reparable")
        
        wsDestino.Range("TablaReparableGeneral").ClearContents
        If wsOrigen.FilterMode Then wsOrigen.ShowAllData
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(4, 1)

        wsDestino.Cells(2, 3).Value = Date
                
        '1) Lima
        
        Set rng = wsOrigen.Range("A1").CurrentRegion
        Set wsDestino = wbDestino.Sheets("Lima - Reparable")
        
        wsDestino.Range("TablaReparableLima").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Lima"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
                     
        '2) Arequipa
        
        Set wsDestino = wbDestino.Sheets("Arequipa - Reparable")
        
        wsDestino.Range("TablaReparableArequipa").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("AREQUIPA"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
        
        '3) Ayacucho
        
        Set wsDestino = wbDestino.Sheets("Ayacucho - Reparable")
        
        wsDestino.Range("TablaReparableAyacucho").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Ayacucho"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
        
        wsOrigen.ShowAllData
        
        MsgBox "Procedimiento terminado", vbCritical, "Final"
        
    Case vbNo
        
        MsgBox "Procedimiento cancelado", vbCritical, "Final"
        
        Exit Sub
        
    End Select
    
End Sub

Sub CuartaAplicacionCuarentena()

Dim wbOrigen, wbDestino As Workbook
Dim wsOrigen, wsDestino As Worksheet
Dim rng As Range
Dim respuesta As String

    ' 1. Configuración de nombres de libros y hojas
    
    Set wbOrigen = Workbooks("069. Reporte de vencimiento preservado con flota - Mejorado con Macros.xlsm")
    Set wbDestino = Workbooks("c) Reporte de control de articulos preservados (cuarentena).xlsx")
    
    ' 2. Validar que los libros estén abiertos
    If wbOrigen Is Nothing Or wbDestino Is Nothing Then
        MsgBox "Asegúrate de que ambos libros estén abiertos antes de ejecutar la macro.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set wsOrigen = wbOrigen.Sheets("Cuarentena")
    
respuesta = MsgBox("Iniciar procedimiento?", vbYesNo + vbQuestion, "Inicio")

Select Case respuesta

    Case vbYes
        
        '0) General

        Set rng = wsOrigen.Range("A1").CurrentRegion
        Set wsDestino = wbDestino.Sheets("GENERAL - Cuarentena")
        If wsOrigen.FilterMode Then wsOrigen.ShowAllData
        
        wsDestino.Range("TablaGeneralCuarentena").ClearContents
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(4, 1)

        wsDestino.Cells(2, 3).Value = Date
                
        '1) Lima
        
        Set rng = wsOrigen.Range("A1").CurrentRegion
        Set wsDestino = wbDestino.Sheets("Lima - Cuarentena")
        
        wsDestino.Range("TablaLimaCuarentena").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Lima"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
                     
        '2) Malvinas
        
        Set wsDestino = wbDestino.Sheets("Malvinas - Cuarentena")
        
        wsDestino.Range("TablaMalvinasCuarentena").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Malvinas"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
                
        '3) Ayacucho
        
        Set wsDestino = wbDestino.Sheets("Ayacucho - Cuarentena")
        
        wsDestino.Range("TablaAyacuchoCuarentena").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Ayacucho"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
       
        '4) Kiteni
        
        Set wsDestino = wbDestino.Sheets("Kiteni - Cuarentena")
        
        wsDestino.Range("TablaKiteniCuarentena").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Kiteni"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)

        '5) NuevoMundo
        
        Set wsDestino = wbDestino.Sheets("NuevoMundo - Cuarentena")

        wsDestino.Range("TablaNuevoMundoCuarentena").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("Nuevo Mundo"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)

        '6) Arequipa
        
        Set wsDestino = wbDestino.Sheets("Arequipa - Cuarentena")
        
        wsDestino.Range("TablaArequipaCuarentena").ClearContents
        rng.AutoFilter Field:=2, Criteria1:=Array("AREQUIPA"), Operator:=xlFilterValues
        On Error Resume Next
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsDestino.Cells(5, 1)
        
        wsOrigen.ShowAllData
        
        MsgBox "Procedimiento terminado", vbCritical, "Final"
        
    Case vbNo
        
        MsgBox "Procedimiento cancelado", vbCritical, "Final"
        
        Exit Sub
        
    End Select
    
End Sub


Sub FormatearTablasTodosLosLibros()
    Dim nombreArchivo As String
    Dim wbObjetivo As Workbook
    Dim i As String
    Dim respuesta As String
   
Do

FormateoStock:
        i = InputBox("Indicar que almacenes aplicar el formato: " & vbCrLf & "(1) Servibles" & vbCrLf & "(2) Reparables" & vbCrLf & "(3) Cuarentena", "Variable i")
        
        ' Validar si el usuario presionó Cancelar
        If i = "" Then Exit Sub
        
        If Not IsNumeric(i) Then
            MsgBox "Por favor, ingresa un número.", vbCritical
            GoTo FormateoStock
        ElseIf i > 3 Or i < 1 Then
            MsgBox "¿Valor incorrecto? Por favor, intenta de nuevo.", vbQuestion, "Error"
            GoTo FormateoStock
        End If
    
    ' 1. ¡ASEGÚRATE DE CAMBIAR ESTO AL NOMBRE REAL DE TU ARCHIVO!
    If i = 1 Then
        nombreArchivo = "a) Reporte de control de articulos preservados (serviciables).xlsx"
    ElseIf i = 2 Then
        nombreArchivo = "b) Reporte de control de articulos preservados (reparables).xlsx"
    ElseIf i = 3 Then
        nombreArchivo = "c) Reporte de control de articulos preservados (cuarentena).xlsx"
    End If
    
    On Error Resume Next
    Set wbObjetivo = Workbooks(nombreArchivo)
    On Error GoTo 0
    
    ' 2. Verificar si el libro está abierto
    If wbObjetivo Is Nothing Then
        MsgBox "El archivo '" & nombreArchivo & "' no está abierto.", vbCritical
        Exit Sub
    End If
    
    ' 3. Aplicar el formato buscando en todas las hojas
    ' Asegúrate de que los nombres de las tablas sean exactos
    
    If i = 1 Then
        AplicarFormatoExterno wbObjetivo, "TablaGeneralServible"
        AplicarFormatoExterno wbObjetivo, "TablaLimaServible"
        AplicarFormatoExterno wbObjetivo, "TablaMalvinasServible"
        AplicarFormatoExterno wbObjetivo, "TablaAyacuchoServible"
        AplicarFormatoExterno wbObjetivo, "TablaKiteniServible"
        AplicarFormatoExterno wbObjetivo, "TablaNuevoMundoServible"
        AplicarFormatoExterno wbObjetivo, "TablaArequipaServible"
    ElseIf i = 2 Then
        AplicarFormatoExterno wbObjetivo, "TablaReparableGeneral"
        AplicarFormatoExterno wbObjetivo, "TablaReparableLima"
        AplicarFormatoExterno wbObjetivo, "TablaReparableArequipa"
        AplicarFormatoExterno wbObjetivo, "TablaReparableAyacucho"
    ElseIf i = 3 Then
        AplicarFormatoExterno wbObjetivo, "TablaGeneralCuarentena"
        AplicarFormatoExterno wbObjetivo, "TablaLimaCuarentena"
        AplicarFormatoExterno wbObjetivo, "TablaMalvinasCuarentena"
        AplicarFormatoExterno wbObjetivo, "TablaAyacuchoCuarentena"
        AplicarFormatoExterno wbObjetivo, "TablaKiteniCuarentena"
        AplicarFormatoExterno wbObjetivo, "TablaNuevoMundoCuarentena"
        AplicarFormatoExterno wbObjetivo, "TablaArequipaCuarentena"
    End If
    
        ' Preguntar al usuario si desea repetir
        respuesta = MsgBox("¿Deseas repetir el proceso de formateado?", _
                           vbQuestion + vbYesNo + vbDefaultButton2, _
                           "Repetir Macro")

    Loop While respuesta = vbYes

    MsgBox "Proceso finalizado.", vbInformation
    
End Sub

Sub AplicarFormatoExterno(wb As Workbook, nombreTabla As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tablaEncontrada As Boolean
    
    tablaEncontrada = False
    
    ' Buscamos en cada hoja del libro objetivo
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(nombreTabla)
        On Error GoTo 0
        
        If Not lo Is Nothing Then
            ' Si encontramos la tabla, aplicamos el formato a todo su rango
            With lo.Range
                With .Font
                    .Name = "Calibri"
                    .Size = 11
                End With
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With
            tablaEncontrada = True
            Exit For ' Salimos del bucle ya que la encontramos
        End If
    Next ws
    
    If Not tablaEncontrada Then
        ' Esto aparecerá en la ventana "Inmediato" (Ctrl+G) si falla
        Debug.Print "Error: No se encontró la tabla '" & nombreTabla & "' en " & wb.Name
    End If
End Sub
