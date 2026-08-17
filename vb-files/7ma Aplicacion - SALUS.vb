Sub SeptimaAplicacion()

Dim wbOrigen, wbDestino As Workbook
Dim wsOrigen, wsDestino As Worksheet
Dim rng As Range
Dim respuesta As String

'Definimos variables para los componentes usados

Dim PartNumber_Usado, Descripcion_Usado, Cantidad_Usada As String


    ' 1. Configuración de nombres de libros y hojas
    
    Set wbOrigen = Workbooks("Insumos Orden Trabajo")
    Set wbDestino = Workbooks("Helisur Parts Feedback Sheet MS TEAMS")
    
    ' 2. Validar que los libros estén abiertos
    If wbOrigen Is Nothing Or wbDestino Is Nothing Then
        MsgBox "Asegúrate de que ambos libros estén abiertos antes de ejecutar la macro.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set wsOrigen = wbOrigen.Sheets("Sheet0")
    Set wsDestino = wbDestino.Sheets("PARTS USED & REMOVED 2026")
    
respuesta = MsgBox("Iniciar procedimiento?", vbYesNo + vbQuestion, "Inicio")

Select Case respuesta

    Case vbYes
        
        wsOrigen.Columns("H").Delete Shift:=xlToLeft
        wsOrigen.Columns("C:F").Delete Shift:=xlToLeft
        
        wsOrigen.Cells(2, 1).Value = PartNumber_Usado
        wsOrigen.Cells(2, 2).Value = Descripcion_Usado
        wsOrigen.Cells(2, 3).Value = Cantidad_Usado
              
        Set rng = wsOrigen.Range("A1").CurrentRegion
        
        Application.CutCopyMode = False
        wsOrigen.ListObjects.Add(xlSrcRange, rng.SpecialCells(xlCellTypeVisible), , xlYes).Name = _
            "Tabla1"
        wsOrigen.ListObjects("Tabla1").Range.AutoFilter Field:=3, Criteria1:="<>0"
                
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Copy
               

End Sub
