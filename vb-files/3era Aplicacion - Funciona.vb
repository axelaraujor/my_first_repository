Sub CTerceraAplicacion()

    Dim MiTabla As ListObject
    Dim NombreTabla As String
    Dim NombreHoja As String
    Dim NuevaColumna As ListColumn
    Dim IndiceTabla As Integer
    Dim RutaGuardado As String
    
    ' 0. Definir nombres

DefinirNombreTabla:

    IndiceTabla = InputBox("Inserte el Indice del nombre de la tabla de Excel" & vbCrLf & "(Entre los siguientes:)" & vbCrLf & vbCrLf & "(1) Exterior2025" & vbCrLf & vbCrLf & "(2) Exterior2026" & vbCrLf & vbCrLf & "(3) Locales2025" & vbCrLf & vbCrLf & "(4) Locales2026", "IndiceTabla")
   
    If IndiceTabla > 5 Or IndiceTabla < 1 Then
        MsgBox "¿Valor incorrecto? Por favor, intenta de nuevo.", vbQuestion, "Error"
        GoTo DefinirNombreTabla
    End If
   
    If IndiceTabla = 1 Then
        
        NombreTabla = "Exterior2025"
        NombreHoja = "Tabla - Exterior 2025"
        RutaGuardado = ActiveWorkbook.Path & "\1) Planificación - Análisis de insumos exterior - 2025.xlsx"
        
    ElseIf IndiceTabla = 2 Then
        
        NombreTabla = "Exterior2026"
        NombreHoja = "Tabla - Exterior 2026"
        RutaGuardado = ActiveWorkbook.Path & "\2) Planificación - Análisis de insumos exterior - 2026.xlsx"

    ElseIf IndiceTabla = 3 Then
        
        NombreTabla = "Locales2025"
        NombreHoja = "Tabla - Locales 2025"
        RutaGuardado = ActiveWorkbook.Path & "\3) Planificación - Análisis de insumos locales - 2025.xlsx"
        
    ElseIf IndiceTabla = 4 Then
        
        NombreTabla = "Locales2026"
        NombreHoja = "Tabla - Locales 2026"
        RutaGuardado = ActiveWorkbook.Path & "\4) Planificación - Análisis de insumos locales - 2026.xlsx"
    
    End If
    
    ' Validar que no se dejen vacíos los inputs
    If NombreTabla = "" Or NombreHoja = "" Then
        MsgBox "Proceso cancelado. Debe ingresar ambos nombres.", vbExclamation
        Exit Sub
    End If
    
    ' 1. Renombrar hoja (si existe Sheet0)
    On Error Resume Next
    Sheets("Sheet0").Name = NombreHoja
    On Error GoTo 0
    
    ' 2. Crear la tabla (equivalente a Ctrl + E)
    If ActiveSheet.ListObjects.Count > 0 Then
        Set MiTabla = ActiveSheet.ListObjects(1)
    Else
        Set MiTabla = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes)
    End If
    MiTabla.Name = NombreTabla
    
    ' 3. Insertar columna para el N° de Parte corregido en la posición 2
    Set NuevaColumna = MiTabla.ListColumns.Add(Position:=2)
    NuevaColumna.Name = "N° Parte_Nuevo"
    
    ' 4. Aplicar fórmula de manera segura usando el rango de la tabla
    ' Usamos intersect o la referencia de la columna asegurándonos de que haya filas
    If Not MiTabla.DataBodyRange Is Nothing Then
        With NuevaColumna.DataBodyRange
            ' Nota: Si tu Excel está en español, a veces FormulaR1C1 requiere los nombres de las columnas exactos.
            ' El uso de [@[N° Parte]] es correcto si la columna 1 se llama así.
            .FormulaR1C1 = "=MID([@[N° Parte]],2,9999)"
            .NumberFormat = "@"
            .Value = .Value ' Convertir a valores para romper la fórmula
        End With
    End If
    
    ' 5. Limpiar la columna original (que es la 1) y renombrar la nueva (que ahora pasa a ser la 1)
    MiTabla.ListColumns(1).Delete
    MiTabla.ListColumns(1).Name = "N° Parte"
    
    ' 6. Mover columnas de Stock
    On Error Resume Next
    Range(NombreTabla & "[[#All],[Stock Minimo]:[Sobrante/Faltante]]").Cut
    Columns("D:D").Insert Shift:=xlToRight
    On Error GoTo 0
    
    ' 7. Formato final
    MiTabla.TableStyle = "TableStyleLight8"
    MiTabla.Range.Font.Name = "Calibri"
    Cells(1, 1).Select
    
    ' Desactivamos alertas para que Excel no pregunte sobre perder el proyecto de VBA (las macros)
    Application.DisplayAlerts = False
    
    ' Guardamos con el formato xlOpenXMLWorkbook (que corresponde a .xlsx)
    ActiveWorkbook.SaveAs Filename:=RutaGuardado, FileFormat:=xlOpenXMLWorkbook
    
    ' Volvemos a activar las alertas del sistema
    Application.DisplayAlerts = True
    
    ' Mensaje de éxito final
    MsgBox "Proceso finalizado con éxito. El archivo se ha guardado como '.xlsx' en la ruta de origen.", vbInformation

End Sub
