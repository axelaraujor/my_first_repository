'## Esta primera macro era un primer intento para completar el reporte de Stock Servible por cada UFA para las aeronaves BK
'Esta macro cumplia y me ayudo a mejorar mi primer tiempo, pero igual era muy lento, por esto ya esta fuera de servicio.

Sub APrimeraAplicacion()

Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim m As Integer

'Stock
Dim StockMalvinas, StockNuevoMundo, StockLima, StockTransferencias As String
Dim Malvinas, NuevoMundo, Lima, Transferencias As String

'Mensaje
Dim Mensaje, MensajeFinal As String

'Valores que nosotros insertamos
Dim RespuestaStock, RazonStock, NumeroInvoice, AWB, CantidadInvoice, NumeroLinea, StockCuarentena As String

Do

    n = InputBox("Insertar desde donde comenzar el llenado", "Variable n")
    m = InputBox("Insertar hasta donde finalizar el llenado", "Variable m")
    j = InputBox("Insertar donde insertar los comentarios", "Variable j")

Loop While n > m Or n < 1

    For i = n To m
        
        Cells(i, 2).Copy
        Rows(i).Select
        
        RespuestaStock = MsgBox("¿Tenemos Stock?", vbYesNoCancel, "RespuestaStock")
        
        Select Case RespuestaStock
            
            Case vbYes
                                   
                StockMalvinas = InputBox("Cuantas unidades servibles hay en Malvinas Serviciables Sin Costo?", "Malvinas")
                
                If StockMalvinas = "" Then
                    MsgBox "Fin del procedimiento", vbCritical, "Final"
                    Exit Sub
                ElseIf StockMalvinas = "0" Then
                    MsgBox "Omitimos Malvinas", vbExclamation, "Malvinas"
                Else
                    Malvinas = vbCrLf & StockMalvinas & " unidades en Malvinas Serviciables sin Costo"
                    Mensaje = Malvinas
                End If
            
                StockNuevoMundo = InputBox("Cuantas unidades servibles hay en Nuevo Mundo Serviciables Sin Costo?", "Nuevo Mundo")
            
                If StockNuevoMundo = "" Then
                    MsgBox "Fin del procedimiento", vbCritical, "Final"
                Exit Sub
                ElseIf StockNuevoMundo = "0" Then
                    MsgBox "Omitimos NuevoMundo", vbExclamation, "NuevoMundo"
                Else
                    NuevoMundo = vbCrLf & StockNuevoMundo & " unidades en Nuevo Mundo Serviciables sin Costo"
                    Mensaje = Mensaje & NuevoMundo
                End If
                
                StockLima = InputBox("Cuantas unidades servibles hay en Callao Serviciables Sin Costo?", "Lima")
                
                If StockLima = "" Then
                    MsgBox "Fin del procedimiento", vbCritical, "Final"
                    Exit Sub
                ElseIf StockLima = "0" Then
                    MsgBox "Omitimos Lima", vbExclamation, "Final"
                Else
                    Lima = vbCrLf & StockLima & " unidades en Callao Serviciables sin Costo"
                    Mensaje = Mensaje & Lima
                End If
              
                StockTransferencias = InputBox("Cuantas unidades servibles hay en Transferencias?", "Transferencias")
                
                If StockTransferencias = "" Then
                    MsgBox "Fin del procedimiento", vbCritical, "Final"
                    Exit Sub
                ElseIf StockTransferencias = "0" Then
                    MsgBox "Omitimos Transferencias", vbExclamation, "Final"
                Else
                    Transferencias = vbCrLf & StockTransferencias & " unidades en Transferencias"
                    Mensaje = Mensaje & Transferencias
                End If
                
                MensajeFinal = "Tenemos stock: " & Mensaje
                
                Cells(i, j).Value = MensajeFinal
                
                Mensaje = ""
            
            Case vbNo
                 
                '¿Etiqueta?

PedirRazonStock:

                RazonStock = InputBox("¿No hay Stock?" & vbCrLf & "(Que le decimos a los demas?)" & vbCrLf & vbCrLf & "(1) No hay Stock" & vbCrLf & "(2) Van a llegar mediante el Siguiente Invoice: S*-XXXXXX y AWB #XXXXXXXXXXX" & vbCrLf & "(3) En proceso de compras por parte de SALUS, linea #XX, hoja Orders2026" & vbCrLf & "(4) Solicitado a SALUS en linea #XX, hoja Orders2026" & vbCrLf & "(5) Hay stock, pero en almacen Cuarentena", "RazonStock")

                If RazonStock > 5 Or RazonStock < 1 Then
                    MsgBox "¿Valor incorrecto? Por favor, intenta de nuevo.", vbQuestion, "Error"
                    GoTo PedirRazonStock
                End If
                
                If RazonStock = 1 Then
                    
                    MensajeFinal = "No hay Stock"
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                
                ElseIf RazonStock = 2 Then
                    
                    CantidadInvoice = InputBox("Cuantas unidades llegaran?", "CantidadInvoice")
                    NumeroInvoice = InputBox("Inserta Numero de Invoice S*-XXXXXX", "NumeroInvoice")
                    AWB = InputBox("Inserta numero de AWB", "AWB")
                        
                    MensajeFinal = "No hay stock, llegaran " & CantidadInvoice & " unidades mediante el Invoice " & NumeroInvoice & vbCrLf & "AWB: #" & AWB
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                    
                ElseIf RazonStock = 3 Then
                    
                    NumeroLinea = InputBox("Que numero de linea esta en el Excel Compartido de SALUS?", "NumeroLinea")
                    MensajeFinal = "No hay stock, en proceso de compras por parte de SALUS, linea #" & NumeroLinea & ", hoja Orders2026"
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                
                ElseIf RazonStock = 4 Then
                    
                    NumeroLinea = InputBox("Que numero de linea esta en el Excel Compartido de SALUS?", "NumeroLinea")
                    MensajeFinal = "No hay stock, hemos solicitado a SALUS, linea #" & NumeroLinea & ", hoja Orders2026"
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                
                ElseIf RazonStock = 5 Then
                    
                    StockCuarentena = InputBox("Cuantas unidades hay en el almacen Nuevo Mundo Cuarentena Sin Costo?", "StockCuarentena")
                    MensajeFinal = "Tenemos Stock: " & vbCrLf & StockCuarentena & " unidades en Nuevo Mundo Cuarentena Sin Costo"
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                                
                End If
                
            Case vbCancel
                MsgBox "Fin del procedimiento", vbCritical, "Final"
                Exit Sub
            
            End Select
            
    Next i
        
End Sub



-------------------------------------------------------------------
(Con columnas numéricas)


Sub APrimeraAplicacion()

Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim m As Integer

'Stock
Dim StockMalvinas, StockNuevoMundo, StockLima, StockTransferencias As String
Dim Malvinas, NuevoMundo, Lima, Transferencias As String

'Mensaje
Dim Mensaje, MensajeFinal As String

'Valores que nosotros insertamos
Dim RespuestaStock, RazonStock, NumeroInvoice, AWB, CantidadInvoice, NumeroLinea, StockCuarentena As String

Do

    n = InputBox("Insertar desde donde comenzar el llenado", "Variable n")
    m = InputBox("Insertar hasta donde finalizar el llenado", "Variable m")
    j = InputBox("Insertar donde insertar los comentarios", "Variable j")

Loop While n > m Or n < 1

    For i = n To m
        
        Cells(i, 2).Copy
        Rows(i).Select
        
        RespuestaStock = MsgBox("¿Tenemos Stock?", vbYesNoCancel, "RespuestaStock")
        
        Select Case RespuestaStock
            
            Case vbYes
                                   
                StockMalvinas = InputBox("Cuantas unidades servibles hay en Malvinas Serviciables Sin Costo?", "Malvinas")
                
                If StockMalvinas = "" Then
                    MsgBox "Fin del procedimiento", vbCritical, "Final"
                    Exit Sub
                ElseIf StockMalvinas = "0" Then
                    MsgBox "Omitimos Malvinas", vbExclamation, "Malvinas"
                Else
                    Malvinas = vbCrLf & StockMalvinas & " unidades en Malvinas Serviciables sin Costo"
                    Mensaje = Malvinas
                    Cells(i, j - 4).Value = StockMalvinas
                End If
            
                StockNuevoMundo = InputBox("Cuantas unidades servibles hay en Nuevo Mundo Serviciables Sin Costo?", "Nuevo Mundo")
            
                If StockNuevoMundo = "" Then
                    MsgBox "Fin del procedimiento", vbCritical, "Final"
                Exit Sub
                ElseIf StockNuevoMundo = "0" Then
                    MsgBox "Omitimos NuevoMundo", vbExclamation, "NuevoMundo"
                Else
                    NuevoMundo = vbCrLf & StockNuevoMundo & " unidades en Nuevo Mundo Serviciables sin Costo"
                    Mensaje = Mensaje & NuevoMundo
                    Cells(i, j - 3).Value = StockNuevoMundo
                End If
                
                StockLima = InputBox("Cuantas unidades servibles hay en Callao Serviciables Sin Costo?", "Lima")
                
                If StockLima = "" Then
                    MsgBox "Fin del procedimiento", vbCritical, "Final"
                    Exit Sub
                ElseIf StockLima = "0" Then
                    MsgBox "Omitimos Lima", vbExclamation, "Final"
                Else
                    Lima = vbCrLf & StockLima & " unidades en Callao Serviciables sin Costo"
                    Mensaje = Mensaje & Lima
                    Cells(i, j - 2).Value = StockLima
                End If
              
                StockTransferencias = InputBox("Cuantas unidades servibles hay en Transferencias?", "Transferencias")
                
                If StockTransferencias = "" Then
                    MsgBox "Fin del procedimiento", vbCritical, "Final"
                    Exit Sub
                ElseIf StockTransferencias = "0" Then
                    MsgBox "Omitimos Transferencias", vbExclamation, "Final"
                Else
                    Transferencias = vbCrLf & StockTransferencias & " unidades en Transferencias"
                    Mensaje = Mensaje & Transferencias
                    Cells(i, j - 1).Value = StockTransferencias
                End If
                
                MensajeFinal = "Tenemos stock: " & Mensaje
                
                Cells(i, j).Value = MensajeFinal
                
                Mensaje = ""
            
            Case vbNo
                 
                '¿Etiqueta?

PedirRazonStock:

                RazonStock = InputBox("¿No hay Stock?" & vbCrLf & "(Que le decimos a los demas?)" & vbCrLf & vbCrLf & "(1) No hay Stock" & vbCrLf & "(2) Van a llegar mediante el Siguiente Invoice: S*-XXXXXX y AWB #XXXXXXXXXXX" & vbCrLf & "(3) En proceso de compras por parte de SALUS, linea #XX, hoja Orders2026" & vbCrLf & "(4) Solicitado a SALUS en linea #XX, hoja Orders2026" & vbCrLf & "(5) Hay stock, pero en almacen Cuarentena", "RazonStock")

                If RazonStock > 5 Or RazonStock < 1 Then
                    MsgBox "¿Valor incorrecto? Por favor, intenta de nuevo.", vbQuestion, "Error"
                    GoTo PedirRazonStock
                End If
                
                If RazonStock = 1 Then
                    
                    MensajeFinal = "No hay Stock"
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                
                ElseIf RazonStock = 2 Then
                    
                    CantidadInvoice = InputBox("Cuantas unidades llegaran?", "CantidadInvoice")
                    NumeroInvoice = InputBox("Inserta Numero de Invoice S*-XXXXXX", "NumeroInvoice")
                    AWB = InputBox("Inserta numero de AWB", "AWB")
                        
                    MensajeFinal = "No hay stock, llegaran " & CantidadInvoice & " unidades mediante el Invoice " & NumeroInvoice & vbCrLf & "AWB: #" & AWB
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                    
                ElseIf RazonStock = 3 Then
                    
                    NumeroLinea = InputBox("Que numero de linea esta en el Excel Compartido de SALUS?", "NumeroLinea")
                    MensajeFinal = "No hay stock, en proceso de compras por parte de SALUS, linea #" & NumeroLinea & ", hoja Orders2026"
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                
                ElseIf RazonStock = 4 Then
                    
                    NumeroLinea = InputBox("Que numero de linea esta en el Excel Compartido de SALUS?", "NumeroLinea")
                    MensajeFinal = "No hay stock, hemos solicitado a SALUS, linea #" & NumeroLinea & ", hoja Orders2026"
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                
                ElseIf RazonStock = 5 Then
                    
                    StockCuarentena = InputBox("Cuantas unidades hay en el almacen Nuevo Mundo Cuarentena Sin Costo?", "StockCuarentena")
                    MensajeFinal = "Tenemos Stock: " & vbCrLf & StockCuarentena & " unidades en Nuevo Mundo Cuarentena Sin Costo"
                    Cells(i, j).Value = MensajeFinal
                    MensajeFinal = ""
                                
                End If
                
            Case vbCancel
                MsgBox "Fin del procedimiento", vbCritical, "Final"
                Exit Sub
            
            End Select
            
    Next i
        
End Sub
