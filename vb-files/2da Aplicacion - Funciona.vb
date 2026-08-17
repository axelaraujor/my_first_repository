Sub BSegundaAplicacion()

Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim m As Integer
Dim NroParte As String
Dim Respuesta As String

Do

    n = InputBox("Insertar desde donde comenzar el llenado", "Variable n")
    m = InputBox("Insertar hasta donde finalizar el llenado", "Variable m")
    j = InputBox("Insertar donde estan los comentarios", "Variable j")

Loop While n > m Or n < 1

    For i = n To m
        
        Respuesta = MsgBox("¿Continuamos para el siguiente N° de parte?", vbYesNo, "Continuar?")
        
        Select Case Respuesta
        
            Case vbYes
            
                Cells(i, 2).Copy
                NroParte = Cells(i, 2).Value
                Rows(i).Select
                
                MsgBox "¿Ya terminaste de buscar?" & vbCrLf & "N° de parte: " & NroParte, vbQuestion, "Busqueda"
                
                Cells(i, j).Copy
                
                MsgBox "¿Ya terminaste de llenar?" & vbCrLf & "N° de parte: " & NroParte, vbExclamation, "Llenado"
        
                Cells(i, j).Font.Bold = True
                Cells(i, j).Interior.Color = 14348258
        
            'Cancelamos en el caso de que hayamos terminado
            
            Case vbNo
                
                MsgBox "Proceso finalizado", vbCritical, "Final"
                Exit Sub
                            
        End Select
            
    Next i

End Sub


