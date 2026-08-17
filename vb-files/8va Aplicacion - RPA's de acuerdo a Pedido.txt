Sub HOctavaAplicacionPedido()

'Valores acumulativos
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim m As Integer

'Valores a rellenar
Dim MsjRequerimiento, MensajeRequerimiento, MsjCriticidad, Criticidad, Respuesta1, Respuesta2 As String
Dim CantidadASolicitar As Double
Dim FechaMaxReq, Fecha1, Fecha2 As Date

Do

    n = InputBox("Insertar desde donde comenzar el llenado", "Variable n")
    m = InputBox("Insertar hasta donde finalizar el llenado", "Variable m")
    j = InputBox("Insertar en que columna llenar los comentarios", "Variable j")

Loop While n > m Or n < 1

    MsgBox "Abre el Modulo de Requerimientos V3 y ve a la primera pestaña: Datos Generales", vbCritical, "Aviso"

    MsgBox "Ve escribiendo en esa pestaña el motivo de tu requerimiento" & vbNewLine & "(De ahi, ve a la segunda pagina 'Seleccion de articulos')", vbInformation + vbOKOnly, "1. Mensaje del Requerimiento"
       
    For i = n To m
        
        Rows(i).Select
                
        'Busqueda del N° de parte a solicitar
        
        Cells(i, 2).Copy
        
        MsgBox "Nro de parte copiado hermano!" & vbNewLine & "(Ve buscandolo en la Seleccion de Articulos)", vbExclamation + vbOKOnly, "2. Seleccion de N° de parte"
               
        'Vamos con la criticidad, en base a esto saldra la Fecha Maxima Requerida
        
        Criticidad = Cells(i, 6).Value
        Fecha2 = Cells(i, 5).Value
        Cells(i, j).Clear
        
        If Criticidad = "Estandar" Then
        
            MsjCriticidad = Cells(i, 7).Value
            Fecha1 = Date + 30
             
        ElseIf Criticidad = "AOG" Then
        
            MsjCriticidad = Cells(i, 7).Value
            Fecha1 = Date + 7
              
        ElseIf Criticidad = "Urgente" Then
        
            MsjCriticidad = Cells(i, 7).Value
            Fecha1 = Date + 15
                          
        End If
        
        If Fecha2 = #12/31/1899# Then
            FechaMaxReq = Fecha1
        Else
            FechaMaxReq = Fecha2
        End If
        
        'Comenzamos con la Cantidad a Solicitar
        
        CantidadASolicitar = Abs(Cells(i, 4).Value)
        Cells(i, j).Value = CantidadASolicitar
        Cells(i, j).Copy
        
        MsgBox "Cantidad a Solicitar copiada hermano!", vbOKOnly + vbExclamation, "1. Cantidad Requerida del articulo"
        
        Cells(i, j).ClearContents
        
        'Luego, vamos a la Fecha Maxima Requerida
        
        Cells(i, j).Value = FechaMaxReq
        Cells(i, j).NumberFormat = "mm/dd/yyyy"
        Cells(i, j).Copy
        
        MsgBox "Fecha Maxima Requerida copiada hermano!", vbOKOnly + vbExclamation, "2. Fecha Maxima Requerida del articulo"

        Cells(i, j).Clear
        
        'Ahora, vamos con la Criticidad / Comentario Articulo
        
        Cells(i, j).Value = MsjCriticidad
        Cells(i, j).Copy
        
        MsgBox "Criticidad copiada hermano!", vbOKOnly + vbExclamation, "3. Especificaciones del articulo a solicitar"
              
        Cells(i, j).ClearContents
              
        'Fin de toda la funcion
        
        Cells(i, j).Value = "X"
        
        Respuesta2 = MsgBox("¿Deseas repetir el proceso de filtrado?", _
                     vbQuestion + vbYesNo + vbDefaultButton2, _
                     "Repetir Macro")
        
        Select Case Respuesta2
        
        Case vbNo
            Exit Sub
        End Select
                            
    Next i
        
End Sub


