Sub HOctavaAplicacion()

'Valores acumulativos
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim m As Integer

'Valores a rellenar
Dim MsjRequerimiento, MensajeRequerimiento, MsjCriticidad, Criticidad, Respuesta1, Respuesta2 As String
Dim CantidadASolicitar As Double
Dim FechaMaxReq As Date

Do

    n = InputBox("Insertar desde donde comenzar el llenado", "Variable n")
    m = InputBox("Insertar hasta donde finalizar el llenado", "Variable m")
    j = 21

Loop While n > m Or n < 1

    MsgBox "Abre el Modulo de Requerimientos V3 y ve a la primera pestaña: Datos Generales", vbCritical, "Aviso"
    
    MsjRequerimiento = _
    InputBox("¿Cual es el motivo de tu requerimiento?" & _
    vbNewLine & "(1) Debido a Planificacion de acuerdo al MGA" & _
    vbNewLine & "(2) Debido a Stock Minimo" & _
    vbNewLine & "(3) Ambos (MGA y Stock Minimo)", _
    "Mensaje del requerimiento")
  
    If MsjRequerimiento = 1 Then
        
        MensajeRequerimiento = "Planificacion: Items solicitados debido a cumplimiento a la Planificacion Programada por Mantenimiento en el MGA - Flota Mi"
        
    ElseIf MsjRequerimiento = 2 Then
        
        MensajeRequerimiento = "Planificacion: Items solicitados debido a cumplimiento de Stock Minimo - Flota Mi"
        
    ElseIf MsjRequerimiento = 3 Then
        
        MensajeRequerimiento = "Planificacion: Items solicitados debido a cumplimiento de la Planificacion reflejada en el MGA y del Nivel de Stock Minimo - Flota Mi"
    
    Else
        
        MsgBox "Fin del procedimiento", vbOKOnly + vbCritical, "Final"
        
        Exit Sub
    
    End If
    
    Cells(1, 5).Value = MensajeRequerimiento
    Cells(1, 5).Copy
    
    MsgBox "Comentario del requerimiento copiado hermano!" & vbNewLine & "(Procede a pegarlo en la primera pestaña: Datos generales)", vbInformation + vbOKOnly, "1. Mensaje del Requerimiento"
       
    Cells(1, 5).ClearContents
    
    MsgBox "Pasemos a la 2da Pestaña: Seleccion de articulos", vbCritical, "Aviso"

    For i = n To m
        
        Rows(i).Select
        
        'Busqueda del N° de parte a solicitar
        
        Cells(i, 2).Copy
        
        MsgBox "Nro de parte copiado hermano!" & vbNewLine & "(Ve buscandolo en la Seleccion de Articulos)", vbExclamation + vbOKOnly, "2. Seleccion de N° de parte"
               
        'Vamos con la criticidad, en base a esto saldra la Fecha Maxima Requerida
        
        Criticidad = Cells(i, 14).Value
        FechaMaxReq = Date
        Cells(i, j).Clear
        
        If Criticidad = "Compra Critica A" Then
        
            MsjCriticidad = "Criticidad: Compra Critica A (Muy Urgente)"
            FechaMaxReq = Date + 15
             
        ElseIf Criticidad = "Compra Critica B" Then
        
            MsjCriticidad = "Criticidad: Compra Critica B (Urgente)"
            FechaMaxReq = Date + 30
              
        ElseIf Criticidad = "Compra Critica C" Then
        
            MsjCriticidad = "Criticidad: Compra Critica C (No muy urgente)"
            FechaMaxReq = Date + 45
        
        ElseIf Criticidad = "Abastecido" Then
        
            MsjCriticidad = "Criticidad: Abastecidos (Aunque para el año estemos abastecidos, necesitamos esta cantidad para stock)"
            FechaMaxReq = Date + 60
            
        ElseIf Criticidad = "Sin movimiento" Then
        
            MsjCriticidad = "Criticidad: Sin movimiento (Este item no tuvo movimiento en años anteriores, a pesar de ello, se necesita comprar este item)"
            FechaMaxReq = Date + 20
            
        End If
        
        'Comenzamos con la Cantidad a Solicitar
        
        CantidadASolicitar = Abs(Cells(i, 16).Value)
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
