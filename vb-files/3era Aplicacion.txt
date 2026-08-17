Sub BSegundaAplicacion()

Dim i As Integer
Dim PartNumber, Alternativo, Descripcion, Cantidad, Prioridad, FechaMaxima, NumeroSerie, Restringido As String

'Todo esto es para moverme a la ultima fila del Excel Compartido

Dim ultimaFila As Long

ultimaFila = Cells(Rows.Count, "A").End(xlUp).Row
Cells(ultimaFila, 1).Select
ActiveWindow.SmallScroll Down:=1
ActiveCell.Offset(1, 0).Select

'Rellenamos la primera celda a la fecha

ActiveCell.Value = Date

'Asignamos todos los valores y rellenamos datos

PartNumber = InputBox("N° de parte", "Part Number")
ActiveCell.Offset(0, 1).Value = PartNumber

Alternativo = InputBox("Alternativos?", "Alternativo")
ActiveCell.Offset(0, 2).Value = Alternativo

Descripcion = InputBox("Descripcion", "Descripcion")
ActiveCell.Offset(0, 3).Value = Descripcion

Cantidad = InputBox("Cantidad", "Cantidad")
ActiveCell.Offset(0, 4).Value = Cantidad

Prioridad = InputBox("Prioridad", "Prioridad")
ActiveCell.Offset(0, 5).Value = Prioridad

FechaMaxima = InputBox("FechaMaxima", "FechaMaxima")
ActiveCell.Offset(0, 6).Value = FechaMaxima

NumeroSerie = InputBox("N° de serie de la aeronave", "Numero de serie")
ActiveCell.Offset(0, 7).Value = NumeroSerie

'Comentarios paso
Restringido = InputBox("Restringido?", "Restringido")
ActiveCell.Offset(0, 9).Value = Restringido

End Sub