Sub ActualizarFechasServible()

Range("B4:E4").Value = "Reporte extraido el " & Date
Range("B36:E36").Value = "Reporte extraido el " & Date
Range("B66:E66").Value = "Reporte extraido el " & Date

Range("A1:AJ2").Value = "SEGUIMIENTO DE VENCIMIENTO DE PRESERVADO (serviciable, desde 01/01/2024 al " & Date & ")"

ActiveSheet.Name = "Seguimiento " & Format(Date, "dd.mm.yyyy")

End Sub

--------------

Sub ActualizarFechasReparable()

Range("B4:E4").Value = "Reporte extraido el " & Date
Range("B36:E36").Value = "Reporte extraido el " & Date
Range("B68:E68").Value = "Reporte extraido el " & Date

Range("A1:AG2").Value = "SEGUIMIENTO DE VENCIMIENTO DE PRESERVADO (reparable, desde 01/04/2024 al " & Date & ")"

ActiveSheet.Name = "Seguimiento " & Format(Date, "dd.mm.yyyy")

End Sub

---------------

Sub ActualizarFechasCuarentena()

Range("B4:E4").Value = "Reporte extraido el " & Date
Range("B36:E36").Value = "Reporte extraido el " & Date
Range("B68:E68").Value = "Reporte extraido el " & Date

Range("A1:O2").Value = "SEGUIMIENTO DE VENCIMIENTO DE PRESERVADO (cuarentena, desde 01/01/2026 al " & Date & ")"

ActiveSheet.Name = "Seguimiento " & Format(Date, "dd.mm.yyyy")

End Sub
