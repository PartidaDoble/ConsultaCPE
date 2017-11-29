Attribute VB_Name = "Consulta"
Option Explicit

Sub ConsultaIndividual()
    On Error GoTo TratarErrores
    Dim Consulta As New ConsultaCPE

    Consulta.Sol [Ruc], [Usuario], [Clave]
    Consulta.Comprobante [RucProveedor], [Tipo], [Serie], [Numero]
    [Respuesta] = Consulta.Enviar()
    
    Exit Sub
TratarErrores:
    If Err.Number = 65535 Then
        MsgBox Err.Description, vbCritical, "ERROR SOL"
    ElseIf Err.Number < 0 Then
        MsgBox "Verifique su conexión a internet.", vbCritical, "ERROR"
    End If
End Sub

Sub ConsultaMasiva()
    On Error GoTo TratarErrores
    Dim Consulta As New ConsultaCPE
    Dim f As Integer
    Dim Ultimafila As Integer

    Ultimafila = Hoja3.Cells(Rows.Count, 2).End(xlUp).Row

    Consulta.Sol [Ruc], [Usuario], [Clave]

    For f = 5 To Ultimafila
        With Hoja3
            Consulta.Comprobante .Cells(f, 2), .Cells(f, 3), .Cells(f, 4), .Cells(f, 5)
            .Cells(f, 6) = Consulta.Enviar()
        End With
    Next f
    
    Exit Sub
TratarErrores:
    If Err.Number = 65535 Then
        MsgBox Err.Description, vbCritical, "ERROR SOL"
    ElseIf Err.Number < 0 Then
        MsgBox "Verifique su conexión a internet.", vbCritical, "ERROR"
    End If
End Sub
