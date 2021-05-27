Attribute VB_Name = "Consulta"
Option Explicit

Private Type CredentialType
    Ruc As String
    Username As String
    Password As String
End Type

Private Type DocumentType
    Ruc As String
    TypeDoc As String
    Serie As String
    Number As String
End Type

Public Sub Consult()
    On Error GoTo HandleErrors
    Dim Credential As CredentialType
    Dim Document As DocumentType
    Dim LastRow As Integer
    Dim Row As Integer
    
    Credential.Ruc = SheetSol.Range("D5")
    Credential.Username = SheetSol.Range("D7")
    Credential.Password = SheetSol.Range("D9")
    
    Application.ScreenUpdating = False
    
    LastRow = SheetDocs.Cells(Rows.Count, 2).End(xlUp).Row
    For Row = 5 To LastRow
        Document.Ruc = SheetDocs.Cells(Row, 2)
        Document.TypeDoc = SheetDocs.Cells(Row, 3)
        Document.Serie = SheetDocs.Cells(Row, 4)
        Document.Number = SheetDocs.Cells(Row, 5)
        
        SheetDocs.Cells(Row, 6) = SendRequest(BuildXmlSoap(Credential, Document))
    Next Row
    
    Application.ScreenUpdating = True
    Exit Sub

HandleErrors:
    If Err.Number = 65535 Then
        MsgBox Err.Description, vbCritical, "Error SOL"
        SheetSol.Activate
    Else
        MsgBox Err.Description, vbCritical
    End If
End Sub

Private Function SendRequest(XmlSoap As String) As String
    On Error GoTo HandleErrors
    Dim ClientHttp As New MSXML2.XMLHTTP60
    Dim XmlResponse As New MSXML2.DOMDocument60
    Const Endpoint As String = "https://ww1.sunat.gob.pe/ol-it-wsconscpegem/billConsultService"
    
    ClientHttp.Open "POST", Endpoint, False
    ClientHttp.send XmlSoap
    XmlResponse.LoadXML ClientHttp.responseText

    SendRequest = XmlResponse.SelectSingleNode("//statusMessage").Text
    Exit Function

HandleErrors:
    If Err.Number = 91 Then
        If XmlResponse.SelectSingleNode("//faultcode").Text = "ns0:0103" Then
            Err.Raise 65535, , "El número de RUC o el nombre de usuario son incorrectos."
        Else
            Err.Raise 65535, , XmlResponse.SelectSingleNode("//faultstring").Text
        End If
    Else
        Err.Raise 65534, , "Verifique su conexión a internet."
    End If
End Function

Private Function BuildXmlSoap(Credential As CredentialType, Document As DocumentType) As String
    BuildXmlSoap = _
        "<soapenv:Envelope xmlns:ser=""http://service.sunat.gob.pe"" " & _
                          "xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
                          "xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd""> " & _
            "<soapenv:Header>" & _
                "<wsse:Security>" & _
                    "<wsse:UsernameToken>" & _
                    "<wsse:Username>" & Credential.Ruc & Credential.Username & "</wsse:Username>" & _
                    "<wsse:Password>" & Credential.Password & "</wsse:Password>" & _
                    "</wsse:UsernameToken>" & _
                "</wsse:Security>" & _
            "</soapenv:Header>" & _
            "<soapenv:Body>" & _
                "<ser:getStatus>" & _
                    "<rucComprobante>" & Document.Ruc & "</rucComprobante>" & _
                    "<tipoComprobante>" & Format(Document.TypeDoc, "00") & "</tipoComprobante>" & _
                    "<serieComprobante>" & Document.Serie & "</serieComprobante>" & _
                    "<numeroComprobante>" & Document.Number & "</numeroComprobante>" & _
                "</ser:getStatus>" & _
            "</soapenv:Body>" & _
        "</soapenv:Envelope>"
End Function

Public Sub CleanTable()
    Rows(3).ClearContents
    Range("B4").CurrentRegion.Offset(1, 0).ClearContents
End Sub
