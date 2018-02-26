VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsultaCPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' ConsultaCPE v1.0.1
' (c) Victor Zevallos - https://github.com/vba-dev/ConsultaCPE
'
' Library used: Microsoft XML v6.0
'
' @author victorzevallos@protonmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

Const Url As String = "https://www.sunat.gob.pe/ol-it-wsconscpegem/billConsultService"

Private XmlSol As String

Private XmlComprobante As String

Private XmlDoc As New MSXML2.DOMDocument60

Public Sub Sol(Ruc As String, Usuario As String, Clave As String)
    XmlSol = "<wsse:Security>" & _
                 "<wsse:UsernameToken>" & _
                 "<wsse:Username>" & Ruc & Usuario & "</wsse:Username>" & _
                 "<wsse:Password>" & Clave & "</wsse:Password>" & _
                 "</wsse:UsernameToken>" & _
             "</wsse:Security>"
End Sub

Public Sub Comprobante(Ruc As String, Tipo As String, Serie As String, Numero As String)
    XmlComprobante = "<ser:getStatus>" & _
                         "<rucComprobante>" & Ruc & "</rucComprobante>" & _
                         "<tipoComprobante>" & Format(Tipo, "00") & "</tipoComprobante>" & _
                         "<serieComprobante>" & Serie & "</serieComprobante>" & _
                         "<numeroComprobante>" & Numero & "</numeroComprobante>" & _
                     "</ser:getStatus>"
End Sub

Public Function Enviar() As String
    Dim XmlHttp As New MSXML2.XMLHTTP60
    
    XmlHttp.Open "POST", Url, False
    XmlHttp.Send Xml()

    XmlDoc.LoadXML XmlHttp.responseText
    
    If ErrorAutenticacion() Then
        Err.Raise 65535, , "Verifique el nombre del usuario y que la clave sea correcta."
    End If

    Enviar = XmlDoc.SelectSingleNode("//statusMessage").Text
End Function

Private Function Xml() As String
    Xml = "<soapenv:Envelope xmlns:ser=""http://service.sunat.gob.pe"" " & _
          "xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
          "xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd""> " & _
              "<soapenv:Header>" & _
                  XmlSol & _
              "</soapenv:Header>" & _
              "<soapenv:Body>" & _
                  XmlComprobante & _
              "</soapenv:Body>" & _
          "</soapenv:Envelope>"
End Function

Private Function ErrorAutenticacion() As Boolean
    On Error GoTo TratarErrores

    If Not IsEmpty(XmlDoc.SelectSingleNode("//faultcode").Text) Then
        ErrorAutenticacion = True
    End If
    
TratarErrores:
    If Err.Number = 91 Then ErrorAutenticacion = False
End Function


