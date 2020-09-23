Attribute VB_Name = "basROPE"
Option Explicit

Dim SOAPCalltxt As String

Private psURIServicesDescription As String
Private psServicesDescription As String

Public Sub GetServicesDescription(URIServicesDescription As String)

  Dim loXMLHTTP As Object 'Microsoft.XMLHTTP
  
  ' Create XML HTTP object
  Set loXMLHTTP = CreateObject("Microsoft.XMLHTTP")

  ' Make request to get services description
  loXMLHTTP.Open "GET", URIServicesDescription, False, "", ""
  loXMLHTTP.Send

  ' If OK, save in private variable
  If Len(loXMLHTTP.ResponseXML.XML) > 0 Then
    psURIServicesDescription = URIServicesDescription
    psServicesDescription = loXMLHTTP.ResponseXML.XML
  Else
 If frmROPEsample.chkShowPackets Then
 MsgBox loXMLHTTP.ResponseText, , "Responso:"
End If
   End If

 SOAPCalltxt = loXMLHTTP.ResponseText
 
End Sub
Public Function SOAPCall(URIServicesDescription As String, Method As String, Arguments As Variant, ReturnName As String, ShowPackets As Boolean) As String

' Call a SOAP service/method and return response.
' IN:  URIServiceDescription, URI to service description
'      Method, method to call
'      Arguments, method parameters as Variant array
'      ReturnName, name of return value
'      ShowPackets, indicates if XML packages (payload) should be shown
' OUT: SOAPCall, return value (from SOAP response)
  
  Dim loXMLHTTP As Object 'Microsoft.XMLHTTP
  Dim lsListener As String
  Dim lsParameterOrder As String
  Dim lasParameterOrder As Variant '() As String
  Dim lsRequest As String
  Dim lsResponse As String
  Dim i As Integer

  ' Get Services Description (if not already loaded)
  If URIServicesDescription <> psURIServicesDescription Then
    GetServicesDescription URIServicesDescription
  End If
 
  SOAPCall = SOAPCalltxt
  
  ' Set Payload
  'lsRequest = ""
 ' lsRequest = lsRequest & "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
  'lsRequest = lsRequest & "<SOAP:Envelope xmlns:SOAP=""http://schemas.xmlsoap.org/soap/envelope/"" SOAP:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">" & vbCrLf
Rem lsRequest = lsRequest & "<SOAP:Envelope xmlns:SOAP=""http://schemas.xmlsoap.org/soap/envelope/2000-03-01"" encodingStyle=""http://schemas.xmlsoap.org/soap/envelope/2000-03-01"">" & vbCrLf
  'lsRequest = lsRequest & "<SOAP:Body>" & vbCrLf
  'lsRequest = lsRequest & "<" & Method & ">" & vbCrLf
  'lsParameterOrder = GetParameterOrder(psServicesDescription, Method)
  'If lsParameterOrder <> "" Then
  '  lasParameterOrder = Split(lsParameterOrder, " ")
  '  For i = 0 To UBound(lasParameterOrder)
  '    lsRequest = lsRequest & "<" & lasParameterOrder(i) & ">" & CStr(Arguments(i + 1)) & "</" & lasParameterOrder(i) & ">" & vbCrLf
  '  Next i
  'End If
  'lsRequest = lsRequest & "</" & Method & ">" & vbCrLf
  'lsRequest = lsRequest & "</SOAP:Body>" & vbCrLf
  'lsRequest = lsRequest & "</SOAP:Envelope>"
  
  'If ShowPackets Then MsgBox lsRequest, , "Chiamata a SOAP"
  
  ' Get Listener
 ' lsListener = GetListener(psServicesDescription)
  
'MsgBox lsListener, , "SOAP Listener"
  
  ' Create XML HTTP object
  'Set loXMLHTTP = CreateObject("Microsoft.XMLHTTP")
    
  ' Make request to SOAP service/method
  'loXMLHTTP.Open "POST", lsListener, False, "", ""
  ' (set header info)
  'loXMLHTTP.setRequestHeader "SOAPAction", Method
  'loXMLHTTP.setRequestHeader "Content-Type", "text/xml"
  'loXMLHTTP.Send lsRequest

  ' If OK, get response
'  If Len(loXMLHTTP.ResponseXML.XML) > 0 Then
 '   lsResponse = loXMLHTTP.ResponseXML.XML

  '  If ShowPackets Then MsgBox lsResponse, , "Responso Chiamata a SOAP"

    ' Find type of call and if "Function" send back return value
   ' If GetTypeOfCall(psServicesDescription, Method) = "Function" Then
    '  If Len(lsResponse) > 0 Then
     '   If Len(ReturnName) > 0 Then
      '    SOAPCall = GetReturnValue(lsResponse, Method, ReturnName)
       ' Else
        '  SOAPCall = GetReturnValue(lsResponse, Method, "return")
        'End If
      'Else
       ' SOAPCall = ""
      'End If
    'Else
     ' SOAPCall = ""
    'End If
  'Else
  '  MsgBox loXMLHTTP.ResponseText
 ' End If
    
End Function
Private Function GetListener(URIServicesDescription As String) As String

' Get listener URL.
' IN:  URIServiceDescription, URI to service description
' OUT: GetListener, listener URL
  
  Dim loXMLDoc As Object 'MSXML.DOMDocument
  Dim loNodelist As Object 'MSXML.IXMLDOMNodeList
  
  ' Create XML document
  Set loXMLDoc = CreateObject("Microsoft.XMLDOM")
  
  ' Load services description
  loXMLDoc.loadXML URIServicesDescription
  
  ' Get listener URL
  Set loNodelist = loXMLDoc.documentElement.selectNodes("//addresses")
  GetListener = loNodelist.Item(0).childnodes(0).Attributes(0).Text
    
End Function
Private Function GetParameterOrder(URIServicesDescription As String, Method As String) As String

' Get parameter order from services description.
' IN:  URIServiceDescription, URI to service description
'      Method, method called
' OUT: GetParameterOrder, string containing parameter names in order
  
  Dim loXMLDoc As Object 'MSXML.DOMDocument
  Dim loNodelist As Object 'MSXML.IXMLDOMNodeList
  Dim lnodX As Object 'MSXML.IXMLDOMNode
  Dim lnodY As Object 'MSXML.IXMLDOMNode
  
  ' Create XML document
  Set loXMLDoc = CreateObject("Microsoft.XMLDOM")
  
  ' Load services description
  loXMLDoc.loadXML URIServicesDescription
  
  ' Get parameter order
  Set loNodelist = loXMLDoc.documentElement.selectNodes("//requestResponse")
  For Each lnodX In loNodelist
    If lnodX.Attributes(0).Text = Method Then
      For Each lnodY In lnodX.childnodes
        If lnodY.baseName = "parameterorder" Then
          GetParameterOrder = lnodY.Text
          Exit Function
        End If
      Next
    End If
  Next
    
End Function
Private Function GetTypeOfCall(URIServicesDescription As String, Method As String) As String

' Get type of call made.
' IN:  URIServiceDescription, URI to service description
'      Method, method called
' OUT: GetTypeOfCall, type of call as string ("Function" or "Procedure")
  
  Dim loXMLDoc As Object 'MSXML.DOMDocument
  Dim loNodelist As Object 'MSXML.IXMLDOMNodeList
  Dim lnodX As Object 'MSXML.IXMLDOMNode
  Dim ls As String

  ' Create XML document
  Set loXMLDoc = CreateObject("Microsoft.XMLDOM")
  
  ' Load services description
  loXMLDoc.loadXML URIServicesDescription
  
  ' Get type of call (if Function)
  Set loNodelist = loXMLDoc.documentElement.selectNodes("//requestResponse")
  For Each lnodX In loNodelist
    If lnodX.Attributes(0).Text = Method Then
      ls = "Function"
      Exit For
    End If
  Next
    
  ' If not Function, check for procedure
  If Len(ls) < 1 Then
    Set loNodelist = loXMLDoc.documentElement.selectNodes("//oneway")
    For Each lnodX In loNodelist
      If lnodX.Attributes(0).Text = Method Then
        ls = "Procedure"
        Exit For
      End If
    Next
  End If
  
  ' Return type of call
  GetTypeOfCall = ls
    
End Function
Private Function GetReturnValue(ResponseData As String, Method As String, ReturnName As String) As String

' Get return value from response.
' IN:  ResponseData, response string
'      Method, method called
'      ReturnName, name of return node
' OUT: GetReturnValue, return value

  Dim loXMLDoc As Object 'MSXML.DOMDocument
  Dim loRootElement As Object 'MSXML.IXMLDOMElement
  Dim loNodelist As Object 'MSXML.IXMLDOMNodeList
  Dim lnodX As Object 'MSXML.IXMLDOMNode
  Dim lnodY As Object 'MSXML.IXMLDOMNode
  Dim loErrXML As Object 'MSXML.IXMLDOMParseError
  Dim llErr As Long
  Dim lsErr As String

  ' Create XML document
  Set loXMLDoc = CreateObject("Microsoft.XMLDOM")
       
  ' Load response data
  loXMLDoc.loadXML ResponseData
  
  ' Check if any XML parser error
  Set loErrXML = loXMLDoc.parseError
  If loErrXML.errorCode <> 0 Then
    Err.Raise loErrXML.errorCode, "XML", loErrXML.reason
  End If
  
  ' Check response
  Set loRootElement = loXMLDoc.documentElement
  For Each lnodX In loRootElement.childnodes(0).childnodes    '<SOAP:Body>
  
    ' Check if SOAP fault response
    If lnodX.nodeName = "SOAP:Fault" Then
      For Each lnodY In lnodX.childnodes
        If lnodY.nodeName = "faultcode" Then
          llErr = CLng(lnodY.Text)
        ElseIf lnodY.nodeName = "faultstring" Then
          lsErr = lnodY.Text
        End If
        
        If llErr <> 0 And lsErr <> "" Then
          Exit For
        End If
      Next
      Err.Raise llErr, "SOAPCall", lsErr
      
    ' If <Method>Reponse node found, get "return" value and send it back
    ElseIf lnodX.nodeName = Method & "Response" Then
      For Each lnodY In lnodX.childnodes
        If lnodY.nodeName = ReturnName Then
          GetReturnValue = lnodY.Text
          Exit Function
        End If
      Next
    End If
  Next
    
End Function
