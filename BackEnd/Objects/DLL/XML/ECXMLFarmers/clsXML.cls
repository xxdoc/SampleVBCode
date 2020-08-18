VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moHTTPrq As MSXML2.XMLHTTP40
Private moXMLdoc As MSXML2.DOMDocument40


Public Function SendHttpReq(psXMLDocVarName As String, _
                            psXMLDocPath As String, _
                            psMethod As String, _
                            psURL As String, _
                            psAsync As String, _
                            psUserName As String, _
                            psPassword As String, _
                            psRequestHeader As String, _
                            psRequestHeaderValue As String, _
                            pbDelXMLDoc As Boolean) As String
    On Error GoTo EH
    Dim sXMLDocVarName As String
    Dim sXMLDocPath As String
    Dim sMethod As String
    Dim sURL As String
    Dim sAsync As String
    Dim sUserName As String
    Dim sPassword As String
    Dim sRequestHeader As String
    Dim sRequestHeaderValue As String
    Dim sXML As String
    Dim oHTTPrq As MSXML2.XMLHTTP40
    Dim oXMLdoc As MSXML2.DOMDocument40
    Dim bFlag As Boolean
    
    'Set local vars to Params
    sXMLDocVarName = psXMLDocVarName
    sXMLDocPath = psXMLDocPath
    sMethod = psMethod
    sURL = psURL
    sAsync = psAsync
    sUserName = psUserName
    sRequestHeader = psRequestHeader
    sRequestHeaderValue = psRequestHeaderValue
    sPassword = psPassword
    
    'Load the XmL Document
    Set oXMLdoc = New MSXML2.DOMDocument40
    bFlag = oXMLdoc.Load(sXMLDocPath)
    'Exit if could not open the XML Document
    If Not bFlag Then
        SendHttpReq = "Could not open " & sXMLDocPath
        Set oXMLdoc = Nothing
        Exit Function
    End If
    
    'Set the xml string to the xml
    sXML = sXMLDocVarName & oXMLdoc.xml
    
    'Create Http req
    Set oHTTPrq = New MSXML2.XMLHTTP40
    
    'Open the Request with correct User Pass and Method
    oHTTPrq.open sMethod, sURL, CBool(sAsync), sUserName, sPassword
    'Send the XML with proper Header
    oHTTPrq.setRequestHeader sRequestHeader, sRequestHeaderValue
    oHTTPrq.send sXML
    'Get the response text
    SendHttpReq = oHTTPrq.responseText
    
    'Delete the XML Doc if flagged to do so
    If pbDelXMLDoc Then
        On Error Resume Next
        Kill sXMLDocPath
        On Error GoTo 0
        On Error GoTo EH
    End If
    
    Set oHTTPrq = Nothing
    Exit Function
EH:
    SendHttpReq = "Error Number: " & Err.Number & vbCrLf
    SendHttpReq = SendHttpReq & Err.Description
End Function

Public Function ValidXML(psXMLDocPath As String, pbDelXMLDoc As Boolean) As String
    On Error GoTo EH
    Dim sXMLDocPath As String
    Dim bFlag As Boolean
    
    'Start Ok
    ValidXML = "True"
    
    'Set local vars to Params
    sXMLDocPath = psXMLDocPath
    
    'Load the XML Document
    Set moXMLdoc = New MSXML2.DOMDocument40
    bFlag = moXMLdoc.Load(sXMLDocPath)
    'Exit if could not open the XML Document
    If Not bFlag Then
        ValidXML = "Invalid XML Document "
        ValidXML = ValidXML & "Parse Error: errorcode: " & moXMLdoc.parseError.errorCode & " "
        ValidXML = ValidXML & "filepos: " & moXMLdoc.parseError.filepos & " "
        ValidXML = ValidXML & "Line: " & moXMLdoc.parseError.Line & " "
        ValidXML = ValidXML & "linepos: " & moXMLdoc.parseError.linepos & " "
        ValidXML = ValidXML & "reason: " & moXMLdoc.parseError.reason
    End If
    
    'Delete the XML Doc if flagged to do so
    If pbDelXMLDoc Then
        On Error Resume Next
        Kill sXMLDocPath
        On Error GoTo 0
        On Error GoTo EH
    End If

    Exit Function
EH:
    ValidXML = "Error Number: " & Err.Number & " Description: " & Err.Description
    Set moXMLdoc = Nothing
End Function

Public Function XMLGetAttribute(psNodeName As String, psAttributeName As String) As String
    On Error GoTo EH
    Dim objNode As IXMLDOMNode
    Dim objAttr As IXMLDOMAttribute
    
    Set objNode = moXMLdoc.selectSingleNode(psNodeName)
    
    If Not objNode Is Nothing Then
        For Each objAttr In objNode.Attributes
            If StrComp(psAttributeName, objAttr.Name, vbTextCompare) = 0 Then
                XMLGetAttribute = objAttr.Value
                Exit For
            End If
        Next
    End If
    
    
    Set objNode = Nothing
    Set objAttr = Nothing
    Exit Function
EH:
    XMLGetAttribute = "Error: " & Err.Number & " Description: " & Err.Description
End Function

Public Function XMLGetNodeValue(psNodeName As String) As String
    On Error GoTo EH
    Dim objNode As IXMLDOMNode
    
    Set objNode = moXMLdoc.documentElement.selectSingleNode(psNodeName)
    XMLGetNodeValue = moXMLdoc.Text
    
    Set objNode = Nothing
    Exit Function
EH:
    XMLGetNodeValue = "Error: " & Err.Number & " Description: " & Err.Description
End Function

Public Function CleanUp()
   Set moHTTPrq = Nothing
   Set moXMLdoc = Nothing
End Function



