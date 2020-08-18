VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsECSKETCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'******************************************************************


'WDDX Serializer (Private mMySer As WDDXSerializer )
' Provides a way to serialize a value into a WDDX Packet. The value could
'be a simple value like a string or number, or it could be a complex data
'structure like an array or recordset.
'
'WDDX Deserializer (Private mMyDeser As WDDXDeserializer )
' Provides a way to take a WDDX Packet and deserialize it into the proper
'simple value, array, structure, or recordset. The datatype of the recordset
'
'WDDX Recordset (Private mMyRS As WDDXRecordset )
' Provides a way to create new recordsets, add columns and rows to the recordsets,
'and receive recordsets from other environments. Since the recordset is maintained by
'the COM object, WDDX Recordset classes can be utilized even by programs that don't
'include a concept of a recordset as part of their native language.
'
'WDDX Struct (Private mMyStruct As WDDXStruct )
' Provides a way to create new "structure" objects, which are based on ColdFusion's
'Structure datatype. Similar conceptually to what some programming languages call
'"associative arrays". Again, since the structure is maintained by the COM object,
'even languages that don't have such a concept natively can work with WDDX-style
'structure objects.
'
'WDDX JSConverter (Not being used by Eberls Development at this time)
' Provides a way to convert a simple value, recordset, array, structure or
'WDDX Packet into dynamically-produced JavaScript code that will recreate the
'appropriate value or object when executed in a browser. Helpful for setting up
'"communications" between a browser and server.
'
'******************************************************************

Public Type crvHistory
    ch000_Type       As String
    ch001_x1         As Double
    ch002_x2         As Double
    ch003_x3         As Double
    ch004_y1         As Double
    ch005_y2         As Double
    ch006_y3         As Double
    ch007_content    As String
    ch008_clr        As String
    ch009_fillclr    As String
End Type

Private mMySer As WDDXSerializer        'Allaire's WDDX serializer
Private mMyDeser As WDDXDeserializer    'Allaire's WDDX deserializer
Private mMyRS As WDDXRecordset          'Allaire's WDDX Recordset
Private mMyStruct As WDDXStruct         'Allaire's WDDX Structure (Cold Fusion Strucuture type)
Private msWddxXml As String
Private mMyCurves() As crvHistory
Private msJPGPath As String
Private mbCancel As Boolean
Private mbSave As Boolean

Public Property Get Cancel() As Boolean
    Cancel = mbCancel
End Property
Public Property Let Cancel(pbFlag As Boolean)
    mbCancel = pbFlag
End Property

Public Property Get Save() As Boolean
    Save = mbSave
End Property
Public Property Let Save(pbFlag As Boolean)
    mbSave = pbFlag
End Property

Friend Property Get myCurves() As Variant
    myCurves = mMyCurves
End Property
Friend Property Let myCurves(pvCurves As Variant)
    mMyCurves = pvCurves
End Property

Public Property Get myJPGPath() As String
    myJPGPath = msJPGPath
End Property
Friend Property Let myJPGPath(psPath As String)
    msJPGPath = psPath
End Property

Public Property Get WddxXml() As String
    WddxXml = msWddxXml
End Property
Public Property Let WddxXml(psWddxXML As String)
    msWddxXml = psWddxXML
End Property

Private Property Get msClassName() As String
    msClassName = App.EXEName & ".clsECSKETCH"
End Property

Public Function SerializeToWddxPacket(Optional pbRaiseError As Boolean = False) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim vCurves As Variant
    
    vCurves = mMyCurves
    If UBound(vCurves, 1) = -1 Then
        If pbRaiseError Then
            lErrNum = -999
            sErrDesc = "MyCurves can not be an empty array!"
            Err.Raise lErrNum, , sErrDesc
        Else
            SerializeToWddxPacket = False
            Exit Function
        End If
    End If
    
    msWddxXml = mMySer.serialize(ConvertToVariantArray(mMyCurves))
    
    SerializeToWddxPacket = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function SerializeToWddxPacket" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function DeSerializeWddxPacket(Optional pbRaiseError As Boolean = False) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim vDataVar As Variant
    
    If Trim(msWddxXml) = vbNullString Then
        If pbRaiseError Then
            lErrNum = -999
            sErrDesc = "Wddx packet data can not be empty!"
            Err.Raise lErrNum, , sErrDesc
        Else
            DeSerializeWddxPacket = False
            Exit Function
        End If
    End If
    
    If IsArray(mMyDeser.deserialize(msWddxXml)) Then
        vDataVar = mMyDeser.deserialize(msWddxXml)
    ElseIf IsObject(mMyDeser.deserialize(msWddxXml)) Then
        Set vDataVar = mMyDeser.deserialize(msWddxXml)
    End If
    
    If TypeOf vDataVar Is WDDXRecordset Then
        Set mMyRS = vDataVar
    ElseIf TypeOf vDataVar Is WDDXStruct Then
        Set mMyStruct = vDataVar
        ListStruct mMyStruct
    ElseIf IsArray(vDataVar) Then
        mMyCurves = ConvertToUDTArray(vDataVar)
    End If
    
    DeSerializeWddxPacket = True
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function DeSerializeWddxPacket" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function ConvertToUDTArray(pvaryCurves As Variant) As Variant
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim aryCurves() As crvHistory
    Dim varyCurves() As Variant
    Dim MyCurve As crvHistory
    Dim varyCurve() As Variant
    Dim lCount As Long
    
    varyCurves = pvaryCurves
    
    'Convert Variant Array to UDT Array
    
    For lCount = LBound(varyCurves, 1) To UBound(varyCurves, 1)
        varyCurve = varyCurves(lCount)
        With MyCurve
            .ch000_Type = varyCurve(0)
            .ch001_x1 = varyCurve(1)
            .ch002_x2 = varyCurve(2)
            .ch003_x3 = varyCurve(3)
            .ch004_y1 = varyCurve(4)
            .ch005_y2 = varyCurve(5)
            .ch006_y3 = varyCurve(6)
            .ch007_content = varyCurve(7)
            .ch008_clr = varyCurve(8)
            .ch009_fillclr = varyCurve(9)
        End With
        ReDim Preserve aryCurves(0 To lCount)
        aryCurves(lCount) = MyCurve
    Next
    
    ConvertToUDTArray = aryCurves
    
    Erase varyCurves
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function ConvertToUDTArray" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Function ConvertToVariantArray(pvaryCurves As Variant) As Variant
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim aryCurves() As crvHistory
    Dim varyCurves() As Variant
    Dim MyCurve As crvHistory
    Dim varyCurve() As Variant
    Dim lCount As Long
    
    aryCurves = pvaryCurves
    
    'Convert UDT Array to Variant Array
    
    For lCount = LBound(aryCurves, 1) To UBound(aryCurves, 1)
        MyCurve = aryCurves(lCount)
        'Need to set all variables types in Mycurve to Variant array
        'Reset this Curve
        Erase varyCurve
        ReDim varyCurve(0 To 9)
        With MyCurve
            varyCurve(0) = .ch000_Type
            varyCurve(1) = .ch001_x1
            varyCurve(2) = .ch002_x2
            varyCurve(3) = .ch003_x3
            varyCurve(4) = .ch004_y1
            varyCurve(5) = .ch005_y2
            varyCurve(6) = .ch006_y3
            varyCurve(7) = .ch007_content
            varyCurve(8) = .ch008_clr
            varyCurve(9) = .ch009_fillclr
        End With
        ReDim Preserve varyCurves(0 To lCount)
        varyCurves(lCount) = varyCurve
    Next
    
    ConvertToVariantArray = varyCurves
    
    Erase varyCurves
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Function ConvertToVariantArray" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Private Sub Class_Initialize()
   ' Create Instances of Allaire objects
    Set mMySer = New WDDXSerializer
    Set mMyDeser = New WDDXDeserializer
    
End Sub

Private Sub Class_Terminate()
    Set mMySer = Nothing
    Set mMyDeser = Nothing
    Set mMyRS = Nothing
    Set mMyStruct = Nothing
End Sub

Private Sub ListStruct(poStruct As WDDXStruct)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim saryProp() As Variant
    Dim lCount As Long
    Dim sMess As String
    
    
    saryProp() = poStruct.getPropNames
    
    For lCount = LBound(saryProp, 1) To UBound(saryProp, 1)
        sMess = sMess & "Property Name: " & saryProp(lCount) & vbCrLf
    Next
    
    If sMess = vbNullString Then
        sMess = "No Properties Found!"
    Else
        sMess = "Properties Found..." & vbCrLf & sMess
    End If
    
    
    MsgBox sMess, vbInformation + vbOKOnly, "Struc Propertied"
    
    
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Private Sub ListStruct" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Sub


Public Function ShowSketch() As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim myfrmSketch As frmSketch
    
    Set myfrmSketch = New frmSketch
    Set myfrmSketch.mySketch = Me
    
    Load myfrmSketch
    'Set basBalloon.mMyFrmSketch = myfrmSketch
    myfrmSketch.Show vbModal
    
    Unload myfrmSketch
    Set myfrmSketch = Nothing

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & msClassName & vbCrLf & "Public Function ShowSketch" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function
