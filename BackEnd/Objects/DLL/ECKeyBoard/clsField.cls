VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msFieldName As String
Private miDataType As VbVarType
Private miLen As Integer
Private mvValue As Variant

Public Property Get ClassName() As String
    ClassName = App.EXEName & ".clsField"
End Property

Public Property Get FieldName() As String
    FieldName = msFieldName
End Property
Public Property Let FieldName(psName As String)
    msFieldName = psName
End Property

Public Property Get DataType() As VbVarType
    DataType = miDataType
End Property
Public Property Let DataType(piType As VbVarType)
    miDataType = piType
End Property

Public Property Get Value() As Variant
    Value = mvValue
End Property
Public Property Let Value(pvValue As Variant)
    mvValue = pvValue
End Property

Public Property Let Length(piLen As Integer)
    miLen = piLen
End Property
Public Property Get Length() As Integer
    Length = miLen
End Property
