VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarUL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Carrier Upload Polymorphism
'Carrier Upload InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarUL in ALL Upload Classes For All Carrier Objects

Public Enum UploadType
    ClosedIB = 0
    RealTime
End Enum

Public Property Get ULType() As V2ECKeyBoard.UploadType
'    ULType = V2ECKeyBoard.ClosedIB 'V2ECKeyBoard.RealTime
End Property

Public Property Let DoDBUpdate(pbFlag As Boolean)
'    mbDoDBUpdate = pbFlag
End Property
Public Property Get DoDBUpdate() As Boolean
'    clsCarUL_DoDBUpdate = mbDoDBUpdate
End Property

Public Property Get ClassName() As String
'    clsCarUL_ClassName = App.EXEName & "." & msClassName
End Property

Public Property Get TableName() As String
'    clsCarUL_TableName = msIBTableName
End Property

Public Property Get IBNumber() As String
'    clsCarUL_IBNumber = msIBNumber
End Property

Public Function PopulateFields(psClaimNo As String, psCheckNum As String, pCurDB As Object, pclsCarUL As V2ECKeyBoard.clsUpload) As Boolean

End Function

Public Function UpdateBatchesUDT(poBatches As V2ECKeyBoard.clsBatches, poUL As V2ECKeyBoard.clsUpload) As Boolean
    
End Function

Public Sub UpdateDB(poBatches As V2ECKeyBoard.clsBatches, poUL As V2ECKeyBoard.clsUpload)
   
End Sub

Public Function GetTotalFee(psClaimNumber As String, psIBNumber As String, poBatch As V2ECKeyBoard.clsBatches, poUL As V2ECKeyBoard.clsUpload) As Currency
    
End Function

Public Function SetUtilObject(oUtilObject As Object) As Boolean
'    On Error Resume Next
'    Set goUtil = oUtilObject
'    If Err.Number = 0 Then
'        SetUtilObject = True
'    Else
'        SetUtilObject = False
'    End If
End Function

