VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Loss Report Polymorphism
'Loss Report InterFace Definiton
'Use Implements V2ECKeyBoard.clsCarLR in ALL Loss Report Classes For All Carrier Objects

Public Property Get AbortProcessRawData() As Boolean
    
End Property
Public Property Let AbortProcessRawData(pbFlag As Boolean)
    
End Property

Public Property Let LossType(pType As Variant)
'    mLossType = pType
End Property
Public Property Get LossType() As Variant
'    LossType = mLossType
End Property

Public Property Get LossReport() As Variant
'    LossReport = mudt???LossReport
End Property

Public Property Let LRs(poLRs As V2ECKeyBoard.clsLossReports)
'    Set moLRs = poLRs
End Property
Public Property Set LRs(poLRs As V2ECKeyBoard.clsLossReports)
'    Set moLRs = poLRs
End Property
Public Property Get LRs() As V2ECKeyBoard.clsLossReports
'    Set LRs = moLRs
End Property

Public Property Let PrnData(psData As String)
'    mudt???LossReport.PrnData = psData
End Property
Public Property Get PrnData() As String
'    PrnData = mudt???LossReport.PrnData
End Property

Public Property Let PrnKey(psKey As String)
'    mudt??LossReport.PrnKey = psKey
End Property
Public Property Get PrnKey() As String
'    PrnKey = mudt???LossReport.PrnKey
End Property

Public Property Get ClassName() As String
'    ClassName = App.EXEName & ".cls???"
End Property

Public Property Get GenTextHeight() As Long
'    GenTextHeight = 165
End Property
Public Property Get GenTextFontName() As String
'    GenTextFontName = "Courier New"
End Property
Public Property Get GenTextFontSize() As Long
'    GenTextFontSize = 11
End Property

Public Property Get OleType() As String
'    OleType = msOleType
End Property
Public Property Let OleType(psOleType As String)
'    msOleType = psOleType
End Property

Public Property Get MaxAllowedInChain() As Long
    
End Property

Public Property Get InsuredName() As String
'    InsuredName = msInsuredName
End Property
Public Property Let InsuredName(psInsuredName As String)
'    msInsuredName = psInsuredName
End Property

Public Property Get HomePhone() As String
'    HomePhone = msHomePhone
End Property
Public Property Let HomePhone(psHomePhone As String)
'    msHomePhone = psHomePhone
End Property

Public Property Get WorkPhone() As String
'    WorkPhone = msWorkPhone
End Property
Public Property Let WorkPhone(psWorkPhone As String)
'    msWorkPhone = psWorkPhone
End Property

Public Property Get DateAssign() As String
'    DateAssign = msDateAssign
End Property
Public Property Let DateAssign(psDateAssign As String)
'    msDateAssign = psDateAssign
End Property

Public Property Get AssignmentType() As String
'    AssignmentType = msAssignmentType
End Property
Public Property Let AssignmentType(psDateAssign As String)
'    msAssignmentType = psAssignmentType
End Property

Public Property Get Status() As String
'    Status = msStatus
End Property
Public Property Let Status(psStatus As String)
'    msStatus = psStatus
End Property

Public Property Get CatName() As String
'    CatName = msCatName
End Property
Public Property Let CatName(psCatName As String)
'    msCatName = psCatName
End Property

Public Property Get CatCode() As String
'    CatCode = msCatCode
End Property
Public Property Let CatCode(psCatCode As String)
'    msCatCode = psCatCode
End Property

Public Property Get Adjuster() As String
'    Adjuster = msAdjuster
End Property
Public Property Let Adjuster(psAdjuster As String)
'    msAdjuster = psAdjuster
End Property

Public Property Get ACID() As String
'    ACID = msACID
End Property
Public Property Let ACID(psACID As String)
'    msACID = psACID
End Property

Public Property Get CLIENTNUM() As String
'    ACID = msCLIENTNUM
End Property
Public Property Let CLIENTNUM(psCLIENTNUM As String)
'    msACID = psCLIENTNUM
End Property

Public Property Get IBNUM() As String
'    ACID = msIBNUM
End Property
Public Property Let IBNUM(psIBNUM As String)
'    msACID = psIBNUM
End Property

Public Property Get TypeOfACID() As String
'    TypeOfACID = msTypeOfACID
End Property
Public Property Let TypeOfACID(psTypeOfACID As String)
'    msTypeOfACID = psTypeOfACID
End Property

Public Function ProcessRawData(psDataPath As String, _
                               psOutPath As String, _
                               Optional poPB As Object, _
                               Optional potxtProgMess As Object) As Boolean
    
End Function

Public Function AdditmX(poitmX As Object, polvw As Object) As Boolean
    
End Function

Public Function PrintMe(psPrnDeviceName As String, plFormat As Long, pbPreview As Boolean, pbPDF As Boolean, _
                        pofrmIncludeDocs As Object, plHwnd As Long, plChainType As Long) As Boolean
    
End Function

Public Function ExportME(poLRs As V2ECKeyBoard.clsLossReports, pXportType As ExportType) As Boolean
    
End Function

Public Function GetLRSQL(psAppName As String, plVersion As Long, Optional psSQLError As String) As String
    
End Function

Public Sub PopulateMe()

End Sub

Public Function SetUtilObject(oUtilObject As Object) As Boolean
'    On Error Resume Next
'    Set goUtil = oUtilObject
'    If Err.Number = 0 Then
'        SetUtilObject = True
'    Else
'        SetUtilObject = False
'    End If
End Function

Public Function CLEANUP() As Boolean
    
End Function


