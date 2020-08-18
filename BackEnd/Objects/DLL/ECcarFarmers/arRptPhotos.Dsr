VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arRptPhotos 
   Caption         =   "Photos"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arRptPhotos.dsx":0000
End
Attribute VB_Name = "arRptPhotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private moLists As V2ECKeyBoard.clsCarLists

'BGS 12.27.2000 Need this to see if the Report is still active.
Private mbActiveFlag As Boolean
Private msClaimNo As String
Private mRS As Recordset
Private mbStretch As Boolean
Private moDB As DAO.Database

'Chain Reports
'Private mbChainPgBrk As Boolean
Private mbChainFlag As Boolean
Private mlChainCount As Long
Private mcolChainReports As Collection ' Contains Reports Chained to it to be added to the Sub Report object
Private moChainReport As Object

Public Property Let ChainReport(poActReport As Object)
    Set moChainReport = poActReport
End Property
Public Property Set ChainReport(poActReport As Object)
    Set moChainReport = poActReport
End Property
Public Property Get ChainReport() As Object
    Set ChainReport = moChainReport
End Property

Public Property Let Lists(poLists As V2ECKeyBoard.clsCarLists)
    Set moLists = poLists
End Property
Public Property Set Lists(poLists As V2ECKeyBoard.clsCarLists)
    Set moLists = poLists
End Property
Public Property Get Lists() As V2ECKeyBoard.clsCarLists
    Set Lists = moLists
End Property

Public Property Let ACCESSDB(poDB As DAO.Database)
    Set moDB = poDB
End Property
Public Property Set ACCESSDB(poDB As DAO.Database)
    Set moDB = poDB
End Property
Public Property Get ACCESSDB() As DAO.Database
    Set ACCESSDB = moDB
End Property
    
Public Property Let Stretch(pbFlag As Boolean)
    mbStretch = pbFlag
End Property
Public Property Get Stretch() As Boolean
    Stretch = mbStretch
End Property

Public Property Let ClaimNo(psClaimNo As String)
    msClaimNo = psClaimNo
End Property
Public Property Get ClaimNo() As String
    ClaimNo = msClaimNo
End Property

Public Property Get ActiveFlag() As Boolean
    ActiveFlag = mbActiveFlag
End Property
Public Property Let ActiveFlag(pbFlag As Boolean)
    mbActiveFlag = pbFlag
End Property

Public Property Get ClassName() As String
    ClassName = App.EXEName & "." & "arRptIBFarmers"
End Property

Private Sub ActiveReport_Initialize()
    mbActiveFlag = True
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim RSHeader As Recordset
    Dim sSQL As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    sSQL = "SELECT A.NewInsuredNames As Insured, "
    sSQL = sSQL & "A.ClientClaimNo, "
    sSQL = sSQL & "A.PolicyNumber, "
    sSQL = sSQL & "A.CatCode, "
    sSQL = sSQL & "A.LossDate, "
    sSQL = sSQL & "A.AdjustorFirstName & ' ' & A.AdjustorLastName As Adjuster, "
    sSQL = sSQL & "A.InspectedDate "
    sSQL = sSQL & "FROM " & DB_ASSIGNMENTS & " A "
    sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(msClaimNo) & "' "
    
    Set RSHeader = moDB.OpenRecordset(sSQL)
    
    With RSHeader
        If Not .EOF Then
            fInsured.Text = IIf(IsNull(!Insured), vbNullString, !Insured)
            fClientClaimNo.Text = IIf(IsNull(!ClientClaimNo), vbNullString, !ClientClaimNo)
            fPolicyNo.Text = IIf(IsNull(!PolicyNumber), vbNullString, !PolicyNumber)
            fCatCode.Text = IIf(IsNull(!CatCode), vbNullString, !CatCode)
            fDateOfLoss.Text = IIf(IsNull(!LossDate), vbNullString, Mid(!LossDate, 5, 2) & "/" & Mid(!LossDate, 7, 2) & "/" & left(!LossDate, 4))
            fAdjuster.Text = IIf(IsNull(!Adjuster), vbNullString, !Adjuster)
            fDateInspected.Text = IIf(IsNull(!InspectedDate), vbNullString, Mid(!InspectedDate, 5, 2) & "/" & Mid(!InspectedDate, 7, 2) & "/" & left(!InspectedDate, 4))
        End If
    End With
    
    'BGS 12.26.2001  need to get the photolog entries recordset
    sSQL = "SELECT * FROM " & DB_PHOTO & " A "
    sSQL = sSQL & "WHERE A.ClaimNo = '" & goUtil.utCleanSQLString(msClaimNo) & "' "
    sSQL = sSQL & "ORDER BY A.SortOrder "
    
    Set mRS = moDB.OpenRecordset(sSQL)
    
    If mbStretch Then
        imgPhoto.SizeMode = ddSMStretch
    Else
        imgPhoto.SizeMode = ddSMClip
    End If
    'Cleanup
    Set RSHeader = Nothing
    
    'Set the Chain flag if we have any
        If Not mcolChainReports Is Nothing Then
            If Not mbChainFlag Then
                mbChainFlag = True
                mlChainCount = 1
            End If
        Else
            mbChainFlag = False
        End If
        
        'If we have Chained Reports...
        If mbChainFlag Then
            Set moChainReport = mcolChainReports(mlChainCount)
            'Start the daisy linking here
            SetNextChainReport mlChainCount, mcolChainReports
            'Set the ref to sub reports in this Report
            Set subChain.object = moChainReport
        Else
            If Not moChainReport Is Nothing Then
                'Set the ref to sub reports in this Report
                Set subChain.object = moChainReport
            End If
            
        End If
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
   
    If Not mRS.EOF Then
    
        fPhotoNo.Text = IIf(IsNull(mRS!SortOrder), vbNullString, mRS!SortOrder)
        fPhotodate.Text = IIf(IsNull(mRS!PhotoDate), vbNullString, Mid(mRS!PhotoDate, 5, 2) & "/" & Mid(mRS!PhotoDate, 7, 2) & "/" & left(mRS!PhotoDate, 4))
        fDesc.Text = IIf(IsNull(mRS!Description), vbNullString, mRS!Description)
        If Not IsNull(mRS!photopath) Then
            If goUtil.utFileExists(mRS!photopath) Then
                imgPhoto.Picture = LoadPicture(mRS!photopath)
                'BGS also add a TOC (Sort number and file Name of the Photo
                TOC.Add fPhotoNo.Text & " " & Mid(mRS!photopath, InStrRev(mRS!photopath, "\") + 1)
            End If
        End If
        Detail.PrintSection
        mRS.MoveNext
    Else
        If Not moChainReport Is Nothing Then
            Me.Detail.NewPage = ddNPBefore
            subChain.Visible = True
            ReportFooter.Visible = True
        Else
            subChain.Visible = False
            ReportFooter.Visible = False
            Exit Sub 'Bail here!
        End If
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub Detail_Format"
End Sub

Private Sub ActiveReport_ReportEnd()
    On Error Resume Next
    Dim oAR As Object
    'Clean up chain reports collection and objects
    If Not mcolChainReports Is Nothing Then
        For Each oAR In mcolChainReports
            Unload oAR
            Set oAR = Nothing
        Next
        Set mcolChainReports = Nothing
        Unload moChainReport
        Set moChainReport = Nothing
    End If
    'Cleanup
    Set mRS = Nothing
    mbActiveFlag = False
    Set moDB = Nothing
    
    Exit Sub

End Sub

Public Function ExportME(psXportPath As String, pXportType As ExportType) As Boolean
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    ExportME = Lists.ExportARReport(Me, psXportPath, pXportType)
    
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function ExportME"
End Function

'For Chained Reports
Public Sub AddChainReport(poActiveReport As Object)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    If mcolChainReports Is Nothing Then
        Set mcolChainReports = New Collection
    End If
    
    mcolChainReports.Add poActiveReport
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub AddChainReport"
End Sub
'For Chained Reports
Public Sub SetNextChainReport(plChainCount As Long, pcolChainReports As Collection)
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If plChainCount + 1 <= pcolChainReports.Count Then
        Set pcolChainReports(plChainCount).ChainReport = pcolChainReports(plChainCount + 1)
        plChainCount = plChainCount + 1
        'Do daisy again
        pcolChainReports(plChainCount - 1).SetNextChainReport plChainCount, pcolChainReports
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub SetNextChainReport"
End Sub





