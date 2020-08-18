VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01cal 
   Caption         =   "ActiveReport2"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01cal.dsx":0000
End
Attribute VB_Name = "arsubLossXML01cal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mActivitiesRS As WDDXRecordset
Private mvComments As Variant
Private mlCount As Long
Private mlComCount As Long
Private mbNextComment As Boolean

Public Property Let ActivitiesRS(pActivitiesRS As WDDXRecordset)
    Set mActivitiesRS = pActivitiesRS
End Property
Public Property Set ActivitiesRS(pActivitiesRS As WDDXRecordset)
    Set mActivitiesRS = pActivitiesRS
End Property

Private Sub ActiveReport_ReportEnd()
    Set mActivitiesRS = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    mlCount = 1
    mlComCount = 0
    mbNextComment = False
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim CAL As udtXML01Activities
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim myvar As Variant
    
    myvar = mActivitiesRS.getColumnNames
    
    If mActivitiesRS Is Nothing Then
        Exit Sub
    End If
NEXT_LOG:
    If mlCount <= mActivitiesRS.getRowCount Then
        If Not mbNextComment Then
            With CAL
                .GMTCreated = IIf(IsNull(mActivitiesRS.getField(mlCount, "GMTCreated")), vbNullString, mActivitiesRS.getField(mlCount, "GMTCreated"))
                .CreatedBy = IIf(IsNull(mActivitiesRS.getField(mlCount, "CreatedBy")), vbNullString, mActivitiesRS.getField(mlCount, "CreatedBy"))
                .Type = IIf(IsNull(mActivitiesRS.getField(mlCount, "Type")), vbNullString, mActivitiesRS.getField(mlCount, "Type"))
                .Description = IIf(IsNull(mActivitiesRS.getField(mlCount, "Description")), vbNullString, mActivitiesRS.getField(mlCount, "Description"))
                .Comment = IIf(IsNull(mActivitiesRS.getField(mlCount, "Comment")), vbNullString, mActivitiesRS.getField(mlCount, "Comment"))
            End With
            f_GMTCreated.Text = CAL.GMTCreated
            f_CreatedBy.Text = CAL.CreatedBy
            f_Type.Text = CAL.Type
            f_Description.Text = CAL.Description
            f_Comment.Text = CAL.Comment
            'Need to split up the multi line comments
            If CAL.Comment <> vbNullString Then
                mvComments = Split(CAL.Comment, vbCrLf)
            Else
                mvComments = Empty
            End If
            'Set the Next comment flag if there is more than one comment line
            If IsArray(mvComments) Then
                If UBound(mvComments) > 0 Then
                    mbNextComment = True
                End If
            End If
        Else
            f_GMTCreated.Text = vbNullString
            f_CreatedBy.Text = vbNullString
            f_Type.Text = vbNullString
            f_Description.Text = vbNullString
        End If
        
        'If we have more than one comment or at least one comment
        If mbNextComment Or IsArray(mvComments) Then
            f_Comment.Text = mvComments(mlComCount)
            'If the last comment is nullstring do not
            If mvComments(mlComCount) = vbNullString Then
                mlCount = mlCount + 1
                mlComCount = 0 'Reset this for the next set of comments
                mbNextComment = False
                GoTo NEXT_LOG
            End If
            'Check the next element
            If mlComCount + 1 > UBound(mvComments, 1) Then
                mbNextComment = False
                'If we do not have any more comments for this Item
                'Then we can move on to the next item count
                mlCount = mlCount + 1
                mlComCount = 0 'Reset this for the next set of comments
            Else
                mlComCount = mlComCount + 1
            End If
        Else
            f_Comment.Text = vbNullString
            mlCount = mlCount + 1
            mlComCount = 0
        End If
        If mlCount Mod 2 = 1 Then
            Detail.BackColor = &HE0E0E0
        Else
            Detail.BackColor = &HFFFFFF
        End If
        Detail.PrintSection
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub


