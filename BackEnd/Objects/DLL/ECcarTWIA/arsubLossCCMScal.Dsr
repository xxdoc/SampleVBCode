VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossCCMScal 
   Caption         =   "CCMS Loss sub CAL Report"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossCCMScal.dsx":0000
End
Attribute VB_Name = "arsubLossCCMScal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcolCal As Collection
Private mvComments As Variant
Private mlCount As Long
Private mlComCount As Long
Private mbNextComment As Boolean

Public Property Let CALcol(pcolCAL As Object)
    Set mcolCal = pcolCAL
End Property
Public Property Set CALcol(pcolCAL As Object)
    Set mcolCal = pcolCAL
End Property

Private Sub ActiveReport_ReportEnd()
    Set mcolCal = Nothing
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
    Dim CAL As udtCommentsActLog
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mcolCal Is Nothing Then
        Exit Sub
    End If
NEXT_LOG:
    If mlCount <= mcolCal.Count Then
        If Not mbNextComment Then
            CAL = mcolCal(mlCount)
            fcal01_CAT.Text = CAL.cal01_CAT
            fcal02_Date.Text = CAL.cal02_Date
            fcal03_Time.Text = CAL.cal03_Time
            fcal04_Action.Text = CAL.cal04_Action
            fcal05_User.Text = CAL.cal05_User
            'Need to split up the multi line comments
            If CAL.cal06_Comments <> vbNullString Then
                mvComments = Split(CAL.cal06_Comments, vbCrLf)
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
            fcal01_CAT.Text = vbNullString
            fcal02_Date.Text = vbNullString
            fcal03_Time.Text = vbNullString
            fcal04_Action.Text = vbNullString
            fcal05_User.Text = vbNullString
        End If
        
        'If we have more than one comment or at least one comment
        If mbNextComment Or IsArray(mvComments) Then
            fcal06_Comments.Text = mvComments(mlComCount)
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
            fcal06_Comments.Text = vbNullString
            mlCount = mlCount + 1
            mlComCount = 0
        End If
        Detail.PrintSection
        
    End If
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub Detail_Format"
End Sub

