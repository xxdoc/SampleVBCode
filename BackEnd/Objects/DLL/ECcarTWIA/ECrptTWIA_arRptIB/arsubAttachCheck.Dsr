VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubAttachCheck 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubAttachCheck.dsx":0000
End
Attribute VB_Name = "arsubAttachCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcolProperty As Collection

Public Property Let PropertyCol(pcolProperty As Object)
    Set mcolProperty = pcolProperty
End Property
Public Property Set PropertyCol(pcolProperty As Object)
    Set mcolProperty = pcolProperty
End Property

Public Function GetProperty(psName As String) As Variant
    On Error GoTo EH
    
    If Not mcolProperty Is Nothing Then
        GetProperty = mcolProperty(psName)
    Else
        GetProperty = psName & ": Property collection not set!"
    End If
    
    Exit Function
EH:
    GetProperty = vbNullString
    Err.Clear
End Function

Private Sub ActiveReport_ReportEnd()
    If Not mcolProperty Is Nothing Then
        Set mcolProperty = Nothing
    End If
End Sub


Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    Dim oField As Object
    Dim sTemp As String
   
    With Me
        .f42_ClassOfClaim = GetProperty("f42_ClassOfClaim")
        .f43_sCauseOfLoss = GetProperty("f43_sCauseOfLoss")
        .f50InsuredPayee = GetProperty("f50_sInsuredPayeeName")
        .f51PayeeNames = GetProperty("f51_sPayeeNames")
        .f52Address = GetProperty("f52_sAddress")
        .f53AmountOfCheck = Format(GetProperty("f53_cAmountOfCheck"), "$#,###,###,##0.00")
        .f54CatCode = GetProperty("f54_sCatCode")
        'Also Loop through fields for Param values
        For Each oField In Me.Detail.Controls
            If TypeOf oField Is DDActiveReports.Field Then
            If StrComp(Left(CStr(oField.Name), 3), "f_p", vbTextCompare) = 0 Then
                sTemp = GetProperty(CStr(oField.Name))
                Select Case UCase(sTemp)
                    Case "YES", "TRUE", "-1"
                        sTemp = "X"
                    Case "NO", "FALSE", "0"
                        sTemp = vbNullString
                End Select
                Select Case oField.Name
                    Case f_p049_dtDate.Name, f_p051_dtApproveDate.Name
                        If sTemp = NULL_DATE Then
                            sTemp = vbNullString
                        End If
                     Case f_p046_cTexasRoofDepreciation.Name
                        sTemp = Format(sTemp, "$#,###,###,##0.00")
                End Select
                oField.Text = sTemp
            End If
            End If
        Next
    End With
    
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub



