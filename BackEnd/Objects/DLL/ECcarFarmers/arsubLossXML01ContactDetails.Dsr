VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsubLossXML01ContactDetails 
   Caption         =   "ActiveReport4"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arsubLossXML01ContactDetails.dsx":0000
End
Attribute VB_Name = "arsubLossXML01ContactDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moLossXML01 As V2ECcarFarmers.clsLossXML01
Private mContactDetailRS As WDDXRecordset
Private mContactRS As WDDXRecordset
Private mAddressRS As WDDXRecordset
Private mlCount As Long
' Detail Heights
Private Const SHOW_DETAIL_HEIGHT As Long = 2130
Private Const HIDE_DETAIL_HEIGHT As Long = 200
' Detail Heights

Public Property Let AddressRS(pAddressRS As WDDXRecordset)
    Set mAddressRS = pAddressRS
End Property
Public Property Set AddressRS(pAddressRS As WDDXRecordset)
    Set mAddressRS = pAddressRS
End Property

Public Property Let ContactRS(pContactRS As WDDXRecordset)
    Set mContactRS = pContactRS
End Property
Public Property Set ContactRS(pContactRS As WDDXRecordset)
    Set mContactRS = pContactRS
End Property

Public Property Let ContactDetailRS(pContactDetailRS As WDDXRecordset)
    Set mContactDetailRS = pContactDetailRS
End Property
Public Property Set ContactDetailRS(pContactDetailRS As WDDXRecordset)
    Set mContactDetailRS = pContactDetailRS
End Property

Public Property Let LossXML01(poLossXML01 As V2ECcarFarmers.clsLossXML01)
    Set moLossXML01 = poLossXML01
End Property
Public Property Set LossXML01(poLossXML01 As V2ECcarFarmers.clsLossXML01)
    Set moLossXML01 = poLossXML01
End Property
Public Property Get LossXML01() As V2ECcarFarmers.clsLossXML01
    Set LossXML01 = moLossXML01
End Property

Private Sub ActiveReport_ReportEnd()
    Set mContactDetailRS = Nothing
    Set mContactRS = Nothing
    Set mAddressRS = Nothing
    Set moLossXML01 = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo EH
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    mlCount = 1
    Exit Sub
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , sErrDesc & vbCrLf & Me.Name & vbCrLf & "Private Sub ActiveReport_ReportStart"
End Sub

Private Sub Detail_Format()
    On Error GoTo EH
    Dim ContactDetail As udtXML01ContactDetail
    Dim AssDetail As udtXML01AssignmentDetail
    Dim DupAssDetail As udtXML01AssignmentDetail
    Dim Contact As udtXML01Contacts
    Dim DupContact As udtXML01Contacts
    Dim PrimaryAddress As udtXML01Address
    Dim lCount As Long
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    If mContactDetailRS Is Nothing Then
       Exit Sub
    End If
    
    If mlCount <= mContactDetailRS.getRowCount Then
        'Populate all the text fields with the main udt
        '1. Populate Contact Detail
        With ContactDetail
            AssDetail.UniqueID = mContactDetailRS.getField(mlCount, "UniqueID")
            AssDetail.UnitNumber = moLossXML01.GetAssignmentDetailItem(AssDetail.UniqueID, "UnitNumber")
            .AgentFirstName = IIf(IsNull(mContactDetailRS.getField(mlCount, "AgentFirstName")), vbNullString, mContactDetailRS.getField(mlCount, "AgentFirstName"))
            .AgentLastName = IIf(IsNull(mContactDetailRS.getField(mlCount, "AgentLastName")), vbNullString, mContactDetailRS.getField(mlCount, "AgentLastName"))
            .AgentPrimaryPhone = IIf(IsNull(mContactDetailRS.getField(mlCount, "AgentPrimaryPhone")), vbNullString, mContactDetailRS.getField(mlCount, "AgentPrimaryPhone"))
            .ContactRowID = IIf(IsNull(mContactDetailRS.getField(mlCount, "ContactRowID")), vbNullString, mContactDetailRS.getField(mlCount, "ContactRowID"))
            Contact.FirstName = moLossXML01.GetContactItem(ContactDetail.ContactRowID, "FirstName")
            Contact.LastName = moLossXML01.GetContactItem(ContactDetail.ContactRowID, "LastName")
            Contact.RelationshipToInsured = moLossXML01.GetContactItem(ContactDetail.ContactRowID, "RelationshipToInsured")
            Contact.PrimaryPhoneNumber = moLossXML01.GetContactItem(ContactDetail.ContactRowID, "PrimaryPhoneNumber")
            Contact.HomePhoneNumber = moLossXML01.GetContactItem(ContactDetail.ContactRowID, "HomePhoneNumber")
            Contact.CellularPhoneNumber = moLossXML01.GetContactItem(ContactDetail.ContactRowID, "CellularPhoneNumber")
            Contact.PrimaryAddressID = moLossXML01.GetContactItem(ContactDetail.ContactRowID, "PrimaryAddressID")
            PrimaryAddress.Type = moLossXML01.GetAddressItem(Contact.PrimaryAddressID, "Type")
            PrimaryAddress.StreetAddress = moLossXML01.GetAddressItem(Contact.PrimaryAddressID, "StreetAddress")
            PrimaryAddress.StreetAddress2 = moLossXML01.GetAddressItem(Contact.PrimaryAddressID, "StreetAddress2")
            PrimaryAddress.City = moLossXML01.GetAddressItem(Contact.PrimaryAddressID, "City")
            PrimaryAddress.State = moLossXML01.GetAddressItem(Contact.PrimaryAddressID, "State")
            PrimaryAddress.PostalCode = moLossXML01.GetAddressItem(Contact.PrimaryAddressID, "PostalCode")
            
            f_UnitNumber.Text = AssDetail.UnitNumber
            'First check to see if this record is a subsequent contact detail record
            'If it is Need to check all previous records for duplicate contact row ID
            'If the row id is the same need to indicate that this particular Contact details is the
            'same as a previous Unit number and hide the duplicate data.  This will save space
            'and make the report more efficient / readable.
            Detail.Height = SHOW_DETAIL_HEIGHT
            f_ContactDetailsSameAsUnitNumber.Text = vbNullString
            f_ContactDetailsSameAsUnitNumber.Visible = False
            If mlCount > 1 Then
                For lCount = 1 To mlCount - 1
                    DupContact.ContactRowID = IIf(IsNull(mContactDetailRS.getField(lCount, "ContactRowID")), vbNullString, mContactDetailRS.getField(lCount, "ContactRowID"))
                    If DupContact.ContactRowID = ContactDetail.ContactRowID Then
                        DupAssDetail.UniqueID = mContactDetailRS.getField(lCount, "UniqueID")
                        DupAssDetail.UnitNumber = moLossXML01.GetAssignmentDetailItem(DupAssDetail.UniqueID, "UnitNumber")
                        f_ContactDetailsSameAsUnitNumber.Text = " (Same Contact Detail as Unit Number: " & DupAssDetail.UnitNumber & ")"
                        f_ContactDetailsSameAsUnitNumber.Visible = True
                        Detail.Height = HIDE_DETAIL_HEIGHT
                        Exit For
                    End If
                Next
            End If
            
            f_AgentFirstName.Text = .AgentFirstName
            F_AgentLastName.Text = .AgentLastName
            f_AgentPrimaryPhone.Text = .AgentPrimaryPhone
            
            f_CDFirstName.Text = Contact.FirstName
            f_CDLastName.Text = Contact.LastName
            f_RelationshipToInsured.Text = Contact.RelationshipToInsured
            f_PrimaryPhoneNumber.Text = Contact.PrimaryPhoneNumber
            f_HomePhoneNumber.Text = Contact.HomePhoneNumber
            f_CellularPhoneNumber.Text = Contact.CellularPhoneNumber
            
            f_CDType.Text = PrimaryAddress.Type
            f_StreetAddress.Text = PrimaryAddress.StreetAddress
            f_StreetAddress2.Text = PrimaryAddress.StreetAddress2
            f_City.Text = PrimaryAddress.City
            f_State.Text = PrimaryAddress.State
            f_PostalCode.Text = PrimaryAddress.PostalCode
        End With
        
        mlCount = mlCount + 1
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






