Attribute VB_Name = "modMain"
Option Explicit

Private mvDevices As Variant            'Printer Devices under a user profile
Private mvPrinterPorts As Variant       'Printer Ports under a user profile
Private mvDefDevices As Variant         'Printer Devices under .default user
Private mvDefPrinterPorts As Variant    'Printer devices under the .default user
Private moReg As V2ECKeyBoard.clsRegSetting

Public Sub Main()
    On Error GoTo EH
    Dim sMess As String
    
    Set moReg = New V2ECKeyBoard.clsRegSetting
    
    RemoveDEFPrinters
    AddDEFPrinters
    SetDefaultPrinterDevice
    ShowWebControlService
    
    'Clean Up
    ErasePrinters
    Set moReg = Nothing
    Exit Sub
EH:
    sMess = "Error #" & Err.Number & vbCrLf
    sMess = sMess & Err.Description & vbCrLf & vbCrLf
    sMess = sMess & App.EXEName
    sMess = sMess & "Public Sub Main" & vbCrLf
    MsgBox sMess, vbCritical + vbOKOnly, "Error"
    
End Sub

Private Sub ErasePrinters()
    On Error Resume Next
    
    If IsArray(mvDevices) Then
        Erase mvDevices
    End If
    If IsArray(mvPrinterPorts) Then
        Erase mvPrinterPorts
    End If
    If IsArray(mvDefDevices) Then
        Erase mvDefDevices
    End If
    If IsArray(mvDefPrinterPorts) Then
        Erase mvDefPrinterPorts
    End If
End Sub

Private Sub RemoveDEFPrinters()
    Dim lCount As Long
    Dim sValue As String
    
    '1.
    'Enumerate all the .Default Printer Device names and Values in the Registry,
    mvDefDevices = moReg.EnumValues(HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows NT\CurrentVersion\Devices")
    '2
    'Also need to do same for printer ports, Printer ports should have duplicate
    'entries as devices but with port info attached to the values
    mvDefPrinterPorts = moReg.EnumValues(HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts")
    
    '3. Remove .Default User Printer Devices
    If DynamicArraySet(mvDefDevices) Then
        For lCount = 0 To UBound(mvDefDevices, 1)
            sValue = mvDefDevices(lCount, 0)
            If sValue <> vbNullString Then
                moReg.Delete_Value HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows NT\CurrentVersion\Devices", sValue
            End If
        Next
    End If
    
    '4. Remove .Default User Printer Ports
    If DynamicArraySet(mvDefPrinterPorts) Then
        For lCount = 0 To UBound(mvDefPrinterPorts, 1)
            sValue = mvDefPrinterPorts(lCount, 0)
            If sValue <> vbNullString Then
                moReg.Delete_Value HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts", sValue
            End If
        Next
    End If

End Sub

Private Sub AddDEFPrinters()
    Dim lCount As Long
    Dim sValue As String
    Dim sSetting As String
    
    '1.
    'Enumerate all the Printer Device names and Values in the Registry,
    'for the current user. Printers must first be added under Control Panel
    'and printers while the current user is logged on,otherwise we can't
    'add any printers to the .default user now can we.
    mvDevices = moReg.EnumValues(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices")
    'Add them to the List
    If DynamicArraySet(mvDevices) Then
        For lCount = 0 To UBound(mvDevices, 1)
            sValue = mvDevices(lCount, 0)
            sSetting = mvDevices(lCount, 1)
            If sValue <> vbNullString Then
                moReg.SaveSetting HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows NT\CurrentVersion\Devices", sValue, sSetting
            End If
        Next
    End If
    '2
    'Also need to do same for printer ports, Printer ports should have duplicate
    'entries as devices but with port info attached to the values
    mvPrinterPorts = moReg.EnumValues(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts")
    If DynamicArraySet(mvPrinterPorts) Then
        For lCount = 0 To UBound(mvPrinterPorts, 1)
            sValue = mvPrinterPorts(lCount, 0)
            sSetting = mvPrinterPorts(lCount, 1)
            If sValue <> vbNullString Then
                moReg.SaveSetting HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts", sValue, sSetting
            End If
        Next
    End If
    
End Sub

Private Sub SetDefaultPrinterDevice()
    Dim sSetting As String
    
    sSetting = moReg.Query_Value(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device", vbNullString, STRING_RESULT)
    moReg.SaveSetting HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device", sSetting
        
End Sub

Private Function DynamicArraySet(pVarArray As Variant) As Boolean
    'Purpose: To see if a Dynamic array has been set
    'Parameters : pVarArray As Variant: Send in any Dynamic array data type
    'Returns: True if has been set, false if not
    'Author : BGS-3/24/2000
    'Revision History:  SMR     Initials    Date    Description
    
    On Error GoTo NOT_SET
    Dim iRet As Integer
    
    If IsArray(pVarArray) Then
        iRet = LBound(pVarArray, 1)
        'if the Lbound call to the first dimension of
        'pVarArray does not error then the dynamic array must
        'be set so...
        DynamicArraySet = True
        Exit Function
    End If
    
NOT_SET:
    DynamicArraySet = False
End Function

Private Sub ShowWebControlService()
    moReg.SaveSetting HKEY_USERS, ".DEFAULT\Software\VB and VBA Program Settings\V2WebControl\Msg", "WebControlVisible", "True"
End Sub

