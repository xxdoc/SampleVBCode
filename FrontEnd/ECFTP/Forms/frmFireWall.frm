VERSION 5.00
Begin VB.Form frmFireWall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fire Wall Settings (Advanced Users)"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFireWall.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cndCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   8520
      TabIndex        =   16
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox txtFireWallHost 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   120
      TabIndex        =   7
      Top             =   4305
      Width           =   2175
   End
   Begin VB.TextBox txtFireWallHostDesc 
      BackColor       =   &H00ECFFFF&
      Height          =   735
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3960
      Width           =   7575
   End
   Begin VB.TextBox txtFireWallLogonName 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   120
      TabIndex        =   10
      Top             =   5145
      Width           =   2175
   End
   Begin VB.TextBox txtFireWallLogonNameDesc 
      BackColor       =   &H00ECFFFF&
      Height          =   735
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4800
      Width           =   7575
   End
   Begin VB.TextBox txtFireWallPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   120
      TabIndex        =   13
      Top             =   5985
      Width           =   2175
   End
   Begin VB.TextBox txtFireWallPasswordDesc 
      BackColor       =   &H00ECFFFF&
      Height          =   735
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   5640
      Width           =   7575
   End
   Begin VB.TextBox txtFireWallPortDesc 
      BackColor       =   &H00ECFFFF&
      Height          =   735
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3120
      Width           =   7575
   End
   Begin VB.TextBox txtFireWallPort 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Text            =   "0"
      Top             =   3465
      Width           =   2175
   End
   Begin VB.TextBox txtFireWallTypesDesc 
      BackColor       =   &H00ECFFFF&
      Height          =   2415
      Left            =   7200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   518
      Width           =   2775
   End
   Begin VB.ListBox lstFirewallTypes 
      BackColor       =   &H00FFFFFF&
      Height          =   2490
      ItemData        =   "frmFireWall.frx":0442
      Left            =   120
      List            =   "frmFireWall.frx":0444
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label Label3 
      Caption         =   "Fire Wall Host"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Fire Wall Logon"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Fire Wall Password"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label lblfirewallPort 
      Caption         =   "Fire Wall port"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblFireWalltypes 
      Caption         =   "Fire Wall types [Select the Fire Wall Type Use 0 if you are not sure!]"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmFireWall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cmdSave_Click()
    On Error GoTo EH
    Dim sValue As String
    Dim lRet As Long
    Dim lCount As Long
    
    For lCount = 0 To lstFirewallTypes.ListCount - 1
        If lstFirewallTypes.Selected(lCount) = True Then
            lRet = lstFirewallTypes.ItemData(lCount)
            sValue = CStr(lRet)
            Exit For
        End If
    Next
    SaveSetting App.EXEName, "FIREWALL_SETTINGS", "FireWallType", sValue
    
    sValue = txtFireWallPort.Text
    SaveSetting App.EXEName, "FIREWALL_SETTINGS", "FireWallPort", sValue
    sValue = txtFireWallHost.Text
    SaveSetting App.EXEName, "FIREWALL_SETTINGS", "FireWallHost", sValue
    sValue = txtFireWallLogonName.Text
    SaveSetting App.EXEName, "FIREWALL_SETTINGS", "FireWallLogonName", sValue
    sValue = txtFireWallPassword.Text
    SaveSetting App.EXEName, "FIREWALL_SETTINGS", "FireWallPassword", sValue
    
    Unload Me
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdSave_Click"
End Sub

Private Sub cndCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim sRet As String
    Dim sTemp As String
    Dim lRet As Long
    Dim lCount As Long
    
    PopulateFireWallTypes
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallType", "0")
    If Not IsNumeric(sRet) Then
        sRet = "0"
    End If
    lRet = CLng(sRet)
    If lRet < 0 Then
        lRet = 0
    End If
    For lCount = 0 To lstFirewallTypes.ListCount - 1
        If lstFirewallTypes.ItemData(lCount) = lRet Then
            lstFirewallTypes.Selected(lCount) = True
            Exit For
        End If
    Next
    
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallPort", "0")
    If Not IsNumeric(sRet) Then
        sRet = "0"
    End If
    lRet = CLng(sRet)
    If lRet < 0 Then
        lRet = 0
    End If
    txtFireWallPort.Text = CLng(sRet)
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallHost", "")
    txtFireWallHost.Text = sRet
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallLogonName", "")
    txtFireWallLogonName.Text = sRet
    sRet = GetSetting(App.EXEName, "FIREWALL_SETTINGS", "FireWallPassword", "")
    txtFireWallPassword.Text = sRet
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Private Sub PopulateFireWallTypes()
    On Error GoTo EH
    Dim lCount As Long
    Dim sTemp As String
    
    sTemp = "This property determines the type of firewall to which FTP/X " & vbCrLf
    sTemp = sTemp & "will connect.  Set this property equal to Firewall Type None (0) " & vbCrLf
    sTemp = sTemp & "if you don't need firewall support.  Otherwise, set it to " & vbCrLf
    sTemp = sTemp & "one of the other type values."

    txtFireWallTypesDesc.Text = sTemp

    lstFirewallTypes.AddItem "0 - Firewall Type None [No firewall]", 0
    lstFirewallTypes.ItemData(lstFirewallTypes.NewIndex) = 0
    lstFirewallTypes.AddItem "1 - Firewall Type Socks 4 [Use Socks4 firewall]", 1
    lstFirewallTypes.ItemData(lstFirewallTypes.NewIndex) = 1
    lstFirewallTypes.AddItem "2 - Firewall Type Socks 5 [Use Socks5 firewall]", 2
    lstFirewallTypes.ItemData(lstFirewallTypes.NewIndex) = 2
    lstFirewallTypes.AddItem "3 - Firewall Type Proxy USER [Use Proxy USER command]", 3
    lstFirewallTypes.ItemData(lstFirewallTypes.NewIndex) = 3
    lstFirewallTypes.AddItem "4 - Firewall Type Proxy SITE [Use Proxy SITE command]", 4
    lstFirewallTypes.ItemData(lstFirewallTypes.NewIndex) = 4
    lstFirewallTypes.AddItem "5 - Firewall Type Proxy PROXY [Use Proxy PROXY command]", 5
    lstFirewallTypes.ItemData(lstFirewallTypes.NewIndex) = 5
    lstFirewallTypes.AddItem "6 - Firewall Type Proxy OPEN [Use Proxy OPEN command]", 6
    lstFirewallTypes.ItemData(lstFirewallTypes.NewIndex) = 6
    lstFirewallTypes.AddItem "7 - Firewall Type Pipe [Use simple Pipe]", 7
    lstFirewallTypes.ItemData(lstFirewallTypes.NewIndex) = 7
    
    sTemp = "Different firewall types usually 'listen' on different port values.  "
    sTemp = sTemp & "A standard FTP proxy usually listens on port 21 (the same as FTP servers), "
    sTemp = sTemp & "and Socks 4 and Socks 5 firewalls listen on port 1080.  "
    sTemp = sTemp & "Other port numbers may be valid.  Check with the server administator.  "
    sTemp = sTemp & "Set this property prior to connecting to a remote firewall server. "
    sTemp = sTemp & "This value can be different from the Port property. "
    sTemp = sTemp & "NOTE:  Only Socks 4 and Socks 5 firewalls allow connections to FTP servers with non-standard ports (i.e., ports other than 21). "
    txtFireWallPortDesc.Text = sTemp


    sTemp = "This property is used when a connection is to be established "
    sTemp = sTemp & "through any supported firewall. All firewalls work on "
    sTemp = sTemp & "a 'gateway' principle.  The client connects to the firewall "
    sTemp = sTemp & "and sends a request to open a connection to any remote "
    sTemp = sTemp & "FTP server.  Therefore, this property is not the same as "
    sTemp = sTemp & "the Host property.  To connect to a remote server using "
    sTemp = sTemp & "any type of firewall, the user should set all parameters "
    sTemp = sTemp & "with default values (as if he is not using a firewall), "
    sTemp = sTemp & "and then set the following properties:  " & vbCrLf
    sTemp = sTemp & "Fire Wall Host = ""firewall.host.com"" " & vbCrLf
    sTemp = sTemp & "Fire Wall Port = 21 " & vbCrLf
    sTemp = sTemp & "Fire Wall Type = 3 - Firewall Type Proxy USER " & vbCrLf
    sTemp = sTemp & "That 's it!  The FTP/X control will make the connection "
    sTemp = sTemp & "through the firewall.  Please note the properties that need "
    sTemp = sTemp & "to be set when using firewalls: FirewallHost as the remote "
    sTemp = sTemp & "firewall server, FirewallPort as the server's port, and the Type of firewall. "

    txtFireWallHostDesc.Text = sTemp

    sTemp = "USER, Socks 4 and Socks 5 are the only types of firewalls that "
    sTemp = sTemp & "support authorization.  When a connection to this type of "
    sTemp = sTemp & "firewall is established, and if authorization is required, "
    sTemp = sTemp & "then this text is sent to the firewall server. " & vbCrLf
    sTemp = sTemp & "If you need to get through a firewall using ""USER no logon"", "
    sTemp = sTemp & "set this property and the FirewallPassword property to "
    sTemp = sTemp & "blank strings. " & vbCrLf
    sTemp = sTemp & "Please note that this property is not the same as the LogonName property. "

    txtFireWallLogonNameDesc.Text = sTemp

    sTemp = "This property is supported only with the USER and Socks 5 firewall "
    sTemp = sTemp & "types.  If the remote server needs authorization, the value of "
    sTemp = sTemp & "this property is sent as identification along with the "
    sTemp = sTemp & "FirewallLogonName property.  If you need to get through a "
    sTemp = sTemp & "firewall using ""USER no logon"", set this property and the "
    sTemp = sTemp & "FirewallLogonName property to blank strings. "

    txtFireWallPasswordDesc.Text = sTemp
    
    Exit Sub
EH:
   goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub PopulateFireWallTypes"
End Sub

