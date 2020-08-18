VERSION 5.00
Begin VB.Form frmClaim 
   AutoRedraw      =   -1  'True
   Caption         =   "Claim Maintenance"
   ClientHeight    =   855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClaim.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optClaim 
      Caption         =   "&9 File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   7560
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "frmPrint"
      Top             =   0
      Width           =   900
   End
   Begin VB.OptionButton optClaim 
      Caption         =   "&8 Misc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   6720
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   900
   End
   Begin VB.OptionButton optClaim 
      Caption         =   "&7 Bill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   5880
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "frmBillingInfo"
      Top             =   0
      Width           =   900
   End
   Begin VB.OptionButton optClaim 
      Caption         =   "&6 Attach"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   5040
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":1298
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "frmAttachments"
      Top             =   0
      Width           =   900
   End
   Begin VB.OptionButton optClaim 
      Caption         =   "&5 Rpt."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   4200
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":16DA
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "frmReports"
      Top             =   0
      Width           =   900
   End
   Begin VB.OptionButton optClaim 
      Caption         =   "&4 Photo "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   3360
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":1B1C
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "frmPhotos"
      Top             =   0
      Width           =   900
   End
   Begin VB.OptionButton optClaim 
      Caption         =   "&3 Indem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   2520
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":1E26
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "frmIndemnity"
      Top             =   0
      Width           =   900
   End
   Begin VB.OptionButton optClaim 
      Caption         =   "&2 I-Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   1680
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":2268
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "frmActivityLog"
      Top             =   0
      Width           =   900
   End
   Begin VB.OptionButton optClaim 
      Caption         =   "&1 Claim"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   840
      MaskColor       =   &H80000017&
      Picture         =   "frmClaim.frx":26AA
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "frmClaimInfo"
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdViewPDFLossReport 
      Caption         =   "&Loss Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmClaim.frx":2AEC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   900
   End
   Begin VB.Timer TimerReSize 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8640
      Top             =   0
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   855
      Left            =   9600
      MaskColor       =   &H00000000&
      Picture         =   "frmClaim.frx":2C36
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer Timer_UnloadForm 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8640
      Top             =   360
   End
End
Attribute VB_Name = "frmClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Begin Color Const
Private Const BG_WHITE As Long = &H80000014
Private Const BG_GRAY As Long = &H8000000F
Private Const NO_SELECTION As Integer = -999

'END Color Const
Private mbLoading As Boolean
Private MyfrmClaimsList As frmClaimsList
Private mfrmClaimInfo As frmClaimInfo
Private mfrmActivityLog As frmActivityLog
Private mfrmIndemnity As frmIndemnity
Private mfrmBillingInfo As frmBillingInfo
Private mfrmPhotos As frmPhotos
Private mfrmReports As frmReports
Private mfrmAttachments As frmAttachments
Private mfrmMiscellaneous As frmMiscellaneous
Private mfrmPrint As frmPrint
Private msAssignmentsID As String
Private moGUI As V2ECKeyBoard.clsCarGUI
Private madoRSAssignments As ADODB.Recordset
Private madoRSPolicyLimits As ADODB.Recordset
Private madoRSBillingCount As ADODB.Recordset 'list of Billing Items for Assignment
Private madoRSRTPhotoLog As ADODB.Recordset
Private madoRSRTActivityLogInfo As ADODB.Recordset
Private madoRSRTActivityLog As ADODB.Recordset
Private madoRSRTIndemnity As ADODB.Recordset
Private madoRSRTChecks As ADODB.Recordset
Private madoRSRTAttachments As ADODB.Recordset
Private madoRSMainReports As ADODB.Recordset
Private madoRSMainReportsHistory As ADODB.Recordset
Private madoRSCarSpecReports As ADODB.Recordset
Private madoRSCarSpecReportsHistory As ADODB.Recordset
Private madoRSWordXLDocs As ADODB.Recordset
Private madoRSIB As ADODB.Recordset 'These are the actual IB Records
Private madoRSBillingReports As ADODB.Recordset 'This is the Billing Report Software
Private madoRSBillingReportsHistory As ADODB.Recordset 'This is the Billing Report Software History
Private madoRSPayment As ADODB.Recordset 'These are the actual Payment Records
Private madoRSPaymentReports As ADODB.Recordset 'This is the Payment Report Software
Private madoRSPaymentReportsHistory As ADODB.Recordset 'This is the Payment Report Software History
Private madoRSRTPhotoReportList As ADODB.Recordset 'Photo Report List (Multi Report Type)
Private madoRSRTWSDiagramList As ADODB.Recordset 'Worksheet Diagram List (Multi Report type)
Private madoRSPackageList As ADODB.Recordset 'Package List (Multi Report)
Private madoRSPackageItem As ADODB.Recordset 'PakageItems
Private madoRSFeeScheduleList As ADODB.Recordset ' List of Available Fee schedules for Current Carrier company
Private madoRSBillingCountItem As ADODB.Recordset 'a Single Billing Item
Private madoRSCLIB As ADODB.Recordset 'Closed IB RS
Private madoRSCLIBFee As ADODB.Recordset ' Closed IB Fee (Service and Expense) Items
Private madoRSRTIB As ADODB.Recordset 'Currently working on IB RS
Private madoRSRTIBFee As ADODB.Recordset ' Currently working on IB Fee (Service and Expesne) Items
Private madoRSClientCOCat As ADODB.Recordset 'Current Client company Cat Info
Private miMyStatus As V2ECKeyBoard.AssgnStatus
Private mArv As V2ARViewer.clsARViewer
Private moForm As Form
'2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
Private msMiscReportParamName As String
'2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30

'2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
Public Property Let MiscReportParamName(psMiscReportParamName As String)
    msMiscReportParamName = psMiscReportParamName
End Property
Public Property Get MiscReportParamName() As String
    MiscReportParamName = msMiscReportParamName
End Property
'2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30

Public Property Let MyStatus(piStatus As V2ECKeyBoard.AssgnStatus)
    miMyStatus = piStatus
End Property
Public Property Get MyStatus() As V2ECKeyBoard.AssgnStatus
    MyStatus = miMyStatus
End Property

'FeeSchedule List
Public Property Let adoRSFeeScheduleList(padoRS As ADODB.Recordset)
    Set madoRSFeeScheduleList = padoRS
End Property
Public Property Set adoRSFeeScheduleList(padoRS As ADODB.Recordset)
    Set madoRSFeeScheduleList = padoRS
End Property
Public Property Get adoRSFeeScheduleList() As ADODB.Recordset
    Set adoRSFeeScheduleList = madoRSFeeScheduleList
End Property

'Package List (Multi Report)
Public Property Let adoRSPackageList(padoRS As ADODB.Recordset)
    Set madoRSPackageList = padoRS
End Property
Public Property Set adoRSPackageList(padoRS As ADODB.Recordset)
    Set madoRSPackageList = padoRS
End Property
Public Property Get adoRSPackageList() As ADODB.Recordset
    Set adoRSPackageList = madoRSPackageList
End Property

'Package Items
Public Property Let adoRSPackageItem(padoRS As ADODB.Recordset)
    Set madoRSPackageItem = padoRS
End Property
Public Property Set adoRSPackageItem(padoRS As ADODB.Recordset)
    Set madoRSPackageItem = padoRS
End Property
Public Property Get adoRSPackageItem() As ADODB.Recordset
    Set adoRSPackageItem = madoRSPackageItem
End Property

'Worksheet Diagram List (Multi Report type)
Public Property Let adoRSRTWSDiagramList(padoRS As ADODB.Recordset)
    Set madoRSRTWSDiagramList = padoRS
End Property
Public Property Set adoRSRTWSDiagramList(padoRS As ADODB.Recordset)
    Set madoRSRTWSDiagramList = padoRS
End Property
Public Property Get adoRSRTWSDiagramList() As ADODB.Recordset
    Set adoRSRTWSDiagramList = madoRSRTWSDiagramList
End Property

'Photo Report List (Multi Report Type)
Public Property Let adoRSRTPhotoReportList(padoRS As ADODB.Recordset)
    Set madoRSRTPhotoReportList = padoRS
End Property
Public Property Set adoRSRTPhotoReportList(padoRS As ADODB.Recordset)
    Set madoRSRTPhotoReportList = padoRS
End Property
Public Property Get adoRSRTPhotoReportList() As ADODB.Recordset
    Set adoRSRTPhotoReportList = madoRSRTPhotoReportList
End Property

'Assignemnts RS
Public Property Let adoRSAssignments(padoRS As ADODB.Recordset)
    Set madoRSAssignments = padoRS
End Property
Public Property Set adoRSAssignments(padoRS As ADODB.Recordset)
    Set madoRSAssignments = padoRS
End Property
Public Property Get adoRSAssignments() As ADODB.Recordset
    Set adoRSAssignments = madoRSAssignments
End Property

'Policy Limits RS
Public Property Let adoRSPolicyLimits(padoRS As ADODB.Recordset)
    Set madoRSPolicyLimits = padoRS
End Property
Public Property Set adoRSPolicyLimits(padoRS As ADODB.Recordset)
    Set madoRSPolicyLimits = padoRS
End Property
Public Property Get adoRSPolicyLimits() As ADODB.Recordset
    Set adoRSPolicyLimits = madoRSPolicyLimits
End Property

'Billing Count RS
Public Property Let adoRSBillingCount(padoRS As ADODB.Recordset)
    Set madoRSBillingCount = padoRS
End Property
Public Property Set adoRSBillingCount(padoRS As ADODB.Recordset)
    Set madoRSBillingCount = padoRS
End Property
Public Property Get adoRSBillingCount() As ADODB.Recordset
    Set adoRSBillingCount = madoRSBillingCount
End Property

'Photo Log  RS
Public Property Let adoRSRTPhotoLog(padoRS As ADODB.Recordset)
    Set madoRSRTPhotoLog = padoRS
End Property
Public Property Set adoRSRTPhotoLog(padoRS As ADODB.Recordset)
    Set madoRSRTPhotoLog = padoRS
End Property
Public Property Get adoRSRTPhotoLog() As ADODB.Recordset
    Set adoRSRTPhotoLog = madoRSRTPhotoLog
End Property

'Activity Log Info RS
Public Property Let adoRSRTActivityLogInfo(padoRS As ADODB.Recordset)
    Set madoRSRTActivityLogInfo = padoRS
End Property
Public Property Set adoRSRTActivityLogInfo(padoRS As ADODB.Recordset)
    Set madoRSRTActivityLogInfo = padoRS
End Property
Public Property Get adoRSRTActivityLogInfo() As ADODB.Recordset
    Set adoRSRTActivityLogInfo = madoRSRTActivityLogInfo
End Property

'Activity Log RS
Public Property Let adoRSRTActivityLog(padoRS As ADODB.Recordset)
    Set madoRSRTActivityLog = padoRS
End Property
Public Property Set adoRSRTActivityLog(padoRS As ADODB.Recordset)
    Set madoRSRTActivityLog = padoRS
End Property
Public Property Get adoRSRTActivityLog() As ADODB.Recordset
    Set adoRSRTActivityLog = madoRSRTActivityLog
End Property

'RTIndemnity RS
Public Property Let adoRSRTIndemnity(padoRS As ADODB.Recordset)
    Set madoRSRTIndemnity = padoRS
End Property
Public Property Set adoRSRTIndemnity(padoRS As ADODB.Recordset)
    Set madoRSRTIndemnity = padoRS
End Property
Public Property Get adoRSRTIndemnity() As ADODB.Recordset
    Set adoRSRTIndemnity = madoRSRTIndemnity
End Property

'RTChecks RS
Public Property Let adoRSRTChecks(padoRS As ADODB.Recordset)
    Set madoRSRTChecks = padoRS
End Property
Public Property Set adoRSRTChecks(padoRS As ADODB.Recordset)
    Set madoRSRTChecks = padoRS
End Property
Public Property Get adoRSRTChecks() As ADODB.Recordset
    Set adoRSRTChecks = madoRSRTChecks
End Property

'Attachments RS
Public Property Let adoRSRTAttachments(padoRS As ADODB.Recordset)
    Set madoRSRTAttachments = padoRS
End Property
Public Property Set adoRSRTAttachments(padoRS As ADODB.Recordset)
    Set madoRSRTAttachments = padoRS
End Property
Public Property Get adoRSRTAttachments() As ADODB.Recordset
    Set adoRSRTAttachments = madoRSRTAttachments
End Property

'Main Reports RS
Public Property Let adoRSMainReports(padoRS As ADODB.Recordset)
    Set madoRSMainReports = padoRS
End Property
Public Property Set adoRSMainReports(padoRS As ADODB.Recordset)
    Set madoRSMainReports = padoRS
End Property
Public Property Get adoRSMainReports() As ADODB.Recordset
    Set adoRSMainReports = madoRSMainReports
End Property

'Main Reports History RS
Public Property Let adoRSMainReportsHistory(padoRS As ADODB.Recordset)
    Set madoRSMainReportsHistory = padoRS
End Property
Public Property Set adoRSMainReportsHistory(padoRS As ADODB.Recordset)
    Set madoRSMainReportsHistory = padoRS
End Property
Public Property Get adoRSMainReportsHistory() As ADODB.Recordset
    Set adoRSMainReportsHistory = madoRSMainReportsHistory
End Property

'Carrier Specific Reports RS
Public Property Let adoRSCarSpecReports(padoRS As ADODB.Recordset)
    Set madoRSCarSpecReports = padoRS
End Property
Public Property Set adoRSCarSpecReports(padoRS As ADODB.Recordset)
    Set madoRSCarSpecReports = padoRS
End Property
Public Property Get adoRSCarSpecReports() As ADODB.Recordset
    Set adoRSCarSpecReports = madoRSCarSpecReports
End Property

'Carrier Specific Reports History RS
Public Property Let adoRSCarSpecReportsHistory(padoRS As ADODB.Recordset)
    Set madoRSCarSpecReportsHistory = padoRS
End Property
Public Property Set adoRSCarSpecReportsHistory(padoRS As ADODB.Recordset)
    Set madoRSCarSpecReportsHistory = padoRS
End Property
Public Property Get adoRSCarSpecReportsHistory() As ADODB.Recordset
    Set adoRSCarSpecReportsHistory = madoRSCarSpecReportsHistory
End Property

'IB RS Actual Bills
Public Property Let adoRSIB(padoRS As ADODB.Recordset)
    Set madoRSIB = padoRS
End Property
Public Property Set adoRSIB(padoRS As ADODB.Recordset)
    Set madoRSIB = padoRS
End Property
Public Property Get adoRSIB() As ADODB.Recordset
    Set adoRSIB = madoRSIB
End Property

'Billing Reports RS
Public Property Let adoRSBillingReports(padoRS As ADODB.Recordset)
    Set madoRSBillingReports = padoRS
End Property
Public Property Set adoRSBillingReports(padoRS As ADODB.Recordset)
    Set madoRSBillingReports = padoRS
End Property
Public Property Get adoRSBillingReports() As ADODB.Recordset
    Set adoRSBillingReports = madoRSBillingReports
End Property

'Billing Reports History RS
Public Property Let adoRSBillingReportsHistory(padoRS As ADODB.Recordset)
    Set madoRSBillingReportsHistory = padoRS
End Property
Public Property Set adoRSBillingReportsHistory(padoRS As ADODB.Recordset)
    Set madoRSBillingReportsHistory = padoRS
End Property
Public Property Get adoRSBillingReportsHistory() As ADODB.Recordset
    Set adoRSBillingReportsHistory = madoRSBillingReportsHistory
End Property

'Payment RS Actual Payments
Public Property Let adoRSPayment(padoRS As ADODB.Recordset)
    Set madoRSPayment = padoRS
End Property
Public Property Set adoRSPayment(padoRS As ADODB.Recordset)
    Set madoRSPayment = padoRS
End Property
Public Property Get adoRSPayment() As ADODB.Recordset
    Set adoRSPayment = madoRSPayment
End Property

'Payment Reports RS
Public Property Let adoRSPaymentReports(padoRS As ADODB.Recordset)
    Set madoRSPaymentReports = padoRS
End Property
Public Property Set adoRSPaymentReports(padoRS As ADODB.Recordset)
    Set madoRSPaymentReports = padoRS
End Property
Public Property Get adoRSPaymentReports() As ADODB.Recordset
    Set adoRSPaymentReports = madoRSPaymentReports
End Property

'Payment Reports History RS
Public Property Let adoRSPaymentReportsHistory(padoRS As ADODB.Recordset)
    Set madoRSPaymentReportsHistory = padoRS
End Property
Public Property Set adoRSPaymentReportsHistory(padoRS As ADODB.Recordset)
    Set madoRSPaymentReportsHistory = padoRS
End Property
Public Property Get adoRSPaymentReportsHistory() As ADODB.Recordset
    Set adoRSPaymentReportsHistory = madoRSPaymentReportsHistory
End Property

'a Single Billing Item
Public Property Let adoRSBillingCountItem(padoRS As ADODB.Recordset)
    Set madoRSBillingCountItem = padoRS
End Property
Public Property Set adoRSBillingCountItem(padoRS As ADODB.Recordset)
    Set madoRSBillingCountItem = padoRS
End Property
Public Property Get adoRSBillingCountItem() As ADODB.Recordset
    Set adoRSBillingCountItem = madoRSBillingCountItem
End Property

'Closed IB RS
Public Property Let adoRSCLIB(padoRS As ADODB.Recordset)
    Set madoRSCLIB = padoRS
End Property
Public Property Set adoRSCLIB(padoRS As ADODB.Recordset)
    Set madoRSCLIB = padoRS
End Property
Public Property Get adoRSCLIB() As ADODB.Recordset
    Set adoRSCLIB = madoRSCLIB
End Property

'Closed IB Fee (Service and Expense) Items
Public Property Let adoRSCLIBFee(padoRS As ADODB.Recordset)
    Set madoRSCLIBFee = padoRS
End Property
Public Property Set adoRSCLIBFee(padoRS As ADODB.Recordset)
    Set madoRSCLIBFee = padoRS
End Property
Public Property Get adoRSCLIBFee() As ADODB.Recordset
    Set adoRSCLIBFee = madoRSCLIBFee
End Property

'Currently working on IB RS
Public Property Let adoRSRTIB(padoRS As ADODB.Recordset)
    Set madoRSRTIB = padoRS
End Property
Public Property Set adoRSRTIB(padoRS As ADODB.Recordset)
    Set madoRSRTIB = padoRS
End Property
Public Property Get adoRSRTIB() As ADODB.Recordset
    Set adoRSRTIB = madoRSRTIB
End Property

' Currently working on IB Fee (Service
Public Property Let adoRSRTIBFee(padoRS As ADODB.Recordset)
    Set madoRSRTIBFee = padoRS
End Property
Public Property Set adoRSRTIBFee(padoRS As ADODB.Recordset)
    Set madoRSRTIBFee = padoRS
End Property
Public Property Get adoRSRTIBFee() As ADODB.Recordset
    Set adoRSRTIBFee = madoRSRTIBFee
End Property

'Current Client company Cat Info
Public Property Let adoRSClientCOCat(padoRS As ADODB.Recordset)
    Set madoRSClientCOCat = padoRS
End Property
Public Property Set adoRSClientCOCat(padoRS As ADODB.Recordset)
    Set madoRSClientCOCat = padoRS
End Property
Public Property Get adoRSClientCOCat() As ADODB.Recordset
    Set adoRSClientCOCat = madoRSClientCOCat
End Property

'Word XL Docs
Public Property Let adoRSWordXLDocs(padoRS As ADODB.Recordset)
    Set madoRSWordXLDocs = padoRS
End Property
Public Property Set adoRSWordXLDocs(padoRS As ADODB.Recordset)
    Set madoRSWordXLDocs = padoRS
End Property
Public Property Get adoRSWordXLDocs() As ADODB.Recordset
    Set adoRSWordXLDocs = madoRSWordXLDocs
End Property

Public Property Let MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Set MyGUI(poGUI As V2ECKeyBoard.clsCarGUI)
    Set moGUI = poGUI
End Property
Public Property Get MyGUI() As V2ECKeyBoard.clsCarGUI
    Set MyGUI = moGUI
End Property

Public Property Let AssignmentsID(psAssignmentsID As String)
    msAssignmentsID = psAssignmentsID
End Property
Public Property Get AssignmentsID() As String
    AssignmentsID = msAssignmentsID
End Property

Public Property Let MyClaimsList(poForm As Object)
    Set MyfrmClaimsList = poForm
End Property
Public Property Set MyClaimsList(poForm As Object)
    Set MyfrmClaimsList = poForm
End Property
Public Property Get MyClaimsList() As Object
    Set MyClaimsList = MyfrmClaimsList
End Property

Public Property Let MyClaimInfo(poForm As Object)
    Set mfrmClaimInfo = poForm
End Property
Public Property Set MyClaimInfo(poForm As Object)
    Set mfrmClaimInfo = poForm
End Property
Public Property Get MyClaimInfo() As Object
    Set MyClaimInfo = mfrmClaimInfo
End Property

Public Property Let MyActivityLog(poForm As Object)
    Set mfrmActivityLog = poForm
End Property
Public Property Set MyActivityLog(poForm As Object)
    Set mfrmActivityLog = poForm
End Property
Public Property Get MyActivityLog() As Object
    Set MyActivityLog = mfrmActivityLog
End Property

Public Property Let MyIndemnity(poForm As Object)
    Set mfrmIndemnity = poForm
End Property
Public Property Set MyIndemnity(poForm As Object)
    Set mfrmIndemnity = poForm
End Property
Public Property Get MyIndemnity() As Object
    Set MyIndemnity = mfrmIndemnity
End Property

Public Property Let MyBillingInfo(poForm As Object)
    Set mfrmBillingInfo = poForm
End Property
Public Property Set MyBillingInfo(poForm As Object)
    Set mfrmBillingInfo = poForm
End Property
Public Property Get MyBillingInfo() As Object
    Set MyBillingInfo = mfrmBillingInfo
End Property

Public Property Let MyPhotos(poForm As Object)
    Set mfrmPhotos = poForm
End Property
Public Property Set MyPhotos(poForm As Object)
    Set mfrmPhotos = poForm
End Property
Public Property Get MyPhotos() As Object
    Set MyPhotos = mfrmPhotos
End Property

Public Property Let MyReports(poForm As Object)
    Set mfrmReports = poForm
End Property
Public Property Set MyReports(poForm As Object)
    Set mfrmReports = poForm
End Property
Public Property Get MyReports() As Object
    Set MyReports = mfrmReports
End Property

Public Property Let MyAttachments(poForm As Object)
    Set mfrmAttachments = poForm
End Property
Public Property Set MyAttachments(poForm As Object)
    Set mfrmAttachments = poForm
End Property
Public Property Get MyAttachments() As Object
    Set MyAttachments = mfrmAttachments
End Property

Public Property Let MyMiscellaneous(poForm As Object)
    Set mfrmMiscellaneous = poForm
End Property
Public Property Set MyMiscellaneous(poForm As Object)
    Set mfrmMiscellaneous = poForm
End Property
Public Property Get MyMiscellaneous() As Object
    Set MyMiscellaneous = mfrmMiscellaneous
End Property

Public Property Let MyPrint(poForm As Object)
    Set mfrmPrint = poForm
End Property
Public Property Set MyPrint(poForm As Object)
    Set mfrmPrint = poForm
End Property
Public Property Get MyPrint() As Object
    Set MyPrint = mfrmPrint
End Property


Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Public Property Get const_NO_SELECTION() As Long
    const_NO_SELECTION = NO_SELECTION
End Property


Private Sub cmdExit_Click()
    On Error GoTo EH

    If Not MyfrmClaimsList Is Nothing Then
        TimerReSize.Enabled = False
        MyfrmClaimsList.UnloadingClaim = True
        MyfrmClaimsList.Timer_UnloadClaim.Enabled = True
    End If

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdExit_Click"
End Sub

Private Sub cmdViewPDFLossReport_Click()
    On Error GoTo EH
    Dim oLR As V2ECKeyBoard.clsLossReports
    Dim MyadoRSAssignments As ADODB.Recordset
    Dim sLRFormat As String
    Dim sLRData As String
    Dim sCaption As String
    Dim sPDFFilePath As String
    Dim sIBNUM As String
    Dim sCLIENTNUM As String
    
    'check to see if this claim is currenlty unloading
    'if it is don' allow this event to occur
    If MyfrmClaimsList.UnloadingClaim Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    cmdViewPDFLossReport.Enabled = False
    Set MyadoRSAssignments = madoRSAssignments
    
    sIBNUM = goUtil.IsNullIsVbNullString(MyadoRSAssignments.Fields("IBNUM"))
    sCLIENTNUM = goUtil.IsNullIsVbNullString(MyadoRSAssignments.Fields("CLIENTNUM"))
    sCaption = "Loss Report " & Chr(160) & " (" & sIBNUM & "_" & sCLIENTNUM & ")"
    
    
    sLRFormat = goUtil.IsNullIsVbNullString(MyadoRSAssignments.Fields("LRFormat"))
    sLRData = goUtil.IsNullIsVbNullString(MyadoRSAssignments.Fields("LossReport"))
    
    If InStr(1, sLRFormat, "OLEType_pdf", vbTextCompare) > 0 Then
        sPDFFilePath = goUtil.gsInstallDir & "\AttachRepos\" & sLRData
        'Need to shell the PDF Loss Report to Adobe Reader
        goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sCaption
    Else
        sPDFFilePath = goUtil.gsInstallDir & "\TempLossReport" & goUtil.utGetTickCount & ".pdf"
        Set oLR = New V2ECKeyBoard.clsLossReports
        If StrComp(sLRFormat, "TEXT", vbTextCompare) <> 0 Then
            sLRData = sLRFormat & vbCrLf & sLRData
        End If
        oLR.CreateExport sLRData, sPDFFilePath, ARPdf
        If goUtil.utFileExists(sPDFFilePath) Then
            goUtil.utShellExecute , "OPEN", sPDFFilePath, , , vbNormalFocus, True, True, True, sCaption
            DoEvents
            Sleep 1000
            goUtil.utDeleteFile sPDFFilePath
        End If
    End If
    
    Set MyadoRSAssignments = Nothing
    Set oLR = Nothing
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    If Not MyfrmClaimsList.UnloadingClaim Then
        cmdViewPDFLossReport.Enabled = True
    End If
    Exit Sub
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdViewPDFLossReport_Click"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If MyfrmClaimsList Is Nothing Then
        Exit Sub
    End If
    MyfrmClaimsList.Visible = False
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
        goUtil.gfrmECTray.Visible = False
    Else
        goUtil.gfrmECTray.Visible = True
    End If
End Sub

Private Sub Form_GotFocus()
    On Error Resume Next
    MyfrmClaimsList.Visible = False
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
        goUtil.gfrmECTray.Visible = False
    Else
        goUtil.gfrmECTray.Visible = True
    End If
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH
    Dim sMess As String
    
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyF1
            If optClaim(GuiClaimOptions.opt01_ClaimInfo).Enabled Then
                optClaim(GuiClaimOptions.opt01_ClaimInfo).Value = True
            End If
        Case KeyCodeConstants.vbKeyF2
            If optClaim(GuiClaimOptions.opt02_ActivityLog).Enabled Then
                optClaim(GuiClaimOptions.opt02_ActivityLog).Value = True
            End If
        Case KeyCodeConstants.vbKeyF3
            If optClaim(GuiClaimOptions.opt03_Indemnity).Enabled Then
                optClaim(GuiClaimOptions.opt03_Indemnity).Value = True
            End If
        Case KeyCodeConstants.vbKeyF4
            If optClaim(GuiClaimOptions.opt05_Photos).Enabled Then
                optClaim(GuiClaimOptions.opt05_Photos).Value = True
            End If
        Case KeyCodeConstants.vbKeyF5
            If optClaim(GuiClaimOptions.opt06_Reports).Enabled Then
                optClaim(GuiClaimOptions.opt06_Reports).Value = True
            End If
        Case KeyCodeConstants.vbKeyF6
            If optClaim(GuiClaimOptions.opt07_Attachments).Enabled Then
                optClaim(GuiClaimOptions.opt07_Attachments).Value = True
            End If
        Case KeyCodeConstants.vbKeyF7
            If optClaim(GuiClaimOptions.opt04_BillingInformation).Enabled Then
                optClaim(GuiClaimOptions.opt04_BillingInformation).Value = True
            End If
        Case KeyCodeConstants.vbKeyF8
           If optClaim(GuiClaimOptions.opt08_Miscellaneous).Enabled Then
                optClaim(GuiClaimOptions.opt08_Miscellaneous).Value = True
            End If
        Case KeyCodeConstants.vbKeyF9
            If optClaim(GuiClaimOptions.opt09_Print).Enabled Then
                optClaim(GuiClaimOptions.opt09_Print).Value = True
            End If
        Case KeyCodeConstants.vbKeyF10
            ' not used yet
        Case KeyCodeConstants.vbKeyF11
            'not used yet
        Case KeyCodeConstants.vbKeyF12
            sMess = "Are you sure you want to Exit Claim Maintenance?"
            If MsgBox(sMess, vbQuestion + vbOKCancel, "Exit Claim Maintenance") = vbOK Then
                If Not MyfrmClaimsList Is Nothing Then
                    MyfrmClaimsList.UnloadingClaim = True
                    MyfrmClaimsList.Timer_UnloadClaim.Enabled = True
                End If
            End If
            
    End Select
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub Form_KeyDown"
End Sub


Private Sub Form_Load()
    On Error GoTo EH
    Dim lOptCount As Long
    Dim iMyStatus As V2ECKeyBoard.AssgnStatus
    Dim bStatusClosed As Boolean
    
    
    'LET FTP know that the currently SELECTED Assignment is LOCKED
    'that Means FTP will exlcude Updates to this ID until the user
    'Exits this Assignment
    SaveSetting "ECFTP", "MSG", "LOCKED_AssignmentsID", msAssignmentsID
    
    'Init the Option selectin
    
    mbLoading = True
    SetadoRSAssignments msAssignmentsID
    iMyStatus = madoRSAssignments.Fields("StatusID").Value
    MyStatus = iMyStatus
    Select Case iMyStatus
        Case V2ECKeyBoard.AssgnStatus.iAssignmentsStatus_CLOSED
            bStatusClosed = True
    End Select
    
    For lOptCount = optClaim.LBound To optClaim.UBound
        If optClaim(lOptCount).Tag = vbNullString Or bStatusClosed Then
            If optClaim(lOptCount).Tag <> "frmClaimInfo" And optClaim(lOptCount).Tag <> "frmPrint" Then
                optClaim(lOptCount).Enabled = False
                optClaim(lOptCount).BackColor = BG_GRAY
            End If
        End If
    Next
    
    
    TimerReSize.Enabled = True
    
    mbLoading = False
    
    optClaim(0).Value = True
    
    Exit Sub
EH:
    mbLoading = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Load"
End Sub

Public Function EnableMyOptions(pbEnable As Boolean, Optional piSkipOption As GuiClaimOptions = -1) As Boolean
    On Error GoTo EH
    
    Dim lOptCount As Long
    
    For lOptCount = optClaim.LBound To optClaim.UBound
        If optClaim(lOptCount).Tag <> vbNullString And lOptCount <> piSkipOption Then
            If MyStatus = iAssignmentsStatus_CLOSED Then
                If lOptCount = GuiClaimOptions.opt01_ClaimInfo Then
                    optClaim(lOptCount).Enabled = pbEnable
                End If
            Else
                optClaim(lOptCount).Enabled = pbEnable
            End If
        End If
    Next
    
    EnableMyOptions = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EnableMyOptions"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo EH
    Dim sMess As String
    
    Select Case UnloadMode
        Case vbFormControlMenu
            Cancel = True
'            sMess = "Are you sure you want to Exit Claim Maintenance?"
'            If MsgBox(sMess, vbQuestion + vbOKCancel, "Exit Claim Maintenance") <> vbOK Then
'                Exit Sub
'            End If
            If Not MyfrmClaimsList Is Nothing Then
                TimerReSize.Enabled = False
                MyfrmClaimsList.UnloadingClaim = True
                MyfrmClaimsList.Timer_UnloadClaim.Enabled = True
            End If
    End Select

    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    ReSizeMe
End Sub

Public Sub ReSizeMe()
    On Error Resume Next
    If MyfrmClaimsList Is Nothing Then
        Exit Sub
    End If
    Dim sNavScreenPos As String
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = False
    End If
    
    Me.top = 10
    Me.left = 10
    Me.Width = Screen.Width - 10
    Me.Height = 1260
    cmdExit.left = Me.Width - 1125
    MyfrmClaimsList.Visible = False
    
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
        goUtil.gfrmECTray.Visible = False
    Else
        goUtil.gfrmECTray.Visible = True
    End If
End Sub


Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    UnloadMyForms True
    
    Set MyfrmClaimsList = Nothing
    Set MyGUI = Nothing
    Set madoRSAssignments = Nothing
    Set madoRSPolicyLimits = Nothing
    Set madoRSBillingCount = Nothing
    Set madoRSRTPhotoLog = Nothing
    Set madoRSRTActivityLog = Nothing
    Set madoRSRTActivityLogInfo = Nothing
    Set madoRSRTIndemnity = Nothing
    Set madoRSRTChecks = Nothing
    Set madoRSRTAttachments = Nothing
    Set madoRSMainReports = Nothing
    Set madoRSMainReportsHistory = Nothing
    Set madoRSCarSpecReports = Nothing
    Set madoRSCarSpecReportsHistory = Nothing
    Set madoRSIB = Nothing
    Set madoRSBillingReports = Nothing
    Set madoRSBillingReportsHistory = Nothing
    Set madoRSPayment = Nothing
    Set madoRSPaymentReports = Nothing
    Set madoRSPaymentReportsHistory = Nothing
    Set madoRSWordXLDocs = Nothing
    Set madoRSRTPhotoReportList = Nothing
    Set madoRSRTWSDiagramList = Nothing
    Set madoRSPackageList = Nothing
    Set madoRSPackageItem = Nothing
    Set madoRSFeeScheduleList = Nothing
    Set madoRSBillingCountItem = Nothing
    Set madoRSCLIB = Nothing
    Set madoRSCLIBFee = Nothing
    Set madoRSRTIB = Nothing
    Set madoRSRTIBFee = Nothing
    Set madoRSClientCOCat = Nothing
    
    If Not moForm Is Nothing Then
        Unload moForm
        Set moForm = Nothing
    End If
        
    If Not mArv Is Nothing Then
        mArv.CLEANUP
        Set mArv = Nothing
    End If
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function CLEANUP"
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    'LET FTP know that there is no currently locked Assignment
    SaveSetting "ECFTP", "MSG", "LOCKED_AssignmentsID", vbNullString
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Form_Unload"
End Sub

Private Sub optClaim_Click(Index As Integer)
    On Error GoTo EH
    Dim lOptCount As Long
    Dim oForm As Form
    Dim sIBNUM As String
    Dim sCLIENTNUM As String
    
    'Populate some Variables here
    sIBNUM = goUtil.IsNullIsVbNullString(madoRSAssignments.Fields("IBNUM"))
    sCLIENTNUM = goUtil.IsNullIsVbNullString(madoRSAssignments.Fields("CLIENTNUM"))
    
    Select Case Index
        Case GuiClaimOptions.opt01_ClaimInfo
            If mfrmClaimInfo Is Nothing Then
                Set mfrmClaimInfo = New frmClaimInfo
                Set mfrmClaimInfo.MyfrmClaim = Me
                Set mfrmClaimInfo.MyGUI = Me.MyGUI
                mfrmClaimInfo.AssignmentsID = Me.AssignmentsID
                Load mfrmClaimInfo
            End If
            Set oForm = mfrmClaimInfo
            
        Case GuiClaimOptions.opt02_ActivityLog
            If mfrmActivityLog Is Nothing Then
                Set mfrmActivityLog = New frmActivityLog
                Set mfrmActivityLog.MyfrmClaim = Me
                Set mfrmActivityLog.MyGUI = Me.MyGUI
                mfrmActivityLog.AssignmentsID = Me.AssignmentsID
                Load mfrmActivityLog
            End If
            Set oForm = mfrmActivityLog
            
        Case GuiClaimOptions.opt03_Indemnity
            If mfrmIndemnity Is Nothing Then
                Set mfrmIndemnity = New frmIndemnity
                Set mfrmIndemnity.MyfrmClaim = Me
                Set mfrmIndemnity.MyGUI = Me.MyGUI
                mfrmIndemnity.AssignmentsID = Me.AssignmentsID
                Load mfrmIndemnity
            End If
            Set oForm = mfrmIndemnity
            
        Case GuiClaimOptions.opt04_BillingInformation
            If mfrmBillingInfo Is Nothing Then
                Set mfrmBillingInfo = New frmBillingInfo
                Set mfrmBillingInfo.MyfrmClaim = Me
                Set mfrmBillingInfo.MyGUI = Me.MyGUI
                mfrmBillingInfo.AssignmentsID = Me.AssignmentsID
                Load mfrmBillingInfo
            End If
            Set oForm = mfrmBillingInfo
            
        Case GuiClaimOptions.opt05_Photos
            If mfrmPhotos Is Nothing Then
                Set mfrmPhotos = New frmPhotos
                Set mfrmPhotos.MyfrmClaim = Me
                Set mfrmPhotos.MyGUI = Me.MyGUI
                mfrmPhotos.AssignmentsID = Me.AssignmentsID
                mfrmPhotos.IBNUM = sIBNUM
                Load mfrmPhotos
            End If
            Set oForm = mfrmPhotos
            
        Case GuiClaimOptions.opt06_Reports
            If mfrmReports Is Nothing Then
                Set mfrmReports = New frmReports
                Set mfrmReports.MyfrmClaim = Me
                Set mfrmReports.MyGUI = Me.MyGUI
                mfrmReports.AssignmentsID = Me.AssignmentsID
                mfrmReports.IBNUM = sIBNUM
                Load mfrmReports
            End If
            Set oForm = mfrmReports
            
        Case GuiClaimOptions.opt07_Attachments
            If mfrmAttachments Is Nothing Then
                Set mfrmAttachments = New frmAttachments
                Set mfrmAttachments.MyfrmClaim = Me
                Set mfrmAttachments.MyGUI = Me.MyGUI
                mfrmAttachments.AssignmentsID = Me.AssignmentsID
                Load mfrmAttachments
            End If
            Set oForm = mfrmAttachments
            
        Case GuiClaimOptions.opt08_Miscellaneous
            If mfrmMiscellaneous Is Nothing Then
                Set mfrmMiscellaneous = New frmMiscellaneous
                Set mfrmMiscellaneous.MyfrmClaim = Me
                Set mfrmMiscellaneous.MyGUI = Me.MyGUI
                mfrmMiscellaneous.AssignmentsID = Me.AssignmentsID
                Load mfrmMiscellaneous
            End If
            Set oForm = mfrmMiscellaneous
            
        Case GuiClaimOptions.opt09_Print
            'unload all other forms
            UnloadMyForms True, opt09_Print
            EnableMyOptions False, opt09_Print
            If mfrmPrint Is Nothing Then
                Set mfrmPrint = New frmPrint
                Set mfrmPrint.MyfrmClaim = Me
                Set mfrmPrint.MyGUI = Me.MyGUI
                mfrmPrint.AssignmentsID = Me.AssignmentsID
                Load mfrmPrint
            End If
            Set oForm = mfrmPrint
    End Select
    
    'check for Max or min
    If oForm.WindowState = vbMaximized Or oForm.WindowState = vbMinimized Then
        oForm.WindowState = vbNormal
    End If
    
    If MyfrmClaimsList Is Nothing Then
        GoTo CLEAN_UP
    End If
    
    MyfrmClaimsList.PopulateFrmCaptionAssignmentInfo oForm, oForm.Tag
    oForm.top = Me.top + Me.Height
    oForm.left = 10
    oForm.Width = Screen.Width - 10
    oForm.Height = Screen.Height - (10 + Me.Height + goUtil.utGetTaskbarHeight)
    oForm.Show vbModeless
    optClaim(Index).BackColor = BG_WHITE
CLEAN_UP:
    'cleanup
    Set oForm = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub optClaim_Click"
End Sub

Private Sub Timer_UnloadForm_Timer()
    On Error GoTo EH
    Timer_UnloadForm.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    UnloadMyForms
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
EH:
    Screen.MousePointer = vbNormal
    Timer_UnloadForm.Enabled = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub Timer_UnloadForm_Timer"
End Sub

Private Function UnloadMyForms(Optional pbForceUnload As Boolean = False, Optional piSkipOption As GuiClaimOptions = -1) As Boolean
    On Error GoTo EH
    'Claim Info
    If Not mfrmClaimInfo Is Nothing And piSkipOption <> opt01_ClaimInfo Then
        If mfrmClaimInfo.UnloadMe Or pbForceUnload Then
            mfrmClaimInfo.UnloadMe = True
            Unload mfrmClaimInfo
            Set mfrmClaimInfo = Nothing
            Me.optClaim(GuiClaimOptions.opt01_ClaimInfo).Value = False
            optClaim(GuiClaimOptions.opt01_ClaimInfo).BackColor = BG_GRAY
        End If
    End If
    
    'Activity Log
    If Not mfrmActivityLog Is Nothing And piSkipOption <> opt02_ActivityLog Then
        If mfrmActivityLog.UnloadMe Or pbForceUnload Then
            mfrmActivityLog.UnloadMe = True
            Unload mfrmActivityLog
            Set mfrmActivityLog = Nothing
            Me.optClaim(GuiClaimOptions.opt02_ActivityLog).Value = False
            optClaim(GuiClaimOptions.opt02_ActivityLog).BackColor = BG_GRAY
        End If
    End If
    
    'Indemnity
    If Not mfrmIndemnity Is Nothing And piSkipOption <> opt03_Indemnity Then
        If mfrmIndemnity.UnloadMe Or pbForceUnload Then
            mfrmIndemnity.UnloadMe = True
            Unload mfrmIndemnity
            Set mfrmIndemnity = Nothing
            Me.optClaim(GuiClaimOptions.opt03_Indemnity).Value = False
            optClaim(GuiClaimOptions.opt03_Indemnity).BackColor = BG_GRAY
        End If
    End If
    
    'Billing Information
    If Not mfrmBillingInfo Is Nothing And piSkipOption <> opt04_BillingInformation Then
        If mfrmBillingInfo.UnloadMe Or pbForceUnload Then
            mfrmBillingInfo.UnloadMe = True
            Unload mfrmBillingInfo
            Set mfrmBillingInfo = Nothing
            Me.optClaim(GuiClaimOptions.opt04_BillingInformation).Value = False
            optClaim(GuiClaimOptions.opt04_BillingInformation).BackColor = BG_GRAY
        End If
    End If
    
    'Photos
    If Not mfrmPhotos Is Nothing And piSkipOption <> opt05_Photos Then
        If mfrmPhotos.UnloadMe Or pbForceUnload Then
            mfrmPhotos.UnloadMe = True
            Unload mfrmPhotos
            Set mfrmPhotos = Nothing
            Me.optClaim(GuiClaimOptions.opt05_Photos).Value = False
            optClaim(GuiClaimOptions.opt05_Photos).BackColor = BG_GRAY
        End If
    End If
    
    'Reports  mfrmReports
    If Not mfrmReports Is Nothing And piSkipOption <> opt06_Reports Then
        If mfrmReports.UnloadMe Or pbForceUnload Then
            mfrmReports.UnloadMe = True
            Unload mfrmReports
            Set mfrmReports = Nothing
            Me.optClaim(GuiClaimOptions.opt06_Reports).Value = False
            optClaim(GuiClaimOptions.opt06_Reports).BackColor = BG_GRAY
        End If
    End If
    
    'Attachments  mfrmAttachments
    If Not mfrmAttachments Is Nothing And piSkipOption <> opt07_Attachments Then
        If mfrmAttachments.UnloadMe Or pbForceUnload Then
            mfrmAttachments.UnloadMe = True
            Unload mfrmAttachments
            Set mfrmAttachments = Nothing
            Me.optClaim(GuiClaimOptions.opt07_Attachments).Value = False
            optClaim(GuiClaimOptions.opt07_Attachments).BackColor = BG_GRAY
        End If
    End If
    
    'Miscellaneous
    If Not mfrmMiscellaneous Is Nothing And piSkipOption <> opt08_Miscellaneous Then
        If mfrmMiscellaneous.UnloadMe Or pbForceUnload Then
            mfrmMiscellaneous.UnloadMe = True
            Unload mfrmMiscellaneous
            Set mfrmMiscellaneous = Nothing
            Me.optClaim(GuiClaimOptions.opt08_Miscellaneous).Value = False
            optClaim(GuiClaimOptions.opt08_Miscellaneous).BackColor = BG_GRAY
        End If
    End If
    
    'Print
    If Not mfrmPrint Is Nothing And piSkipOption <> opt09_Print Then
        If mfrmPrint.UnloadMe Or pbForceUnload Then
            mfrmPrint.UnloadMe = True
            Unload mfrmPrint
            Set mfrmPrint = Nothing
            Me.optClaim(GuiClaimOptions.opt09_Print).Value = False
            optClaim(GuiClaimOptions.opt09_Print).BackColor = BG_GRAY
            EnableMyOptions True
        End If
    End If
    
    UnloadMyForms = True
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Function UnloadMyForms"
End Function

Public Function ShowCalendar(poTextBox As Object) As Boolean
    On Error GoTo EH
    
    ShowCalendar = MyGUI.ShowCalendar(poTextBox)

    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function ShowCalendar"
End Function

Public Function SetadoRSAssignments(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    'rese the typeof loss rs
    If Not madoRSAssignments Is Nothing Then
        Set madoRSAssignments = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSAssignments = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM Assignments "
    sSQL = sSQL & "WHERE ID = " & psIDAssignments & " "
    
    madoRSAssignments.CursorLocation = adUseClient
    madoRSAssignments.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSAssignments.ActiveConnection = Nothing
    
    SetadoRSAssignments = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSAssignments = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSAssignments"
End Function

Public Function RefreshMe() As Boolean
    On Error GoTo EH
    
    MyfrmClaimsList.RefreshMe
    SetadoRSAssignments msAssignmentsID
    'Claim Info
    If Not mfrmClaimInfo Is Nothing Then
        mfrmClaimInfo.LoadMe
    End If
    
    'Activity Log
    If Not mfrmActivityLog Is Nothing Then
        mfrmActivityLog.LoadMe
    End If
    
    'Indemnity
    If Not mfrmIndemnity Is Nothing Then
        mfrmIndemnity.LoadMe
    End If
    
    'Billing Information
    If Not mfrmBillingInfo Is Nothing Then
        mfrmBillingInfo.LoadMe
    End If
    
    'Photos
    If Not mfrmPhotos Is Nothing Then
        mfrmPhotos.LoadMe
    End If
    
    'Reports  mfrmReports
    If Not mfrmReports Is Nothing Then
        mfrmReports.LoadMe
    End If
    
    'Attachments  mfrmAttachments
    If Not mfrmAttachments Is Nothing Then
        mfrmAttachments.LoadMe
    End If
    
    'Miscellaneous
    If Not mfrmMiscellaneous Is Nothing Then
        mfrmMiscellaneous.LoadMe
    End If
    
    'Print
    If Not mfrmPrint Is Nothing Then
        mfrmPrint.LoadMe
    End If
    
    RefreshMe = True
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function RefreshMe"
End Function

Public Function SetadoRSPolicyLimits(psIDAssignments As String, _
                                        Optional pbUseAppDedClassTypeIDOrder As Boolean = False) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sAppDedClassTypeIDOrder As String
    Dim RS As ADODB.Recordset
    
    'reset RS
    If Not madoRSPolicyLimits Is Nothing Then
        Set madoRSPolicyLimits = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSPolicyLimits = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
     If pbUseAppDedClassTypeIDOrder Then
        Set RS = madoRSAssignments
        If RS.RecordCount = 1 Then
            RS.MoveFirst
            sAppDedClassTypeIDOrder = goUtil.IsNullIsVbNullString(RS.Fields("AppDedClassTypeIDOrder"))
            Set RS = Nothing
        End If
     End If
    
    
    sSQL = "SELECT RetPolicyLimits.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[PolicyLimitsID], "
    sSQL = sSQL & "[AssignmentsID] , "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[ClassTypeID] , "
    sSQL = sSQL & "(SELECT    Class "
    sSQL = sSQL & "FROM   ClassType "
    sSQL = sSQL & "WHERE  ClassTypeID = S.[ClassTypeID]) As [ClassTypeClass],  "
    sSQL = sSQL & "(SELECT    Description "
    sSQL = sSQL & "FROM   ClassType "
    sSQL = sSQL & "WHERE  ClassTypeID = S.[ClassTypeID]) As [ClassTypeDescription],  "
    sSQL = sSQL & "[LimitAmount] , "
    sSQL = sSQL & "[RCSaidProp] , "
    sSQL = sSQL & "[Reserves] , "
    sSQL = sSQL & "[IsDeleted] , "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = S.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID]  "
    sSQL = sSQL & "FROM PolicyLimits S "
    If pbUseAppDedClassTypeIDOrder Then
        If sAppDedClassTypeIDOrder = vbNullString Then
            sSQL = sSQL & "WHERE S.[ClassTypeID] Is Null "
        Else
            sSQL = sSQL & "WHERE S.[ClassTypeID] IN (" & sAppDedClassTypeIDOrder & ") "
        End If
    End If
    sSQL = sSQL & ") RetPolicyLimits "
    sSQL = sSQL & "WHERE IDAssignments = " & psIDAssignments & " "
    sSQL = sSQL & "ORDER BY ClassTypeClass "

    'Use Disconnected RS Only
    madoRSPolicyLimits.CursorLocation = adUseClient
    madoRSPolicyLimits.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSPolicyLimits.ActiveConnection = Nothing
    
    SetadoRSPolicyLimits = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSPolicyLimits = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSPolicyLimits"
End Function

Public Function SetadoRSBillingCount(psIDAssignments As String, _
                                    Optional pbSumActLogServiceTime As Boolean, _
                                    Optional pbCountPhotoLog As Boolean, _
                                    Optional pbSumAmountOfCheck As Boolean) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    'reset RS
    If Not madoRSBillingCount Is Nothing Then
        Set madoRSBillingCount = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSBillingCount = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetBillingCount.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[AssignmentsID] , "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "'IB' & IIF([Supplement] > 0 ,'S' & Format([Supplement],'00') ,'' ) & IIF([Rebill] > 0 ,'R' & Format([Rebill],'00') ,'' ) "
    sSQL = sSQL & ") As IB, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   ( "
                            sSQL = sSQL & "IIF((Not IsNull(RTIB.IDBillingCount) And Not IsDate(RTIB.RT06_dtDateClosed)), "
                            'then
                            sSQL = sSQL & "'Current', " '
                            'else
                            sSQL = sSQL & "'Closed - [' & Format(RTIB.RT06_dtDateClosed, 'MM/DD/YYYY') & ']') "
                            'If Flagged need to add some stuff to the Description.
                            If pbSumActLogServiceTime Then
                                'Activity log Need to Sum the Service Time
                                'Associated with the IB.
                                sSQL = sSQL & "& ' Service Time - [' & "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   (IIF(IsNull(SUM(RTAL.[ServiceTime])),'0.00', Format(SUM(RTAL.[ServiceTime]),'0.00'))) As SumOfServiceTime "
                                sSQL = sSQL & "FROM     RTActivityLog RTAL "
                                sSQL = sSQL & "WHERE    RTAL.IDAssignments = BC.IDAssignments "
                                sSQL = sSQL & "AND      RTAL.IDBillingCount = BC.ID "
                                sSQL = sSQL & "AND      RTAL.IsDeleted = False "
                                sSQL = sSQL & ") "
                                sSQL = sSQL & "& ']' "
                            ElseIf pbCountPhotoLog Then
                                'Photo log need to count up the number
                                'of photos associated with the IB.
                                sSQL = sSQL & "& ' Photo Count - [' & "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   (IIF(IsNull(Count(RTPL.[ID])), '0',Count(RTPL.[ID]))) As CountOfID "
                                sSQL = sSQL & "FROM     RTPhotoLog RTPL "
                                sSQL = sSQL & "WHERE    RTPL.IDAssignments = BC.IDAssignments "
                                sSQL = sSQL & "AND      RTPL.IDBillingCount = BC.ID "
                                sSQL = sSQL & "AND      RTPL.IsDeleted = False "
                                sSQL = sSQL & ") "
                                sSQL = sSQL & "& ']' "
                            ElseIf pbSumAmountOfCheck Then
                                'Payment Request (RTChecks) Need to Sum the Amount Of Check
                                'Associated with the IB.
                                sSQL = sSQL & "& ' Pay Reqs Total - [' & "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   (IIF(IsNull(SUM(RTC.[RT53_cAmountOfCheck])),'0.00', Format(SUM(RTC.[RT53_cAmountOfCheck]),'0.00'))) As SumOfPayReqs "
                                sSQL = sSQL & "FROM     RTChecks RTC "
                                sSQL = sSQL & "WHERE    RTC.IDAssignments = BC.IDAssignments "
                                sSQL = sSQL & "AND      RTC.IDBillingCount = BC.ID "
                                sSQL = sSQL & "AND      RTC.IsDeleted = False "
                                sSQL = sSQL & ") "
                                sSQL = sSQL & "& ']' "
                            End If
            sSQL = sSQL & ") As Desc1 "
    sSQL = sSQL & "FROM     RTIB "
    sSQL = sSQL & "WHERE    RTIB.IDBillingCount = BC.ID "
    sSQL = sSQL & "AND      RTIB.IDAssignments = BC.IDAssignments "
    sSQL = sSQL & ") As IBDescription, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT   ( "
                            sSQL = sSQL & "IIF((Not IsNull(IB.IDBillingCount) And Not IsDate(IB.IB06_dtDateClosed)), "
                            'then
                            sSQL = sSQL & "'Current', " '
                            'else
                            sSQL = sSQL & "'Closed - [' & Format(IB.IB06_dtDateClosed, 'MM/DD/YYYY') & ']') "
                            'If Flagged need to add some stuff to the Description.
                            If pbSumActLogServiceTime Then
                                'Activity log Need to Sum the Service Time
                                'Associated with the IB.
                                sSQL = sSQL & "& ' Service Time - [' & "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   (IIF(IsNull(SUM(RTAL.[ServiceTime])),'0.00', Format(SUM(RTAL.[ServiceTime]),'0.00'))) As SumOfServiceTime "
                                sSQL = sSQL & "FROM     RTActivityLog RTAL "
                                sSQL = sSQL & "WHERE    RTAL.IDAssignments = BC.IDAssignments "
                                sSQL = sSQL & "AND      RTAL.IDBillingCount = BC.ID "
                                sSQL = sSQL & "AND      RTAL.IsDeleted = False "
                                sSQL = sSQL & ") "
                                sSQL = sSQL & "& ']' "
                            ElseIf pbCountPhotoLog Then
                                'Photo log need to count up the number
                                'of photos associated with the IB.
                                sSQL = sSQL & "& ' Photo Count - [' & "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   (IIF(IsNull(Count(RTPL.[ID])), '0',Count(RTPL.[ID]))) As CountOfID "
                                sSQL = sSQL & "FROM     RTPhotoLog RTPL "
                                sSQL = sSQL & "WHERE    RTPL.IDAssignments = BC.IDAssignments "
                                sSQL = sSQL & "AND      RTPL.IDBillingCount = BC.ID "
                                sSQL = sSQL & "AND      RTPL.IsDeleted = False "
                                sSQL = sSQL & ") "
                                sSQL = sSQL & "& ']' "
                            ElseIf pbSumAmountOfCheck Then
                                'Payment Request (RTChecks) Need to Sum the Amount Of Check
                                'Associated with the IB.
                                sSQL = sSQL & "& ' Pay Reqs Total - [' & "
                                sSQL = sSQL & "( "
                                sSQL = sSQL & "SELECT   (IIF(IsNull(SUM(RTC.[RT53_cAmountOfCheck])),'0.00', Format(SUM(RTC.[RT53_cAmountOfCheck]),'0.00'))) As SumOfPayReqs "
                                sSQL = sSQL & "FROM     RTChecks RTC "
                                sSQL = sSQL & "WHERE    RTC.IDAssignments = BC.IDAssignments "
                                sSQL = sSQL & "AND      RTC.IDBillingCount = BC.ID "
                                sSQL = sSQL & "AND      RTC.IsDeleted = False "
                                sSQL = sSQL & ") "
                                sSQL = sSQL & "& ']' "
                            End If
    sSQL = sSQL & ") As Desc2 "
    sSQL = sSQL & "FROM     IB "
    sSQL = sSQL & "WHERE    IB.IDBillingCount = BC.ID "
    sSQL = sSQL & "AND      IB.IDAssignments = BC.IDAssignments "
    sSQL = sSQL & ") As IBDescription2, "
    sSQL = sSQL & "[Rebill], "
    sSQL = sSQL & "[Supplement], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = BC.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID]  "
    sSQL = sSQL & "FROM BillingCount BC "
    sSQL = sSQL & ") RetBillingCount "
    sSQL = sSQL & "WHERE [IDAssignments] = " & psIDAssignments & " "
    sSQL = sSQL & "ORDER BY [Supplement], [Rebill] "

    'Use Disconnected RS Only
    madoRSBillingCount.CursorLocation = adUseClient
    madoRSBillingCount.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSBillingCount.ActiveConnection = Nothing
    
    SetadoRSBillingCount = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSBillingCount = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSBillingCount"
End Function

Public Function SetadoRSRTPhotoLog(psIDAssignments As String, psIDRTPhotoReportID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSRTPhotoLog Is Nothing Then
        Set madoRSRTPhotoLog = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTPhotoLog = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetRTPhotoLog.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[RTPhotoLogID], "
    sSQL = sSQL & "[RTPhotoReportID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDRTPhotoReport], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDBillingCount], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT 'IB' & IIF(BC.Supplement > 0 ,'S' & Format(BC.Supplement,'00') ,'' ) & IIF(BC.Rebill > 0 ,'R' & Format(BC.Rebill,'00') ,'' ) "
    sSQL = sSQL & "FROM BillingCount BC "
    sSQL = sSQL & "WHERE    BC.ID = P.IDBillingCount "
    sSQL = sSQL & "AND      BC.IDAssignments = P.IDAssignments "
    sSQL = sSQL & ") As IB, "
    sSQL = sSQL & "[PhotoDate], "
    sSQL = sSQL & "[SortOrder], "
    sSQL = sSQL & "TRIM([Description]) As [Description], "
    sSQL = sSQL & "TRIM([PhotoName]) AS [PhotoName], "
    sSQL = sSQL & "[Photo], "
    sSQL = sSQL & "[DownloadPhoto], "
    sSQL = sSQL & "[UpLoadPhoto], "
    sSQL = sSQL & "[PhotoThumb], "
    sSQL = sSQL & "[DownloadPhotoThumb], "
    sSQL = sSQL & "[UpLoadPhotoThumb], "
    sSQL = sSQL & "[PhotoHighRes], "
    sSQL = sSQL & "[DownloadPhotoHighRes], "
    sSQL = sSQL & "[UploadPhotoHighRes], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "TRIM([AdminComments]) As [AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = P.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "(SELECT    IBNUM "
    sSQL = sSQL & "FROM   Assignments "
    sSQL = sSQL & "WHERE  ID = P.[IDAssignments]) As [IBNUM],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM RTPhotoLog P "
    sSQL = sSQL & ") RetRTPhotoLog "
    sSQL = sSQL & "WHERE IDAssignments = " & psIDAssignments & " "
    If psIDRTPhotoReportID <> "0" Then
        sSQL = sSQL & "AND P.IDRTPhotoReport = " & psIDRTPhotoReportID & " "
    End If
    If CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True)) Then
        sSQL = sSQL & "AND P.IsDeleted = False "
    End If
    sSQL = sSQL & "ORDER BY SortOrder "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTPhotoLog.CursorLocation = adUseClient
    madoRSRTPhotoLog.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTPhotoLog.ActiveConnection = Nothing
    
    SetadoRSRTPhotoLog = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTPhotoLog = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTPhotoLog"
End Function

Public Function SetadoRSRTActivityLogInfo(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
     If Not madoRSRTActivityLogInfo Is Nothing Then
        Set madoRSRTActivityLogInfo = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTActivityLogInfo = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetRTActivityLogInfo.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[AL01_sPresentDurringInspection], "
    sSQL = sSQL & "[AL02_sExplainedEstimate], "
    sSQL = sSQL & "[AL03_sExplainedRCV], "
    sSQL = sSQL & "[AL03_sExplainedRCVNA], "
    sSQL = sSQL & "[AL04_sConfirmMortgageeIsCorrect], "
    sSQL = sSQL & "[AL04_sConfirmMortgageeIsCorrectNA], "
    sSQL = sSQL & "[AL05_sExplainedMortgageeChecks], "
    sSQL = sSQL & "[AL05_sExplainedMortgageeChecksNA], "
    sSQL = sSQL & "[AL06_sConfirmedCoverage], "
    sSQL = sSQL & "[AL07_sPriorLoss], "
    sSQL = sSQL & "[AL07_sPriorLossNA], "
    sSQL = sSQL & "[AL08_sSalvage], "
    sSQL = sSQL & "[AL09_sSubrogation], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "TRIM([AdminComments]) As [AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = A.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "(SELECT    IBNUM "
    sSQL = sSQL & "FROM   Assignments "
    sSQL = sSQL & "WHERE  ID = A.[IDAssignments]) As [IBNUM],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM RTActivityLogInfo A "
    sSQL = sSQL & ") RetRTActivityLogInfo "
    sSQL = sSQL & "WHERE [IDAssignments] = " & psIDAssignments & " "
    sSQL = sSQL & "AND [IsDeleted] = False "


    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTActivityLogInfo.CursorLocation = adUseClient
    madoRSRTActivityLogInfo.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTActivityLogInfo.ActiveConnection = Nothing
    
    SetadoRSRTActivityLogInfo = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTActivityLogInfo = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTActivityLogInfo"
End Function

Public Function SetadoRSRTActivityLog(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSRTActivityLog Is Nothing Then
        Set madoRSRTActivityLog = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTActivityLog = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetRTActivityLog.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[RTActivityLogID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[IDBillingCount], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT 'IB' & IIF(BC.Supplement > 0 ,'S' & Format(BC.Supplement,'00') ,'' ) & IIF(BC.Rebill > 0 ,'R' & Format(BC.Rebill,'00') ,'' ) "
    sSQL = sSQL & "FROM BillingCount BC "
    sSQL = sSQL & "WHERE    BC.ID = A.IDBillingCount "
    sSQL = sSQL & "AND      BC.IDAssignments = A.IDAssignments "
    sSQL = sSQL & ") As IB, "
    sSQL = sSQL & "[ServiceTime], "
    sSQL = sSQL & "[ActDate], "
    sSQL = sSQL & "[ActText], "
    sSQL = sSQL & "[ActTime], "
    sSQL = sSQL & "[PageBreakAfter], "
    sSQL = sSQL & "[BlankPageAfter], "
    sSQL = sSQL & "[BlankRowsAfter], "
    sSQL = sSQL & "[IsMgrEntry], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "TRIM([AdminComments]) As [AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = A.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "(SELECT    IBNUM "
    sSQL = sSQL & "FROM   Assignments "
    sSQL = sSQL & "WHERE  ID = A.[IDAssignments]) As [IBNUM],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM RTActivityLog A "
    sSQL = sSQL & ") RetRTActivityLog "
    sSQL = sSQL & "WHERE [IDAssignments] = " & psIDAssignments & " "
    If CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True)) Then
        sSQL = sSQL & "AND [IsDeleted] = False "
    End If
    sSQL = sSQL & "ORDER BY [ActTime], ABS([ID]) "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTActivityLog.CursorLocation = adUseClient
    madoRSRTActivityLog.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTActivityLog.ActiveConnection = Nothing
    
    SetadoRSRTActivityLog = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTActivityLog = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTActivityLog"
End Function

Public Function SetadoRSRTIndemnity(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    SetadoRSPaymentReports
    Set RS = madoRSPaymentReports ' Software
    If RS.RecordCount = 1 Then
        RS.MoveFirst
    Else
        Exit Function
    End If
    
    If Not madoRSRTIndemnity Is Nothing Then
        Set madoRSRTIndemnity = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTIndemnity = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetRTIndemnity.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "RTI.[RTIndemnityID], "
    sSQL = sSQL & "RTI.[AssignmentsID], "
    sSQL = sSQL & "RTI.[RTChecksID], "
    sSQL = sSQL & "RTI.[ID], "
    sSQL = sSQL & "RTI.[IDAssignments], "
    sSQL = sSQL & "RTI.[IDRTChecks], "
    sSQL = sSQL & "( "
'    sSQL = sSQL & "IIF(RTC.CheckNum Is Not Null, "
'    'Then
'    sSQL = sSQL & "Cstr(RTC.CheckNum) & ' Of ' & "
'    sSQL = sSQL & "( "
'    sSQL = sSQL & "SELECT MAX(CheckNum) "
'    sSQL = sSQL & "FROM RTChecks "
'    sSQL = sSQL & "WHERE [AssignmentsID] = " & psIDAssignments & " "
'    sSQL = sSQL & "), "
'    'else
'    sSQL = sSQL & "'') "
    sSQL = sSQL & "IIF(RTC.CheckNum Is Not Null, "
    'Then
    sSQL = sSQL & "Cstr(RTC.CheckNum) "
    sSQL = sSQL & ", "
    'else
    sSQL = sSQL & "'') "
    sSQL = sSQL & ") As [PaymentRequest], "
    'Sort Payment Request
    sSQL = sSQL & "( "
    sSQL = sSQL & "IIF(RTC.CheckNum Is Not Null,  "
    'Then
    sSQL = sSQL & "Format(Cstr(RTC.[CheckNum] * 100000000000) ,'0000000000000000') "
    sSQL = sSQL & ", "
    'else
    sSQL = sSQL & "Format('1000000000000000','0000000000000000') ) "
    sSQL = sSQL & ") As [SortPaymentRequest], "
    'Sort Me
    sSQL = sSQL & "( "
    sSQL = sSQL & "IIF(RTI.[RTIndemnityID] < 0, "
    'Then
    sSQL = sSQL & " Format(Cstr(ABS(RTI.[RTIndemnityID]) * 100000000000 ) ,'0000000000000000') "
    sSQL = sSQL & ", "
    'else
    sSQL = sSQL & "Format(Cstr(RTI.[RTIndemnityID]),'0000000000000000') ) "
    sSQL = sSQL & ") As [SortME], "
    sSQL = sSQL & "RTC.[BillingCountID], "
    sSQL = sSQL & "RTC.[IDBillingCount], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT [Supplement] "
    sSQL = sSQL & "FROM BillingCount "
    sSQL = sSQL & "WHERE [BillingCountID] = RTC.[BillingCountID] "
    sSQL = sSQL & ") As [Supplement], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT [Rebill] "
    sSQL = sSQL & "FROM BillingCount "
    sSQL = sSQL & "WHERE [BillingCountID] = RTC.[BillingCountID] "
    sSQL = sSQL & ") As [Rebill], "
    sSQL = sSQL & "RTI.[ACVClaim], "
    sSQL = sSQL & "RTI.[ACVLessExcessLimits], "
    sSQL = sSQL & "RTI.[SpecialLimits], "
    sSQL = sSQL & "RTI.[ExcessLimits], "
    sSQL = sSQL & "RTI.[Miscellaneous], "
    sSQL = sSQL & "RTI.[MiscDescription], "
    sSQL = sSQL & "RTI.[IsAddAmountOfInsurance], "
    sSQL = sSQL & "RTI.[ExcessAbsorbsDeductible], "
    sSQL = sSQL & "RTI.[AppliedDeductible], "
    sSQL = sSQL & "RTI.[NonRecoverableDep], "
    sSQL = sSQL & "RTI.[RecoverableDep], "
    sSQL = sSQL & "RTI.[ReplacementCost], "
    sSQL = sSQL & "RTI.[TypeOfLossID], "
    sSQL = sSQL & "RTI.[ClassOfLossID], "
    sSQL = sSQL & "(CT.Description + ' (' + COL.Description + ')') As ClassOfLoss, "
    sSQL = sSQL & "COL.Code As ClassOfLossCode, "
    sSQL = sSQL & "TOL.TypeOfLoss + ' (' + TOL.Description + ')' As TypeOfLoss, "
    sSQL = sSQL & "TOL.Code As typeOfLossCode, "
    sSQL = sSQL & "RTI.[Description], "
    sSQL = sSQL & "RTI.[IsPreviousPayment], "
    sSQL = sSQL & "RTI.[PPayDatePaid], "
    sSQL = sSQL & "RTI.[PPayAmountPaid], "
    sSQL = sSQL & "RTI.[PPayCheckNumber], "
    sSQL = sSQL & "RTI.[IsDeleted], "
    sSQL = sSQL & "RTI.[DownLoadMe], "
    sSQL = sSQL & "RTI.[UpLoadMe], "
    sSQL = sSQL & "TRIM(RTI.[AdminComments]) As [AdminComments], "
    sSQL = sSQL & "RTI.[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = RTI.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "RTI.[UpdateByUserID] "
    sSQL = sSQL & "FROM (((RTIndemnity AS RTI LEFT JOIN RTChecks AS RTC ON RTI.RTChecksID = RTC.RTChecksID) "
    sSQL = sSQL & "LEFT JOIN TypeOfLoss AS TOL ON RTI.TypeOfLossID = TOL.TypeOfLossID) "
    sSQL = sSQL & "LEFT JOIN ClassOfLoss AS COL ON RTI.ClassOfLossID = COL.ClassOfLossID) "
    sSQL = sSQL & "LEFT JOIN ClassType AS CT ON COL.ClassTypeID = CT.ClassTypeID "
    sSQL = sSQL & "WHERE RTI.AssignmentsID = " & psIDAssignments & " "
    sSQL = sSQL & ") RetRTIndemnity "
    sSQL = sSQL & "WHERE [IDAssignments] = " & psIDAssignments & " "
    If CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True)) Then
        sSQL = sSQL & "AND [IsDeleted] = False "
    End If
    sSQL = sSQL & "ORDER BY [SortPaymentRequest], [SortME], [IsPreviousPayment], [ClassOfLoss] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTIndemnity.CursorLocation = adUseClient
    madoRSRTIndemnity.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTIndemnity.ActiveConnection = Nothing
    
    SetadoRSRTIndemnity = True
    
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTIndemnity = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTIndemnity"
End Function

Public Function SetadoRSRTChecks(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSRTChecks Is Nothing Then
        Set madoRSRTChecks = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTChecks = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetRTChecks.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "RTC.[RTChecksID], "
    sSQL = sSQL & "RTC.[AssignmentsID], "
    sSQL = sSQL & "RTC.[BillingCountID], "
    sSQL = sSQL & "RTC.[ID], "
    sSQL = sSQL & "RTC.[IDAssignments], "
    sSQL = sSQL & "RTC.[IDBillingCount], "
    sSQL = sSQL & "RTC.[CheckNum], "
    sSQL = sSQL & "MRP.[ParamValue] As NoOfRequests, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT 'IB' & IIF(BC.Supplement > 0 ,'S' & Format(BC.Supplement,'00') ,'' ) & IIF(BC.Rebill > 0 ,'R' & Format(BC.Rebill,'00') ,'' ) "
    sSQL = sSQL & "FROM BillingCount BC "
    sSQL = sSQL & "WHERE    BC.ID = RTC.IDBillingCount "
    sSQL = sSQL & "AND      BC.IDAssignments = RTC.IDAssignments "
    sSQL = sSQL & ") As IB, "
    sSQL = sSQL & "RTC.[PrintOnIB], "
    sSQL = sSQL & "RTC.[PrintedDate], "
    sSQL = sSQL & "RTC.[RT42_ClassOfLossID], "
    sSQL = sSQL & "RTC.[RT43_TypeOfLossID], "
    sSQL = sSQL & "(CT.Description + ' (' + COL.Description + ')') As ClassOfLoss, "
    sSQL = sSQL & "COL.Code As ClassOfLossCode, "
    sSQL = sSQL & "TOL.TypeOfLoss + ' (' + TOL.Description + ')' As TypeOfLoss, "
    sSQL = sSQL & "TOL.Code As typeOfLossCode, "
    sSQL = sSQL & "RTC.[RT50_sInsuredPayeeName], "
    sSQL = sSQL & "RTC.[RT51_sPayeeNames], "
    sSQL = sSQL & "RTC.[RT52_sAddress], "
    sSQL = sSQL & "RTC.[RT53_cAmountOfCheck], "
    sSQL = sSQL & "RTC.[AppliedDeductible], "
    sSQL = sSQL & "RTC.[RT54_CompanyCatSpecID], "
    sSQL = sSQL & "( "
    sSQL = sSQL & "SELECT CCCS.[CatCode] "
    sSQL = sSQL & "FROM ClientCompanyCatSpec CCCS "
    sSQL = sSQL & "WHERE CCCS.ClientCompanyCatSpecID = RTC.[RT54_CompanyCatSpecID] "
    sSQL = sSQL & ") As CatCode, "
    sSQL = sSQL & "RTC.[tempCHeckName], "
    sSQL = sSQL & "RTC.[IsDeleted], "
    sSQL = sSQL & "RTC.[DownLoadMe], "
    sSQL = sSQL & "RTC.[UpLoadMe], "
    sSQL = sSQL & "TRIM(RTC.[AdminComments]) As [AdminComments], "
    sSQL = sSQL & "RTC.[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = RTC.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "(SELECT    IBNUM "
    sSQL = sSQL & "FROM   Assignments "
    sSQL = sSQL & "WHERE  ID = RTC.[IDAssignments]) As [IBNUM],  "
    sSQL = sSQL & "RTC.[UpdateByUserID] "
    sSQL = sSQL & "FROM (((RTChecks AS RTC LEFT JOIN TypeOfLoss AS TOL ON RTC.RT43_TypeOfLossID = TOL.TypeOfLossID) "
    sSQL = sSQL & "LEFT JOIN MiscReportParam AS MRP ON (MRP.IDAssignments = RTC.IDAssignments And MRP.Number = RTC.CheckNum And MRP.ParamName = 'f_p00_sNumberOfRequests' )) "
    sSQL = sSQL & "LEFT JOIN ClassOfLoss AS COL ON RTC.RT42_ClassOfLossID = COL.ClassOfLossID) "
    sSQL = sSQL & "LEFT JOIN ClassType AS CT ON COL.ClassTypeID = CT.ClassTypeID "
    sSQL = sSQL & "WHERE RTC.AssignmentsID = " & psIDAssignments & " "
    sSQL = sSQL & ") RetRTChecks "
    sSQL = sSQL & "WHERE [IDAssignments] = " & psIDAssignments & " "
    If CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True)) Then
        sSQL = sSQL & "AND [IsDeleted] = False "
    End If
    sSQL = sSQL & "ORDER BY [CheckNum], ABS([ID]) "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTChecks.CursorLocation = adUseClient
    madoRSRTChecks.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTChecks.ActiveConnection = Nothing
    
    SetadoRSRTChecks = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTChecks = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTChecks"
End Function

Public Function SetadoRSRTAttachments(psIDAssignments As String, Optional pbHideDeleted As Boolean = False) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSRTAttachments Is Nothing Then
        Set madoRSRTAttachments = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTAttachments = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetRTAttachments.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[RTAttachmentsID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[AttachDate], "
    sSQL = sSQL & "[SortOrder], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[AttachName], "
    sSQL = sSQL & "[Attachment], "
    sSQL = sSQL & "[DownloadAttachment], "
    sSQL = sSQL & "[UpLoadAttachment], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "TRIM([AdminComments]) As [AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = A.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM RTAttachments A "
    sSQL = sSQL & ") RetRTAttachments "
    sSQL = sSQL & "WHERE [IDAssignments] = " & psIDAssignments & " "
    If CBool(GetSetting(App.EXEName, "GENERAL", "HIDE_DELETED", True)) Then
        sSQL = sSQL & "AND [IsDeleted] = False "
    ElseIf pbHideDeleted Then
        sSQL = sSQL & "AND [IsDeleted] = False "
    End If
    sSQL = sSQL & "ORDER BY [SortOrder] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTAttachments.CursorLocation = adUseClient
    madoRSRTAttachments.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTAttachments.ActiveConnection = Nothing
    
    SetadoRSRTAttachments = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTAttachments = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTAttachments"
End Function

Public Function SetadoRSMainReports(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    
    If Not madoRSMainReports Is Nothing Then
        Set madoRSMainReports = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSMainReports = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  Application A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Claim Information' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Reports' "
    sSQL = sSQL & "AND      A.[SectionLevel03] Like 'Main Reports' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSMainReports.CursorLocation = adUseClient
    madoRSMainReports.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSMainReports.ActiveConnection = Nothing
    
    SetadoRSMainReports = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSMainReports = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSMainReports"
End Function

Public Function SetadoRSMainReportsHistory(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    
    If Not madoRSMainReportsHistory Is Nothing Then
        Set madoRSMainReportsHistory = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSMainReportsHistory = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  ApplicationHistory A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Claim Information' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Reports' "
    sSQL = sSQL & "AND      A.[SectionLevel03] Like 'Main Reports' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSMainReportsHistory.CursorLocation = adUseClient
    madoRSMainReportsHistory.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSMainReportsHistory.ActiveConnection = Nothing
    
    SetadoRSMainReportsHistory = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSMainReportsHistory = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSMainReportsHistory"
End Function

Public Function SetadoRSCarSpecReports(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    
    If Not madoRSCarSpecReports Is Nothing Then
        Set madoRSCarSpecReports = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSCarSpecReports = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  Application A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Claim Information' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Reports' "
    sSQL = sSQL & "AND      A.[SectionLevel03] Like 'Carrier Specific Reports' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSCarSpecReports.CursorLocation = adUseClient
    madoRSCarSpecReports.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSCarSpecReports.ActiveConnection = Nothing
    
    SetadoRSCarSpecReports = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSCarSpecReports = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSCarSpecReports"
End Function

Public Function SetadoRSCarSpecReportsHistory(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long

    If Not madoRSCarSpecReportsHistory Is Nothing Then
        Set madoRSCarSpecReportsHistory = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSCarSpecReportsHistory = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  ApplicationHistory A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Claim Information' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Reports' "
    sSQL = sSQL & "AND      A.[SectionLevel03] Like 'Carrier Specific Reports' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSCarSpecReportsHistory.CursorLocation = adUseClient
    madoRSCarSpecReportsHistory.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSCarSpecReportsHistory.ActiveConnection = Nothing
    
    SetadoRSCarSpecReportsHistory = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSCarSpecReportsHistory = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSCarSpecReportsHistory"
End Function

Public Function SetadoRSWordXLDocs(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    
    If Not madoRSWordXLDocs Is Nothing Then
        Set madoRSWordXLDocs = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSWordXLDocs = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT D.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageDocument SPD "
    sSQL = sSQL & "ON SPD.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  Document D "
    sSQL = sSQL & "ON D.[DocumentID] = SPD.[DocumentID]) "
    sSQL = sSQL & "WHERE    D.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= D.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= D.[SPVersion] "
    sSQL = sSQL & "AND      SPD.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[ClientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      D.[SectionLevel01] Like 'WordXL' "
    sSQL = sSQL & "Order by D.[Description] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSWordXLDocs.CursorLocation = adUseClient
    madoRSWordXLDocs.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSWordXLDocs.ActiveConnection = Nothing
    
    SetadoRSWordXLDocs = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSWordXLDocs = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSWordXLDocs"
End Function

Public Function SetadoRSIB(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSIB Is Nothing Then
        Set madoRSIB = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSIB = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name

    sSQL = "SELECT IB.*, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "IIF( IB.[IB14a_sSupplement] > 0 And IB.[IB14b_sRebilled] > 0, "
        'Then
        sSQL = sSQL & "IB.[IB02_sIBNumber] + 'S' + Cstr(IB.[IB14a_sSupplement]) + 'R' + Cstr(IB.[IB14b_sRebilled]), "
        'else
                    sSQL = sSQL & "IIF(IB.[IB14a_sSupplement] > 0, "
                    'Then
                    sSQL = sSQL & "IB.[IB02_sIBNumber] + 'S' + Cstr(IB.[IB14a_sSupplement]) , "
                    'else
                                sSQL = sSQL & "IIF(IB.[IB14b_sRebilled] > 0, "
                                'Then
                                sSQL = sSQL & "IB.[IB02_sIBNumber] + 'R' + Cstr(IB.[IB14b_sRebilled]) , "
                                'else
                                sSQL = sSQL & "IB.[IB02_sIBNumber]) "
                    sSQL = sSQL & ") "
        sSQL = sSQL & ") "
    sSQL = sSQL & ") As sIBNumber "
    sSQL = sSQL & "FROM IB "
    sSQL = sSQL & "WHERE [IDAssignments] = " & psIDAssignments & " "
    sSQL = sSQL & "AND Void = 0 "
    sSQL = sSQL & "ORDER BY [IB14a_sSupplement], [IB14b_sRebilled] "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSIB.CursorLocation = adUseClient
    madoRSIB.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSIB.ActiveConnection = Nothing
    
    SetadoRSIB = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSIB = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSIB"
End Function

Public Function SetadoRSBillingReports() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    
    If Not madoRSBillingReports Is Nothing Then
        Set madoRSBillingReports = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSBillingReports = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  Application A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Claim Information' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Billings' "
    sSQL = sSQL & "AND      A.[SectionLevel03] Like 'Internal Billing Information' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSBillingReports.CursorLocation = adUseClient
    madoRSBillingReports.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSBillingReports.ActiveConnection = Nothing
    
    SetadoRSBillingReports = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSBillingReports = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSBillingReports"
End Function

Public Function SetadoRSBillingReportsHistory() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    
    If Not madoRSBillingReportsHistory Is Nothing Then
        Set madoRSBillingReportsHistory = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSBillingReportsHistory = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  ApplicationHistory A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Claim Information' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Billings' "
    sSQL = sSQL & "AND      A.[SectionLevel03] Like 'Internal Billing Information' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSBillingReportsHistory.CursorLocation = adUseClient
    madoRSBillingReportsHistory.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSBillingReportsHistory.ActiveConnection = Nothing
    
    SetadoRSBillingReportsHistory = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSBillingReportsHistory = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSBillingReportsHistory"
End Function

Public Function SetadoRSPayment(psIDAssignments As String, _
                                Optional pbSumNetActualCashValueClaim As Boolean = False) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSPayment Is Nothing Then
        Set madoRSPayment = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSPayment = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name

    sSQL = "SELECT RTC.*, "
    sSQL = sSQL & "(CT.Description + ' (' + COL.Description + ')') As ClassOfLoss, "
    sSQL = sSQL & "COL.Code As ClassOfLossCode, "
    sSQL = sSQL & "TOL.TypeOfLoss + ' (' + TOL.Description + ')' As TypeOfLoss, "
    sSQL = sSQL & "TOL.Code As typeOfLossCode "
    If pbSumNetActualCashValueClaim Then
        sSQL = sSQL & ", "
        sSQL = sSQL & "( "
                        'Payment Request (RTChecks) Need to Sum the Amount Of Check
                        'Associated with the IB.
                        sSQL = sSQL & "' Net ACVC Total - [' & "
                        sSQL = sSQL & "( "
                        sSQL = sSQL & "SELECT   (IIF(IsNull(SUM(RTI.[ACVLessExcessLimits])),'0.00', Format(SUM(RTI.[ACVLessExcessLimits]),'0.00'))) As SumOfNetACVC "
                        sSQL = sSQL & "FROM     RTIndemnity RTI "
                        sSQL = sSQL & "WHERE    RTI.IDAssignments = RTC.IDAssignments "
                        sSQL = sSQL & "AND      RTI.IDRTChecks = RTC.ID "
                        sSQL = sSQL & "AND      RTI.IsDeleted = False "
                        sSQL = sSQL & ") "
                        sSQL = sSQL & "& ']' "
                        sSQL = sSQL & "& "
                        sSQL = sSQL & "( "
                        sSQL = sSQL & "IIF(IsDate(RTC.[PrintedDate]),' Printed - ' & RTC.[PrintedDate] & ' ', '') "
                        sSQL = sSQL & ") "
        sSQL = sSQL & ") As PayReqDescription "
    End If
    sSQL = sSQL & "FROM (((RTChecks AS RTC LEFT JOIN TypeOfLoss AS TOL ON RTC.RT43_TypeOfLossID = TOL.TypeOfLossID) "
    sSQL = sSQL & "LEFT JOIN MiscReportParam AS MRP ON (MRP.IDAssignments = RTC.IDAssignments And MRP.Number = RTC.CheckNum And MRP.ParamName = 'f_p00_sNumberOfRequests' )) "
    sSQL = sSQL & "LEFT JOIN ClassOfLoss AS COL ON RTC.RT42_ClassOfLossID = COL.ClassOfLossID) "
    sSQL = sSQL & "LEFT JOIN ClassType AS CT ON COL.ClassTypeID = CT.ClassTypeID "
    sSQL = sSQL & "WHERE RTC.[IDAssignments] = " & psIDAssignments & " "
    sSQL = sSQL & "AND RTC.[IsDeleted] = 0 "
    sSQL = sSQL & "ORDER BY RTC.[CheckNum] "


    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSPayment.CursorLocation = adUseClient
    madoRSPayment.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSPayment.ActiveConnection = Nothing
    
    SetadoRSPayment = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSPayment = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSPayment"
End Function

Public Function SetadoRSPaymentReports() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    
    If Not madoRSPaymentReports Is Nothing Then
        Set madoRSPaymentReports = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSPaymentReports = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  Application A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Claim Information' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Payments' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSPaymentReports.CursorLocation = adUseClient
    madoRSPaymentReports.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSPaymentReports.ActiveConnection = Nothing
    
    SetadoRSPaymentReports = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSPaymentReports = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSPaymentReports"
End Function

Public Function SetadoRSPaymentReportsHistory() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    
    If Not madoRSPaymentReportsHistory Is Nothing Then
        Set madoRSPaymentReportsHistory = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSPaymentReportsHistory = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  ApplicationHistory A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[SectionLevel01] Like 'Claim Information' "
    sSQL = sSQL & "AND      A.[SectionLevel02] Like 'Payments' "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSPaymentReportsHistory.CursorLocation = adUseClient
    madoRSPaymentReportsHistory.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSPaymentReportsHistory.ActiveConnection = Nothing
    
    SetadoRSPaymentReportsHistory = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSPaymentReportsHistory = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSPaymentReportsHistory"
End Function

'PackageItem
Public Function SetadoRSPackageItem(psIDAssignments As String, psPackageID As String, _
                                    Optional psFindFeeBillSupplement As String, _
                                    Optional psFindPaymentRequestCheckNum As String, _
                                    Optional psFindRTAttachmentsID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSPackageItem Is Nothing Then
        Set madoRSPackageItem = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSPackageItem = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT "
    sSQL = sSQL & "[PackageItemID], "
    sSQL = sSQL & "[PackageID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDPackage], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[ReportFormat], "
    sSQL = sSQL & "[RTAttachmentsID], "
    sSQL = sSQL & "[IDRTAttachments], "
    sSQL = sSQL & "[Number], "
    sSQL = sSQL & "[AttachmentName], "
    sSQL = sSQL & "[SortOrder], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[IsCoApprove], "
    sSQL = sSQL & "[CoApproveDate], "
    sSQL = sSQL & "[CoApproveDesc], "
    sSQL = sSQL & "[IsClientCoReject], "
    sSQL = sSQL & "[ClientCoRejectDate], "
    sSQL = sSQL & "[ClientCoRejectDesc], "
    sSQL = sSQL & "[IsClientCoDelete], "
    sSQL = sSQL & "[ClientCoDeleteDate], "
    sSQL = sSQL & "[ClientCoDeleteDesc], "
    sSQL = sSQL & "[IsClientCoApprove], "
    sSQL = sSQL & "[ClientCoApproveDate], "
    sSQL = sSQL & "[ClientCoApproveDesc], "
    sSQL = sSQL & "[PackageItemGUID], "
    sSQL = sSQL & "[SendMe],"
    sSQL = sSQL & "[SentDate],"
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM PackageItem "
    sSQL = sSQL & "WHERE    [IDAssignments] = " & psIDAssignments & " "
    If psFindFeeBillSupplement <> vbNullString Then
        sSQL = sSQL & "AND  [Number]  = " & psFindFeeBillSupplement & " "
        sSQL = sSQL & "AND Instr(1, [ReportFormat], '" & "ECrpt" & goUtil.gsCurCarDBName & "_arRptIB" & "') > 0 "
    ElseIf psFindPaymentRequestCheckNum <> vbNullString Then
        sSQL = sSQL & "AND  [Number]  = " & psFindPaymentRequestCheckNum & " "
        sSQL = sSQL & "AND Instr(1, [ReportFormat], '" & "ECrpt" & goUtil.gsCurCarDBName & "_arRptAddlChk" & "') > 0 "
    ElseIf psFindRTAttachmentsID <> vbNullString Then
        sSQL = sSQL & "AND  [RTAttachmentsID]  = " & psFindRTAttachmentsID & " "
    Else
        sSQL = sSQL & "AND [PackageID] = " & psPackageID & " "
        sSQL = sSQL & "ORDER BY [SortOrder] "
    End If
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSPackageItem.CursorLocation = adUseClient
    madoRSPackageItem.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSPackageItem.ActiveConnection = Nothing
    
    SetadoRSPackageItem = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSPackageItem = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSPackageItem"
End Function

'Package List (Multi Report)
Public Function SetadoRSPackageList(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSPackageList Is Nothing Then
        Set madoRSPackageList = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSPackageList = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT "
    sSQL = sSQL & "[PackageID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[CreateDate], "
    sSQL = sSQL & "[PackageStatus], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[Number], "
    sSQL = sSQL & "[SendMe], "
    sSQL = sSQL & "[SentDate], "
    sSQL = sSQL & "[SentToEmail], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM Package "
    sSQL = sSQL & "WHERE    [IDAssignments] = " & psIDAssignments & " "
    sSQL = sSQL & "ORDER BY [Number] "
    


    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSPackageList.CursorLocation = adUseClient
    madoRSPackageList.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSPackageList.ActiveConnection = Nothing
    
    SetadoRSPackageList = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSPackageList = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSPackageList"
End Function


Public Function SetadoRSFeeScheduleList() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSFeeScheduleList Is Nothing Then
        Set madoRSFeeScheduleList = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSFeeScheduleList = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetFeeSchedule.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[ClientCompanyID], "
    sSQL = sSQL & "[ScheduleName], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[NumOfLevels], "
    sSQL = sSQL & "[NumOfFeeTypes], "
    sSQL = sSQL & "[FeeServiceHourlyRate], "
    sSQL = sSQL & "[TaxPercent], "
    sSQL = sSQL & "[InitialOptions], "
    sSQL = sSQL & "[Options], "
    sSQL = sSQL & "[DefaultAppDedClassTypeIDOrder], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = F.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM FeeSchedule F "
    sSQL = sSQL & ") RetFeeSchedule "
    sSQL = sSQL & "WHERE [ClientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "ORDER BY [ScheduleName] "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSFeeScheduleList.CursorLocation = adUseClient
    madoRSFeeScheduleList.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSFeeScheduleList.ActiveConnection = Nothing
    
    SetadoRSFeeScheduleList = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSFeeScheduleList = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSFeeScheduleList"
End Function

'Worksheet Diagram List (Multi Report type)
Public Function SetadoRSRTWSDiagramList(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSRTWSDiagramList Is Nothing Then
        Set madoRSRTWSDiagramList = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTWSDiagramList = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT "
    sSQL = sSQL & "[RTWSDiagramID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[Number], "
    sSQL = sSQL & "[DiagramPhotoName], "
    sSQL = sSQL & "[DownloadDiagramPhoto], "
    sSQL = sSQL & "[UploadDiagramPhoto], "
    sSQL = sSQL & "[DiagramXML], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM RTWSDiagram "
    sSQL = sSQL & "WHERE    IDAssignments = " & psIDAssignments & " "
    sSQL = sSQL & "ORDER BY [Number] "


    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTWSDiagramList.CursorLocation = adUseClient
    madoRSRTWSDiagramList.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTWSDiagramList.ActiveConnection = Nothing
    
    SetadoRSRTWSDiagramList = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTWSDiagramList = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTWSDiagramList"
End Function

'Photo Report List (Multi Report Type)
Public Function SetadoRSRTPhotoReportList(psIDAssignments As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSRTPhotoReportList Is Nothing Then
        Set madoRSRTPhotoReportList = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTPhotoReportList = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name

    sSQL = "SELECT "
    sSQL = sSQL & "[RTPhotoReportID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[Name], "
    sSQL = sSQL & "[Description], "
    sSQL = sSQL & "[Number], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM RTPhotoReport "
    sSQL = sSQL & "WHERE    IDAssignments = " & psIDAssignments & " "
    sSQL = sSQL & "ORDER BY [Number] "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTPhotoReportList.CursorLocation = adUseClient
    madoRSRTPhotoReportList.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTPhotoReportList.ActiveConnection = Nothing
    
    SetadoRSRTPhotoReportList = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTPhotoReportList = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTPhotoReportList"
End Function

'madoRSBillingCountItem 'a Single Billing Item
Public Function SetadoRSBillingCountItem(psAssignmentsID As String, psIDBillingCount As String, Optional pbFindMaxSupplement As Boolean = False) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    'reset RS
    If Not madoRSBillingCountItem Is Nothing Then
        Set madoRSBillingCountItem = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSBillingCountItem = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RetBillingCount.* "
    sSQL = sSQL & "FROM( "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & "[BillingCountID], "
    sSQL = sSQL & "[AssignmentsID] , "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[Rebill], "
    sSQL = sSQL & "[Supplement], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated] , "
    sSQL = sSQL & "(SELECT    USERNAME "
    sSQL = sSQL & "FROM   USERS "
    sSQL = sSQL & "WHERE  USERSID = BC.[UpdateByUserID]) As [UpdateByUserName],  "
    sSQL = sSQL & "[UpdateByUserID]  "
    sSQL = sSQL & "FROM BillingCount BC "
    sSQL = sSQL & ") RetBillingCount "
    If pbFindMaxSupplement Then
        sSQL = sSQL & "WHERE BC.[Supplement] = ( "
                        sSQL = sSQL & "SELECT Max(Supplement) As MaxSupplement "
                        sSQL = sSQL & "FROM BillingCount "
                        sSQL = sSQL & "WHERE AssignmentsID = " & psAssignmentsID & " "
                        sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "WHERE BC.[ID] = " & psIDBillingCount & " "
    End If
    sSQL = sSQL & "AND AssignmentsID = " & psAssignmentsID & " "

    'Use Disconnected RS Only
    madoRSBillingCountItem.CursorLocation = adUseClient
    madoRSBillingCountItem.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSBillingCountItem.ActiveConnection = Nothing
    
    SetadoRSBillingCountItem = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSBillingCountItem = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSBillingCountItem"
End Function

'madoRSCLIB 'Closed IB RS
Public Function SetadoRSCLIB(psAssignmentsID As String, psIDBillingCount As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSCLIB Is Nothing Then
        Set madoRSCLIB = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSCLIB = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name

    sSQL = "SELECT IB.*, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "IIF( IB.[IB14a_sSupplement] > 0 And IB.[IB14b_sRebilled] > 0, "
        'Then
        sSQL = sSQL & "IB.[IB02_sIBNumber] + 'S' + Cstr(IB.[IB14a_sSupplement]) + 'R' + Cstr(IB.[IB14b_sRebilled]), "
        'else
                    sSQL = sSQL & "IIF(IB.[IB14a_sSupplement] > 0, "
                    'Then
                    sSQL = sSQL & "IB.[IB02_sIBNumber] + 'S' + Cstr(IB.[IB14a_sSupplement]) , "
                    'else
                                sSQL = sSQL & "IIF(IB.[IB14b_sRebilled] > 0, "
                                'Then
                                sSQL = sSQL & "IB.[IB02_sIBNumber] + 'R' + Cstr(IB.[IB14b_sRebilled]) , "
                                'else
                                sSQL = sSQL & "IB.[IB02_sIBNumber]) "
                    sSQL = sSQL & ") "
        sSQL = sSQL & ") "
    sSQL = sSQL & ") As sIBNumber "
    sSQL = sSQL & "FROM IB "
    sSQL = sSQL & "WHERE [IDBillingCount] = " & psIDBillingCount & " "
    sSQL = sSQL & "AND AssignmentsID = " & psAssignmentsID & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSCLIB.CursorLocation = adUseClient
    madoRSCLIB.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSCLIB.ActiveConnection = Nothing
    
    SetadoRSCLIB = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSCLIB = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSCLIB"
End Function

'madoRSCLIBFee ' Closed IB Fee (Service and Expense) Items
Public Function SetadoRSCLIBFee(psAssignmentsID As String, psIDBillingCount As String, Optional psFeeScheduleID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSCLIBFee Is Nothing Then
        Set madoRSCLIBFee = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSCLIBFee = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT IBFee.[IBFeeID], "
    sSQL = sSQL & "IBFee.[AssignmentsID], "
    sSQL = sSQL & "IBFee.[IBID], "
    sSQL = sSQL & "IBFee.[ID], "
    sSQL = sSQL & "IBFee.[IDAssignments], "
    sSQL = sSQL & "IBFee.[IDIB], "
    sSQL = sSQL & "IBFee.[FeeScheduleFeeTypesID], "
    sSQL = sSQL & "FSFT.[FeeScheduleID] As [fsftFeeScheduleID], "
    sSQL = sSQL & "FSFT.[TypeNum] As [fsftTypeNum], "
    sSQL = sSQL & "FSFT.[Name] As [fsftName], "
    sSQL = sSQL & "FSFT.[Description] As [fsftDescription], "
    sSQL = sSQL & "FSFT.[FeeAmount] As [fsftFeeAmount], "
    sSQL = sSQL & "FSFT.[IsExpense] As [fsftIsExpense], "
    sSQL = sSQL & "FSFT.[MaxNumberOfItems] As [fsftMaxNumberOfItems], "
    sSQL = sSQL & "FSFT.[MaxFeeAmount] As [fsftMaxFeeAmount], "
    sSQL = sSQL & "FSFT.[IsMiscAmount] As [fsftIsMiscAmount], "
    sSQL = sSQL & "FSFT.[UseFormula] As [fsftUseFormula], "
    sSQL = sSQL & "FSFT.[VBFormula] As [fsftVBFormula], "
    sSQL = sSQL & "FSFT.[IsDeleted] As [fsftIsDeleted], "
    sSQL = sSQL & "FSFT.[DateLastUpdated] As [fsftDateLastUpdated], "
    sSQL = sSQL & "FSFT.[UpdateByUserID] As [fsftUpdateByUserID], "
    sSQL = sSQL & "IBFee.[NumberOfItems], "
    sSQL = sSQL & "IBFee.[Amount], "
    sSQL = sSQL & "IBFee.[Comment], "
    sSQL = sSQL & "IBFee.[DownLoadMe], "
    sSQL = sSQL & "IBFee.[UpLoadMe], "
    sSQL = sSQL & "IBFee.[AdminComments], "
    sSQL = sSQL & "IBFee.[DateLastUpdated], "
    sSQL = sSQL & "IBFee.[UpdateByUserID] "
    sSQL = sSQL & "FROM IBFee INNER JOIN FeeScheduleFeeTypes FSFT ON IBFee.[FeeScheduleFeeTypesID] = FSFT.[FeeScheduleFeeTypesID] "
    sSQL = sSQL & "WHERE IBFee.[IBID] = ( "
                        sSQL = sSQL & "SELECT IB.[IBID] "
                        sSQL = sSQL & "FROM IB "
                        sSQL = sSQL & "WHERE [IDBillingCount] = " & psIDBillingCount & " "
                        sSQL = sSQL & ") "
    sSQL = sSQL & "AND IBFee.[AssignmentsID] = " & psAssignmentsID & " "
    If psFeeScheduleID = vbNullString Then
        sSQL = sSQL & "AND FSFT.[FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "AND [FeeScheduleID] = ( "
                                    sSQL = sSQL & "SELECT   " & psFeeScheduleID & " As [FeeScheduleID] "
                                    sSQL = sSQL & "FROM     ClientCompanyCat "
                                    sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                    sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                    sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "ORDER BY FSFT.[TypeNum] "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSCLIBFee.CursorLocation = adUseClient
    madoRSCLIBFee.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSCLIBFee.ActiveConnection = Nothing
    
    SetadoRSCLIBFee = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSCLIBFee = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSCLIBFee"
End Function

'madoRSRTIB 'Currently working on IB RS
Public Function SetadoRSRTIB(psAssignmentsID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSRTIB Is Nothing Then
        Set madoRSRTIB = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTIB = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name

    sSQL = "SELECT RTIB.*, "
    sSQL = sSQL & "( "
    sSQL = sSQL & "IIF( RTIB.[RT14a_sSupplement] > 0 And RTIB.[RT14b_sRebilled] > 0, "
        'Then
        sSQL = sSQL & "RTIB.[RT02_sIBNumber] + 'S' + Cstr(RTIB.[RT14a_sSupplement]) + 'R' + Cstr(RTIB.[RT14b_sRebilled]), "
        'else
                    sSQL = sSQL & "IIF(RTIB.[RT14a_sSupplement] > 0, "
                    'Then
                    sSQL = sSQL & "RTIB.[RT02_sIBNumber] + 'S' + Cstr(RTIB.[RT14a_sSupplement]) , "
                    'else
                                sSQL = sSQL & "IIF(RTIB.[RT14b_sRebilled] > 0, "
                                'Then
                                sSQL = sSQL & "RTIB.[RT02_sIBNumber] + 'R' + Cstr(RTIB.[RT14b_sRebilled]) , "
                                'else
                                sSQL = sSQL & "RTIB.[RT02_sIBNumber]) "
                    sSQL = sSQL & ") "
        sSQL = sSQL & ") "
    sSQL = sSQL & ") As sIBNumber "
    sSQL = sSQL & "FROM RTIB "
    sSQL = sSQL & "WHERE [AssignmentsID] = " & psAssignmentsID & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTIB.CursorLocation = adUseClient
    madoRSRTIB.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTIB.ActiveConnection = Nothing
    
    SetadoRSRTIB = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTIB = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTIB"
End Function

'madoRSRTIBFee ' Currently working on IB Fee (Service and Expesne) Items
Public Function SetadoRSRTIBFee(psAssignmentsID As String, Optional psFeeScheduleID As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSRTIBFee Is Nothing Then
        Set madoRSRTIBFee = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSRTIBFee = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name

    sSQL = "SELECT RTIBFee.[RTIBFeeID], "
    sSQL = sSQL & "RTIBFee.[AssignmentsID], "
    sSQL = sSQL & "RTIBFee.[ID], "
    sSQL = sSQL & "RTIBFee.[IDAssignments], "
    sSQL = sSQL & "RTIBFee.[FeeScheduleFeeTypesID], "
    sSQL = sSQL & "FSFT.[FeeScheduleID] As [fsftFeeScheduleID], "
    sSQL = sSQL & "FSFT.[TypeNum] As [fsftTypeNum], "
    sSQL = sSQL & "FSFT.[Name] As [fsftName], "
    sSQL = sSQL & "FSFT.[Description] As [fsftDescription], "
    sSQL = sSQL & "FSFT.[FeeAmount] As [fsftFeeAmount], "
    sSQL = sSQL & "FSFT.[IsExpense] As [fsftIsExpense], "
    sSQL = sSQL & "FSFT.[MaxNumberOfItems] As [fsftMaxNumberOfItems], "
    sSQL = sSQL & "FSFT.[MaxFeeAmount] As [fsftMaxFeeAmount], "
    sSQL = sSQL & "FSFT.[IsMiscAmount] As [fsftIsMiscAmount], "
    sSQL = sSQL & "FSFT.[UseFormula] As [fsftUseFormula], "
    sSQL = sSQL & "FSFT.[VBFormula] As [fsftVBFormula], "
    sSQL = sSQL & "FSFT.[IsDeleted] As [fsftIsDeleted], "
    sSQL = sSQL & "FSFT.[DateLastUpdated] As [fsftDateLastUpdated], "
    sSQL = sSQL & "FSFT.[UpdateByUserID] As [fsftUpdateByUserID], "
    sSQL = sSQL & "RTIBFee.[NumberOfItems], "
    sSQL = sSQL & "RTIBFee.[Amount], "
    sSQL = sSQL & "RTIBFee.[Comment], "
    sSQL = sSQL & "RTIBFee.[DownLoadMe], "
    sSQL = sSQL & "RTIBFee.[UpLoadMe], "
    sSQL = sSQL & "RTIBFee.[AdminComments], "
    sSQL = sSQL & "RTIBFee.[DateLastUpdated], "
    sSQL = sSQL & "RTIBFee.[UpdateByUserID] "
    sSQL = sSQL & "FROM RTIBFee INNER JOIN FeeScheduleFeeTypes FSFT ON RTIBFee.[FeeScheduleFeeTypesID] = FSFT.[FeeScheduleFeeTypesID] "
    sSQL = sSQL & "WHERE RTIBFee.[AssignmentsID] = " & psAssignmentsID & " "
    If psFeeScheduleID = vbNullString Then
        sSQL = sSQL & "AND FSFT.[FeeScheduleID] = ( "
                                sSQL = sSQL & "SELECT   [FeeScheduleID] "
                                sSQL = sSQL & "FROM     ClientCompanyCat "
                                sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                sSQL = sSQL & ") "
    Else
        sSQL = sSQL & "AND [FeeScheduleID] = ( "
                                    sSQL = sSQL & "SELECT   " & psFeeScheduleID & " As [FeeScheduleID] "
                                    sSQL = sSQL & "FROM     ClientCompanyCat "
                                    sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
                                    sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
                                    sSQL = sSQL & ") "
    End If
    sSQL = sSQL & "ORDER BY FSFT.[TypeNum] "
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSRTIBFee.CursorLocation = adUseClient
    madoRSRTIBFee.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSRTIBFee.ActiveConnection = Nothing
    
    SetadoRSRTIBFee = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSRTIBFee = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSRTIBFee"
End Function

'Private madoRSClientCOCat As ADODB.Recordset 'Current Client company Cat Info
Public Function SetadoRSClientCOCat() As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    If Not madoRSClientCOCat Is Nothing Then
        Set madoRSClientCOCat = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set madoRSClientCOCat = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name

    
    'Get Client Company Cat info
    sSQL = "SELECT   [ClientCompanyID], "
    sSQL = sSQL & "[CATID], "
    sSQL = sSQL & "[BillingCode], "
    sSQL = sSQL & "[TypeOfLossID], "
    sSQL = sSQL & "[FeeScheduleID], "
    sSQL = sSQL & "[SiteAddress], "
    sSQL = sSQL & "[SACity], "
    sSQL = sSQL & "[SAState], "
    sSQL = sSQL & "[SAZip], "
    sSQL = sSQL & "[SAZip4], "
    sSQL = sSQL & "[SAOtherPostCode], "
    sSQL = sSQL & "[ActiveDate], "
    sSQL = sSQL & "[InactiveDate], "
    sSQL = sSQL & "[AssignByZipDefault], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & "FROM     ClientCompanyCat "
    sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
    
    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    madoRSClientCOCat.CursorLocation = adUseClient
    madoRSClientCOCat.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set madoRSClientCOCat.ActiveConnection = Nothing
    
    SetadoRSClientCOCat = True
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    SetadoRSClientCOCat = False
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function SetadoRSClientCOCat"
End Function

Public Function PopulateLookUp(pLookUpRS As ADODB.Recordset, _
                                pAssTableRS As ADODB.Recordset, _
                                pcboBox As Object, _
                                psIDName As String, _
                                psIDNameAssTable As String, _
                                psItemName As String, _
                                psItemDescName As String, _
                                Optional pbUseItemNameAsID As Boolean, _
                                Optional psAssTableItemName As String, _
                                Optional pbIfItemDescNameIsNullUse2 As Boolean, _
                                Optional psItemDescName2 As String, _
                                Optional pClassTypeIDInRS As ADODB.Recordset) As Boolean
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim adoRS As ADODB.Recordset
    Dim cboBox As ComboBox
    Dim sTemp As String
    Dim sTemp2 As String
    Dim lId As Long
    Dim lID2 As Long
    Dim lSelIndex As Long
    Dim sItemNameValue As String
    Dim sAssTableItemNameValue As String
    Dim sSQL As String
    Dim oConn As ADODB.Connection
    Dim MyRS As ADODB.Recordset
    Dim bClassTypeIN As Boolean
    
    'If the ClassTypeID in is passed that means populating
    'Class Type DropDown box.  Only include items that are
    'In this RS
    If Not pClassTypeIDInRS Is Nothing Then
        If pClassTypeIDInRS.RecordCount > 0 Then
            bClassTypeIN = True
        Else
            Exit Function
        End If
    End If
    
    Set RS = pLookUpRS
    Set adoRS = pAssTableRS
    
    Set oConn = New ADODB.Connection
    Set MyRS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    If TypeOf pcboBox Is ComboBox Then
        Set cboBox = pcboBox
    Else
        Exit Function
    End If
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        lSelIndex = -1
        Do Until RS.EOF
            lId = RS.Fields(psIDName).Value
            'Check Class Type If Applicable
            If bClassTypeIN Then
                pClassTypeIDInRS.MoveFirst
                Do Until pClassTypeIDInRS.EOF
                    lID2 = pClassTypeIDInRS.Fields("ClassTypeID")
                    If lId = lID2 Then
                        GoTo ADD_ITEM
                    End If
                    pClassTypeIDInRS.MoveNext
                Loop
                GoTo NEXT_ITEM
            End If
ADD_ITEM:
            sTemp = RS.Fields(psItemName).Value
            sTemp = sTemp & " ("
            If pbIfItemDescNameIsNullUse2 Then
                sTemp2 = goUtil.IsNullIsVbNullString(RS.Fields(psItemDescName))
                If sTemp2 = vbNullString Then
                    sTemp = sTemp & RS.Fields(psItemDescName2).Value
                Else
                    sTemp = sTemp & sTemp2
                End If
            Else
                sTemp = sTemp & RS.Fields(psItemDescName).Value
            End If
            sTemp = sTemp & ")"
            cboBox.AddItem sTemp
            'Set the Record Id to the Itemdata of the element just added
            cboBox.ItemData(cboBox.NewIndex) = lId
            
            If pbUseItemNameAsID Then
                sItemNameValue = RS.Fields(psItemName).Value
                sAssTableItemNameValue = adoRS.Fields(psAssTableItemName).Value
                If StrComp(sItemNameValue, sAssTableItemNameValue, vbTextCompare) = 0 Then
                    lSelIndex = cboBox.NewIndex
                End If
            Else
                If Not adoRS Is Nothing Then
                    sTemp = goUtil.IsNullIsVbNullString(adoRS.Fields(psIDNameAssTable))
                    If IsNumeric(sTemp) Then
                        lID2 = sTemp
                    Else
                        lID2 = 0
                    End If
                    'See if this ID matches
                    If (StrComp(psIDNameAssTable, "RT42_ClassOfLossID") = 0 Or StrComp(psIDNameAssTable, "ClassOfLossID") = 0) And lID2 <> 0 Then
                        'If looking up Class of Loss Need to Get the Class of loss
                        'ClassType to match
                        sSQL = "SELECT  ClassTypeID "
                        sSQL = sSQL & "FROM ClassOfLoss "
                        sSQL = sSQL & "WHERE ClassOfLossID = " & lID2 & " "
                        Set MyRS = Nothing
                        Set MyRS = New ADODB.Recordset
                        MyRS.CursorLocation = adUseClient
                        MyRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
                        Set MyRS.ActiveConnection = Nothing
                        If MyRS.RecordCount = 1 Then
                            lID2 = MyRS.Fields("ClassTypeID").Value
                        End If
                    End If
                    
                    
                    If lId = lID2 Then
                        lSelIndex = cboBox.NewIndex
                    End If
                End If
            End If
NEXT_ITEM:
            RS.MoveNext
        Loop
        'Now Select the id that is in Assignments Tabel for this claim
        If lSelIndex > -1 Then
            cboBox.ListIndex = lSelIndex
        End If
        
    End If
    
    PopulateLookUp = True
    
    Set RS = Nothing
    Set adoRS = Nothing
    Set oConn = Nothing
    Set MyRS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PopulateLookUp"
End Function

Public Function RemoveVBCRLF(poTextBox As Object) As Boolean
    On Error GoTo EH
    Dim oTextBox As TextBox
    Dim lPos As Long
    
    If Not TypeOf poTextBox Is TextBox Then
        Exit Function
    Else
        Set oTextBox = poTextBox
    End If
    
    If InStr(1, oTextBox.Text, vbCrLf, vbBinaryCompare) > 0 Then
        lPos = oTextBox.SelStart
        oTextBox.Text = Replace(oTextBox.Text, vbCrLf, vbNullString)
        oTextBox.SelStart = lPos
    End If
    
    RemoveVBCRLF = True
    
    'cleanup
    Set oTextBox = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function RemoveVBCRLF"
End Function


Public Function HighlightOpt(poOptButton As Object) As Boolean
    On Error GoTo EH
    Dim oMyOpt As OptionButton
    
    If TypeOf poOptButton Is OptionButton Then
        Set oMyOpt = poOptButton
    Else
        Exit Function
    End If

    'Need to Highlight the opt button to Gray or White
    'White = Selected
    'Gray = not
    If oMyOpt.Value = True Then
        oMyOpt.BackColor = BG_WHITE
    Else
        oMyOpt.BackColor = BG_GRAY
    End If
    
    HighlightOpt = True
    
    'cleanup
    Set oMyOpt = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function HighlightOpt"
End Function

Private Sub TimerReSize_Timer()
    On Error GoTo EH
    TimerReSize.Enabled = False
    If MyfrmClaimsList Is Nothing Then
        Exit Sub
    End If
    'Claim
    If Me.top <> 10 Or Me.left <> 10 Then
        ReSizeMe
    End If
    
    'Claim Info
    If Not mfrmClaimInfo Is Nothing Then
        ResizeChildForm mfrmClaimInfo
    End If
    
    'Activity Log
    If Not mfrmActivityLog Is Nothing Then
        ResizeChildForm mfrmActivityLog
    End If
    
    'Indemnity
    If Not mfrmIndemnity Is Nothing Then
        ResizeChildForm mfrmIndemnity
    End If
    
    'Billing Information
    If Not mfrmBillingInfo Is Nothing Then
        ResizeChildForm mfrmBillingInfo
    End If
    
    'Photos
    If Not mfrmPhotos Is Nothing Then
        ResizeChildForm mfrmPhotos
    End If
    
    'Reports  mfrmReports
    If Not mfrmReports Is Nothing Then
        ResizeChildForm mfrmReports
    End If
    
    'Attachments  mfrmAttachments
    If Not mfrmAttachments Is Nothing Then
        ResizeChildForm mfrmAttachments
    End If
    
    'Miscellaneous
    If Not mfrmMiscellaneous Is Nothing Then
        ResizeChildForm mfrmMiscellaneous
    End If
    
    'Print
    If Not mfrmPrint Is Nothing Then
        ResizeChildForm mfrmPrint
    End If
    
    TimerReSize.Enabled = True
    
    Exit Sub
EH:
     goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub TimerReSize_Timer"
End Sub

Public Sub ResizeChildForm(poForm As Object)
    On Error GoTo EH
    Dim oForm As Form
    
    If Not TypeOf poForm Is Form Then
        Exit Sub
    Else
        Set oForm = poForm
    End If
    
    If oForm.WindowState <> vbMinimized And oForm.WindowState <> vbMaximized Then
        If oForm.top <> (Me.top + Me.Height) Or oForm.left <> 10 Then
            oForm.ReSizeMe
        End If
    End If
    
    'cleanup
    Set oForm = Nothing
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub ResizeChildForm"
End Sub

Public Function AddMultiReport(psTableName As String, psName As String, psDesc As String, plNumber As Long) As Boolean
    On Error GoTo EH
    Dim sMess As String
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    Dim sTableName As String
    Dim sName As String
    Dim sDescription As String
    Dim sNumber As String
    
    'Set from Passed in params
    sTableName = psTableName
    sName = psName
    sDescription = psDesc
    sNumber = CStr(plNumber)
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    Screen.MousePointer = vbHourglass
    
    If StrComp(psTableName, "RTPhotoReport", vbTextCompare) = 0 Then
        sID = goUtil.GetAccessDBUID("ID", "RTPhotoReport")

        sSQL = "INSERT INTO  RTPhotoReport "
        sSQL = sSQL & "( "
        sSQL = sSQL & "[RTPhotoReportID], "
        sSQL = sSQL & "[AssignmentsID], "
        sSQL = sSQL & "[ID], "
        sSQL = sSQL & "[IDAssignments], "
        sSQL = sSQL & "[Name], "
        sSQL = sSQL & "[Description], "
        sSQL = sSQL & "[Number], "
        sSQL = sSQL & "[IsDeleted], "
        sSQL = sSQL & "[DownLoadMe], "
        sSQL = sSQL & "[UpLoadMe], "
        sSQL = sSQL & "[AdminComments], "
        sSQL = sSQL & "[DateLastUpdated], "
        sSQL = sSQL & "[UpdateByUserID] "
        sSQL = sSQL & ") "
        sSQL = sSQL & "SELECT "
        sSQL = sSQL & sID & " As [RTPhotoReportID], "
        sSQL = sSQL & msAssignmentsID & " As [AssignmentsID], "
        sSQL = sSQL & sID & " As [ID], "
        sSQL = sSQL & msAssignmentsID & " As [IDAssignments], "
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(sName) & "' As [Name], "
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(sDescription) & "' As [Description], "
        sSQL = sSQL & sNumber & " As [Number], "
        sSQL = sSQL & "False As [IsDeleted], "
        sSQL = sSQL & "False As [DownLoadMe], "
        sSQL = sSQL & "True As [UpLoadMe], "
        sSQL = sSQL & "'' As [AdminComments], "
        sSQL = sSQL & "#" & Now() & "# As [DateLastUpdated], "
        sSQL = sSQL & goUtil.gsCurUsersID & " As [UpdateByUserID]"
    ElseIf StrComp(psTableName, "RTWSDiagram", vbTextCompare) = 0 Then
        sID = goUtil.GetAccessDBUID("ID", "RTWSDiagram")

        sSQL = "INSERT INTO  RTWSDiagram "
        sSQL = sSQL & "( "
        sSQL = sSQL & "[RTWSDiagramID], "
        sSQL = sSQL & "[AssignmentsID], "
        sSQL = sSQL & "[ID], "
        sSQL = sSQL & "[IDAssignments], "
        sSQL = sSQL & "[Name], "
        sSQL = sSQL & "[Description], "
        sSQL = sSQL & "[Number], "
        sSQL = sSQL & "[DiagramPhotoName], "
        sSQL = sSQL & "[DownloadDiagramPhoto], "
        sSQL = sSQL & "[UploadDiagramPhoto], "
        sSQL = sSQL & "[DiagramXML], "
        sSQL = sSQL & "[IsDeleted], "
        sSQL = sSQL & "[DownLoadMe], "
        sSQL = sSQL & "[UpLoadMe], "
        sSQL = sSQL & "[AdminComments], "
        sSQL = sSQL & "[DateLastUpdated], "
        sSQL = sSQL & "[UpdateByUserID] "
        sSQL = sSQL & ") "
        sSQL = sSQL & "SELECT "
        sSQL = sSQL & sID & " As [RTWSDiagramID], "
        sSQL = sSQL & msAssignmentsID & " As [AssignmentsID], "
        sSQL = sSQL & sID & " As [ID], "
        sSQL = sSQL & msAssignmentsID & " As [IDAssignments], "
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(sName) & "' As [Name], "
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(sDescription) & "' As [Description], "
        sSQL = sSQL & sNumber & " As [Number], "
        sSQL = sSQL & "'' As [DiagramPhotoName], "
        sSQL = sSQL & "False As [DownloadDiagramPhoto], "
        sSQL = sSQL & "False As [UploadDiagramPhoto], "
        sSQL = sSQL & "'' As [DiagramXML], "
        sSQL = sSQL & "False As [IsDeleted], "
        sSQL = sSQL & "False As [DownLoadMe], "
        sSQL = sSQL & "True As [UpLoadMe], "
        sSQL = sSQL & "'' As [AdminComments], "
        sSQL = sSQL & "#" & Now() & "# As [DateLastUpdated], "
        sSQL = sSQL & goUtil.gsCurUsersID & " As [UpdateByUserID]"
    ElseIf StrComp(psTableName, "Package", vbTextCompare) = 0 Then
        sID = goUtil.GetAccessDBUID("ID", "Package")

        sSQL = "INSERT INTO  Package "
        sSQL = sSQL & "( "
        sSQL = sSQL & "[PackageID], "
        sSQL = sSQL & "[AssignmentsID], "
        sSQL = sSQL & "[ID], "
        sSQL = sSQL & "[IDAssignments], "
        sSQL = sSQL & "[CreateDate], "
        sSQL = sSQL & "[PackageStatus], "
        sSQL = sSQL & "[Name], "
        sSQL = sSQL & "[Description], "
        sSQL = sSQL & "[Number], "
        sSQL = sSQL & "[SendMe], "
        sSQL = sSQL & "[SentDate], "
        sSQL = sSQL & "[SentToEmail], "
        sSQL = sSQL & "[IsDeleted], "
        sSQL = sSQL & "[DownLoadMe], "
        sSQL = sSQL & "[UpLoadMe], "
        sSQL = sSQL & "[AdminComments], "
        sSQL = sSQL & "[DateLastUpdated], "
        sSQL = sSQL & "[UpdateByUserID] "
        sSQL = sSQL & ") "
        sSQL = sSQL & "SELECT "
        sSQL = sSQL & sID & " As [PackageID], "
        sSQL = sSQL & msAssignmentsID & " As [AssignmentsID], "
        sSQL = sSQL & sID & " As [ID], "
        sSQL = sSQL & msAssignmentsID & " As [IDAssignments], "
        sSQL = sSQL & "#" & Now() & "# As [CreateDate], "
        sSQL = sSQL & "'NEW' As [PackageStatus], "
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(sName) & "' As [Name], "
        sSQL = sSQL & "'" & goUtil.utCleanSQLString(sDescription) & "' As [Description], "
        sSQL = sSQL & sNumber & " As [Number], "
        sSQL = sSQL & "False As [SendMe], "
        sSQL = sSQL & "Null As [SentDate], "
        sSQL = sSQL & "Null As [SentToEmail], "
        sSQL = sSQL & "False As [IsDeleted], "
        sSQL = sSQL & "False As [DownLoadMe], "
        sSQL = sSQL & "True As [UpLoadMe], "
        sSQL = sSQL & "'' As [AdminComments], "
        sSQL = sSQL & "#" & Now() & "# As [DateLastUpdated], "
        sSQL = sSQL & goUtil.gsCurUsersID & " As [UpdateByUserID]"
    Else
        AddMultiReport = False
        GoTo CLEAN_UP
    End If

    oConn.Execute sSQL
    
    Sleep 500
    
    RefreshMe
    
    Screen.MousePointer = vbNormal
    
    AddMultiReport = True
CLEAN_UP:
    'cleanup
    Set oConn = Nothing
    Exit Function
EH:
    Screen.MousePointer = vbNormal
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddMultiReport"
End Function

Public Function AddRptParamItem(pudtRptParam As V2ECKeyBoard.MiscReportParam) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    sID = goUtil.GetAccessDBUID("ID", msMiscReportParamName)
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    
    With pudtRptParam
        .MiscReportParamID = sID
        .AssignmentsID = msAssignmentsID
        .ID = sID
        .IDAssignments = msAssignmentsID
        If .Number = vbNullString Then
            .Number = "null"
        End If
    End With
    
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    sSQL = "INSERT INTO " & msMiscReportParamName & " "
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    
    sSQL = sSQL & "( "
    sSQL = sSQL & "[MiscReportParamID], "
    sSQL = sSQL & "[AssignmentsID], "
    sSQL = sSQL & "[ID], "
    sSQL = sSQL & "[IDAssignments], "
    sSQL = sSQL & "[Number], "
    sSQL = sSQL & "[ProjectName], "
    sSQL = sSQL & "[ClassName], "
    sSQL = sSQL & "[ParamName], "
    sSQL = sSQL & "[ParamCaption], "
    sSQL = sSQL & "[ParamDataType], "
    sSQL = sSQL & "[ParamValue], "
    sSQL = sSQL & "[SortMe], "
    sSQL = sSQL & "[IsDeleted], "
    sSQL = sSQL & "[DownLoadMe], "
    sSQL = sSQL & "[UpLoadMe], "
    sSQL = sSQL & "[AdminComments], "
    sSQL = sSQL & "[DateLastUpdated], "
    sSQL = sSQL & "[UpdateByUserID] "
    sSQL = sSQL & ") "
    sSQL = sSQL & "SELECT "
    sSQL = sSQL & pudtRptParam.MiscReportParamID & " As [MiscReportParamID], "
    sSQL = sSQL & pudtRptParam.AssignmentsID & " As [AssignmentsID], "
    sSQL = sSQL & pudtRptParam.ID & " As [ID], "
    sSQL = sSQL & pudtRptParam.IDAssignments & " As [IDAssignments], "
    sSQL = sSQL & pudtRptParam.Number & " As [Number], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRptParam.ProjectName) & "'" & " As [ProjectName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRptParam.ClassName) & "'" & " As [ClassName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRptParam.ParamName) & "'" & " As [ParamName], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRptParam.ParamCaption) & "'" & " As [ParamCaption], "
    sSQL = sSQL & pudtRptParam.ParamDataType & " As [ParamDataType], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRptParam.ParamValue) & "'" & " As [ParamValue], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRptParam.SortMe) & "'" & " As [SortMe], "
    sSQL = sSQL & pudtRptParam.IsDeleted & " As [IsDeleted], "
    sSQL = sSQL & pudtRptParam.DownLoadMe & " As [DownLoadMe], "
    sSQL = sSQL & pudtRptParam.UpLoadMe & " As [UpLoadMe], "
    sSQL = sSQL & "'" & goUtil.utCleanSQLString(pudtRptParam.AdminComments) & "'" & " As [AdminComments], "
    sSQL = sSQL & "#" & pudtRptParam.DateLastUpdated & "#" & " As [DateLastUpdated], "
    sSQL = sSQL & pudtRptParam.UpdateByUserID & " As [UpdateByUserID] "
   
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    
    AddRptParamItem = True
    
    'Clean up
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function AddRptParamItem"
End Function

Public Function EditRptParamItem(pudtRptParam As V2ECKeyBoard.MiscReportParam) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sID As String
    
    With pudtRptParam
        If .Number = vbNullString Then
            .Number = "null"
        End If
    End With
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    sSQL = "UPDATE " & msMiscReportParamName & " SET "
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    sSQL = sSQL & "[MiscReportParamID] = " & pudtRptParam.MiscReportParamID & ", "
    sSQL = sSQL & "[AssignmentsID] = " & pudtRptParam.AssignmentsID & ", "
    sSQL = sSQL & "[ID] = " & pudtRptParam.ID & ", "
    sSQL = sSQL & "[IDAssignments] = " & pudtRptParam.IDAssignments & ", "
    sSQL = sSQL & "[Number] = " & pudtRptParam.Number & ", "
    sSQL = sSQL & "[ProjectName] = '" & goUtil.utCleanSQLString(pudtRptParam.ProjectName) & "', "
    sSQL = sSQL & "[ClassName] = '" & goUtil.utCleanSQLString(pudtRptParam.ClassName) & "', "
    sSQL = sSQL & "[ParamName] = '" & goUtil.utCleanSQLString(pudtRptParam.ParamName) & "', "
    sSQL = sSQL & "[ParamCaption] = '" & goUtil.utCleanSQLString(pudtRptParam.ParamCaption) & "', "
    sSQL = sSQL & "[ParamDataType] = " & pudtRptParam.ParamDataType & ", "
    sSQL = sSQL & "[ParamValue] = '" & goUtil.utCleanSQLString(pudtRptParam.ParamValue) & "', "
    sSQL = sSQL & "[SortMe] = '" & goUtil.utCleanSQLString(pudtRptParam.SortMe) & "', "
    sSQL = sSQL & "[IsDeleted] = " & pudtRptParam.IsDeleted & ", "
    sSQL = sSQL & "[DownLoadMe] = " & pudtRptParam.DownLoadMe & ", "
    sSQL = sSQL & "[UpLoadMe] = " & pudtRptParam.UpLoadMe & ", "
    sSQL = sSQL & "[AdminComments] = '" & goUtil.utCleanSQLString(pudtRptParam.AdminComments) & "', "
    sSQL = sSQL & "[DateLastUpdated] = #" & pudtRptParam.DateLastUpdated & "#, "
    sSQL = sSQL & "[UpdateByUserID]  = " & pudtRptParam.UpdateByUserID & " "
    sSQL = sSQL & "WHERE [MiscReportParamID] = " & pudtRptParam.MiscReportParamID
   
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    oConn.Execute sSQL
    
    
    EditRptParamItem = True
    
    'Clean up
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function EditRptParamItem"
End Function

Public Function GetPaymentsParams(psRTChecksID As String, psCheckNum As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sRTChecksID As String
    Dim sCheckNum As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT RTC.[RTChecksID], "
    sSQL = sSQL & "RTC.[CheckNum] "
    sSQL = sSQL & "FROM RTChecks RTC "
    sSQL = sSQL & "WHERE RTC.[AssignmentsID] = " & msAssignmentsID & " "
    If psRTChecksID = vbNullString Then
        sSQL = sSQL & "AND RTC.[CheckNum] = " & psCheckNum & " "
    Else
        sSQL = sSQL & "AND RTC.[RTChecksID] = " & psRTChecksID & " "
    End If
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        sRTChecksID = RS.Fields("RTChecksID").Value
        sCheckNum = RS.Fields("CheckNum").Value
    End If
    
    
    psRTChecksID = sRTChecksID
    psCheckNum = sCheckNum
    GetPaymentsParams = True
    
    'Cleanup
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetPaymentsParams"
End Function

Public Function GetIBParams(psIBID As String, psSupplement As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sIBID As String
    Dim sSupplement As String
    
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT IB.[IBID], "
    sSQL = sSQL & "BC.[Supplement] "
    sSQL = sSQL & "FROM IB "
    sSQL = sSQL & "INNER JOIN BillingCount BC ON IB.BillingCountID = BC.BillingCountID "
    sSQL = sSQL & "WHERE IB.[AssignmentsID] = " & msAssignmentsID & " "
    If psIBID = vbNullString Then
        sSQL = sSQL & "AND IB.[IB14a_sSupplement] = " & psSupplement & " "
    Else
        sSQL = sSQL & "AND IB.[IBID] = " & psIBID & " "
    End If
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        sIBID = RS.Fields("IBID").Value
        sSupplement = RS.Fields("Supplement").Value
    End If
    
    
    psIBID = sIBID
    psSupplement = sSupplement
    GetIBParams = True
    
    'Cleanup
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetIBParams"
End Function

Public Function GetRTIBParams(psSupplement As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    Dim sSupplement As String
    
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT BC.[Supplement] "
    sSQL = sSQL & "FROM RTIB "
    sSQL = sSQL & "INNER JOIN BillingCount BC ON RTIB.BillingCountID = BC.BillingCountID "
    sSQL = sSQL & "WHERE RTIB.[AssignmentsID] = " & msAssignmentsID & " "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If Not RS.EOF Then
        sSupplement = RS.Fields("Supplement").Value
    End If
    
    psSupplement = sSupplement
    GetRTIBParams = True
    
    'Cleanup
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetRTIBParams"
End Function

Public Function PrintActiveReport(poReportItem As Object, _
                                Optional piMode As VBRUN.FormShowConstants = vbModeless, _
                                Optional psCopyName As String = vbNullString, _
                                Optional pbPrintPreview As Boolean = True, _
                                Optional psSaveToFilePath As String, _
                                Optional psSaveToFileName As String, _
                                Optional pbExportXML As Boolean, _
                                Optional pbExportXMLOnly As Boolean) As Boolean
    PrintActiveReport = MyfrmClaimsList.PrintActiveReport(poReportItem, _
                                                          piMode, _
                                                          psCopyName, _
                                                          pbPrintPreview, _
                                                          psSaveToFilePath, _
                                                          psSaveToFileName, _
                                                          pbExportXML, _
                                                          pbExportXMLOnly)
    
    Exit Function
EH:
    Screen.MousePointer = vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PrintActiveReport"
End Function

Public Function GetRptParamColAndLoadLvw(pocboReport As Object, _
                                        polvwRptParams As ListView, _
                                        Optional poframMultiUpdate As Object = Nothing) As Boolean
    On Error GoTo EH
    Dim sParams As String
    Dim sReportName As String
    Dim srptProjectName As String
    Dim srptClassName As String
    Dim lrptVersion As Long
    Dim sData As String
    Dim saryData() As String
    Dim ocboReport As Object
    Dim colParams As Collection
    Dim oCarList As V2ECKeyBoard.clsCarLists
    'Some Reports need extra Params passed to them
    'Payments
    Dim sRTChecksID As String
    Dim sCheckNum As String
    'Internal Billing
    Dim sIBID As String
    Dim sSupplement As String
    'Photo Reports (Multi Report)
    Dim sPhotoReportNumber As String
    'Worksheet Diagram (Multi Report)
    Dim sDiagramNumber As String
    Dim sDiagramPhotoName As String
    
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    Dim vMiscReportParamName As Variant
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    
    If Not TypeOf pocboReport Is ListBox And Not TypeOf pocboReport Is ComboBox Then
        Exit Function
    Else
        Set ocboReport = pocboReport
    End If
    
    sData = ocboReport.Text
    
    If sData <> vbNullString Then
        sReportName = Trim(left(sData, 200))
        sData = Mid(sData, 200)
        sData = Trim(sData)
        'Do not Get Param for Loss Report
        If StrComp(sData, "LRFormat", vbTextCompare) = 0 Then
            polvwRptParams.ListItems.Clear
            GetRptParamColAndLoadLvw = True
            Exit Function
        End If
        saryData() = Split(sData, "|", , vbBinaryCompare)
        srptProjectName = saryData(0)
        srptClassName = saryData(1)
        lrptVersion = saryData(2)
        'Check For Multi Reports Here
        If InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
            'Photo Reports (Multi Report)
            sPhotoReportNumber = saryData(3)
        ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
            'Worksheet Diagram (Multi Report)
            sDiagramNumber = saryData(3)
            sDiagramPhotoName = saryData(4)
        End If
    Else
        Exit Function
    End If
    
    'Build Params List to be passed in to Create Report Object
    'This Object will have list of Report Parameters it requires
    sParams = vbNullString
    sParams = sParams & "psAssignmentsID=" & msAssignmentsID & "|"
    sParams = sParams & "pbGetObjectOnly=" & "True" & "|"
    
    'Certain Reports Need to have some more Params Passed in
    If InStr(1, srptProjectName, "_arRptAddlChk", vbTextCompare) > 0 Then
        'Need to Get the ChecksID and Check Number
        sRTChecksID = CStr(ocboReport.ItemData(ocboReport.ListIndex))
        If Not GetPaymentsParams(sRTChecksID, sCheckNum) Then
            GoTo CLEAN_UP
        End If
        sParams = sParams & "pRTChecksID=" & sRTChecksID & "|"
        sParams = sParams & "psCheckNum=" & sCheckNum & "|"
    ElseIf InStr(1, srptProjectName, "_arRptIB", vbTextCompare) > 0 Then
        sIBID = CStr(ocboReport.ItemData(ocboReport.ListIndex))
        If Not GetIBParams(sIBID, sSupplement) Then
            GoTo CLEAN_UP
        End If
        sParams = sParams & "pIBID=" & sIBID & "|"
        sParams = sParams & "pSupplement=" & sSupplement & "|"
    ElseIf InStr(1, srptProjectName, "_arRptPhotos", vbTextCompare) > 0 Then
        'Photo Reports (Multi Report)
        sParams = sParams & "pNumber=" & sPhotoReportNumber & "|"
    ElseIf InStr(1, srptProjectName, "_arWorkSheetDiag", vbTextCompare) > 0 Then
        'Worksheet Diagram (Multi Report)
        sParams = sParams & "pNumber=" & sDiagramNumber & "|"
    
    End If
    

    sReportName = srptProjectName & "." & srptClassName
    
    Set oCarList = CreateObject("V2ECcar" & goUtil.gsCurCarDBName & ".clsLists")
   
    Set colParams = oCarList.GetARMiscDelimParamsCol(sReportName, lrptVersion, sParams)
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
    If Not colParams Is Nothing Then
        vMiscReportParamName = goUtil.GetItemFromCollection(colParams, "sMiscReportParamName")
        If vMiscReportParamName = vbNullString Then
            msMiscReportParamName = "MiscReportParam"
        Else
            msMiscReportParamName = vMiscReportParamName
            goUtil.RemoveItemFromCollection colParams, "sMiscReportParamName"
        End If
    Else
        msMiscReportParamName = vbNullString
    End If
    '2/8/2004 MiscReportParam , and MiscReportParam01 to MiscReportParam30
     
    PopulatelvwRptParams colParams, polvwRptParams, poframMultiUpdate
    
    'If this is a diagram need to show the Diagram Button
    'And Set the memeber Diaram. If not need to Hide and clear
    If Not mfrmReports Is Nothing Then
        If IsNumeric(sDiagramNumber) Then
            mfrmReports.cmdEditDiagram.Enabled = True
            mfrmReports.cmdEditDiagram.Visible = True
            mfrmReports.EditDiagramNumber = CLng(sDiagramNumber)
            If goUtil.utFileExists(goUtil.PhotoReposPath & "\" & sDiagramPhotoName) Then
                mfrmReports.imgDiagram.Picture = LoadPicture(goUtil.PhotoReposPath & "\" & sDiagramPhotoName)
            Else
                mfrmReports.imgDiagram.Picture = LoadPicture()
            End If
        Else
            mfrmReports.cmdEditDiagram.Enabled = False
            mfrmReports.cmdEditDiagram.Visible = False
            mfrmReports.EditDiagramNumber = 0
            mfrmReports.imgDiagram.Picture = LoadPicture()
        End If
    End If
    
    GetRptParamColAndLoadLvw = True
    oCarList.CLEANUP
    Set oCarList = Nothing
CLEAN_UP:
    'cleanup
    Set ocboReport = Nothing
    Set colParams = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetRptParamColAndLoadLvw"
End Function

Public Sub PopulatelvwRptParams(pColParams As Collection, _
                                polvwRptParams As ListView, _
                                Optional poframMultiUpdate As Object = Nothing)
    On Error GoTo EH
    Dim iMyIcon As Long
    Dim sFlagText As String
    Dim sTemp As String
    Dim itmX As ListItem
    Dim oListView As MSComctlLib.ListView
    Dim colParams As Collection
    Dim MyParam As V2ECKeyBoard.MiscReportParam
    Dim vParam As Variant
    
    'Clear the List view
    Set oListView = polvwRptParams
    oListView.ListItems.Clear
   

    Set colParams = pColParams
    Set pColParams = Nothing
    
    If colParams Is Nothing Then
        If TypeOf poframMultiUpdate Is Frame Then
            poframMultiUpdate.Visible = False
        End If
        GoTo CLEAN_UP
    End If
    
    If colParams.Count = 0 Then
        If TypeOf poframMultiUpdate Is Frame Then
            poframMultiUpdate.Visible = False
        End If
        GoTo CLEAN_UP
    End If
    
    oListView.Visible = False
    
    'Before Populating the list view with this collection of Params...
    'need to see if the ID is Null string.  If it is then need to
    'Insert a record into MiscReportParam table for each Parameter that has never been
    'added in to the table
    Screen.MousePointer = MousePointerConstants.vbHourglass
    For Each vParam In colParams
      MyParam = vParam
      If MyParam.ID = vbNullString Then
            With MyParam
                .UpLoadMe = "True"
                .DateLastUpdated = Now()
                .UpdateByUserID = goUtil.gsCurUsersID
            End With
            If Not AddRptParamItem(MyParam) Then
                Screen.MousePointer = MousePointerConstants.vbDefault
                GoTo CLEAN_UP
            End If
            'Remove the old Param and add the updated one
            RemoveParam MyParam.ParamName, colParams
            colParams.Add MyParam, MyParam.ParamName
      End If
    Next
    Screen.MousePointer = MousePointerConstants.vbDefault
    For Each vParam In colParams
        MyParam = vParam
        'ParamCaption
        Set itmX = oListView.ListItems.Add(, """" & MyParam.ID & """", MyParam.ParamCaption)
        'ParamValue
        If MyParam.ParamDataType = vbBoolean Then
            'Check for Boolean values
            If CBool(MyParam.ParamValue) Then
                iMyIcon = GuiRptParamsStatusList.ValueIsChecked
            Else
                iMyIcon = GuiRptParamsStatusList.ValueIsUnchecked
            End If
            sFlagText = goUtil.GetFlagText(CBool(MyParam.ParamValue))
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = sFlagText
            itmX.ListSubItems(GuiRptParamsListView.ParamValue - 1).ReportIcon = iMyIcon
        ElseIf MyParam.ParamDataType = vbDate Then
            iMyIcon = GuiRptParamsStatusList.ShowCalendarBox
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = MyParam.ParamValue
            itmX.ListSubItems(GuiRptParamsListView.ParamValue - 1).ReportIcon = iMyIcon
        ElseIf MyParam.ParamDataType = vbUserDefinedType Then
            iMyIcon = GuiRptParamsStatusList.ShowUserDefinedToolBox
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = MyParam.ParamValue
            itmX.ListSubItems(GuiRptParamsListView.ParamValue - 1).ReportIcon = iMyIcon
        Else
            itmX.SubItems(GuiRptParamsListView.ParamValue - 1) = MyParam.ParamValue
        End If
       
        'IsDeleted
        If CBool(MyParam.IsDeleted) Then
            iMyIcon = GuiRptParamsStatusList.IsDeleted
        Else
            iMyIcon = Empty
        End If
        sFlagText = goUtil.GetFlagText(CBool(MyParam.IsDeleted))
        itmX.SubItems(GuiRptParamsListView.IsDeleted - 1) = sFlagText
        itmX.ListSubItems(GuiRptParamsListView.IsDeleted - 1).ReportIcon = iMyIcon
        'UpLoadMe
        '8. UpLoad Me
        If CBool(MyParam.UpLoadMe) Then
            iMyIcon = GuiRptParamsStatusList.UpLoadMe
        Else
            iMyIcon = Empty
        End If
        sFlagText = goUtil.GetFlagText(CBool(MyParam.UpLoadMe))
        itmX.SubItems(GuiRptParamsListView.UpLoadMe - 1) = sFlagText
        itmX.ListSubItems(GuiRptParamsListView.UpLoadMe - 1).ReportIcon = iMyIcon
        'DateLastUpdated
        If IsDate(MyParam.DateLastUpdated) Then
            itmX.SubItems(GuiRptParamsListView.DateLastUpdated - 1) = Format(MyParam.DateLastUpdated, "MM/DD/YYYY HH:MM:SS")
            'DateLastUpdatedSort ' Hidden
            itmX.SubItems(GuiRptParamsListView.DateLastUpdatedSort - 1) = Format(MyParam.DateLastUpdated, "YYYY/MM/DD HH:MM:SS")
        Else
            itmX.SubItems(GuiRptParamsListView.DateLastUpdated - 1) = vbNullString
            'DateLastUpdatedSort ' Hidden
            itmX.SubItems(GuiRptParamsListView.DateLastUpdatedSort - 1) = vbNullString
        End If
        'AdminComments
        itmX.SubItems(GuiRptParamsListView.AdminComments - 1) = MyParam.AdminComments
        'Number ' hidden
        itmX.SubItems(GuiRptParamsListView.Number - 1) = MyParam.Number
        'ProjectName ' hidden
        itmX.SubItems(GuiRptParamsListView.ProjectName - 1) = MyParam.ProjectName
        'ClassName ' hidden
        itmX.SubItems(GuiRptParamsListView.ClassName - 1) = MyParam.ClassName
        'ParamName ' hidden
        itmX.SubItems(GuiRptParamsListView.ParamName - 1) = MyParam.ParamName
        'ParamDataType ' hidden
        itmX.SubItems(GuiRptParamsListView.ParamDataType - 1) = MyParam.ParamDataType
        'SortMe ' hidden
        itmX.SubItems(GuiRptParamsListView.SortMe - 1) = MyParam.SortMe
        'SortMeSort ' hidden
        itmX.SubItems(GuiRptParamsListView.SortMeSort - 1) = goUtil.utNumInTextSortFormat(MyParam.SortMe)
        'ID ' Hidden
        itmX.SubItems(GuiRptParamsListView.ID - 1) = MyParam.ID
        'IDAssignments ' Hidden
        itmX.SubItems(GuiRptParamsListView.IDAssignments - 1) = MyParam.IDAssignments
        'MiscReportParamID ' Hidden
        itmX.SubItems(GuiRptParamsListView.MiscReportParamID - 1) = MyParam.MiscReportParamID
        'AssignmentsID ' Hidden
        itmX.SubItems(GuiRptParamsListView.AssignmentsID - 1) = MyParam.AssignmentsID
        'DownLoadMe ' hidden
        itmX.SubItems(GuiRptParamsListView.DownLoadMe - 1) = MyParam.DownLoadMe
        'UpdateByUserID ' Hidden
        itmX.SubItems(GuiRptParamsListView.UpdateByUserID - 1) = MyParam.UpdateByUserID
    
        itmX.Selected = False
    Next
    
    'Be sure this one is sorted by SortMe
    oListView.SortKey = GuiRptParamsListView.SortMe
    oListView.Sorted = True
    
    'Be sure to select the top one
    If oListView.ListItems.Count > 0 Then
        oListView.ListItems(1).Selected = True
    End If
    
    
    oListView.Visible = True
    If TypeOf poframMultiUpdate Is Frame Then
'        poframMultiUpdate.Visible = True
    End If
     
CLEAN_UP:
    'Cleanup
    Set itmX = Nothing
    Set colParams = Nothing
    Set oListView = Nothing
    
    Exit Sub
EH:
    Screen.MousePointer = MousePointerConstants.vbDefault
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub PopulatelvwRptParams"
    oListView.Visible = True
End Sub

Public Function UpdateAssgnStatus(pStatusID As AssgnStatus) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "UPDATE Assignments SET "
    sSQL = sSQL & "[StatusID] = " & pStatusID & ", "
    sSQL = sSQL & "[UpLoadMe] = True, "
    sSQL = sSQL & "[DateLastUpdated] = #" & Now() & "#, "
    sSQL = sSQL & "[UpdateByUserID] = " & goUtil.gsCurUsersID & " "
    sSQL = sSQL & "WHERE    [AssignmentsID] = " & msAssignmentsID & " "
 
    oConn.Execute sSQL
    
    UpdateAssgnStatus = True
    
    'CLEANUP
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function UpdateAssgnStatus"
End Function

Public Function GetCurrentActLogTime(psAssignmentsID As String, psBillingCountID As String) As Double
    On Error GoTo EH
    Dim RSBillingCount As ADODB.Recordset
    Dim sBillingCountID As String
    Dim sTemp As String
    Dim sTemp2 As String
    Dim cCurrActLogTime As Double
    
    SetadoRSBillingCount psAssignmentsID, True
    Set RSBillingCount = adoRSBillingCount
    
    If RSBillingCount.RecordCount = 0 Then
        GetCurrentActLogTime = 0#
        Exit Function
    End If
    
    RSBillingCount.MoveFirst
    
    Do Until RSBillingCount.EOF
        sBillingCountID = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("BillingCountID"))
        If StrComp(sBillingCountID, psBillingCountID, vbTextCompare) = 0 Then
            sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("IBDescription"))
            If sTemp = vbNullString Then
                sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("IBDescription2"))
            End If
            If sTemp = vbNullString Then
                GetCurrentActLogTime = 0#
            Else
                sTemp = Mid(sTemp, InStrRev(sTemp, "[", , vbBinaryCompare) + 1)
                sTemp = Replace(sTemp, "]", vbNullString)
                GetCurrentActLogTime = CDbl(sTemp)
            End If
            Exit Do
        End If
        RSBillingCount.MoveNext
    Loop
    
    'cleanup
    
    Set RSBillingCount = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetCurrentActLogTime"
End Function

Public Function GetCurrentPhotoCount(psAssignmentsID As String, psBillingCountID As String) As Double
    On Error GoTo EH
    Dim RSBillingCount As ADODB.Recordset
    Dim sBillingCountID As String
    Dim sTemp As String
    Dim sTemp2 As String
    Dim cCurrActLogTime As Double
    
    SetadoRSBillingCount psAssignmentsID, , True
    Set RSBillingCount = adoRSBillingCount
    
    If RSBillingCount.RecordCount = 0 Then
        GetCurrentPhotoCount = 0
        Exit Function
    End If
    
    RSBillingCount.MoveFirst
    
    Do Until RSBillingCount.EOF
        sBillingCountID = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("BillingCountID"))
        If StrComp(sBillingCountID, psBillingCountID, vbTextCompare) = 0 Then
            sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("IBDescription"))
            If sTemp = vbNullString Then
                sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("IBDescription2"))
            End If
            If sTemp = vbNullString Then
                GetCurrentPhotoCount = 0
            Else
                sTemp = Mid(sTemp, InStrRev(sTemp, "[", , vbBinaryCompare) + 1)
                sTemp = Replace(sTemp, "]", vbNullString)
                GetCurrentPhotoCount = CLng(sTemp)
            End If
            Exit Do
        End If
        RSBillingCount.MoveNext
    Loop
    
    'cleanup
    
    Set RSBillingCount = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetCurrentPhotoCount"
End Function

Public Function GetCurrentAmountOfCheck(psAssignmentsID As String, psBillingCountID As String) As Double
    On Error GoTo EH
    Dim RSBillingCount As ADODB.Recordset
    Dim sBillingCountID As String
    Dim sTemp As String
    Dim sTemp2 As String
    Dim cCurrActLogTime As Double
    
    SetadoRSBillingCount psAssignmentsID, , , True
    Set RSBillingCount = adoRSBillingCount
    
    If RSBillingCount.RecordCount = 0 Then
        GetCurrentAmountOfCheck = 0#
        Exit Function
    End If
    
    RSBillingCount.MoveFirst
    
    Do Until RSBillingCount.EOF
        sBillingCountID = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("BillingCountID"))
        If StrComp(sBillingCountID, psBillingCountID, vbTextCompare) = 0 Then
            sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("IBDescription"))
            If sTemp = vbNullString Then
                sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("IBDescription2"))
            End If
            If sTemp = vbNullString Then
                GetCurrentAmountOfCheck = 0#
            Else
                sTemp = Mid(sTemp, InStrRev(sTemp, "[", , vbBinaryCompare) + 1)
                sTemp = Replace(sTemp, "]", vbNullString)
                GetCurrentAmountOfCheck = CCur(sTemp)
            End If
            Exit Do
        End If
        RSBillingCount.MoveNext
    Loop
    
    'cleanup
    
    Set RSBillingCount = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetCurrentAmountOfCheck"
End Function

Public Function GetCurrentFeeScheduleItem(psFieldName As String, Optional psFeeScheduleID As String) As Variant
    On Error GoTo EH
    Dim RSFeeSched As ADODB.Recordset
    
    moGUI.SetadoRSFeeSchedule psFeeScheduleID
    
    Set RSFeeSched = moGUI.adoFeeSchedule
    
    If RSFeeSched.RecordCount <> 1 Then
        GetCurrentFeeScheduleItem = 0
        Exit Function
    End If
    
    RSFeeSched.MoveFirst
    
    GetCurrentFeeScheduleItem = goUtil.IsNullIsVbNullString(RSFeeSched.Fields(psFieldName))
    
    'cleanup
    
    Set RSFeeSched = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetCurrentFeeScheduleItem"
End Function

Public Function GetCurrentBillingCountID(pbGetIDOnly As Boolean) As String
    On Error GoTo EH
    Dim RSBillingCount As ADODB.Recordset
    Dim sTemp As String
    
    SetadoRSBillingCount msAssignmentsID, True
    Set RSBillingCount = adoRSBillingCount
    
    If RSBillingCount.RecordCount = 0 Then
        GetCurrentBillingCountID = "Null"
        Exit Function
    End If
    
    RSBillingCount.MoveFirst
    
    Do Until RSBillingCount.EOF
        sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("IBDescription"))
        If sTemp = vbNullString Then
            sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("IBDescription2"))
        End If
        If sTemp <> vbNullString Then
            If InStr(1, sTemp, "Current", vbTextCompare) > 0 Then
                If pbGetIDOnly Then
                    sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("ID"))
                Else
                    sTemp = goUtil.IsNullIsVbNullString(RSBillingCount.Fields("BillingCountID"))
                End If
                If sTemp = vbNullString Or sTemp = "0" Then
                    GetCurrentBillingCountID = "Null"
                Else
                    GetCurrentBillingCountID = sTemp
                End If
                Exit Function
            End If
        End If
        RSBillingCount.MoveNext
    Loop
    
    GetCurrentBillingCountID = "Null"
    
    'cleanup
    
    Set RSBillingCount = Nothing
    
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetCurrentBillingCountID"
End Function

Public Function GetClassOfLossCode(psClassOfLossID As String) As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT  Code As ClassOfLossCode "
    sSQL = sSQL & "FROM     ClassOfLoss "
    sSQL = sSQL & "WHERE    ClassOfLossID = " & psClassOfLossID & " "
    sSQL = sSQL & "AND      ClientCompanyID = " & goUtil.gsCurCar & " "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
        
    If RS.RecordCount <> 1 Then
        Exit Function
    End If
    
    RS.MoveFirst
    
    GetClassOfLossCode = goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode"))
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetClassOfLossCode"
End Function

Public Function GetIBID(plCurrentBillingcountID As Long) As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT [IBID] "
    sSQL = sSQL & "FROM IB "
    sSQL = sSQL & "WHERE [BillingCountID] = " & plCurrentBillingcountID & " "
    
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        GetIBID = goUtil.IsNullIsVbNullString(RS.Fields("IBID"))
    Else
        GetIBID = "0"
    End If
    
    'cleanup
    
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetIBID"
End Function

Public Function GetFeeScheduleID(psAssignmentsID As String, psBillingCountID, pbClosedIB) As String
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT [FeeScheduleID] "
    If pbClosedIB Then
        sSQL = sSQL & "FROM IB "
    Else
        sSQL = sSQL & "FROM RTIB "
    End If
    sSQL = sSQL & "WHERE AssignmentsID = " & psAssignmentsID & " "
    sSQL = sSQL & "AND BillingCountID  = " & psBillingCountID & " "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        GetFeeScheduleID = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleID"))
    Else
        GetFeeScheduleID = "0"
    End If
    
    If GetFeeScheduleID = "0" Then
    
        Set RS = Nothing
        Set RS = New ADODB.Recordset
        
        sSQL = "SELECT   [FeeScheduleID] "
        sSQL = sSQL & "FROM     ClientCompanyCat "
        sSQL = sSQL & "WHERE    [ClientCompanyID] = " & goUtil.gsCurCar & " "
        sSQL = sSQL & "AND      [CATID] = " & goUtil.gsCurCat & " "
        
        RS.CursorLocation = adUseClient
        RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
        Set RS.ActiveConnection = Nothing
        
        If RS.RecordCount = 1 Then
            GetFeeScheduleID = goUtil.IsNullIsVbNullString(RS.Fields("FeeScheduleID"))
        Else
            GetFeeScheduleID = "0"
        End If
    End If
    
    'cleanup
    Set RS = Nothing
    Set oConn = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetFeeScheduleID"
End Function

Public Function GetSupplement(plCurrentBillingcountID As Long) As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT [Supplement] "
    sSQL = sSQL & "FROM BillingCount "
    sSQL = sSQL & "WHERE [BillingCountID] = " & plCurrentBillingcountID & " "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        GetSupplement = goUtil.IsNullIsVbNullString(RS.Fields("Supplement"))
    Else
        GetSupplement = "0"
    End If
    
    'cleanup
    
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetSupplement"
End Function

Public Function GetRebill(plCurrentBillingcountID As Long) As String
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim sSQL As String
    
    Set oConn = New ADODB.Connection
    Set RS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    sSQL = "SELECT [Rebill] "
    sSQL = sSQL & "FROM BillingCount "
    sSQL = sSQL & "WHERE [BillingCountID] = " & plCurrentBillingcountID & " "
    
    RS.CursorLocation = adUseClient
    RS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set RS.ActiveConnection = Nothing
    
    If RS.RecordCount = 1 Then
        GetRebill = goUtil.IsNullIsVbNullString(RS.Fields("Rebill"))
    Else
        GetRebill = "0"
    End If
    
    'cleanup
    
    Set oConn = Nothing
    Set RS = Nothing
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetRebill"
End Function

Public Function GetFullCostOfRepair(plCurrentBillingcountID As Long) As Currency
    On Error GoTo EH
    Dim RSBillingCountItem As ADODB.Recordset
    Dim RS As ADODB.Recordset
    Dim cFullCostOfRepair As Currency
    Dim lCurrentSupplement As Long
    Dim lSupplement As Long
    
    'Set Current Billing
    If Not SetadoRSBillingCountItem(msAssignmentsID, CStr(plCurrentBillingcountID)) Then
        Exit Function
    End If
    
    'Set the Indemnity
    If Not SetadoRSRTIndemnity(msAssignmentsID) Then
        Exit Function
    End If
    
    Set RSBillingCountItem = adoRSBillingCountItem
    
    'Get the Indemnity RS
    Set RS = adoRSRTIndemnity
    
     'Set the Current Supplement
    lCurrentSupplement = goUtil.IsNullIsVbNullString(RSBillingCountItem.Fields("Supplement"))

    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            'Only include this Supplement and any previous Supplements
            lSupplement = goUtil.IsNullIsVbNullString(RS.Fields("Supplement"))
            If lSupplement <= lCurrentSupplement Then
                If Not RS.Fields("IsPreviousPayment") And Not RS.Fields("IsDeleted") And StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other", vbTextCompare) <> 0 Then
                    cFullCostOfRepair = cFullCostOfRepair + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ReplacementCost")), "#,###,###,##0.00"))
                End If
            End If
            RS.MoveNext
        Loop
    End If

    GetFullCostOfRepair = cFullCostOfRepair
    
    'cleanup
    Set RSBillingCountItem = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetFullCostOfRepair"
End Function

Public Function GetDepreciation(plCurrentBillingcountID As Long) As Currency
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim RSBillingCountItem As ADODB.Recordset
    Dim cRecoverableDepreciation As Currency
    Dim cNonRecovDepr As Currency
    Dim lCurrentSupplement As Long
    Dim lSupplement As Long
    
    'Set Current Billing
    If Not SetadoRSBillingCountItem(msAssignmentsID, CStr(plCurrentBillingcountID)) Then
        Exit Function
    End If
    
    'Set the Indemnity
    If Not SetadoRSRTIndemnity(msAssignmentsID) Then
        Exit Function
    End If
    
    Set RSBillingCountItem = adoRSBillingCountItem
    
    'Get the Indemnity RS
    Set RS = adoRSRTIndemnity
    
    'Set the Current Supplement
    lCurrentSupplement = goUtil.IsNullIsVbNullString(RSBillingCountItem.Fields("Supplement"))

    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            'Only include this Supplement and any previous Supplements
            lSupplement = goUtil.IsNullIsVbNullString(RS.Fields("Supplement"))
            If lSupplement <= lCurrentSupplement Then
                If Not RS.Fields("IsPreviousPayment") And Not RS.Fields("IsDeleted") And StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other", vbTextCompare) <> 0 Then
                    cRecoverableDepreciation = cRecoverableDepreciation + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("RecoverableDep")), "#,###,###,##0.00"))
                    cNonRecovDepr = cNonRecovDepr + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("NonRecoverableDep")), "#,###,###,##0.00"))
                End If
            End If
            RS.MoveNext
        Loop
    End If

    GetDepreciation = cRecoverableDepreciation + cNonRecovDepr
    
    'cleanup
    Set RSBillingCountItem = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetDepreciation"
End Function

Public Function GetLessExcessLimits(plCurrentBillingcountID As Long) As Currency
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim RSBillingCountItem As ADODB.Recordset
    Dim cLessExcessLimits As Currency
    Dim lCurrentSupplement As Long
    Dim lSupplement As Long
    
    'Set Current Billing
    If Not SetadoRSBillingCountItem(msAssignmentsID, CStr(plCurrentBillingcountID)) Then
        Exit Function
    End If
    
    'Set the Indemnity
    If Not SetadoRSRTIndemnity(msAssignmentsID) Then
        Exit Function
    End If
    
    Set RSBillingCountItem = adoRSBillingCountItem
    
    'Get the Indemnity RS
    Set RS = adoRSRTIndemnity
    
    'Set the Current Supplement
    lCurrentSupplement = goUtil.IsNullIsVbNullString(RSBillingCountItem.Fields("Supplement"))

    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            'Only include this Supplement and any previous Supplements
            lSupplement = goUtil.IsNullIsVbNullString(RS.Fields("Supplement"))
            If lSupplement <= lCurrentSupplement Then
                If Not RS.Fields("IsPreviousPayment") And Not RS.Fields("IsDeleted") And StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other", vbTextCompare) <> 0 Then
                    cLessExcessLimits = cLessExcessLimits + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ExcessLimits")), "#,###,###,##0.00"))
                End If
            End If
            RS.MoveNext
        Loop
    End If

    GetLessExcessLimits = cLessExcessLimits
    
    'cleanup
    Set RSBillingCountItem = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetLessExcessLimits"
End Function

Public Function GetLessMiscellaneous(plCurrentBillingcountID As Long) As Currency
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim RSBillingCountItem As ADODB.Recordset
    Dim cLessMiscellaneous As Currency
    Dim lCurrentSupplement As Long
    Dim lSupplement As Long
    
    
    'Set Current Billing
    If Not SetadoRSBillingCountItem(msAssignmentsID, CStr(plCurrentBillingcountID)) Then
        Exit Function
    End If
    
    'Set the Indemnity
    If Not SetadoRSRTIndemnity(msAssignmentsID) Then
        Exit Function
    End If
    
    Set RSBillingCountItem = adoRSBillingCountItem
    
    'Get the Indemnity RS
    Set RS = adoRSRTIndemnity
    
    'Set the Current Supplement
    lCurrentSupplement = goUtil.IsNullIsVbNullString(RSBillingCountItem.Fields("Supplement"))

    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do Until RS.EOF
            'Only include this Supplement and any previous Supplements
            lSupplement = goUtil.IsNullIsVbNullString(RS.Fields("Supplement"))
            If lSupplement <= lCurrentSupplement Then
                If Not RS.Fields("IsPreviousPayment") And Not RS.Fields("IsDeleted") And StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other", vbTextCompare) <> 0 Then
                    cLessMiscellaneous = cLessMiscellaneous + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("Miscellaneous")), "#,###,###,##0.00"))
                End If
            End If
            RS.MoveNext
        Loop
    End If

    GetLessMiscellaneous = cLessMiscellaneous
    
    'cleanup
    Set RSBillingCountItem = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetLessMiscellaneous"
End Function

Public Function GetNetActualCashValueClaim(plCurrentBillingcountID As Long, _
                                            Optional pbUseOverrides As Boolean, _
                                            Optional pcOverrideGrossLoss As Currency, _
                                            Optional pcOverrideDepreciation As Currency, _
                                            Optional pcOverrideExcessLimit As Currency, _
                                            Optional pcOverrideMiscellaneous As Currency) As Currency
    On Error GoTo EH
    Dim RS As ADODB.Recordset
    Dim RSBillingCountItem As ADODB.Recordset
    Dim cACVLoss As Currency
    Dim cAppliedDeductible As Currency
    Dim cLessExcessLimits As Currency
    Dim cLessExcessLimitsAbsorbDed As Currency
    Dim cLessMiscellaneous As Currency
    Dim cDeductible As Currency
    Dim lCurrentSupplement As Long
    Dim lSupplement As Long
    
    'Set Current Billing
    If Not SetadoRSBillingCountItem(msAssignmentsID, CStr(plCurrentBillingcountID)) Then
        GoTo CLEAN_UP
    End If
    
    'Set the Deductible
    'Get the Assignments RS to get Deductible
    If Not SetadoRSAssignments(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    Set RS = adoRSAssignments
    
    If RS.RecordCount = 1 Then
        RS.MoveFirst
        cDeductible = CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("Deductible")), "#,###,###,##0.00"))
    End If
    
    Set RS = Nothing
    
    'Set Current Billing
    If Not SetadoRSBillingCountItem(msAssignmentsID, CStr(plCurrentBillingcountID)) Then
        GoTo CLEAN_UP
    End If
    
    'Set the Indemnity
    If Not SetadoRSRTIndemnity(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    
    Set RSBillingCountItem = adoRSBillingCountItem
    
    'Get the Indemnity RS
    Set RS = adoRSRTIndemnity
    
    'Set the Current Supplement
    lCurrentSupplement = goUtil.IsNullIsVbNullString(RSBillingCountItem.Fields("Supplement"))
    
    If Not pbUseOverrides Then
        If RS.RecordCount > 0 Then
            RS.MoveFirst
            Do Until RS.EOF
                'Only include this Supplement and any previous Supplements
                lSupplement = goUtil.IsNullIsVbNullString(RS.Fields("Supplement"))
                If lSupplement <= lCurrentSupplement Then
                    If Not RS.Fields("IsPreviousPayment") And Not RS.Fields("IsDeleted") And StrComp(goUtil.IsNullIsVbNullString(RS.Fields("ClassOfLossCode")), "Other", vbTextCompare) <> 0 Then
                        cACVLoss = cACVLoss + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ACVClaim")), "#,###,###,##0.00"))
                        cAppliedDeductible = cAppliedDeductible + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("AppliedDeductible")), "#,###,###,##0.00"))
                        cLessExcessLimits = cLessExcessLimits + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ExcessLimits")), "#,###,###,##0.00"))
                        If RS.Fields("ExcessAbsorbsDeductible") Then
                            cLessExcessLimitsAbsorbDed = cLessExcessLimitsAbsorbDed + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("ExcessLimits")), "#,###,###,##0.00"))
                        End If
                        cLessMiscellaneous = cLessMiscellaneous + CCur(Format(goUtil.IsNullIsVbNullString(RS.Fields("Miscellaneous")), "#,###,###,##0.00"))
                    End If
                End If
                RS.MoveNext
            Loop
        End If
    Else
        cACVLoss = pcOverrideGrossLoss - pcOverrideDepreciation
        cAppliedDeductible = cDeductible
        cLessExcessLimits = pcOverrideExcessLimit
        cLessMiscellaneous = pcOverrideMiscellaneous
    End If
    
    If cLessExcessLimitsAbsorbDed > 0 Then
        If cAppliedDeductible <> cDeductible - cLessExcessLimitsAbsorbDed Then
            cLessExcessLimitsAbsorbDed = cDeductible - cAppliedDeductible
        End If
    End If

    GetNetActualCashValueClaim = cACVLoss - (cAppliedDeductible + cLessExcessLimits + cLessMiscellaneous)
    
CLEAN_UP:
    'cleanup
    Set RSBillingCountItem = Nothing
    Set RS = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetNetActualCashValueClaim"
End Function

Public Function GetCurrentServiceFee(plCurrentBillingcountID As Long, _
                                     Optional psFeeScheduleID As String, _
                                     Optional pbUseOverrides As Boolean, _
                                     Optional pcOverrideGrossLoss As Currency, _
                                     Optional pcOverrideDepreciation As Currency, _
                                     Optional pcOverrideExcessLimit As Currency, _
                                     Optional pcOverrideMiscellaneous As Currency) As Currency
    On Error GoTo EH
    Dim RSFeeSchedLevels As ADODB.Recordset
    Dim RSIB As ADODB.Recordset
    Dim RSBillingCountItem As ADODB.Recordset
    Dim cFullCostOfRepair As Currency
    Dim cCurrentServiceFee As Currency
    Dim cPrevServiceFees As Currency
    Dim lCurrentSupplement As Long
    Dim lSupplement As Long
    Dim cLevelMax As Currency
    Dim dblPctApp As Double
    Dim cLevelMin As Currency
    
    'Set Current Billing
    If Not SetadoRSBillingCountItem(msAssignmentsID, CStr(plCurrentBillingcountID)) Then
        GoTo CLEAN_UP
    End If
    
    If Not moGUI.SetadoRSFeeScheduleLevels(psFeeScheduleID) Then
        GoTo CLEAN_UP
    End If
    
    If Not SetadoRSIB(msAssignmentsID) Then
        GoTo CLEAN_UP
    End If
    
    Set RSBillingCountItem = adoRSBillingCountItem
    Set RSFeeSchedLevels = moGUI.adoFeeScheduleLevels
    Set RSIB = adoRSIB
    
    'Set the Current Supplement
    lCurrentSupplement = goUtil.IsNullIsVbNullString(RSBillingCountItem.Fields("Supplement"))
    
    'Need to figure the Previous Service Fees
    If RSIB.RecordCount > 0 Then
        RSIB.MoveFirst
        Do Until RSIB.EOF
            'Do not include the current supplement in previous totals
            'and only include those supplement Previous to the current
            lSupplement = goUtil.IsNullIsVbNullString(RSIB.Fields("IB14a_sSupplement"))
            If lSupplement < lCurrentSupplement Then
                cPrevServiceFees = cPrevServiceFees + goUtil.IsNullIsVbNullString(RSIB.Fields("IB17_cServiceFee"))
            End If
            RSIB.MoveNext
        Loop
    End If
    
    If Not pbUseOverrides Then
        cFullCostOfRepair = GetFullCostOfRepair(plCurrentBillingcountID)
    Else
        cFullCostOfRepair = pcOverrideGrossLoss
    End If
    
    'Need to figure out what the cCurrentServiceFee is according to the FeeSchedule
    If RSFeeSchedLevels.RecordCount > 0 Then
        RSFeeSchedLevels.MoveFirst
        Do Until RSFeeSchedLevels.EOF
            cLevelMax = goUtil.IsNullIsVbNullString(RSFeeSchedLevels.Fields("LevelMax"))
            dblPctApp = goUtil.IsNullIsVbNullString(RSFeeSchedLevels.Fields("LevelPctApp"))
            cLevelMin = goUtil.IsNullIsVbNullString(RSFeeSchedLevels.Fields("LevelMin"))
            'If the full cost of repair falls within the levelmax then use the minimum
            'Fee amount unless the pctApp is greater than the Minimum.
            If cFullCostOfRepair <= cLevelMax Then
                cCurrentServiceFee = cLevelMin
                If cCurrentServiceFee < cFullCostOfRepair * (dblPctApp / 100) Then
                    cCurrentServiceFee = cFullCostOfRepair * (dblPctApp / 100)
                End If
                'Now that got the correct one can exit
                Exit Do
            End If
            RSFeeSchedLevels.MoveNext
        Loop
    Else
        GoTo CLEAN_UP
    End If
    
    'Now that have the CUrrent service Fee need to subtract any previous payments
    
    If cCurrentServiceFee >= cPrevServiceFees Then
        cCurrentServiceFee = cCurrentServiceFee - cPrevServiceFees
    Else
        cCurrentServiceFee = 0
    End If
    
    GetCurrentServiceFee = cCurrentServiceFee
    
CLEAN_UP:
    'cleanup
    Set RSBillingCountItem = Nothing
    Set RSFeeSchedLevels = Nothing
    Set RSIB = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetCurrentServiceFee"
End Function

Public Function PopulateSectionLevels(psIDAssignments As String, _
                                    psProjectName As String, _
                                    Optional psSL01 As String, _
                                    Optional psSL02 As String, _
                                    Optional psSL03 As String, _
                                    Optional psSL04 As String, _
                                    Optional psSL05 As String, _
                                    Optional psSL06 As String, _
                                    Optional psSL07 As String, _
                                    Optional psSL08 As String, _
                                    Optional psSL09 As String, _
                                    Optional psSL10 As String) As Boolean
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim sSQL As String
    Dim lMainSPVersion As Long
   
    
    If Not madoRSMainReports Is Nothing Then
        Set madoRSMainReports = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  Application A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "AND      A.[ProjectName] Like '%" & psProjectName & "%' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    adoRS.CursorLocation = adUseClient
    adoRS.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set adoRS.ActiveConnection = Nothing
    
    If adoRS.RecordCount = 1 Then
        psSL01 = adoRS.Fields("SectionLevel01").Value
        psSL02 = adoRS.Fields("SectionLevel02").Value
        psSL03 = adoRS.Fields("SectionLevel03").Value
        psSL04 = adoRS.Fields("SectionLevel04").Value
        psSL05 = adoRS.Fields("SectionLevel05").Value
'            psSL06 = adoRS.Fields("").Value
'            psSL07 = adoRS.Fields("").Value
'            psSL08 = adoRS.Fields("").Value
'            psSL09 = adoRS.Fields("").Value
'            psSL10 = adoRS.Fields("").Value
    End If
    
    PopulateSectionLevels = True
    
    Set oConn = Nothing
    Set adoRS = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    Set adoRS = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function PopulateSectionLevels"
End Function

Public Function GetadoRSApplication(psIDAssignments As String, _
                                    Optional psSL01 As String, _
                                    Optional psSL02 As String, _
                                    Optional psSL03 As String, _
                                    Optional psSL04 As String, _
                                    Optional psSL05 As String, _
                                    Optional psSL06 As String, _
                                    Optional psSL07 As String, _
                                    Optional psSL08 As String, _
                                    Optional psSL09 As String, _
                                    Optional psSL10 As String) As ADODB.Recordset
    On Error GoTo EH
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim lMainSPVersion As Long
    Dim sSL01 As String
    Dim sSL02 As String
    Dim sSL03 As String
    Dim sSL04 As String
    Dim sSL05 As String
    Dim sSL06 As String
    Dim sSL07 As String
    Dim sSL08 As String
    Dim sSL09 As String
    Dim sSL10 As String
    
    sSL01 = Trim(psSL01)
    sSL02 = Trim(psSL02)
    sSL03 = Trim(psSL03)
    sSL04 = Trim(psSL04)
    sSL05 = Trim(psSL05)
    sSL06 = Trim(psSL06)
    sSL07 = Trim(psSL07)
    sSL08 = Trim(psSL08)
    sSL09 = Trim(psSL09)
    sSL10 = Trim(psSL10)
    
    If Not madoRSMainReports Is Nothing Then
        Set madoRSMainReports = Nothing
    End If
    
    Set oConn = New ADODB.Connection
    Set GetadoRSApplication = New ADODB.Recordset
    goUtil.utOpenDatabaseADOConn oConn, goUtil.gMainDB.Name
    
    lMainSPVersion = madoRSAssignments.Fields("SPVersion").Value
    
    sSQL = "SELECT A.* "
    sSQL = sSQL & "FROM ((SoftwarePackage SP "
    sSQL = sSQL & "INNER JOIN  SoftwarePackageApplication SPA "
    sSQL = sSQL & "ON SPA.[SoftWarePackageID] = SP.[SoftWarePackageID]) "
    sSQL = sSQL & "INNER JOIN  Application A "
    sSQL = sSQL & "ON A.[ApplicationID] = SPA.[ApplicationID]) "
    sSQL = sSQL & "WHERE    A.[IsDeleted] = 0 "
    sSQL = sSQL & "AND " & lMainSPVersion & " >= A.[SPVersionBase] "
    sSQL = sSQL & "AND " & lMainSPVersion & " <= A.[SPVersion] "
    sSQL = sSQL & "AND      SPA.[IsDeleted] = 0 "
    sSQL = sSQL & "AND      SP.[CATID] = " & goUtil.gsCurCat & " "
    sSQL = sSQL & "AND      SP.[CLientCompanyID] = " & goUtil.gsCurCar & " "
    If sSL01 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel01] Like '" & sSL01 & "' "
    End If
    If sSL02 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel02] Like '" & sSL02 & "' "
    End If
    If sSL03 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel03] Like '" & sSL03 & "' "
    End If
    If sSL04 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel04] Like '" & sSL04 & "' "
    End If
    If sSL05 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel05] Like '" & sSL05 & "' "
    End If
    If sSL06 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel06] Like '" & sSL06 & "' "
    End If
    If sSL07 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel07] Like '" & sSL07 & "' "
    End If
    If sSL08 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel08] Like '" & sSL08 & "' "
    End If
    If sSL09 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel09] Like '" & sSL09 & "' "
    End If
    If sSL10 <> vbNullString Then
        sSQL = sSQL & "AND      A.[SectionLevel10] Like '" & sSL10 & "' "
    End If
    sSQL = sSQL & "AND      A.[ProjectName] Is Not Null "
    sSQL = sSQL & "AND      A.[ProjectName] <> '' "
    sSQL = sSQL & "Order by A.[Description], A.[ProjectName], A.[ClassName] "

    'Use Disconnected Record Set on asUseClient Cusor ONLY !
    GetadoRSApplication.CursorLocation = adUseClient
    GetadoRSApplication.Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set GetadoRSApplication.ActiveConnection = Nothing
    
    
    Set oConn = Nothing
    Exit Function
EH:
    Set oConn = Nothing
    Set GetadoRSApplication = Nothing
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function GetadoRSApplication"
End Function

Public Function RequestClaimApproval(pocmdRequestApproval As Object, _
                                     Optional polvwAssignments As Object, _
                                     Optional polvwPackageItem As Object) As Boolean
    On Error GoTo EH
    
    RequestClaimApproval = MyfrmClaimsList.RequestClaimApproval(pocmdRequestApproval, polvwAssignments, polvwPackageItem)
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function RequestClaimApproval"
End Function
