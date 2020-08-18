VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmECUpdate 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "UpdateBatches"
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   Icon            =   "frmDummy.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framRecords 
      Caption         =   "EC Update Batches Record Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin MSComctlLib.ProgressBar PBarRecord 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Tag             =   "Refresh"
         Top             =   600
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "2 of 2002"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Tag             =   "Refresh"
         Top             =   240
         Width           =   7185
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   7920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      Index           =   1
      X1              =   7680
      X2              =   7680
      Y1              =   -120
      Y2              =   2160
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      Index           =   0
      X1              =   10
      X2              =   10
      Y1              =   -120
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   -120
      X2              =   7800
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmECUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

