VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form dlgStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2985
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer timUnload 
      Left            =   255
      Top             =   2490
   End
   Begin MSComctlLib.ProgressBar prgOverallStatus 
      Height          =   330
      Left            =   255
      TabIndex        =   2
      Top             =   1980
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3270
      TabIndex        =   1
      Top             =   2415
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2445
      Width           =   1215
   End
   Begin VB.Label lblText 
      Caption         =   "(status text)"
      Height          =   1035
      Left            =   270
      TabIndex        =   4
      Top             =   510
      Width           =   5625
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "(title)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   255
      TabIndex        =   3
      Top             =   15
      Width           =   5625
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "dlgStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim m_bCancelled As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = m_bCancelled
End Property


Private Sub CloseMe()
    Me.Visible = False
    Unload Me
End Sub


Private Sub cmdCancel_Click()
m_bCancelled = True
End Sub

Private Sub timUnload_Timer()
    CloseMe
End Sub


