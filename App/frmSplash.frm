VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7485
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0442
   ScaleHeight     =   4320
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7380
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   120
         Picture         =   "frmSplash.frx":0884
         ScaleHeight     =   1245
         ScaleWidth      =   3135
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Change:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   2970
         Width           =   1695
      End
      Begin VB.Label lblLastChange 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   2970
         Width           =   3255
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status..."
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Tag             =   "Warning"
         Top             =   3900
         Width           =   7215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Means"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   555
         Left            =   540
         TabIndex        =   8
         Tag             =   "CompanyProduct"
         Top             =   1665
         Width           =   1515
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "LicenseTo"
         Height          =   255
         Left            =   3390
         TabIndex        =   1
         Tag             =   "LicenseTo"
         Top             =   120
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Construction Cost Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   180
         TabIndex        =   7
         Tag             =   "Product"
         Top             =   2160
         Width           =   5055
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "RS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   180
         TabIndex        =   6
         Tag             =   "CompanyProduct"
         Top             =   1675
         Width           =   375
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version X.X.X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Tag             =   "Version"
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Tag             =   "Warning"
         Top             =   3315
         Width           =   6855
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copyright 1999 © R.S. Means"
         Height          =   1095
         Left            =   4080
         TabIndex        =   4
         Tag             =   "Copyright"
         Top             =   480
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    
    OutputView False

End Sub

Private Sub Form_Load()
    Dim Found As Boolean
    Dim szFileName As String
    Dim strTimeStamp As String
    
    'szFileName = "C:\Program Files\Construction Cost Database\ConstructionCostDatabase.exe"
    szFileName = App.Path
    If Right(szFileName, 1) <> "\" Then szFileName = szFileName & "\"
    szFileName = szFileName & App.EXEName & ".exe"
    Found = FileExist(szFileName)
    If Found = True Then
        strTimeStamp = FileDateTime(szFileName)
    Else
        strTimeStamp = "n/a"
    End If

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.title
    lblCopyright.Caption = App.LegalCopyright
    lblDescription.Caption = App.FileDescription
    lblLastChange.Caption = strTimeStamp
    DoEvents
    
End Sub

Public Sub SetStatus(strStatus As String)
    
    lblStatus.Caption = strStatus
    lblStatus.Refresh

End Sub

Private Function FileExist(ByVal szFileName As String) As Boolean
    Dim nFileNumber As Integer
    
    On Error Resume Next
    nFileNumber = FreeFile
    'Try to open the file
    Open szFileName For Input As nFileNumber
    'If it fails the file doesn't exist
    If Err.Number <> 0 Then
        FileExist = False
    Else
        FileExist = True
    End If
    Close nFileNumber
    
End Function

