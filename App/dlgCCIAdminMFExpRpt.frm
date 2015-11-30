VERSION 5.00
Begin VB.Form dlgCCIADminMFExpRpt 
   Caption         =   "MF Exc Rpt Parameters"
   ClientHeight    =   6540
   ClientLeft      =   4875
   ClientTop       =   2130
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   4455
   Begin VB.TextBox txtPct 
      Height          =   495
      Index           =   5
      Left            =   2520
      TabIndex        =   11
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtPct 
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   10
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtPct 
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtPct 
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtPct 
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtPct 
      Height          =   495
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblTotals 
      Caption         =   "TOTALS"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label lblInstallations 
      Caption         =   "INSTALLATIONS"
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblMaterials 
      Caption         =   "MATERIALS"
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblPct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< Percentage 3 (start)"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   13
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblPct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Percentage 3 (end) >>"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   12
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblPct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Percentage 2 (end) >>"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblPct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< Percentage 2 (start)"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblPct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Percentage 1 (end) >>"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblPct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<<  Percentage 1 (start)"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "dlgCCIADminMFExpRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Mode = False
Unload Me
End Sub

Private Sub cmdGo_Click()
If CDbl(Me.txtPct(0)) = 0 Then
    PCT1 = "1.0"
Else
    PCT1 = Me.txtPct(0)
End If
If CDbl(Me.txtPct(2)) = 0 Then
    PCT3 = "1.0"
Else
    PCT3 = Me.txtPct(2)
End If
If CDbl(Me.txtPct(4)) = 0 Then
    PCT5 = "1.0"
Else
    PCT5 = Me.txtPct(4)
End If

PCT2 = Me.txtPct(1)
'PCT3 = Me.txtPct(2)
PCT4 = Me.txtPct(3)
'PCT5 = Me.txtPct(4)
PCT6 = Me.txtPct(5)

'Save settings(percentages) to the registry

'HOW TO FIND IT:
'HKEY_CURRENT_USER
'SOFTWARE
'MICROSOFT
'VB and VBA Program Settings
'Construction Cost Database
'Settings

Call SaveSetting(App.title, "SETTINGS", "PCT1", PCT1)
Call SaveSetting(App.title, "SETTINGS", "PCT2", PCT2)
Call SaveSetting(App.title, "SETTINGS", "PCT3", PCT3)
Call SaveSetting(App.title, "SETTINGS", "PCT4", PCT4)
Call SaveSetting(App.title, "SETTINGS", "PCT5", PCT5)
Call SaveSetting(App.title, "SETTINGS", "PCT6", PCT6)

Mode = True

Unload Me

End Sub

Private Sub Form_Load()
'Me.txtPct(0) = "0.00"
'Me.txtPct(1) = "1.10"
'Me.txtPct(2) = "0.90"
'Me.txtPct(3) = "1.10"
'Me.txtPct(4) = "0.90"
'Me.txtPct(5) = "1.10"


If GetSetting(App.title, "SETTINGS", "PCT1") <> "" Then
    Me.txtPct(0) = GetSetting(App.title, "SETTINGS", "PCT1")
Else
    Me.txtPct(0) = "0.00"
End If

If GetSetting(App.title, "SETTINGS", "PCT2") <> "" Then
    Me.txtPct(1) = GetSetting(App.title, "SETTINGS", "PCT2")
Else
    Me.txtPct(1) = "1.10"
End If

If GetSetting(App.title, "SETTINGS", "PCT3") <> "" Then
    Me.txtPct(2) = GetSetting(App.title, "SETTINGS", "PCT3")
Else
    Me.txtPct(2) = "0.90"
End If


If GetSetting(App.title, "SETTINGS", "PCT4") <> "" Then
    Me.txtPct(3) = GetSetting(App.title, "SETTINGS", "PCT4")
Else
    Me.txtPct(3) = "1.10"
End If

If GetSetting(App.title, "SETTINGS", "PCT5") <> "" Then
    Me.txtPct(4) = GetSetting(App.title, "SETTINGS", "PCT5")
Else
    Me.txtPct(4) = "0.90"
End If

If GetSetting(App.title, "SETTINGS", "PCT6") <> "" Then
    Me.txtPct(5) = GetSetting(App.title, "SETTINGS", "PCT6")
Else
    Me.txtPct(5) = "1.10"
End If



End Sub
