VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReleaseNotes 
   Caption         =   "Release Notes"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   Icon            =   "frmReleaseNotes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   9180
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9180
      TabIndex        =   2
      Top             =   5310
      Width           =   9180
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   200
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
   End
   Begin RichTextLib.RichTextBox RTBox 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5953
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmReleaseNotes.frx":0442
   End
End
Attribute VB_Name = "frmReleaseNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Unload Me
    Set frmReleaseNotes = Nothing
    
End Sub

Private Sub Form_Activate()
    
    OutputView False

End Sub

Private Sub Form_Load()
    
    On Error GoTo Error_Processing
    RTBox.FileName = App.Path + "\ReleaseNotes.rtf"
    Me.Height = 6300
    Me.Width = 9300

Exit_Sub:
    Exit Sub

Error_Processing:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume Exit_Sub
    
End Sub

Private Sub Form_Resize()

    If Not Me.WindowState = vbMinimized Then
        If Me.Height < 1095 Then
            Me.Height = 1095
        End If
        RTBox.Height = Me.Height - (350 + Picture1.Height) - (RTBox.top * 2)
        RTBox.Width = Me.Width - 100 - (RTBox.left * 2)
    End If
    
End Sub

Private Sub Picture1_Resize()

    cmdClose.left = (Picture1.Width - cmdClose.Width) / 2

End Sub

Private Sub RTBox_GotFocus()
    cmdClose.SetFocus
End Sub
