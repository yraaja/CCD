VERSION 5.00
Begin VB.Form dlgFactor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update by Factor"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4635
   Icon            =   "dlgFactor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkTRACES 
      Caption         =   "TRACES"
      Height          =   255
      Left            =   2460
      TabIndex        =   8
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CheckBox chkList 
      Caption         =   "List"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txtComment 
      Height          =   615
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Tag             =   "2S"
      Top             =   600
      Width           =   3435
   End
   Begin VB.TextBox txtFactor 
      DataField       =   "contact_id"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2460
      TabIndex        =   3
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      Height          =   255
      Left            =   1620
      TabIndex        =   6
      Top             =   180
      Width           =   195
   End
   Begin VB.Label lblComment 
      Alignment       =   1  'Right Justify
      Caption         =   "Latest Price Comment::"
      Height          =   495
      Left            =   60
      TabIndex        =   5
      Top             =   660
      Width           =   915
   End
   Begin VB.Label lblFactor 
      Alignment       =   1  'Right Justify
      Caption         =   "Factor:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "dlgFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub GetFactor(ByRef dblFactor As Double, ByRef strComment As String, ByRef intColumns As Integer)
    Show (vbModal)
    dblFactor = Val(txtFactor)
    strComment = txtComment
    If chkList.Value = 1 Then
        intColumns = intColumns Or 1
    End If
    If chkTRACES.Value = 1 Then
        intColumns = intColumns Or 2
    End If
    Clear
End Sub

Private Sub CancelButton_Click()
    Clear
    Me.Hide
End Sub

Private Sub chkList_Click()
    chkTRACES.Value = chkList.Value
    If chkList.Value = 0 Then
        chkTRACES.Enabled = True
    Else
        chkTRACES.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    OutputView False

End Sub

Private Sub OKButton_Click()
    If Trim(txtFactor) = "" Or Trim(txtFactor) = "0" Then
        MsgBox "Please enter a factor."
        txtFactor.SetFocus
    ElseIf chkTRACES.Value = 0 And chkList.Value = 0 Then
            MsgBox "Please select traces or list price."
            chkList.SetFocus
    Else
        Me.Hide
    End If
End Sub

Private Sub Clear()
    txtFactor = ""
    txtComment = ""
End Sub
