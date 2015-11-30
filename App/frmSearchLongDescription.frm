VERSION 5.00
Begin VB.Form frmSearchLongDescription 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   Icon            =   "frmSearchLongDescription.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Matc&h Case?"
      Height          =   195
      Left            =   8160
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox chkReplaceAll 
      Caption         =   "Replace All?"
      Height          =   195
      Left            =   8160
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkFindAll 
      Caption         =   "Find &All?"
      Height          =   195
      Left            =   8160
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox chkReplace 
      Caption         =   "&Replace?"
      Height          =   195
      Left            =   8160
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtReplacementText 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.TextBox txtSearchText 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label lblReplaceText 
      Caption         =   "Replace &With:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblFindWhat 
      Caption         =   "&Find What:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearchLongDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SearchText As String
Public ReplacementText As String
Public Replace As Boolean
Public FindAll As Boolean
Public ReplaceAll As Boolean
Public Cancel As Boolean
Public MatchCase As Boolean

Private Sub chkFindAll_Click()
    txtSearchText.SetFocus
End Sub

Private Sub chkMatchCase_Click()
    txtSearchText.SetFocus
End Sub

Private Sub chkReplace_Click()

    If chkReplace Then
        lblReplaceText.Visible = True
        txtReplacementText.Visible = True
        chkReplaceAll.Visible = True
        With Me
            .Caption = "Search and Replace"
            .Height = .Height + 780
            cmdCancel.top = cmdCancel.top + 720
            cmdOK.top = cmdOK.top + 720
        End With
    Else
        lblReplaceText.Visible = False
        txtReplacementText.Visible = False
        chkReplaceAll.Visible = False
        
        With Me
            Me.Caption = "Search"
            .Height = .Height - 780
            cmdCancel.top = cmdCancel.top - 720
            cmdOK.top = cmdOK.top - 720
        End With
    End If
    txtSearchText.SetFocus
    
End Sub

Private Sub chkReplaceAll_Click()
    txtSearchText.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Cancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Cancel = False
    SearchText = txtSearchText
    ReplacementText = txtReplacementText
    Replace = chkReplace
    FindAll = chkFindAll
    ReplaceAll = chkReplaceAll
    MatchCase = chkMatchCase
    Unload Me
End Sub

Private Sub Form_Load()
    txtSearchText = SearchText
    txtReplacementText = ReplacementText
End Sub
