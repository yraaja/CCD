VERSION 5.00
Begin VB.Form dlgUserRole 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set User Roles"
   ClientHeight    =   3225
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Select Role(s)"
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.CheckBox chkRole 
         Caption         =   "Admin"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   8
         Tag             =   "128"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox chkRole 
         Caption         =   "Role 16"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   7
         Tag             =   "16"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox chkRole 
         Caption         =   "Role 8"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Tag             =   "8"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CheckBox chkRole 
         Caption         =   "Role 4"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Tag             =   "4"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CheckBox chkRole 
         Caption         =   "Role 2"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Tag             =   "2"
         Top             =   720
         Width           =   3375
      End
      Begin VB.CheckBox chkRole 
         Caption         =   "User"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Tag             =   "0"
         Top             =   360
         Value           =   1  'Checked
         Width           =   3375
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "dlgUserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oReturnResult As VbMsgBoxResult
Private m_iUserRole As Integer

Private Enum UserRoleBits
    ROLE_BIT_USER = 0
    ROLE_BIT_ADMIN = 128
End Enum
'

Private Sub SetCheckDescriptions()
    
    chkRole(0).Caption = "Standard User"
    chkRole(1).Caption = "Undefined Role 2"
    chkRole(2).Caption = "Undefined Role 4"
    chkRole(3).Caption = "Undefined Role 8"
    chkRole(4).Caption = "Undefined Role 16"
    
    chkRole(5).Caption = "Administrator"
    chkRole(5).Tag = 128
    
End Sub

Private Sub UpdateCheckBoxes()
    
    If (m_iUserRole And UserRoleBits.ROLE_BIT_USER) = UserRoleBits.ROLE_BIT_USER Then
        Me.chkRole(0).Value = vbChecked
    End If
    If (m_iUserRole And UserRoleBits.ROLE_BIT_ADMIN) = UserRoleBits.ROLE_BIT_ADMIN Then
        Me.chkRole(5).Value = vbChecked
    End If

End Sub

Private Sub UpdateUserRoleProperty()
    Dim I As Integer
    Dim V As Integer
    
    For I = 0 To 5
        If chkRole(I).Value = vbChecked Then
            V = V + chkRole(I).Tag
        End If
    Next
    m_iUserRole = V

End Sub

Public Property Get UserRole() As Integer
    UserRole = m_iUserRole
End Property
Public Property Let UserRole(NewValue As Integer)
    m_iUserRole = NewValue
    UpdateCheckBoxes
End Property

Public Property Get ReturnResult() As VbMsgBoxResult
    ReturnResult = m_oReturnResult
End Property

Private Sub chkRole_Click(Index As Integer)
    UpdateUserRoleProperty
End Sub

Private Sub Form_Initialize()
    m_oReturnResult = vbCancel
    m_iUserRole = 0
End Sub

Private Sub Form_Load()
    SetCheckDescriptions
End Sub

Private Sub OKButton_Click()
    UpdateUserRoleProperty
    m_oReturnResult = vbOK
    Unload Me
End Sub

Private Sub CancelButton_Click()
    m_oReturnResult = vbCancel
    Unload Me
End Sub


