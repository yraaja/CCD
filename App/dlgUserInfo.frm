VERSION 5.00
Begin VB.Form dlgUserInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information"
   ClientHeight    =   3855
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6105
   Icon            =   "dlgUserInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      Height          =   3495
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   4215
      Begin VB.TextBox last_logon 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Tag             =   "ignore"
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox ccd_version 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "ignore"
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox user_fax 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox user_extension 
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox user_name 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox user_id 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lbl_last_logon 
         Caption         =   "Last Logon"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lbl_ccd_version 
         Caption         =   "CCD Ver."
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "&Fax #"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "&Extension"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "&Name"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "User ID"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Update"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "dlgUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DISPLAY CURRENT APPLICATION USER PROFILE

Option Explicit
Private objUserInfo As New cUserInfo
Private m_rec As ADODB.RecordSet
Private m_sUserId As String
Private m_bExtendedInfo As Boolean
Private m_bReadOnly As Boolean
Private m_bNewUserMode As Boolean

'THE USER-ID TO DISPLAY
Public Property Get UserID() As String
    UserID = m_sUserId
End Property
Public Property Let UserID(NewValue As String)
    m_sUserId = NewValue
End Property

'IF READONLY - USER CAN VIEW INFORMATION, BUT NOT EDIT OR UPDATE
Public Property Get ReadOnly() As Boolean
    ReadOnly = m_bReadOnly
End Property
Public Property Let ReadOnly(NewValue As Boolean)
    m_bReadOnly = NewValue
    SetReadOnlyControls
End Property

'IF TRUE - SHOW EXTENDED USER PROPERTIES
'USEFUL IF VIEWING ANOTHER USER'S INFORMATION
Public Property Get ShowExtendedInfo() As Boolean
    ReadOnly = m_bExtendedInfo
End Property
Public Property Let ShowExtendedInfo(NewValue As Boolean)
    m_bExtendedInfo = NewValue
    SetFormView
End Property

Public Sub NewUser()

    m_bNewUserMode = True
    ShowExtendedInfo = False
    ReadOnly = False
    UserID = ""
    RetrieveData
    UnLockField Me, "user_id"
    
End Sub

Private Sub SetFormView()

    If m_bExtendedInfo Then
        Me.Height = 4230
        Frame1.Height = 3495
    Else
        Me.Height = 3270
        Frame1.Height = 2535
    End If
    lbl_ccd_version.Visible = m_bExtendedInfo
    ccd_version.Visible = m_bExtendedInfo
    lbl_last_logon.Visible = m_bExtendedInfo
    last_logon.Visible = m_bExtendedInfo
    
End Sub

Private Sub SetReadOnlyControls()

    If m_bReadOnly Then
        LockField Me, "user_name"
        LockField Me, "user_extension"
        LockField Me, "user_fax"
        cmdOK.Enabled = False
    Else
        UnLockField Me, "user_name"
        UnLockField Me, "user_extension"
        UnLockField Me, "user_fax"
        cmdOK.Enabled = True
    End If

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    Set dlgUserInfo = Nothing
    
End Sub

Private Sub cmdOK_Click()

    If UpdateData Then
        EventSubscriberNotify esnUserRecordupdated, m_sUserId
        'Unload Me
        'Set dlgUserInfo = Nothing
    End If
    
End Sub

Private Sub Form_Load()

    RetrieveData
    SetFormView
    SetReadOnlyControls
    
End Sub

Private Function UpdateData() As Boolean

    On Error GoTo Err_Handler
    objUserInfo.UserID = Me.user_id.Text
    objUserInfo.UserName = Me.user_name.Text
    objUserInfo.UserExtension = Me.user_extension.Text
    objUserInfo.UserFaxNumber = Me.user_fax.Text
    objUserInfo.ApplicationVersion = Me.ccd_version.Text
    objUserInfo.LastLoginTimestamp = Me.last_logon.Text
    If objUserInfo.UpdateData Then
        MsgBox "Record updated successfully.", vbOKOnly + vbInformation
        UpdateData = True
    Else
        'MsgBox "An error occurred while updating the database:" & vbCrLf & objUserInfo.LastError, vbCritical, "Error"
        UpdateData = False
    End If
    
    Exit Function

Err_Handler:
    Screen.MousePointer = vbDefault
    MsgBox "An error occurred while updating the database:" & vbCrLf & Err.Description, vbCritical, "Error #" & Err.Number
    UpdateData = False
    Exit Function

End Function

Private Sub RetrieveData()
    
    If (m_sUserId = "") And Not m_bNewUserMode Then
        ' GET CURRENT USER ID IF NONE WAS PASSED
        m_sUserId = strUserName
    End If
    objUserInfo.UserID = m_sUserId
    objUserInfo.GetData
    Me.user_id.Text = objUserInfo.UserID
    Me.user_name.Text = objUserInfo.UserName
    Me.user_extension.Text = objUserInfo.UserExtension
    Me.user_fax.Text = objUserInfo.UserFaxNumber
    Me.ccd_version.Text = objUserInfo.ApplicationVersion
    Me.last_logon.Text = objUserInfo.LastLoginTimestamp
    cmdOK.Enabled = Not (Me.user_id.Text = "")

End Sub

Private Sub user_fax_LostFocus()
    user_fax.Text = FormatPhoneNumber(user_fax.Text)
End Sub

Private Sub user_id_Change()
    cmdOK.Enabled = (user_id.Text <> "")
End Sub
