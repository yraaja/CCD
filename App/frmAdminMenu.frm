VERSION 5.00
Begin VB.Form frmAdminMenu 
   Caption         =   "System Administration"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdminMenu.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   8085
   Begin VB.CommandButton cmdReportsAdmin 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   480
      Picture         =   "frmAdminMenu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdVersionAdmin 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   480
      Picture         =   "frmAdminMenu.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdUserAdmin 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   480
      Picture         =   "frmAdminMenu.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblButtonSubheader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage the Reports Menu"
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   9
      Top             =   2880
      Width           =   1905
   End
   Begin VB.Label lblButtonHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Manager"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label lblButtonHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version Administration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   2220
   End
   Begin VB.Label lblButtonSubheader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage the CCD version table"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Width           =   2190
   End
   Begin VB.Label lblButtonSubheader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage database users"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   1725
   End
   Begin VB.Label lblButtonHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Administration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1920
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CCD Admin Control Panel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   3660
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11100
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmAdminMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_sngYCoord As Single
'
'   Keeps up with the field that last had focus when form
'   is deactivate, so when activated can set focus.
Dim m_strCurrentFormControl As String
'
'   Notifies that it wants to see changes.
Dim sEventSubscriberID As String
'

Private Function LockButtons() As Boolean
' DISABLE ALL COMMAND BUTTONS IF THE USER IS NOT AN ADMIN
    Dim bEnabled As Boolean
    Dim ctl As Control
    
    bEnabled = g_blnIsUserAdmin
    For Each ctl In Me.Controls
        If TypeOf ctl Is CommandButton Then
            ctl.Enabled = bEnabled
        End If
    Next
    
End Function

Private Sub cmdReportsAdmin_Click()
    Dim frm As New frmAdminReports
    frm.Show
    
End Sub

Private Sub cmdUserAdmin_Click()
    Dim frm As New frmAdminUsers
    frm.Show
    
End Sub

Private Sub cmdVersionAdmin_Click()
    Dim frm As New frmAdminVersions
    frm.Show
    
End Sub

Private Sub Form_Activate()
    Dim ctl As Control
    
    If Me.WindowState <> vbMinimized Then
        If Len(m_strCurrentFormControl) > 0 Then
            For Each ctl In Me.Controls
                If ctl.Name = m_strCurrentFormControl Then
                    ctl.SetFocus
                    Exit For
                End If
            Next ctl
        End If
        OutputView True
        'ShowToolbarIcons True
    End If
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    'ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()
    
    Status ("Loading Admin Control Panel...")
    Screen.MousePointer = vbHourglass
    sEventSubscriberID = EventSubscriberAdd(Me)
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim blnReturn As Boolean
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    LockButtons
    
    Status ("")
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '
    '   Need to place in common routine for all forms.
    '   Possibly place all buttons in a frame like frame1 with
    '   common name and can just place it.
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 10500 Then
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 10500
        End If
        
        If Me.Height >= 6135 Then
        Else
            Me.Height = 6135
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'ShowToolbarIcons False
    EventSubscriberRemove sEventSubscriberID
    
End Sub
