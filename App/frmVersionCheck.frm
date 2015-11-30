VERSION 5.00
Begin VB.Form frmVersionCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Construction Cost Database"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVersionCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpgrade 
      Caption         =   "Upgrade"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmVersionCheck.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblHeader 
      Caption         =   "lblHeader"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label lblContent 
      Caption         =   "lblContent"
      Height          =   1455
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
End
Attribute VB_Name = "frmVersionCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_result As VbMsgBoxResult
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Property Get Result() As VbMsgBoxResult
    Result = m_result
End Property
Public Property Let Result(NewValue As VbMsgBoxResult)
    m_result = NewValue
End Property

Public Property Get Header() As String
    Header = lblHeader.Caption
End Property
Public Property Let Header(NewValue As String)
    lblHeader.Caption = NewValue
End Property

Public Property Get Content() As String
    Content = lblContent.Caption
End Property
Public Property Let Content(NewValue As String)
    lblContent.Caption = NewValue
End Property

Private Sub cmdCancel_Click()
    m_result = vbCancel
    Unload Me
    Set frmVersionCheck = Nothing
End Sub

Private Sub cmdContinue_Click()
    m_result = vbIgnore
    Unload Me
    Set frmVersionCheck = Nothing
End Sub

Private Sub cmdUpgrade_Click()
    '*************************
    ' Display Upgrade URL
    ' AKD - 9/15/2006
    '*************************
    Dim sURL As String
    Dim res
    m_result = vbCancel
    
    sURL = "\\binnwldatp001\rsmeans$\CCD\Help\upgrade.htm"

    Screen.MousePointer = vbHourglass
    res = ShellExecute(0&, "open", sURL, vbNullString, vbNullString, vbMaximizedFocus)
    If res > 32 Then
        'Call BringWindowToTop(res)
    Else
        MsgBox "The Upgrade page failed to start.", vbCritical + vbOKOnly
    End If
    Screen.MousePointer = vbDefault
    Unload Me
    Set frmVersionCheck = Nothing
End Sub

Private Sub Form_Initialize()
    m_result = vbCancel
End Sub

