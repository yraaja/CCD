VERSION 5.00
Begin VB.Form dlgListSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selection"
   ClientHeight    =   7650
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5685
   Icon            =   "dlgListSelection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox picReportOptions 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   5685
      TabIndex        =   14
      Top             =   6300
      Visible         =   0   'False
      Width           =   5685
      Begin VB.CheckBox Check2 
         Caption         =   "Report Option 2"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Report Option 1"
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox picSingle 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   5685
      TabIndex        =   10
      Top             =   0
      Width           =   5685
      Begin VB.ComboBox cmbSelection 
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox lstSingleSelection 
         Height          =   2985
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblComboSelection 
         Caption         =   "Select Item:"
         Height          =   495
         Left            =   2520
         TabIndex        =   13
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.PictureBox picMulti 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   5685
      TabIndex        =   3
      Top             =   3135
      Width           =   5685
      Begin VB.ListBox lstSelected 
         Height          =   2595
         Left            =   3240
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   165
         Width           =   2295
      End
      Begin VB.ListBox lstAvailable 
         Columns         =   1
         Height          =   2595
         ItemData        =   "dlgListSelection.frx":0442
         Left            =   120
         List            =   "dlgListSelection.frx":0444
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdSelection 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   2520
         TabIndex        =   7
         Top             =   525
         Width           =   615
      End
      Begin VB.CommandButton cmdSelection 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   1005
         Width           =   615
      End
      Begin VB.CommandButton cmdSelection 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   2520
         TabIndex        =   5
         Top             =   1485
         Width           =   615
      End
      Begin VB.CommandButton cmdSelection 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   2520
         TabIndex        =   4
         Top             =   1965
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   7155
      Width           =   5685
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdFinished 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "dlgListSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_strSingleItemValue As String
Dim m_lngSingleItemData As Long
Dim m_SelectType As Integer
Dim m_strComboSelectionCaption As String
Dim m_strCheck1Caption  As String
Dim m_strCheck2Caption  As String
Dim m_blnCancel As Boolean

Property Let Check1Caption(Caption As String)
    m_strCheck1Caption = Caption
    Check1.Caption = m_strCheck1Caption
End Property

Property Let Check2Caption(Caption As String)
    m_strCheck2Caption = Caption
    Check2.Caption = m_strCheck2Caption
End Property
Property Get SingleItemdata() As Long
    SingleItemdata = m_lngSingleItemData
End Property
Property Get Cancel() As Boolean
    Cancel = m_blnCancel
End Property
Property Get SingleItemValue() As String
    SingleItemValue = m_strSingleItemValue
End Property
Property Let SingleItemValue(sValue As String)
    m_strSingleItemValue = sValue
    Dim I As Integer
    cmbSelection.Text = sValue
End Property
Property Let ComboSelectionCaption(Caption As String)
    m_strComboSelectionCaption = Caption
    lblComboSelection.Caption = m_strComboSelectionCaption
End Property
Property Let SelectType(intType As Integer)
    m_SelectType = intType
    
End Property

Private Sub cmdCancel_Click()
m_strSingleItemValue = Empty
cmbSelection.ListIndex = -1
m_blnCancel = True
Me.Visible = False
End Sub

Private Sub cmdFinished_Click()

Select Case m_SelectType

Case SINGLE_LIST
    If lstSingleSelection.ListIndex = -1 Then
        m_strSingleItemValue = ""
        m_lngSingleItemData = 0
    Else
        m_strSingleItemValue = lstSingleSelection.List(lstSingleSelection.ListIndex)
        m_lngSingleItemData = lstSingleSelection.ItemData(lstSingleSelection.ListIndex)
    End If

Case COMBO_BOX
    If Len(cmbSelection.Text) = 0 Then
        m_lngSingleItemData = 0
    Else
        If cmbSelection.ListIndex <> -1 Then
            m_lngSingleItemData = cmbSelection.ItemData(cmbSelection.ListIndex)
        End If
    End If
    m_strSingleItemValue = cmbSelection.Text
End Select

Me.Visible = False

End Sub


Private Sub Form_Resize()
If Me.Width > 400 Then
    Check1.Width = Me.Width - 400
    cmbSelection.Width = Me.Width - 400
End If
End Sub

