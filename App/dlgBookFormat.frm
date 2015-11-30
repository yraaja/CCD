VERSION 5.00
Begin VB.Form dlgBookFormat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book Preview"
   ClientHeight    =   5355
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgBookFormat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstBooks 
      Height          =   4740
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame fraUnitCost 
      Caption         =   "Unit Cost Range"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
      Begin VB.ComboBox cboMasterFormat 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtUnitCostEnd 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtUnitCostStart 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblMasterFormat 
         Caption         =   "MasterFormat Version"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Preview"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Book Edition:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "dlgBookFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oReturnResult As VbMsgBoxResult
Private m_sXmlReportName As String
Private m_sXmlFileName As String
Private m_sBookDescription As String
Private m_iBookID As Long
Private m_sUnitCostStart As String
Private m_sUnitCostEnd As String
Private m_iMasterFormatVersion As Long
Private m_bAllowEditing As Boolean
Private m_colBookPreviewInfo As New Collection

Private Const XML_REPORT_STD As String = "rptBookFormat.xml"
Private Const XML_REPORT_OPN As String = "rptBookFormat_OPN.xml"
Private Const XML_REPORT_RES As String = "rptBookFormat_RES.xml"
Private Const XML_REPORT_RR As String = "rptBookFormat_RR.xml"
'


Public Property Get Result() As VbMsgBoxResult
    Result = m_oReturnResult
End Property

Public Property Get XMLReportName() As String
    XMLReportName = m_sXmlReportName
End Property
Public Property Let XMLReportName(Value As String)
    m_sXmlReportName = Value
End Property

Public Property Get XMLFileName() As String
    XMLFileName = m_sXmlFileName
End Property
Public Property Let XMLFileName(Value As String)
    m_sXmlFileName = Value
End Property

Public Property Get BookDescription() As String
    BookDescription = m_sBookDescription
End Property

Public Property Get bookid() As Long
    bookid = m_iBookID
End Property

Public Property Get UnitCostIdStart() As String
    UnitCostIdStart = m_sUnitCostStart
End Property
Public Property Let UnitCostIdStart(Value As String)
    m_sUnitCostStart = Value
    txtUnitCostStart.Text = m_sUnitCostStart
End Property

Public Property Get UnitCostIdEnd() As String
    UnitCostIdEnd = m_sUnitCostEnd
End Property
Public Property Let UnitCostIdEnd(Value As String)
    m_sUnitCostEnd = Value
    txtUnitCostEnd.Text = m_sUnitCostEnd
End Property

Public Property Get MasterFormatVersion() As Long
    MasterFormatVersion = m_iMasterFormatVersion
End Property
Public Property Let MasterFormatVersion(Value As Long)
    m_iMasterFormatVersion = Value
    SelectMasterFormat m_iMasterFormatVersion
    UpdateXmlReportName
End Property

Public Property Get AllowEditing() As Boolean
    AllowEditing = m_bAllowEditing
End Property
Public Property Let AllowEditing(Value As Boolean)
    m_bAllowEditing = Value
    If m_bAllowEditing Then
        fraUnitCost.Enabled = True
        UnLockField Me, "cboMasterFormat"
        UnLockField Me, "txtUnitCostStart"
        UnLockField Me, "txtUnitCostEnd"
    Else
        fraUnitCost.Enabled = False
        LockField Me, "cboMasterFormat"
        LockField Me, "txtUnitCostStart"
        LockField Me, "txtUnitCostEnd"
    End If
End Property

Private Sub UpdateXmlReportName()

    m_sXmlReportName = "MF-" & MasterFormatVersion & " "
    If bookid = 81 Or bookid = 91 Then
        m_sXmlReportName = m_sXmlReportName + "Metric Book Format"
    Else
        m_sXmlReportName = m_sXmlReportName + "Standard Book Format"
    End If
    
End Sub

' This routine will add the book info to our collection and the lstBooks control.
' The strBookName will be displayed to the user.
' iBookID is the numeric identifier passed to the stored proc to retrieve the book's data.
' strReportName is the name of the xml report file that will be used to display the report.

Private Function AddBook(strBookName As String, iBookID As Integer, strReportName As String) As Long
   
    Dim bookPreviewInfo As New CBkPrevInfo
    bookPreviewInfo.BookName = strBookName
    bookPreviewInfo.bookid = iBookID
    bookPreviewInfo.ReportName = strReportName
    
    m_colBookPreviewInfo.Add bookPreviewInfo
    
    lstBooks.AddItem strBookName
    lstBooks.ItemData(lstBooks.NewIndex) = lstBooks.NewIndex + 1
    
    AddBook = lstBooks.NewIndex
    
End Function
' Removes any books already in the list of the lstBooks control.
Private Sub ClearBookList()
    
    Dim Index As Integer
    For Index = m_colBookPreviewInfo.Count To 1 Step -1
        m_colBookPreviewInfo.Remove Index
    Next
    
End Sub

Private Sub LoadBookList()
    Dim iBCCD As Long
    
    ' Reset the index into the lstBooks control to 0
    ClearBookList
    
    iBCCD = AddBook("Building Construction Cost Data", 1, XML_REPORT_STD)
    AddBook "Assemblies Construction Cost Data", 6, XML_REPORT_STD
    AddBook "Concrete & Masonry Cost Data", 9, XML_REPORT_STD
    AddBook "Electrical Construction Cost Data", 3, XML_REPORT_STD
    AddBook "Green Building Cost Data", 57, XML_REPORT_STD
    AddBook "Heavy Construction Cost Data", 11, XML_REPORT_STD
    AddBook "Interior Construction Cost Data", 8, XML_REPORT_STD
    AddBook "Mechanical Construction Cost Data", 2, XML_REPORT_STD
    AddBook "Plumbing Construction Cost Data", 16, XML_REPORT_STD
    AddBook "Site Work & Landscape Cost Data", 7, XML_REPORT_STD
    AddBook "Square Foot Costs", 5, XML_REPORT_STD
    AddBook "", 0, XML_REPORT_STD
    
    AddBook "Open Shop Construction Cost Data", 10, XML_REPORT_OPN
    AddBook "Light Commercial Cost Data", 13, XML_REPORT_OPN
    AddBook "", 0, XML_REPORT_STD
    
    AddBook "Residential Construction Cost Data", 12, XML_REPORT_RES
    AddBook "", 0, XML_REPORT_STD
    
    AddBook "Facilities Construction Cost Data", 15, XML_REPORT_RR
    AddBook "Commercial Renovation Cost Data", 4, XML_REPORT_RR
    AddBook "Facility Maintenance & Repair Cost Data", 27, XML_REPORT_RR
    AddBook "", 0, XML_REPORT_STD
    
    AddBook "Building Construction Cost Data (Metric)", 81, XML_REPORT_STD
    AddBook "Heavy Construction Cost Data (Metric)", 91, XML_REPORT_STD
    
    lstBooks.Selected(iBCCD) = True
    
End Sub

Public Function SelectMasterFormat(iMasterFormat As Long) As Boolean
'SET THE MASTERFORMAT COMBO BOX TO THE NEW SELECTION
'ADDED 8/2/2005 RTD
    Dim i As Long
    
    cboMasterFormat.ListIndex = -1
    For i = 0 To cboMasterFormat.listcount - 1
        If cboMasterFormat.ItemData(i) = iMasterFormat Then
            cboMasterFormat.ListIndex = i
            SelectMasterFormat = True
            Exit For
        End If
    Next
    
End Function


Private Sub CancelButton_Click()
    m_oReturnResult = vbCancel
    Unload Me
End Sub

Private Sub cboMasterFormat_Change()
    UpdateXmlReportName
End Sub

Private Sub Form_Initialize()
    m_oReturnResult = vbCancel
    m_sXmlReportName = ""
    m_sXmlFileName = "rptBookFormat.xml"
    m_iBookID = 0
    m_sBookDescription = ""
End Sub

Private Sub Form_Load()
    AllowEditing = False
    LoadMasterFormatCombo cboMasterFormat, True
    LoadBookList
End Sub

Private Sub lstBooks_Click()
    
    m_sBookDescription = lstBooks.List(lstBooks.ListIndex)
    m_iBookID = m_colBookPreviewInfo(lstBooks.ItemData(lstBooks.ListIndex)).bookid
    m_sXmlFileName = m_colBookPreviewInfo(lstBooks.ItemData(lstBooks.ListIndex)).ReportName
    UpdateXmlReportName
    OKButton.Enabled = (BookDescription <> "")
    
End Sub

Private Sub OKButton_Click()
    m_oReturnResult = vbOK
    Unload Me
End Sub
