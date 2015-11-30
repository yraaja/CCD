VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBkPrevInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is used to keep track of book ids and xml report file names.
' It is used in the dlgBookFormat form.

Private m_strBookName As String
Private m_iBookID As Integer
Private m_strReportName As String

Public Property Get BookName() As String
    BookName = m_strBookName
End Property


Public Property Let BookName(strBookName As String)
    m_strBookName = strBookName
End Property

Public Property Get Bookid() As Integer
    Bookid = m_iBookID
End Property


Public Property Let Bookid(iBookID As Integer)
    m_iBookID = iBookID
End Property

Public Property Get ReportName() As String
    ReportName = m_strReportName
End Property


Public Property Let ReportName(strReportName As String)
    m_strReportName = strReportName
End Property



