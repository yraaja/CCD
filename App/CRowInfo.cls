VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_vntBookmark As Variant
Private m_strErrorText As String
Private m_strRowSQL As String  'Used for row refresh of data
Private m_blnError As Boolean
Private m_intState As Integer
Private m_lngSkey As Long
Private m_strSubKey As String  'Sub-key - save for deletes after grid record has been deleted


Public Property Get SubKey() As String
    SubKey = m_strSubKey
End Property

Public Property Get Bookmark() As Variant
    Bookmark = m_vntBookmark
End Property

Public Property Get RowSQL() As String
    RowSQL = m_strRowSQL
End Property


Public Property Let Bookmark(ByVal vNewValue As Variant)
    m_vntBookmark = vNewValue
End Property

Public Property Get ErrorText() As Variant
    ErrorText = m_strErrorText
End Property

Public Property Let ErrorText(ByVal vNewValue As Variant)
    m_strErrorText = vNewValue
End Property

Public Property Let RowSQL(ByVal sNewValue As String)
    m_strRowSQL = sNewValue
End Property
Public Property Get Error() As Variant
    Error = m_blnError
End Property

Public Property Let Error(ByVal vNewValue As Variant)
    m_blnError = vNewValue
End Property

Public Property Get State() As Variant
    State = m_intState
End Property

Public Property Get skey() As Long
    skey = m_lngSkey
End Property

Public Property Let State(ByVal vNewValue As Variant)
    m_intState = vNewValue
    If vNewValue = STATE_NONE Then
        m_blnError = False
        m_strErrorText = ""
    End If
End Property

Public Property Let skey(ByVal lNewSkey As Long)
    m_lngSkey = lNewSkey
End Property

Public Property Let SubKey(ByVal strSubKey As String)
    m_strSubKey = strSubKey
End Property
