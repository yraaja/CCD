VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Description As String
Dim m_Col2_Description As String
Dim m_Col3_Description As String
Dim m_SortOrder As Long
Dim m_ListID As Integer
Dim m_ItemData As Long

Property Let Description(strDesc As String)
    m_Description = strDesc
End Property
Property Get Description() As String
    Description = m_Description
End Property

Property Let SortOrder(lngOrder As Long)
    m_SortOrder = lngOrder
End Property
Property Get SortOrder() As Long
SortOrder = m_SortOrder
End Property
Property Let ListID(lngListID As Long)
    m_ListID = lngListID
End Property
Property Get ListID() As Long
ListID = m_ListID
End Property
Property Let ItemData(lngItemData As Long)
    m_ItemData = lngItemData
End Property



Property Get ItemData() As Long
    ItemData = m_ItemData
End Property
Private Sub Class_Initialize()
m_ItemData = 0
m_SortOrder = 0
End Sub


