VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_colAttributes As New Collection
Dim m_strObjID As String
Dim m_strRowMeasSys As String
Dim m_intAttrs As Integer
Dim m_strObjKey As String
Public Property Get Object_Attribute(strAttributeCaption As String) As CObjAttribute
    On Error Resume Next
    Set Object_Attribute = m_colAttributes.Item(strAttributeCaption)
End Property
Public Property Get RowMeasSys() As String
    RowMeasSys = m_strRowMeasSys
End Property
Public Sub Add_Object_Attribute(strAttrFieldName As String, strSysID As Variant)
    Dim objAttribute As CObjAttribute

    Set objAttribute = New CObjAttribute
    With objAttribute
        .AttrFieldName = strAttrFieldName
        If IsNull(strSysID) Then
            .SysID = "A"
        Else
            .SysID = strSysID
        End If
        If m_strRowMeasSys = "M" Then ' Metric Row
            If .SysID = "A" Then 'All for system code
                .Locked = True
            End If
        End If
        .Dirty = False  'Initialize as unchanged - data matches recordset
    End With
    m_colAttributes.Add objAttribute, strAttrFieldName
    m_intAttrs = m_intAttrs + 1
End Sub

Public Property Get Attribute_Count() As Integer
Attribute_Count = m_intAttrs
End Property

Public Property Get Object_ID() As String
Object_ID = m_strObjID
End Property
Public Property Get Object_Key() As String
Object_Key = m_strObjKey
End Property
Public Property Let Object_ID(ByVal strObjID As String)
    m_strObjID = strObjID
End Property
Public Property Let Object_Key(ByVal strObjKey As String)
    m_strObjKey = strObjKey
End Property
Public Property Let RowMeasSys(ByVal strRowMeasSys As String)
    m_strRowMeasSys = strRowMeasSys
End Property
