VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObjAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_strSysID As String
Dim m_strAttrFieldName As String
Dim m_blnLocked As Boolean
Dim m_blnDirty As Boolean     'Indicate if data has changed from associated record
Property Let SysID(strSysID As String)
m_strSysID = strSysID
End Property


Property Let AttrFieldName(strAttrFieldName As String)
m_strAttrFieldName = strAttrFieldName
End Property
Property Let Locked(blnLocked As Boolean)
m_blnLocked = blnLocked
End Property
Property Get Locked() As Boolean
Locked = m_blnLocked
End Property
Property Get SysID() As String
SysID = m_strSysID
End Property

Property Get AttrFieldName() As String
AttrFieldName = m_strAttrFieldName
End Property
Property Let Dirty(blnDirty As Boolean)
m_blnDirty = blnDirty
End Property
Property Get Dirty() As Boolean
Dirty = m_blnDirty
End Property

