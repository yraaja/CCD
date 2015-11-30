VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMatHistoryMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Material grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "MaterialHistory" ' Identity of grid
' Names of table used by the grid
Const MATPRICE_TABLE = "Material_price_audit"
Const MATMAN_TABLE = "Material_manufacturer"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' recordset that populates grid
Private m_objDAL As New CRSMDataAccess ' Data access layer to interface with DB
Private m_blnPreferences As Boolean ' Set to True if on Grid Preferences screen

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub
Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub

Public Property Let RecordSet(RecSet As ADODB.RecordSet)
    ' Pass-through to base class
    m_objGridMap.RecordSet = RecSet
    Set m_rec = RecSet
End Property

'*** APEX Migration Utility Code Change ***
'Public Sub SetGrid(grd As TrueOleDBGrid60.TDBGrid)
'*** APEX Migration Utility Code Change ***
'Public Sub SetGrid(grd As TrueOleDBGrid70.TDBGrid)
Public Sub SetGrid(grd As TrueOleDBGrid80.TDBGrid)
    ' Pass-through to base class
    m_objGridMap.SetGrid grd
    Set TDBGrid = grd
End Sub

Public Property Let Preferences(bln As Boolean)
    ' Pass-through to base class
    m_blnPreferences = bln
End Property

Public Sub InitGrid()
    m_objGridMap.GRIDTYPE = GRIDTYPE
'    m_objGridMap.SetFrozenCols (0)
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "Price", "list_price", MATPRICE_TABLE, 1000, True, True, False, FORMAT_PRICE
    m_objGridMap.AddCol "TRACES Price", "traces_list_price", MATPRICE_TABLE, 1200, True, True, False, FORMAT_PRICE
    m_objGridMap.AddCol "Final Price", "final_price", MATPRICE_TABLE, 1200, True, True, False, FORMAT_PRICE
    m_objGridMap.AddCol "Start Date", "start_date", MATPRICE_TABLE, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Contact ID", "contact_id", MATPRICE_TABLE, 1500, True
    m_objGridMap.AddCol "Manuf. ID", "manufacturer_id", MATPRICE_TABLE, 1500, True
    m_objGridMap.AddCol "Manuf. Desc", "manufacturer_desc", MATMAN_TABLE, 2000, True
    m_objGridMap.AddCol "Use", "use_ind", MATPRICE_TABLE, 1000, True
    m_objGridMap.AddCol "Term Date", "term_date", vbNullString, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Update Date", "matprice_last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "matprice_last_update_person", vbNullString, 1500, True
    m_objGridMap.AddCol "mat_skey", "mat_skey", MATPRICE_TABLE, 800, True, False
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
End Sub

Public Function Update() As Boolean
    ' Screen is read-only
End Function

Public Property Get GRIDTYPE() As String
    GRIDTYPE = m_GRIDTYPE
End Property


Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub

