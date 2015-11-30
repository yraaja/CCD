VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUCostHistMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Unit Cost History grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "UnitCostHistory" ' Identity of grid
' Names of table used by the grid
Const PUBUNITCOST_TABLE = "Published_unit_cost"

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

    ' Add the columns to the grid
    m_objGridMap.AddCol "Start Date", "start_date", PUBUNITCOST_TABLE, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Term Date", "term_date", vbNullString, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Op Code", "op_code", PUBUNITCOST_TABLE, 1000, True, True, False
    m_objGridMap.AddCol "Country", "country_code", PUBUNITCOST_TABLE, 1100, True, False, False
    m_objGridMap.AddCol "Region", "region_code", PUBUNITCOST_TABLE, 1100, True, False, False
    ' For M records
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Crew Qty", "crew_qty", PUBUNITCOST_TABLE, 900, True, True, False
    m_objGridMap.AddCol "Crew Id", "crew_id", PUBUNITCOST_TABLE, 1000, True, True, False
    m_objGridMap.AddCol "Daily Out", "daily_output", PUBUNITCOST_TABLE, 1000, True, True, False
    m_objGridMap.AddCol "Unit", "unit", PUBUNITCOST_TABLE, 1000, True, True, False
    m_objGridMap.AddCol "Metric Daily Out", "metric_daily_output", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Metric Unit", "metric_unit", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Labor Hrs", "labor_hour", PUBUNITCOST_TABLE, 1000, True, True, False
    m_objGridMap.AddCol "Met Labor Hrs", "metric_labor_hour", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Mat Cost", "mat_cost", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Labor Cost", "labor_cost", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Equip Cost", "equip_cost", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Mat Cost Op", "mat_cost_op", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Labor Cost Op", "labor_cost_op", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Equip Cost Op", "equip_cost_op", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Total Cost", "total_cost", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Total Cost Op", "total_cost_op", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Met Mat Cost", "metric_mat_cost", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Met Labor Cost", "metric_labor_cost", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Met Equip Cost", "metric_equip_cost", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Met Mat Cost Op", "metric_mat_cost_op", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Met Labor Cost Op", "metric_labor_cost_op", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Met Equip Cost Op", "metric_equip_cost_op", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Met Total Cost", "metric_total_cost", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Met Total Cost Op", "metric_total_cost_op", PUBUNITCOST_TABLE, 1000, True, False, False
    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "last_update_person", vbNullString, 1500, True
    ' For Exception records
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Percent", "pct_ind", vbNullString, 800
    m_objGridMap.AddCol "Unit", "unit_x", vbNullString, 1000, True, True, False
    m_objGridMap.AddCol "Metric Unit", "metric_unit_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Mat Cost", "mat_cost_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Labor Cost", "labor_cost_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Equip Cost", "equip_cost_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Total Cost", "total_cost_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Total Cost Op", "total_cost_op_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Met Mat Cost", "metric_mat_cost_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Met Labor Cost", "metric_labor_cost_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Met Equip Cost", "metric_equip_cost_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Met Total Cost", "metric_total_cost_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Met Total Cost Op", "metric_total_cost_op_x", vbNullString, 1000, True, False, False
    m_objGridMap.AddCol "Update Date", "last_update_date_x", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "last_update_person_x", vbNullString, 1500, True
    
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



