VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUnitCostOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class handles Unit Cost output to CostWorks
Option Explicit

Private m_rec As ADODB.RecordSet
Private m_clsCWOutput As New CCWOutput

Public Property Let RecordSet(ByVal vNewValue As Variant)
    Set m_rec = vNewValue
End Property

Private Sub Class_Initialize()
    m_clsCWOutput.CreateField "unit_cost_id", 10
    m_clsCWOutput.CreateField "type_code", 1
    m_clsCWOutput.CreateField "format_code", 2
    m_clsCWOutput.CreateField "format_type", 5
    m_clsCWOutput.CreateField "format_characters", 2, True
    m_clsCWOutput.CreateField "indent_code", 2
    m_clsCWOutput.CreateField "book_desc", 75
    m_clsCWOutput.CreateField "metric_book_desc", 75
    m_clsCWOutput.CreateField "index_code", 2
    m_clsCWOutput.CreateField "index_desc", 75
    m_clsCWOutput.CreateField "tech_desc", 75
    m_clsCWOutput.CreateField "metric_tech_desc", 75
    m_clsCWOutput.CreateField "bk_bccd", 2
    m_clsCWOutput.CreateField "bk_mech", 2
    m_clsCWOutput.CreateField "bk_elec", 2
    m_clsCWOutput.CreateField "bk_rr", 2
    m_clsCWOutput.CreateField "bk_site", 2
    m_clsCWOutput.CreateField "bk_int", 2
    m_clsCWOutput.CreateField "bk_conc", 2
    m_clsCWOutput.CreateField "bk_open", 2
    m_clsCWOutput.CreateField "bk_heav", 2
    m_clsCWOutput.CreateField "bk_faci", 2
    m_clsCWOutput.CreateField "bk_plum", 2
    m_clsCWOutput.CreateField "bk_mbccd", 2
    m_clsCWOutput.CreateField "bk_mheav", 2
    m_clsCWOutput.CreateField "bk_ass_missing", 2
    m_clsCWOutput.CreateField "imperial_ind", 2
    m_clsCWOutput.CreateField "metric_ind", 2
    m_clsCWOutput.CreateField "crew_qty", 14
    m_clsCWOutput.CreateField "crew_id", 5
    m_clsCWOutput.CreateField "unit", 8
    m_clsCWOutput.CreateField "daily_output", 14
    m_clsCWOutput.CreateField "std_labor_hour", 14
    m_clsCWOutput.CreateField "std_mat_cost", 20
    m_clsCWOutput.CreateField "std_labor_cost", 20
    m_clsCWOutput.CreateField "std_equip_cost", 20
    m_clsCWOutput.CreateField "std_total_cost", 20
    m_clsCWOutput.CreateField "std_mat_cost_op", 14
    m_clsCWOutput.CreateField "std_labor_cost_op", 14
    m_clsCWOutput.CreateField "std_equip_cost_op", 14
    m_clsCWOutput.CreateField "std_total_cost_op", 20
    m_clsCWOutput.CreateField "rr_labor_hour", 14
    m_clsCWOutput.CreateField "rr_mat_cost", 20
    m_clsCWOutput.CreateField "rr_labor_cost", 20
    m_clsCWOutput.CreateField "rr_equip_cost", 20
    m_clsCWOutput.CreateField "rr_total_cost", 20
    m_clsCWOutput.CreateField "rr_mat_cost_op", 14
    m_clsCWOutput.CreateField "rr_labor_cost_op", 14
    m_clsCWOutput.CreateField "rr_equip_cost_op", 14
    m_clsCWOutput.CreateField "rr_total_cost_op", 20
    m_clsCWOutput.CreateField "opn_labor_hour", 14
    m_clsCWOutput.CreateField "opn_mat_cost", 20
    m_clsCWOutput.CreateField "opn_labor_cost", 20
    m_clsCWOutput.CreateField "opn_equip_cost", 20
    m_clsCWOutput.CreateField "opn_total_cost", 20
    m_clsCWOutput.CreateField "opn_mat_cost_op", 14
    m_clsCWOutput.CreateField "opn_labor_cost_op", 14
    m_clsCWOutput.CreateField "opn_equip_cost_op", 14
    m_clsCWOutput.CreateField "opn_total_cost_op", 20
    m_clsCWOutput.CreateField "metric_unit", 8
    m_clsCWOutput.CreateField "metric_daily_output", 14
    m_clsCWOutput.CreateField "metric_labor_hour", 14
    m_clsCWOutput.CreateField "metric_mat_cost", 20
    m_clsCWOutput.CreateField "metric_labor_cost", 20
    m_clsCWOutput.CreateField "metric_equip_cost", 20
    m_clsCWOutput.CreateField "metric_total_cost", 20
    m_clsCWOutput.CreateField "metric_mat_cost_op", 14
    m_clsCWOutput.CreateField "metric_labor_cost_op", 14
    m_clsCWOutput.CreateField "metric_equip_cost_op", 14
    m_clsCWOutput.CreateField "metric_total_cost_op", 20
    m_clsCWOutput.CreateField "percent_flag", 1
    m_clsCWOutput.CreateField "graphic_ref_num", 12
    m_clsCWOutput.CreateField "circle_ref_num", 12
    m_clsCWOutput.CreateField "circle_ref_num1", 12
    m_clsCWOutput.CreateField "circle_ref_num2", 12
    m_clsCWOutput.CreateField "circle_ref_num3", 12
    m_clsCWOutput.CreateField "circle_ref_num4", 12
    m_clsCWOutput.CreateField "circle_ref_num5", 12
    m_clsCWOutput.CreateField "circle_ref_num6", 12
    m_clsCWOutput.CreateField "circle_ref_num7", 12
    m_clsCWOutput.CreateField "circle_ref_num8", 12
End Sub

Public Sub Add(intUnitCostSKey As Long)
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rec As ADODB.RecordSet
    
    strSelect = "select * from published_unit_cost_costworks where unit_cost_skey = " + CStr(intUnitCostSKey)
        
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
    If blnReturn = False Then
        MsgBox "An error occurred while searching."
        Exit Sub
    End If

    m_clsCWOutput.AddRecord rec
End Sub

Private Function Special(strName As String) As String

End Function

Public Sub Done()
    m_clsCWOutput.Flush
End Sub
