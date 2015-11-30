VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAssemblyMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Unit Cost Usage grid
Option Explicit

''' <modulename> CAssemblyMap.cls</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Helper class module for frmUnitCostGrid.frm
''' </summary>
''' <seealso>frmAssemblyGrid</seealso>
''' <datastruct> TDBGRID</datastruct>
''' <datastruct> m_objGridMap</datastruct>
''' <datastruct> m_rec</datastruct>
'''
''' <storedprocedurename> sp_update_assembly_driver</storedprocedurename>
''' <storedprocedurename> sp_delete_assembly</storedprocedurename>
'''
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''    exec sp_update_assembly_driver @type_code='M', @assembly_skey= 20776, @assembly_id='E10101100100', @alt_assembly_id='1112001100', @rev_uni2_L3='     ',
'''    @rev_uni2_L5='    ', @rev_uni2_L6='    ', @book_desc='Bank equipment, drive up window, drawer & mike, no glazing, economy', @metric_book_desc='',
'''    @tech_desc='Architectural equipment, bank equipment drive up window, drawer & mike, no glazing, economy', @metric_tech_desc='Architectural equipment, bank equipment drive up window, drawer and mike, no glazing, economy',
'''    @coml_ind= 1, @resi_ind= 0, @labor_equip_ind= 0, @comment='', @unit='Day', @metric_unit='Ea.', @ad_change_ind= 1, @std_mat_cost= 4950, @std_inst_cost= 660,
'''    @std_equip_cost= 0, @std_labor_cost= 660, @std_total_cost= 5610, @std_mat_cost_op= 5450, @std_inst_cost_op= 1200, @std_equip_cost_op= 0, @std_labor_cost_op= 1200,
'''    @std_total_cost_op= 6650, @std_labor_hour= 16, @std_change_ind= 0, @opn_mat_cost= 4950, @opn_inst_cost= 495, @opn_equip_cost= 0, @opn_labor_cost= 495, @opn_total_cost= 5445,
'''    @opn_mat_cst_op= 5450, @opn_inst_cost_op= 970, @opn_equip_cost_op= 0, @opn_labor_cost_op= 970, @opn_total_cost_op= 6420, @opn_labor_hour= 16, @opn_change_ind= 0,
'''    @rr_mat_cost= 4950, @rr_inst_cost= 660, @rr_equip_cost= 0, @rr_labor_cost= 660, @rr_total_cost= 5610,
'''    @rr_mat_cost_op= 5450, @rr_inst_cost_op= 1250, @rr_equip_cost_op= 0, @rr_labor_cost_op= 1250, @rr_total_cost_op= 6700,
'''    @rr_labor_hour= 16, @rr_change_ind= 0, @metric_mat_cost= 4950, @metric_inst_cost= 660, @metric_equip_cost= 0, @metric_labor_cost= 660,
'''    @metric_total_cost= 5610, @metric_mat_cost_op= 5450, @metric_inst_cost_op= 1200, @metric_equip_cost_op= 0, @metric_labor_cost_op= 1200,
'''    @metric_total_cost_op= 6650, @metric_labor_hour= 16, @pct_ind= 0, @ad_last_update_id= 3, @std_last_update_id= 1, @opn_last_update_id= 1,
'''    @rr_last_update_id= 1, @last_update_person='Hancockrl',  @update_unitcost_usage_ind=0, @cost_change_ind=0</code>
''' <code>
''' </code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  CGridMap.cls
'''2.  CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''</dependson>


'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "Assembly" ' Identity of grid
' Names of some table this grid uses
Const ASSEMBLYDETAIL_TABLE = "assembly_detail"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private bln_AssemblyID As Boolean
Private bln_UnitCostID As Boolean
Private bln_UseInd As Boolean
Public Function Delete() As Boolean
Dim varSelectedBookmarks() As Variant
Dim varCurrentSelectedBookmark As Variant
Dim varButton
Dim blnResult As Boolean
Dim i As Integer
Dim i1 As Integer

'On Error Resume Next

If TDBGrid.SelBookmarks.Count > 0 Then
        ReDim varSelectedBookmarks(0 To TDBGrid.SelBookmarks.Count)
        i = 0
        For Each varCurrentSelectedBookmark In TDBGrid.SelBookmarks
            varSelectedBookmarks(i) = varCurrentSelectedBookmark
            i = i + 1
        Next varCurrentSelectedBookmark
        For i = 0 To UBound(varSelectedBookmarks) - 1
            TDBGrid.Bookmark = varSelectedBookmarks(i)
            If check_error(TDBGrid.Bookmark, TDBGrid.Columns("Assembly Skey")) = True Then
                For i1 = 0 To TDBGrid.SelBookmarks.Count
                    If TDBGrid.SelBookmarks(i1) = varSelectedBookmarks(i) Then
                        TDBGrid.SelBookmarks.Remove i1
                        Exit For
                    End If
                Next i1
            End If
        Next i
        If TDBGrid.SelBookmarks.Count > 0 Then
            Delete = m_objGridMap.Delete
        Else
            Delete = False
        End If
    Else
        If check_error(TDBGrid.Bookmark, TDBGrid.Columns("Assembly Skey")) = True Then
            Delete = False
        Else
            Delete = m_objGridMap.Delete
        End If
    End If
End Function

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub
Public Function CloneRowRecordset() As ADODB.RecordSet
    Set CloneRowRecordset = m_objGridMap.CloneRowRecordset
End Function

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

    'TDBGrid.AllowAddNew = True  'rlh ccd 8.4  (ksr asked to remove)
    
    m_objGridMap.GRIDTYPE = m_GRIDTYPE

    ' Add the columns to the grid
    m_objGridMap.AddCol "Type", "type_code", ALL_TABLES, 500, True, , True
    m_objGridMap.AddCol "Assembly Skey", "assembly_skey", ALL_TABLES, 1200, True, False
    m_objGridMap.AddCol "Assembly ID", "assembly_id", ALL_TABLES, 1200, True
    m_objGridMap.AddCol "Alt Assembly ID", "alt_assembly_id", ALL_TABLES, 1200, False, True, True
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Uni2 Lvl 3", "rev_uni2_L3", ALL_TABLES, 800
    m_objGridMap.AddCol "Uni2 Lvl 5", "rev_uni2_L5", ALL_TABLES, 600
    m_objGridMap.AddCol "Uni2 Lvl 6", "rev_uni2_L6", ALL_TABLES, 600
    m_objGridMap.AddCol "Book Desc", "book_desc", ALL_TABLES, 2000
    m_objGridMap.AddCol "Metric Book Desc", "metric_book_desc", ALL_TABLES, 2000, , False
    m_objGridMap.AddCol "Tech Desc", "tech_desc", ALL_TABLES, 2000, , False
    m_objGridMap.AddCol "Metric Tech Desc", "metric_tech_desc", ALL_TABLES, 2000, , False
    m_objGridMap.AddCol "Comm'l Ind", "coml_ind", ALL_TABLES, 1000
    m_objGridMap.AddCol "Resi Ind", "resi_ind", ALL_TABLES, 1000
    m_objGridMap.AddCol "Labor Equip Ind", "labor_equip_ind", ALL_TABLES, 1800, False, True
    m_objGridMap.AddCol "Comment", "comment", ALL_TABLES, 2000, , False
    m_objGridMap.AddCol "Unit", "unit", ALL_TABLES, 1000, , , True, , True
    m_objGridMap.AddCol "Metric Unit", "metric_unit", ALL_TABLES, 1000, , False, , , True
    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1700, True, False, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "last_update_person", vbNullString, 1500, True, False, False
    m_objGridMap.AddCol "AD Change Ind", "ad_change_ind", ALL_TABLES, 650, True, False, False
    
    m_objGridMap.AddCol "Std Mat", "std_mat_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Inst", "std_inst_cost", ALL_TABLES, 900, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Std Equip", "std_equip_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Labor", "std_labor_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Total", "std_total_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Mat OP", "std_mat_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Inst OP", "std_inst_cost_op", ALL_TABLES, 1200, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Std Equip OP", "std_equip_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Labor OP", "std_labor_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Total OP", "std_total_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Hours", "std_labor_hour", ALL_TABLES, 900, , , , , True
    m_objGridMap.AddCol "Std Change Ind", "std_change_ind", ALL_TABLES, 900, True, False, False
    
    m_objGridMap.AddCol "Opn Mat", "opn_mat_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Inst", "opn_inst_cost", ALL_TABLES, 900, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Opn Equip", "opn_equip_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Labor", "opn_labor_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Total", "opn_total_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Mat OP", "opn_mat_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Inst OP", "opn_inst_cost_op", ALL_TABLES, 1200, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Opn Equip OP", "opn_equip_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Labor OP", "opn_labor_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Total OP", "opn_total_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Hours", "opn_labor_hour", ALL_TABLES, 900, , , , , True
    m_objGridMap.AddCol "Opn Change Ind", "opn_change_ind", ALL_TABLES, 200, True, False, False

    m_objGridMap.AddCol "R&R Mat", "rr_mat_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Inst", "rr_inst_cost", ALL_TABLES, 900, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "R&R Equip", "rr_equip_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Labor", "rr_labor_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Total", "rr_total_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Mat OP", "rr_mat_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Inst OP", "rr_inst_cost_op", ALL_TABLES, 1200, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "R&R Equip OP", "rr_equip_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Labor OP", "rr_labor_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Total OP", "rr_total_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Hours", "rr_labor_hour", ALL_TABLES, 900, , , , , True
    m_objGridMap.AddCol "R&R Change Ind", "rr_change_ind", ALL_TABLES, 200, True, False, False

    m_objGridMap.AddCol "Met Mat", "metric_mat_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Inst", "metric_inst_cost", ALL_TABLES, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Equip", "metric_equip_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Labor", "metric_labor_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Total", "metric_total_cost", ALL_TABLES, 900, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Mat OP", "metric_mat_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Inst OP", "metric_inst_cost_op", ALL_TABLES, 1200, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Met Equip OP", "metric_equip_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Labor OP", "metric_labor_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Total OP", "metric_total_cost_op", ALL_TABLES, 1200, , , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Hours", "metric_labor_hour", ALL_TABLES, 900, , , , , True
    m_objGridMap.AddCol "Percent", "pct_ind", ALL_TABLES, 900, , , , , True
    m_objGridMap.AddCol "AssDtl Last Update ID", "ad_last_update_id", ALL_TABLES, 200, True, False
    m_objGridMap.AddCol "Std Last Update ID", "std_last_update_id", ALL_TABLES, 200, True, False
    m_objGridMap.AddCol "Opn Last Update ID", "opn_last_update_id", ALL_TABLES, 200, True, False
    m_objGridMap.AddCol "RR Last Update ID", "rr_last_update_id", ALL_TABLES, 200, True, False
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
    If Not m_blnPreferences Then
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid60.ValueItem
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid70.ValueItem
        Dim Item As New TrueOleDBGrid80.ValueItem
        Dim bRet As Boolean
        Dim rec As ADODB.RecordSet

        ' Unit drop downs
        bRet = g_objDAL.GetRecordset(vbNullString, "select unit from unit_of_measure order by unit", rec)
        If bRet Then
            While Not rec.EOF
                Item.Value = rec.Fields(0)
                TDBGrid.Columns("Unit").ValueItems.Add Item
                TDBGrid.Columns("Metric Unit").ValueItems.Add Item
                rec.MoveNext
            Wend
            TDBGrid.Columns("Unit").ValueItems.Presentation = dbgComboBox
            TDBGrid.Columns("Unit").ValueItems.Validate = True
            TDBGrid.Columns("Metric Unit").ValueItems.Presentation = dbgComboBox
            TDBGrid.Columns("Metric Unit").ValueItems.Validate = True
        End If
        rec.Close

        ' Type Code drop down
        Item.Value = "M"
        TDBGrid.Columns("Type").ValueItems.Add Item
        Item.Value = "E"
        TDBGrid.Columns("Type").ValueItems.Add Item
    
        TDBGrid.Columns("Type").ValueItems.Presentation = dbgComboBox
        TDBGrid.Columns("Type").ValueItems.Validate = True
        TDBGrid.Columns("Type").AutoDropDown = True

        'Percent Flag Checkbox
        Item.Value = "1"
        Item.DisplayValue = True
        TDBGrid.Columns("Percent").ValueItems.Add Item
        Item.Value = "0"
        Item.DisplayValue = False
        TDBGrid.Columns("Percent").ValueItems.Add Item
        TDBGrid.Columns("Percent").ValueItems.Translate = True
        TDBGrid.Columns("Percent").ValueItems.Presentation = dbgCheckBox
        TDBGrid.Columns("Percent").ValueItems.Validate = True
      
        'Commercial Flag Checkbox
        Item.Value = "1"
        Item.DisplayValue = True
        TDBGrid.Columns("coml_ind").ValueItems.Add Item
        Item.Value = "0"
        Item.DisplayValue = False
        TDBGrid.Columns("coml_ind").ValueItems.Add Item
        TDBGrid.Columns("coml_ind").ValueItems.Translate = True
        TDBGrid.Columns("coml_ind").ValueItems.Presentation = dbgCheckBox
        TDBGrid.Columns("coml_ind").ValueItems.Validate = True
        
        'Residential Flag Checkbox
        Item.Value = "1"
        Item.DisplayValue = True
        TDBGrid.Columns("resi_ind").ValueItems.Add Item
        Item.Value = "0"
        Item.DisplayValue = False
        TDBGrid.Columns("resi_ind").ValueItems.Add Item
        TDBGrid.Columns("resi_ind").ValueItems.Translate = True
        TDBGrid.Columns("resi_ind").ValueItems.Presentation = dbgCheckBox
        TDBGrid.Columns("resi_ind").ValueItems.Validate = True
        
        'Labor Equipment Flag Checkbox
        Item.Value = "1"
        Item.DisplayValue = True
        TDBGrid.Columns("labor_equip_ind").ValueItems.Add Item
        Item.Value = "0"
        Item.DisplayValue = False
        TDBGrid.Columns("labor_equip_ind").ValueItems.Add Item
        TDBGrid.Columns("labor_equip_ind").ValueItems.Translate = True
        TDBGrid.Columns("labor_equip_ind").ValueItems.Presentation = dbgCheckBox
        TDBGrid.Columns("labor_equip_ind").ValueItems.Validate = True
    End If
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    
    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0
    
    m_rec.MoveFirst
    While Not m_rec.EOF
        ' Skip the record if it didn't change
        Dim clsRowInfo As CRowInfo
        Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
        If Not clsRowInfo Is Nothing Then
        
            If Not clsRowInfo.State = STATE_NONE Then
            ' If record was added, or if both assembly_id and unit_cost_id were changed
            If clsRowInfo.State = STATE_NEW Or clsRowInfo.State = STATE_MODIFIED Then
                blnReturn = False
                strError = ""
                UpdateCostChangeFlags
                
                strUpdate = "exec sp_update_assembly_driver "
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                strUpdate = strUpdate + "@last_update_person='" + strUserName + "', "
                strUpdate = strUpdate + " @update_unitcost_usage_ind=0,"
                strUpdate = strUpdate + " @cost_change_ind=0"
                blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                ' Reset on success
                If blnReturn = True Then
                    intSuccess = intSuccess + 1
                    clsRowInfo.State = STATE_NONE
                    m_rec.Fields("ad_last_update_id").Value = m_rec.Fields("ad_last_update_id").Value + 1
                    m_rec.Fields("std_last_update_id").Value = m_rec.Fields("std_last_update_id").Value + 1
                    m_rec.Fields("opn_last_update_id").Value = m_rec.Fields("opn_last_update_id").Value + 1
                    m_rec.Fields("rr_last_update_id").Value = m_rec.Fields("rr_last_update_id").Value + 1
                    m_rec.Fields("last_update_date").Value = Now
                    m_rec.Fields("last_update_person").Value = strUserName
                   
                   Reset_Orig_Values m_rec
                   
                Else
                    intErrors = intErrors + 1
                    clsRowInfo.Error = True
                    clsRowInfo.ErrorText = strError
                End If
                
            End If
            End If
        End If
        m_rec.MoveNext
        
    Wend

    
    m_rec.UpdateBatch
    
    Dim strMsg As String
    strMsg = ""
    If intSuccess > 0 Then
        strMsg = str(intSuccess) + " rows updated successfully." + Chr(10)
    End If
    If intErrors > 0 Then
        strMsg = strMsg + str(intErrors) + " errors occurred."
        ' Return value will be False
        Update = False
    End If
    If Len(strMsg) > 0 Then
        MsgBox strMsg
    End If
   
    TDBGrid.ReBind
    Exit Function
    
CleanUp:
    MsgBox "Error in Update"
End Function

Private Sub UpdateCostChangeFlags()
Dim CurFields As ADODB.Fields
Dim CurrentField As ADODB.Field
Set CurFields = m_rec.Fields
'If any field has changed for an exception type row, flag which record was modified: STD, OPN, or RR

m_rec.Fields("ad_change_ind") = 0

For Each CurrentField In CurFields
    If CurrentField.Value <> CurrentField.OriginalValue Then
        Select Case CurrentField.Name
            'Assembly Detail fields
            Case "book_desc", _
                "metric_book_desc", _
                "tech_desc", _
                "metric_tech_desc", _
                "comment", _
                "unit", _
                "metric_unit", _
                "type_code", _
                "coml_ind", _
                "resi_ind", _
                "labor_equip_ind", _
                "alt_assembly_id", _
                "rev_uni2_L3", _
                "rev_uni2_L5", _
                "rev_uni2_L6", _
                "assembly_id"

                
                m_rec.Fields("ad_change_ind") = 1
                Exit For
        End Select
    End If
Next CurrentField

If m_rec.Fields("type_code") = "E" Then
    m_rec.Fields("std_change_ind") = 0
    m_rec.Fields("opn_change_ind") = 0
    m_rec.Fields("rr_change_ind") = 0

    For Each CurrentField In CurFields
        If CurrentField.Value <> CurrentField.OriginalValue Then
            Select Case CurrentField.Name
                Case "pct_ind"
                    m_rec.Fields("std_change_ind") = 1
                    m_rec.Fields("opn_change_ind") = 1
                    m_rec.Fields("rr_change_ind") = 1
                    Exit For
                'STD Exception fields
                Case "std_mat_cost", _
                        "std_inst_cost", _
                        "std_equip_cost", _
                        "std_labor_cost", _
                        "std_total_cost", _
                        "std_mat_cost_op", _
                        "std_inst_cost_op", _
                        "std_equip_cost_op", _
                        "std_labor_cost_op", _
                        "std_total_cost_op", _
                        "metric_inst_cost", _
                        "metric_equip_cost", _
                        "metric_labor_cost", _
                        "metric_total_cost", _
                        "metric_mat_cost_op", _
                        "metric_inst_cost_op", _
                        "metric_equip_cost_op", _
                        "metric_labor_cost_op", _
                        "metric_total_cost_op"
                        m_rec.Fields("std_change_ind") = 1
                    Exit For
                'OPN Exception fields
                Case "opn_mat_cost", _
                        "opn_inst_cost", _
                        "opn_equip_cost", _
                        "opn_labor_cost", _
                        "opn_total_cost", _
                        "opn_mat_cost_op", _
                        "opn_inst_cost_op", _
                        "opn_equip_cost_op", _
                        "opn_labor_cost_op", _
                        "opn_total_cost_op"
                    m_rec.Fields("opn_change_ind") = 1
                    Exit For
                'R & R Exception fields
                Case "rr_mat_cost", _
                        "rr_inst_cost", _
                        "rr_equip_cost", _
                        "rr_labor_cost", _
                        "rr_total_cost", _
                        "rr_mat_cost_op", _
                        "rr_inst_cost_op", _
                        "rr_equip_cost_op", _
                        "rr_labor_cost_op", _
                        "rr_total_cost_op"
                    m_rec.Fields("rr_change_ind") = 1
                    Exit For
            End Select
        End If
    Next CurrentField
End If

End Sub
Private Sub Class_Initialize()
    m_objGridMap.m_colCloneCols.Add 1, "comment"
    m_objGridMap.m_colCloneCols.Add 1, "metric_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "coml_ind"
    m_objGridMap.m_colCloneCols.Add 1, "resi_ind"
    m_objGridMap.m_colCloneCols.Add 1, "labor_equip_ind"
    m_objGridMap.m_colCloneCols.Add 1, "unit"
    m_objGridMap.m_colCloneCols.Add 1, "metric_unit"
    m_objGridMap.m_colCloneCols.Add 1, "type_code"
    m_objGridMap.m_colCloneCols.Add 1, "assembly_skey"
    m_objGridMap.m_colCloneCols.Add 1, "assembly_id"
    m_objGridMap.m_colCloneCols.Add 1, "alt_assembly_id"
    m_objGridMap.m_colCloneCols.Add 1, "book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "pac_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "std_mat_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_inst_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_equip_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_labor_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_total_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_mat_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "std_inst_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "std_equip_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "std_labor_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "std_total_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_mat_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_inst_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_equip_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_total_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_mat_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_inst_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_equip_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_total_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_mat_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_inst_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_equip_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_labor_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_total_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_mat_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_inst_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_equip_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_labor_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_total_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_mat_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_inst_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_equip_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_labor_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_total_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_mat_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_inst_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_equip_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_labor_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_total_cost_op"
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
Dim strSelect As String
Dim rec As ADODB.RecordSet
Dim blnReturn  As Boolean
Dim i As Integer

Select Case TDBGrid.Columns(ColIndex).Caption
        Case "Std Equip", "Std Labor"
            TDBGrid.Columns("Std Inst") = Val(TDBGrid.Columns("Std Equip")) + Val(TDBGrid.Columns("Std Labor"))
        Case "Std Equip OP", "Std Labor OP"
            TDBGrid.Columns("Std Inst OP") = Val(TDBGrid.Columns("Std Equip OP")) + Val(TDBGrid.Columns("Std Labor OP"))
        Case "Opn Equip", "Opn Labor"
            TDBGrid.Columns("Opn Inst") = Val(TDBGrid.Columns("Opn Equip")) + Val(TDBGrid.Columns("Opn Labor"))
        Case "Opn Equip OP", "Opn Labor OP"
            TDBGrid.Columns("Opn Inst OP") = Val(TDBGrid.Columns("Opn Equip OP")) + Val(TDBGrid.Columns("Opn Labor OP"))
        Case "R&R Equip", "R&R Labor"
            TDBGrid.Columns("R&R Inst") = Val(TDBGrid.Columns("R&R Equip")) + Val(TDBGrid.Columns("R&R Labor"))
        Case "R&R Equip OP", "R&R Labor OP"
            TDBGrid.Columns("R&R Inst OP") = Val(TDBGrid.Columns("R&R Equip OP")) + Val(TDBGrid.Columns("R&R Labor OP"))
        Case "Met Equip", "Met Labor"
            TDBGrid.Columns("Met Inst") = Val(TDBGrid.Columns("Met Equip")) + Val(TDBGrid.Columns("Met Labor"))
        Case "Met Equip OP", "Met Labor OP"
            TDBGrid.Columns("Met Inst OP") = Val(TDBGrid.Columns("Met Equip OP")) + Val(TDBGrid.Columns("Met Labor OP"))
        Case "Type"
            TDBGrid.PostMsg 2
        Case "Assembly ID"
            If InStr(1, TDBGrid.Columns(ColIndex).Value, "X") > 0 Then
                TDBGrid.Columns("Comm'l Ind").Value = 1
                TDBGrid.Columns("Resi Ind").Value = 0
            Else
                TDBGrid.Columns("Comm'l Ind").Value = 0
                TDBGrid.Columns("Resi Ind").Value = 1
            End If
End Select

End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
If TDBGrid.DataChanged = True Then

If TDBGrid.AddNewMode > 0 Then
    If Screen.ActiveControl.Name <> "cmdDelete" Then 'delete not pending
        If TDBGrid.Columns("Assembly Skey") = 0 Then
            bln_AssemblyID = True
            Cancel = True
        Else
            bln_AssemblyID = False
        End If
    End If
End If
End If
If TDBGrid.Columns("Comm'l Ind").Value = 0 And TDBGrid.Columns("Resi Ind").Value = 0 Then
            bln_UseInd = True
            Cancel = True
        Else
            bln_UseInd = False
End If
End Sub

Private Sub TDBGrid_Change()
    With TDBGrid
        If .Columns(.Col).Caption = "Book Desc" _
            Or .Columns(.Col).Caption = "Metric Book Desc" _
            Or .Columns(.Col).Caption = "Tech Desc" _
            Or .Columns(.Col).Caption = "Metric Tech Desc" _
                Then
            If Len(.Text) > 75 Then .Text = Left(.Text, 75)
        End If
    End With
End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
If bln_AssemblyID = True Then
    MsgBox "Please enter a valid Assembly ID"
    Response = 0
    TDBGrid.SetFocus
    bln_AssemblyID = False
End If
If bln_UseInd = True Then
    MsgBox "Commercial or Residential Use indicator must be selected."
    Response = 0
    TDBGrid.SetFocus
    bln_UseInd = False
End If
    Response = 0

End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    On Error Resume Next

    Dim strType As String
    Dim objRowInfo As CRowInfo
    Dim bLocked As Boolean

    objRowInfo = Nothing
    Set objRowInfo = m_objGridMap.rowinfo(str(Bookmark))
    If Split = 1 Then   'cost split
        If IsNumeric(Bookmark) Then
            m_rec.Bookmark = Bookmark
            strType = m_rec.Fields("type_code").Value
            With TDBGrid.Columns(Col)
    'Process Exception record fields
                If strType = "E" Then
                    Select Case .Caption
                        Case "Std Hours"
                            bLocked = True
                        Case "Opn Hours"
                            bLocked = True
                        Case "R&R Hours"
                            bLocked = True
                        Case "Met Hours"
                            bLocked = True
                    End Select
                ElseIf strType = "M" Then
                    Select Case .Caption
                        Case "Percent"
                            bLocked = True
                        Case "Std Mat"
                            bLocked = True
                        Case "Std Equip"
                            bLocked = True
                        Case "Std Labor"
                            bLocked = True
                        Case "Std Total"
                            bLocked = True
                        Case "Std Mat OP"
                            bLocked = True
                        Case "Std Equip OP"
                            bLocked = True
                        Case "Std Labor OP"
                            bLocked = True
                        Case "Std Total OP"
                            bLocked = True
                        Case "Std Hours"
                            bLocked = True
                        Case "Opn Mat"
                            bLocked = True
                        Case "Opn Equip"
                            bLocked = True
                        Case "Opn Labor"
                            bLocked = True
                        Case "Opn Total"
                            bLocked = True
                        Case "Opn Mat OP"
                            bLocked = True
                        Case "Opn Equip OP"
                            bLocked = True
                        Case "Opn Labor OP"
                            bLocked = True
                        Case "Opn Total OP"
                            bLocked = True
                        Case "Opn Hours"
                            bLocked = True
                        Case "R&R Mat"
                            bLocked = True
                        Case "R&R Inst"
                            bLocked = True
                        Case "R&R Equip"
                            bLocked = True
                        Case "R&R Labor"
                            bLocked = True
                        Case "R&R Hours"
                            bLocked = True
                        Case "R&R Total"
                            bLocked = True
                        Case "R&R Mat OP"
                            bLocked = True
                        Case "R&R Equip OP"
                            bLocked = True
                        Case "R&R Labor OP"
                            bLocked = True
                        Case "R&R Total OP"
                            bLocked = True
                        Case "Met Mat"
                            bLocked = True
                        Case "Met Equip"
                            bLocked = True
                        Case "Met Labor"
                            bLocked = True
                        Case "Met Total"
                            bLocked = True
                        Case "Met Mat OP"
                            bLocked = True
                        Case "Met Equip OP"
                            bLocked = True
                        Case "Met Labor OP"
                            bLocked = True
                        Case "Met Total OP"
                            bLocked = True
                        Case "Met Hours"
                            bLocked = True
                        Case "Metric Unit"
                            bLocked = True
                    End Select
                End If
                If bLocked = True Then
                    bLocked = True
                    CellStyle.Locked = True
                    CellStyle.ForeColor = vbGrayText
                Else
                    CellStyle.ForeColor = vbBlack
                    CellStyle.Locked = False
                End If
            End With
        End If
    End If
    If (Condition And dbgSelectedRow) = 8 Then
        CellStyle.ForeColor = vbWhite
    End If
End Sub

Private Sub TDBGrid_LostFocus()
'TDBGrid.Update
End Sub

Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)
Select Case MsgId
    Case 2
'        m_rec.Bookmark = TDBGrid.Bookmark
'        m_rec.Fields("type_code") = TDBGrid.Columns("Type").Value
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Mat").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Inst").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Equip").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Labor").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Total").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Mat OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Inst OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Equip OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Labor OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Total OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Std Hours").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Mat").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Inst").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Equip").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Labor").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Total").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Mat OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Inst OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Equip OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Labor OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Total OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Opn Hours").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Mat").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Inst").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Equip").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Labor").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Total").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Mat OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Inst OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Equip OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Labor OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Total OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("R&R Hours").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Mat").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Inst").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Equip").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Labor").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Total").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Mat OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Inst OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Equip OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Labor OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Total OP").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Met Hours").ColIndex)
    End Select
End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
   
    ' Position to the desired record and execute delete
    If check_error(Bookmark, TDBGrid.Columns("Assembly Skey")) = True Then Exit Sub
    
    m_rec.Bookmark = Bookmark
    Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
    If Not clsRowInfo Is Nothing Then
        If clsRowInfo.State = STATE_NEW Then
            TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
            m_rec.Delete ' Remove it from recordset so it doesn't show.
            Exit Sub
        End If
    End If

    ' Build SQL statement
    strUpdate = "exec sp_delete_assembly "
    strUpdate = strUpdate + "@assembly_skey=" + str(m_rec.Fields("assembly_skey"))

    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetError m_rec.Bookmark, strError
        m_rec.Fields("Error").Value = "Error"
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
    End If
'    TDBGrid.ReBind
    Exit Sub

NoDelete:
    MsgBox Err.Description
    Bookmark = Null
End Sub
Private Function check_error(varBookmark As Variant, lngSkey As Long) As Boolean
Dim clsRowInfo As CRowInfo
Dim intErrors As Integer ' Tracks if any errors have occurred
Dim rsTemp As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean
Dim strError As String
Dim iAssemblyBookRecordcount As Integer

    strSelect = "select count(*) as abd_count from assembly_book_detail where assembly_skey = " + CStr(lngSkey)
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsTemp)
    If blnReturn = True Then
        If rsTemp.EOF Then
            iAssemblyBookRecordcount = 0
        Else
            iAssemblyBookRecordcount = rsTemp.Fields("abd_count")
        End If
        rsTemp.Close
    End If

    If iAssemblyBookRecordcount > 0 Then
        strError = "This assembly is in use on " + CStr(iAssemblyBookRecordcount) + " assembly book record(s).  Please remove them prior to deleting this assembly."
        Set clsRowInfo = m_objGridMap.rowinfo(CInt(varBookmark))
        If clsRowInfo Is Nothing Then
            m_objGridMap.SetRowState CInt(varBookmark), STATE_NONE
            Set clsRowInfo = m_objGridMap.rowinfo(CInt(varBookmark))
        End If
        intErrors = intErrors + 1
        clsRowInfo.Error = True
        clsRowInfo.ErrorText = strError
        check_error = True
    Else
        check_error = False
    End If
    Set rsTemp = Nothing
End Function
' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim strSelect As String
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnReturn As Boolean
    Dim vntMyBookmark As Variant
    Dim i As Integer
    If TDBGrid.DataChanged = True Then
        Select Case TDBGrid.Columns(ColIndex).Caption
        Case "Assembly ID"
            If Trim(TDBGrid.Text) = "" Then
                MsgBox "The Assembly ID may not be blank."
                Cancel = True ' Cancel so old value is replaced
            Else
                ' Check to see if the assembly_id entered exists already
                strSelect = "Select assembly_id, tech_desc as assembly_tech_desc, metric_tech_desc as assembly_metric_tech_desc, assembly_skey from Assembly_detail where assembly_id='" + TDBGrid.Text + "'"
                 blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
                If rec.RecordCount > 0 Then
                    MsgBox "The Assembly ID " + TDBGrid.Text + " already exists."
                    Cancel = True ' Cancel so old value is replaced
                End If
                rec.Close
                Set rec = Nothing
            End If
        Case "Uni2 Lvl 3"       'Validate the text against a level 3 record in the Uniformt2 ID Hierarchy table
            strSelect = "select * from uniformat2_id_hierarchy where uni2_level = 3 and uni2_category_id='" + TDBGrid.Text + "'"
            ' Use DAL to perform select
             blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
            If rec.RecordCount = 0 Then
                MsgBox "The Level 3 Uniformat ID " + TDBGrid.Text + " does not exist."
                Cancel = True ' Cancel so old value is replaced
            End If
            rec.Close
            Set rec = Nothing
        Case "Uni2 Lvl 5"       'Validate the text against a level 5 record in the Uniformt2 ID Hierarchy table
            strSelect = "select * from uniformat2_id_hierarchy where uni2_level = 5 and uni2_category_id='" + TDBGrid.Columns("Uni2 Lvl 3") + TDBGrid.Text + "'"
            ' Use DAL to perform select
             blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
            If rec.RecordCount = 0 Then
                MsgBox "The Level 5 Uniformat ID " + TDBGrid.Text + " does not exist."
                Cancel = True ' Cancel so old value is replaced
            End If
            rec.Close
            Set rec = Nothing
        Case "Uni2 Lvl 6"        'Validate the text len - must be 4 for the last part of the ID
            If Len(TDBGrid.Text) <> 4 Then
                MsgBox "The third section of the Uniformat 2 ID must be 4 positions."
            End If
        End Select
    End If
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("Assembly Skey").Value = 0
    TDBGrid.Columns("last_update_id") = 0
End Sub

Public Function CloneRow() As ADODB.RecordSet
    ' Pass to base class
    m_objGridMap.CloneRow (True)
    Set CloneRow = Nothing
End Function

Public Property Get GRIDTYPE() As String
    GRIDTYPE = m_GRIDTYPE
End Property

Public Function GetError(intBookmark As Integer) As String
    On Error Resume Next
    Dim objRowInfo As CRowInfo
    Set objRowInfo = Nothing
    Set objRowInfo = m_objGridMap.rowinfo(intBookmark)
    If Not objRowInfo Is Nothing Then
        GetError = objRowInfo.ErrorText
    Else
        GetError = ""
    End If
End Function

Public Function IsPendingChange() As Boolean
    IsPendingChange = m_objGridMap.IsPendingChange
End Function

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub

