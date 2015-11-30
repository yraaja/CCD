VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCrewMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Crew grid
Option Explicit

' <modulename> CCrewMap </modulename>
' <functionname>General (Main) </functionname>
'
' <summary>CLASS HELPER FOR "Crew Grid"
'
'Crews are groupings of laborers and equipment.  The crew parts are defined as:
'"   Trade_id
'"   Equipment_id
'
'groupings
'
'Key FUNCTIONS / SUBS:
'
'"   InitGrid()
'Adds columns to the datagrid dynamically.  Each column usually maps a record/result set from an ADO query or exec stored procedure.
'
'"   TDBGrid_AfterColUpdate
'
'Upon tabbing across the datagrid checking/validation occurs as the tab passes through each grid cell.  In this case the "crew_skey" cell is auto populated with the "crew_skey" value as the user tabs through the "Crew_Id" cell.
'
'"   TDBGrid_BeforeColUpdate
'
'Upon tabbing across the datagrid checking/validation occurs as the tab passes through each grid cell.  For "crew" processing the following columns/cells are validated:
'"Crew_ID"
'"Type"
'"STD Qty", "OPN Qty", "R&R Qty", "Metric Qty"
'
'If something is found to be illegitimate, CANCEL is set to true and and column updating is skipped
'</summary>
'
' <seealso>frmCrewGrid</seealso>
' <datastruct> m_rec </datastruct>
'
'
' <storedprocedurename> sp_update_crew_usage_driver </storedprocedurename>
' <storedprocedurename sp_delete_crew_usage </storedprocedurename>
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'o frmCrewGrid.frm
'o CGridMap.Cls
'o CCDdal.CRSMDataAccess
'(Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>


'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "CrewUsage" ' Identity of grid
' Names of some table this grid uses
Const CREW = "Crew"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private m_blnCrew_ID_Error As Boolean
Private m_blnType_Error As Boolean
Private m_blnLaborEquip_ID_Error As Boolean
Private m_iPostMsgID As Integer

Public strSource As String   'Source of form - Unit Cost, Material, or None

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
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
'    m_objGridMap.SetFrozenCols (3)
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "crew_skey", "crew_skey", ALL_TABLES, 1000, True, False
    m_objGridMap.AddCol "Crew ID", "crew_id", vbNullString, 800, , , , , True
    m_objGridMap.AddCol "Type", "skey_type", ALL_TABLES, 700, , True, True, , True
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "TracesCrew ID", "traces_crew_id", vbNullString, 1500, , , , , True
    m_objGridMap.AddCol "Trade/Equip ID", "trade_equip_id", ALL_TABLES, 1500, , , , , True
    m_objGridMap.AddCol "Trade/Equip Skey", "skey", ALL_TABLES, 1500, True, False, True, , True
    m_objGridMap.AddCol "Sort Order", "sort_order", ALL_TABLES, 1000, , , , , True
    
    m_objGridMap.AddCol "Desc", "descr", vbNullString, 2000, True, , , , True
    m_objGridMap.AddCol "Plural Desc", "plural_desc", vbNullString, 2000, True, , , , True
    m_objGridMap.AddCol "Metric Desc", "metric_desc", vbNullString, 2000, True, , , , True
    m_objGridMap.AddCol "Plural Metric Desc", "plural_metric_desc", vbNullString, 2000, True, , , , True
    m_objGridMap.AddCol "OverRide Desc", "override_desc", ALL_TABLES, 2000, , , , , True
    m_objGridMap.AddCol "OverRide Desc Plural", "override_desc_plural", ALL_TABLES, 2000, , , , , True
    m_objGridMap.AddCol "OverRide Metric Desc", "override_metric_desc", ALL_TABLES, 2000, , , , , True

    m_objGridMap.AddCol "STD Qty", "std_qty", ALL_TABLES, 1500, , , , FORMAT_DECIMAL, True
    m_objGridMap.AddCol "STD Labor Wage (hourly)", "std_labor_wage_hrly", vbNullString, 2000, True, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "STD (daily)", "std_daily", vbNullString, 1000, True, True, True, FORMAT_DECIMAL, True
    m_objGridMap.AddCol "STD OP (daily)", "std_op_dly", vbNullString, 2000, True, , , FORMAT_PRICE, True
    m_objGridMap.AddCol "STD Labor Wage OP (hourly)", "std_labor_wage_op_hrly", vbNullString, 1000, True, True, True, FORMAT_DECIMAL, True
    m_objGridMap.AddCol "OPN Qty", "OPN_qty", ALL_TABLES, 1500, , , , FORMAT_DECIMAL, True
    m_objGridMap.AddCol "BIA Qty", "BIA_qty", ALL_TABLES, 1500, , , , FORMAT_DECIMAL, True
    m_objGridMap.AddCol "TR Qty", "TR_qty", ALL_TABLES, 1500, , , , FORMAT_DECIMAL, True
    m_objGridMap.AddCol "OPN Labor Wage (hourly)", "opn_labor_wage_hrly", vbNullString, 1100, True, False, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "OPN (daily)", "opn_daily", vbNullString, 1000, True, True, True, FORMAT_DECIMAL, True
    m_objGridMap.AddCol "OPN Labor Wage OP (hourly)", "opn_labor_wage_op_hrly", vbNullString, 1100, True, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "OPN Labor Wage OP (daily)", "opn_labor_wage_op_dly", vbNullString, 2000, True, , , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Qty", "RR_qty", ALL_TABLES, 1000, , , , FORMAT_DECIMAL, True
    m_objGridMap.AddCol "R&R Labor Wage (hourly)", "rr_labor_wage_hrly", vbNullString, 1100, True, False, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R (daily)", "rr_daily", vbNullString, 1100, True, True, True, FORMAT_DECIMAL, True
    m_objGridMap.AddCol "R&R Labor Wage OP (hourly)", "rr_labor_wage_op_hrly", vbNullString, 800, True, True, , FORMAT_PRICE, True
    m_objGridMap.AddCol "R&R Labor Wage OP (daily)", "rr_labor_wage_op_dly", vbNullString, 2000, True, , , FORMAT_PRICE, True
    m_objGridMap.AddCol "Metric Qty", "metric_qty", ALL_TABLES, 1500, True, False, , FORMAT_DECIMAL, True
    m_objGridMap.AddCol "Metric Labor Wage (hourly)", "metric_labor_wage_hrly", vbNullString, 1100, True, False, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Metric (daily)", "metric_daily", vbNullString, 1000, True, False, True, FORMAT_PRICE, True
    m_objGridMap.AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False, , True
    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, , True
    m_objGridMap.AddCol "last_update_id", "last_update_id", ALL_TABLES, 800, True, False, False, , True
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
    
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid60.ValueItem
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid70.ValueItem
    Dim Item As New TrueOleDBGrid80.ValueItem
    
    ' Replace display value for total line
    Item.Value = "99999999"
    Item.DisplayValue = "---------------------"
    TDBGrid.Columns("Sort Order").ValueItems.Add Item
    TDBGrid.Columns("Sort Order").ValueItems.Translate = True
 
    ' Replace display value for total line
    Item.Value = "1"
    Item.DisplayValue = "Labor"
    TDBGrid.Columns("Type").ValueItems.Add Item
    Item.Value = "2"
    Item.DisplayValue = "Equip"
    TDBGrid.Columns("Type").ValueItems.Add Item
    Item.Value = "3"
    Item.DisplayValue = "Total"
    TDBGrid.Columns("Type").ValueItems.Add Item
    TDBGrid.Columns("Type").DefaultValue = "1"
    TDBGrid.Columns("Type").ValueItems.Translate = True
    TDBGrid.Columns("Type").ValueItems.MaxComboItems = 2
    TDBGrid.Columns("Type").ValueItems.Presentation = dbgComboBox

    
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    Dim colUnitCostID As New Collection
    
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
            ' If record was added, or if both mat_id and unit_cost_id were changed
                If clsRowInfo.State = STATE_NEW Or clsRowInfo.State = STATE_MODIFIED Then
                    strUpdate = "exec sp_update_crew_usage_driver "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "',"
                    strUpdate = strUpdate + "@UseCommitControl = 1"
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If strError = "" Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                        m_rec.Fields("last_update_date").Value = Now
                        m_rec.Fields("last_update_person").Value = strUserName
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

Private Sub Class_Initialize()
    m_objGridMap.m_colCloneCols.Add 1, "crew_skey"
   m_objGridMap.m_colCloneCols.Add 1, "crew_id"
    m_objGridMap.m_colCloneCols.Add 1, "skey_type"
    m_objGridMap.m_colCloneCols.Add 1, "sort_order"
    m_objGridMap.m_colCloneCols.Add 1, "trade_desc"
    m_objGridMap.m_colCloneCols.Add 1, "crew_equip_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_crew_equip_desc"
    m_objGridMap.m_colCloneCols.Add 1, "trade-equip_id"
    m_objGridMap.m_colCloneCols.Add 1, "TR_qty"
    m_objGridMap.m_colCloneCols.Add 1, "bia_qty"
    m_objGridMap.m_colCloneCols.Add 1, "std_qty"
    m_objGridMap.m_colCloneCols.Add 1, "opn_qty"
    m_objGridMap.m_colCloneCols.Add 1, "rr_qty"
    m_objGridMap.m_colCloneCols.Add 1, "std_labor_wage_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_wage_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "rr_labor_wage_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "std_daily"
    m_objGridMap.m_colCloneCols.Add 1, "opn_daily"
    m_objGridMap.m_colCloneCols.Add 1, "rr_daily"
    m_objGridMap.m_colCloneCols.Add 1, "std_labor_wage_op_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_wage_op_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "rr_labor_wage_op_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "std-labor_wage_op_dly"
    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_wage_op_dly"
    m_objGridMap.m_colCloneCols.Add 1, "rr_labor_wage_op_dly"
'    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_wage_op_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "metric_qty"
    m_objGridMap.m_colCloneCols.Add 1, "metric_labor_wage-hrly"
    m_objGridMap.m_colCloneCols.Add 1, "metric_daily"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_date"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_person"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_id"
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
Dim rec As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean

Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Crew ID"
            strSelect = "select crew_skey from crew where crew_id = '" + TDBGrid.Columns("Crew ID").Value + "'"
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        If blnReturn = True Then
            If Not (rec.BOF And rec.EOF) Then
                TDBGrid.Columns("crew_skey") = CStr(rec.Fields("crew_skey"))
            End If
        End If
        rec.Close
        Set rec = Nothing
    End Select
End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
Dim rec As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean

If Trim(TDBGrid.Columns("Crew Id").Text) = "" Or IsNull(TDBGrid.Columns("Crew Id").Text) Then
    m_blnCrew_ID_Error = True
    Cancel = True
Else
    strSelect = "select * from crew where crew_id = '" + Trim(TDBGrid.Columns("Crew Id").Value) + "'"
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
    If blnReturn = False Then
        m_blnCrew_ID_Error = True
        Cancel = True
    Else
        If rec.RecordCount = 0 Then  'Not found
            m_blnCrew_ID_Error = True
            Cancel = True
        End If
    End If
    rec.Close
    Set rec = Nothing
End If

If Cancel = False Then
    If Trim(UCase(TDBGrid.Columns("Type").Value)) <> "2" And Trim(UCase(TDBGrid.Columns("Type").Value)) <> "1" Then
        m_blnType_Error = True
        Cancel = True
    End If
End If
'Validate the ID - Labor Trade or Equipment
If Cancel = False Then
    If Trim(TDBGrid.Columns("Trade/Equip ID").Value) = "" Or IsNull(TDBGrid.Columns("Trade/Equip ID").Value) Then
        m_blnLaborEquip_ID_Error = True
        Cancel = True
    Else
        If UCase(TDBGrid.Columns("Type").Value) = "E" Then  'Equipment line
            strSelect = "select * from equipment where equip_id = '" + Trim(TDBGrid.Columns("Trade/Equip ID").Value) + "'"
        ElseIf UCase(TDBGrid.Columns("Type").Value) = "L" Then  'Equipment line
            strSelect = "select * from labor_trade where trade_id = '" + Trim(TDBGrid.Columns("Trade/Equip ID").Value) + "'"
        End If
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        If blnReturn = False Then
            m_blnLaborEquip_ID_Error = True
            Cancel = True
        Else
            If rec.RecordCount = 0 Then  'Not found
                m_blnLaborEquip_ID_Error = True
                Cancel = True
            End If
        End If
        rec.Close
        Set rec = Nothing
    End If
End If

End Sub
Private Sub TDBGrid_Change()
Dim intStart As Integer
    If TDBGrid.Columns(TDBGrid.Col).Caption = "Material ID" Then
        intStart = TDBGrid.SelStart
        TDBGrid.Text = UCase(TDBGrid.Text)
        TDBGrid.SelStart = intStart
    End If

End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)

If m_blnCrew_ID_Error = True Then
    MsgBox "Please enter a valid Crew ID"
    TDBGrid.SetFocus
    m_blnCrew_ID_Error = False
End If
If m_blnType_Error = True Then
    MsgBox "Please enter a valid Crew Usage Type - 'L'(Labor) or 'E'(Equipment)."
    TDBGrid.SetFocus
    m_blnType_Error = False
End If
If m_blnLaborEquip_ID_Error = True Then
    If UCase(TDBGrid.Columns("Type").Value) = "2" Then
        MsgBox "Please enter a valid Equipment ID"
    End If
    If UCase(TDBGrid.Columns("Type").Value) = "1" Then
        MsgBox "Please enter a valid Trade ID"
    End If
    TDBGrid.SetFocus
    m_blnLaborEquip_ID_Error = False
End If

Response = 0

End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
Dim rs_m_rec_clone As ADODB.RecordSet
    Dim strType As String
    Dim objRowInfo As CRowInfo
    Dim bLocked As Boolean
    Dim blnNew As Boolean
'    MsgBox TDBGrid.AllowAddNew
'    MsgBox TDBGrid.AddNewMode
    Set objRowInfo = m_objGridMap.rowinfo(str(Bookmark))
        If IsNumeric(Bookmark) Then
            Set rs_m_rec_clone = m_rec.Clone
            rs_m_rec_clone.Bookmark = Bookmark
                With TDBGrid.Columns(Col)
                Select Case .Caption
                    Case "Crew ID"
                        If rs_m_rec_clone.Fields("skey") = 0 Then   'New record
                            CellStyle.ForeColor = vbBlack
                            CellStyle.Locked = False
                        Else
                            CellStyle.ForeColor = vbGrayText
                            CellStyle.Locked = True
                        End If
                    Case "Type"
                        If rs_m_rec_clone.Fields("skey") = 0 Then   'New record
                            CellStyle.ForeColor = vbBlack
                            CellStyle.Locked = False
                        Else
                            CellStyle.ForeColor = vbGrayText
                            CellStyle.Locked = True
                        End If
                    End Select
                End With
            If rs_m_rec_clone.Fields("skey_type") = "3" Then
                With TDBGrid.Columns(Col)
                    CellStyle.ForeColor = vbBlue
                End With
            ElseIf rs_m_rec_clone.Fields("skey_type") = "2" Then
                With TDBGrid.Columns(Col)
                Select Case .Caption
                    Case "Crew ID"
                        If rs_m_rec_clone.Fields("skey") = 0 Then   'New record
                            CellStyle.ForeColor = vbBlack
                            CellStyle.Locked = False
                        Else
                            CellStyle.ForeColor = vbGrayText
                            CellStyle.Locked = True
                        End If
                    Case "Type"
                        If rs_m_rec_clone.Fields("skey") = 0 Then   'New record
                            CellStyle.ForeColor = vbBlack
                            CellStyle.Locked = False
                        Else
                            CellStyle.ForeColor = vbGrayText
                            CellStyle.Locked = True
                        End If
                    Case "Crew Equip Description"
                        CellStyle.ForeColor = vbBlack
                        CellStyle.Locked = False
                    Case "Metric Crew Equip Desc"
                        CellStyle.ForeColor = vbBlack
                        CellStyle.Locked = False
                    Case "OverRide Desc Plural"
                        CellStyle.ForeColor = vbGrayText
                        CellStyle.Locked = True
                End Select
                End With
            ElseIf rs_m_rec_clone.Fields("skey_type") = "1" Then
                With TDBGrid.Columns(Col)
                Select Case .Caption
                    Case "Crew ID"
                        If rs_m_rec_clone.Fields("skey") = 0 Then   'New record
                            CellStyle.ForeColor = vbBlack
                            CellStyle.Locked = False
                        Else
                            CellStyle.ForeColor = vbGrayText
                            CellStyle.Locked = True
                        End If
                    Case "Type"
                        If rs_m_rec_clone.Fields("skey") = 0 Then   'New record
                            CellStyle.ForeColor = vbBlack
                            CellStyle.Locked = False
                        Else
                            CellStyle.ForeColor = vbGrayText
                            CellStyle.Locked = True
                        End If
                    Case "Crew Equip Description"
                        CellStyle.Locked = True
                        CellStyle.ForeColor = vbGrayText
                    Case "Metric Crew Equip Desc"
                        CellStyle.Locked = True
                        CellStyle.ForeColor = vbGrayText
                    Case "OverRide Desc Plural"
                        CellStyle.ForeColor = vbBlack
                        CellStyle.Locked = False
                    Case "OverRide Metric Desc"
                        CellStyle.Locked = True
                        CellStyle.ForeColor = vbGrayText
                End Select
                End With
            End If
            Set rs_m_rec_clone = Nothing
        End If
'    End If
    
    ' If the row is highlighted, then let it be
    If (Condition And dbgSelectedRow) = 8 Then
        Exit Sub
    End If
    ' This col is always white unless an error occurred
    ' NOTE: In this release of the grid, definied styles (like "Normal") cannot be applied to cells
    If Not objRowInfo Is Nothing Then
        If objRowInfo.Error = True Then
            CellStyle.BackColor = vbRed
        Else
        End If
    Else
    End If

End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
'If the row is not the last row for the trade group/start date, lock it.
On Error Resume Next
Dim rs_m_rec_clone As ADODB.RecordSet
Set rs_m_rec_clone = m_rec.Clone
rs_m_rec_clone.Bookmark = Bookmark
'If m_GroupDate(Bookmark) <> rsLaborClone.Fields("start_date") And rsLaborClone.Fields("trade_group_code") > " " Then
If rs_m_rec_clone.Fields("skey_type") = "3" Then
    RowStyle = "LockedStyle"
    RowStyle.ForeColor = vbBlue
    RowStyle.Locked = True
Else
    RowStyle = "ActiveRow2"
End If
Set rs_m_rec_clone = Nothing
End Sub

Private Sub FillFields(strSQL As String)
Dim rec As New ADODB.RecordSet ' Recordset to hold query results
Dim blnReturn As Boolean
Dim i As Integer
On Error Resume Next
     ' Use DAL to perform select
     blnReturn = g_objDAL.GetRecordset(vbNullString, strSQL, rec)

    For i = 0 To rec.Fields.Count - 1
        m_rec.Fields(rec.Fields(i).Name).Value = rec.Fields(i).Value
    Next i
    rec.Close
    Set rec = Nothing
End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
    Dim colUnitCostID As New Collection
   
    ' Position to the desired record and execute delete
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
    
    If m_rec.Fields("skey_type").Value = 3 Then 'total line
        m_objGridMap.SetError m_rec.Bookmark, "Total lines may not be deleted."
    Else
    
        ' Build SQL statement
        strUpdate = "exec sp_delete_crew_usage "
        strUpdate = strUpdate + "@crew_skey=" + CStr(m_rec.Fields("crew_skey")) + ","
        strUpdate = strUpdate + "@skey=" + str(m_rec.Fields("skey")) + ","
        strUpdate = strUpdate + " @skey_type=" + m_rec.Fields("skey_type")
        strUpdate = strUpdate + ", @UseCommitControl=1"
        
        blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
        If Not blnReturn Then
            m_objGridMap.SetError m_rec.Bookmark, strError
            Bookmark = Null
        Else
            TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
            'SaveUnitCostID colUnitCostID, m_rec.Fields("unit_cost_id")
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
            m_rec.Delete ' Remove it from recordset so it doesn't show.
        End If
    End If
    Exit Sub

NoDelete:
    MsgBox Err.Description
    Bookmark = Null
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

    On Error Resume Next
    Dim strSelect As String
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnReturn As Boolean
    Dim i As Integer
    Dim j As Integer
'*** APEX Migration Utility Code Change ***
'    Dim Cols As TrueOleDBGrid60.Columns
'*** APEX Migration Utility Code Change ***
'    Dim Cols As TrueOleDBGrid70.Columns
    Dim Cols As TrueOleDBGrid80.Columns
'*** APEX Migration Utility Code Change ***
'    Dim col As TrueOleDBGrid60.Column
'*** APEX Migration Utility Code Change ***
'    Dim col As TrueOleDBGrid70.Column
    Dim Col As TrueOleDBGrid80.Column
    Dim vntMyBookmark As Variant
    Dim strError As String
    
    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Crew ID"
        strSelect = "select crew_skey from crew where crew_id = '" + TDBGrid.Text + "'"
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        If blnReturn = True Then
            If rec.BOF And rec.EOF Then
                MsgBox "Please enter a valid Crew ID."
                Cancel = True ' Cancel so old value is replaced
            End If
        End If
    Case "Type"
        If Len(TDBGrid.Columns(ColIndex).Value) > 0 Then
            If UCase(TDBGrid.Columns(ColIndex).Value) <> "1" And UCase(TDBGrid.Columns(ColIndex).Value) <> "2" Then
                MsgBox "The type must be 'Equipment' or 'Labor'."
                Cancel = True ' Cancel so old value is replaced
            End If
        End If
    Case "STD Qty", "OPN Qty", "R&R Qty", "Metric Qty"
        If Trim(TDBGrid.Text) = "" Then
            MsgBox TDBGrid.Columns(ColIndex).Caption + " may not be blank."
            Cancel = True ' Cancel so old value is replaced
        Else
            If Not IsNumeric(TDBGrid.Text) Then
                MsgBox TDBGrid.Columns(ColIndex).Caption + " must be a number."
                Cancel = True ' Cancel so old value is replaced
            Else
                If TDBGrid.Text <= 0 Then
                    MsgBox "Please enter a valid number."
                    Cancel = True ' Cancel so old value is replaced
                End If
            End If
        End If
    End Select

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
