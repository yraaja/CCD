VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUCostUsageMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Unit Cost Usage grid
Option Explicit

' <modulename> CUCostUsageMap </modulename>
' <functionname>General (Main) </functionname>
'
' <summary>CLASS HELPER FOR "Unit Cost Usage Grid"
'
'"   Unit Cost Usage:
'Mappings of references or uses of lines across the entire family of Assemblies
'
'groupings
'
'Key FUNCTIONS / SUBS:
'
'"   InitGrid()
'Adds columns to the datagrid dynamically.  Each column usually maps a record/result set from an ADO query or exec stored procedure.
'
'FillUnitCost
'Builds a query containing needed unit cost information from which the grid will be populated
'
'FillAssembly
'
'Builds a query containing needed assembly information from which the grid will be populated
'
'FillFields
'
'Populates the datagrid cells/columns from a generic recordset built from the parameter sql string built by either "fillUnitCost()" or  "FillAssembly()"
'(see just above)
'
'
'"   TDBGrid_AfterColUpdate
'
'Upon tabbing across the datagrid the "Unit Cost ID" and/or "Assembly ID" cells are auto-populated
'
'"   TDBGrid_BeforeColUpdate
'
'Upon tabbing across the datagrid checking/validation occurs as the tab passes through each grid cell.  In this case the "Assembly ID" and the "Unit Cost ID" cells are validated and auto populated as the user tabs through either of these two cells
'If something is found to be illegitimate, CANCEL is set to true and and column updating is skipped
'
'"   UPDATE()
'Either the "update" or "insert" stored procedure is built and run  (see stored procedures, below)
'</summary>
'
' <seealso>frmCrewGrid</seealso>
' <datastruct> m_rec </datastruct>
'
'
' <storedprocedurename> sp_update_unit_cost_usage</storedprocedurename>
' <storedprocedurename > sp_insert_unit_cost_usage</storedprocedurename>
'<storedprocedurename > sp_delete_unit_cost_usage</storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'N/A
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'o frmUnitCostUsage.frm
'o CGridMap.Cls
'o CCDdal.CRSMDataAccess
'(Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>


'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid  ' Access to grid events
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid  ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid   ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "UnitCostUsage"      ' Identity of grid
' Names of some table this grid uses
Const ASSEMBLYDETAIL_TABLE = "Assembly_detail"
Const UNITCOSTUSAGE_TABLE = "Assembly_unit_cost_usage"
Const PUB_ASSEMBLY_COSTWORKS_TABLE = "published_assembly_costworks"
Const PUB_UC_COSTWORKS_TABLE = "published_unit_cost_costworks"
Const UNITCOSTDTL_TABLE = "unit_cost_detail"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap    ' "Parent" class
Private m_rec As ADODB.RecordSet        ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private m_blnAssemblyID As Boolean
Private m_blnUnitCostID As Boolean
Private m_blnSortOrder As Boolean
Private m_intMasterFormat As Long

Public Property Get MasterFormat() As Long
    MasterFormat = m_intMasterFormat
End Property

Public Property Let MasterFormat(NewValue As Long)
    Dim objColumnDef As CColumnDef
    Dim Col As TrueOleDBGrid80.Column
    
    m_intMasterFormat = NewValue
    If m_intMasterFormat = 1995 Then
        Set objColumnDef = m_objGridMap.GetCol("unit_cost_id")
        If Not (objColumnDef Is Nothing) Then
            objColumnDef.Format = FORMAT_UNIT_COST
        End If
    Else
        Set objColumnDef = m_objGridMap.GetCol("unit_cost_id")
        If Not (objColumnDef Is Nothing) Then
            objColumnDef.Format = FORMAT_UNIT_COST_04
        End If
    End If
    Set objColumnDef = Nothing
    
End Property

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Sub FillUnitCost(strSelUCID As String)

    FillFields ("select unit_cost_id, unit_cost_skey, tech_desc as unit_cost_tech_desc, metric_tech_desc as " + _
        "unit_cost_metric_tech_desc, book_desc as unit_cost_book_desc, metric_book_desc as unit_cost_metric_book_desc, unit " + _
        "From published_unit_cost_costworks as pucc " + _
        "where unit_cost_id = '" + strSelUCID + "'")
        
End Sub

Public Sub FillAssembly(strSelAssemblyID As String)
    Dim strSELECT As String
    
    strSELECT = "select assembly_id as parent_id, assembly_skey as parent_skey,  tech_desc as parent_tech_description, " + _
        "metric_tech_desc as parent_met_tech_desc, book_desc as parent_book_desc, " + _
        "metric_book_desc as parent_metric_book_desc, ad.unit " + _
        "From assembly_detail as ad " + _
        "where assembly_id = '" + strSelAssemblyID + "'"
    FillFields (strSELECT)
    
End Sub

Private Sub FillFields(strSQL As String)
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnReturn As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim varCurBookmark As Variant
    
    On Error Resume Next
    varCurBookmark = m_rec.Bookmark
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSQL, rec)
    For i = 0 To rec.Fields.Count - 1
        m_rec.Fields(rec.Fields(i).Name).Value = rec.Fields(i).Value
    Next i
    For i = 0 To rec.Fields.Count - 1
        For j = 0 To TDBGrid.Columns.Count
            If TDBGrid.Columns(j).DataField = rec.Fields(i).Name Then
                TDBGrid.Columns(j).Value = rec.Fields(i).Value
                Exit For
            End If
        Next j
    Next i
    rec.Close
    Set rec = Nothing
    m_rec.Bookmark = varCurBookmark
    
End Sub

Public Sub SetRowState(intBookmark As Integer, intState As Integer)
    m_objGridMap.SetRowState intBookmark, intState
End Sub

Public Function SetRowStateNone() As Boolean
    If m_rec.RecordCount > 0 Then
        m_rec.MoveFirst
        While Not m_rec.EOF
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
            m_rec.MoveNext
        Wend
    End If

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
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
    
    ' Add the columns to the grid
    'rlh 02/17/2009
    If MASTER_FORMAT_ASSEMBLIES = 1995 Then
        m_objGridMap.AddCol "Unit Cost ID", "unit_cost_id", vbNullString, 1450, , True, True, FORMAT_UNIT_COST
    Else  '2004
        'm_objGridMap.AddCol "Unit Cost ID 04", "ext_unit_cost_id", vbNullString, 1500, False, , , FORMAT_UNIT_COST_04
        'rlh 02/17/2009 Having trouble fitting ext_unit_cost_id in...??!!
        m_objGridMap.AddCol "Unit Cost ID", "unit_cost_id", vbNullString, 1450, , True, True, FORMAT_UNIT_COST
    End If
    'end of rlh
    
    'm_objGridMap.AddCol "Unit Cost ID", "unit_cost_id", vbNullString, 1450, , True, True, FORMAT_UNIT_COST
    m_objGridMap.AddCol "Assembly ID", "parent_id", vbNullString, 1400
    m_objGridMap.AddSplit
    
    m_objGridMap.AddCol "Parent Skey", "parent_skey", UNITCOSTUSAGE_TABLE, 1200, True, False
    m_objGridMap.AddCol "Unit Cost Skey", "unit_cost_skey", UNITCOSTUSAGE_TABLE, 1200, True, False
    m_objGridMap.AddCol "Assembly Tech Desc", "parent_tech_description", ASSEMBLYDETAIL_TABLE, 2000, True
    m_objGridMap.AddCol "Assembly Metric Desc", "parent_met_tech_desc", ASSEMBLYDETAIL_TABLE, 2000, True
    m_objGridMap.AddCol "Assembly Book Desc", "parent_book_desc", ASSEMBLYDETAIL_TABLE, 2000, True, False
    m_objGridMap.AddCol "Assembly Met Book Desc", "parent_metric_book_desc", ASSEMBLYDETAIL_TABLE, 2000, True, False
    m_objGridMap.AddCol "UC Book Desc", "unit_cost_book_desc", UNITCOSTDTL_TABLE, 2000, True
    m_objGridMap.AddCol "UC Met Book Desc", "unit_cost_metric_book_desc", UNITCOSTDTL_TABLE, 2000, True, False
    m_objGridMap.AddCol "Override Book Desc", "override_book_desc", UNITCOSTUSAGE_TABLE, 2000
    m_objGridMap.AddCol "Override Met Book Desc", "override_metric_book_desc", UNITCOSTUSAGE_TABLE, 2000, , False
    m_objGridMap.AddCol "Override Book Qty", "override_book_qty", UNITCOSTUSAGE_TABLE, 1000, , , , FORMAT_DECIMAL
    m_objGridMap.AddCol "Override Met Book Qty", "override_metric_book_qty", UNITCOSTUSAGE_TABLE, 1000, , False, , FORMAT_DECIMAL
    m_objGridMap.AddCol "Callout ID", "call_out_id", UNITCOSTUSAGE_TABLE, 1000, , False
    m_objGridMap.AddCol "UC Tech Desc", "unit_cost_tech_desc", PUB_UC_COSTWORKS_TABLE, 2000, True, False
    m_objGridMap.AddCol "UC Met Desc", "unit_cost_metric_tech_desc", PUB_UC_COSTWORKS_TABLE, 2000, True, False
    m_objGridMap.AddCol "Unit", "usage_unit", UNITCOSTUSAGE_TABLE, 1000
    m_objGridMap.AddCol "Qty", "usage_unit_qty", UNITCOSTUSAGE_TABLE, 650, , , , FORMAT_DECIMAL
    m_objGridMap.AddCol "Adj Factor", "adj_factor", UNITCOSTUSAGE_TABLE, 800
    m_objGridMap.AddCol "Met Unit", "usage_metric_unit", UNITCOSTUSAGE_TABLE, 1000
    m_objGridMap.AddCol "Met Qty", "usage_metric_unit_qty", UNITCOSTUSAGE_TABLE, 1000, , False, , FORMAT_DECIMAL
    m_objGridMap.AddCol "Met Adj", "metric_adj_factor", UNITCOSTUSAGE_TABLE, 1000, , False
    m_objGridMap.AddCol "Sort", "sort_order", UNITCOSTUSAGE_TABLE, 450
    
    m_objGridMap.AddCol "Material", "mat", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Labor", "labor", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Equip", "equip", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Inst", "inst", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Total", "total", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Mat O&P", "mat_op", vbNullString, 1200, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Labor O&P", "labor_op", vbNullString, 1200, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Equip O&P", "equip_op", vbNullString, 1200, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Inst O&P", "inst_op", vbNullString, 1200, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Total O&P", "total_op", vbNullString, 1200, True, , , FORMAT_PRICE

    m_objGridMap.AddCol "Assembly Total O&P", "assembly_total", vbNullString, 1000, True, , , FORMAT_PRICE

    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1700, True, False, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "last_update_person", vbNullString, 1500, True, False, False
    m_objGridMap.AddCol "last_update_id", "last_update_id", UNITCOSTUSAGE_TABLE, 800, True, False, False

    m_objGridMap.AddCol "Met Mat", "met_mat", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Labor", "met_labor", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Equip", "met_equip", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Inst", "met_inst", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Total", "met_total", vbNullString, 900, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Mat O&P", "met_mat_op", vbNullString, 1200, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Labor O&P", "met_labor_op", vbNullString, 1200, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Equip O&P", "met_equip_op", vbNullString, 1200, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Inst O&P", "met_inst_op", vbNullString, 1200, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Total O&P", "met_total_op", vbNullString, 1200, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Met Assembly Total O&P", "metric_assembly_total", vbNullString, 1500, True, , , FORMAT_PRICE

    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
    
    If Not m_blnPreferences Then
        Dim bRet As Boolean, bRet2 As Boolean
        Dim rec As ADODB.RecordSet, rec2 As ADODB.RecordSet
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid60.ValueItem
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid70.ValueItem
        Dim Item As New TrueOleDBGrid80.ValueItem
        
        ' Unit drop downs
        bRet = g_objDAL.GetRecordset(vbNullString, "select unit from unit_of_measure order by unit", rec)
        If bRet Then
            While Not rec.EOF
                Item.Value = rec.Fields(0)
                TDBGrid.Columns("Unit").ValueItems.Add Item
                rec.MoveNext
            Wend
            TDBGrid.Columns("Unit").ValueItems.Presentation = dbgComboBox
            TDBGrid.Columns("Unit").ValueItems.Validate = True
        End If
        rec.Close
        
         ' (Metric) Unit drop downs   'rlh
        bRet = g_objDAL.GetRecordset(vbNullString, "SELECT distinct [unit], [metric_unit] FROM [METRIC_CONVERSION]", rec)
        If bRet Then
            Item.Value = ""     ' Add one blank item at start of list
            TDBGrid.Columns("Met Unit").ValueItems.Add Item
            While Not rec.EOF
                'Item.Value = rec.Fields(0)  'rlh - commented out on 08/27/08
                Item.Value = rec.Fields(1)   'rlh - added on 08/27/08
                TDBGrid.Columns("Met Unit").ValueItems.Add Item
                rec.MoveNext
            Wend
            TDBGrid.Columns("Met Unit").ValueItems.Presentation = dbgComboBox
            TDBGrid.Columns("Met Unit").ValueItems.Validate = True
        End If
        rec.Close
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
    Dim colAssemblyID As New Collection
    
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
            If clsRowInfo.State = STATE_NEW Then
                strUpdate = "exec sp_insert_unit_cost_usage "
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, UNITCOSTUSAGE_TABLE
                strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                strUpdate = strUpdate + ", @skey_type='A'"
                
                blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                ' Reset on success
                If blnReturn = True Then
                    intSuccess = intSuccess + 1
                    clsRowInfo.State = STATE_NONE
                    m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                    m_rec.Fields("last_update_date").Value = Now
                    m_rec.Fields("last_update_person").Value = strUserName
'                    If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
                        SaveAssemblyID colAssemblyID, m_rec.Fields("parent_ID")
'                    End If
                Else
                    intErrors = intErrors + 1
                    clsRowInfo.Error = True
                    clsRowInfo.ErrorText = strError
                End If
            ' If the record was changed
            ElseIf clsRowInfo.State = STATE_MODIFIED Then
                blnReturn = False
                strError = ""
                strUpdate = "exec sp_update_unit_cost_usage "
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, UNITCOSTUSAGE_TABLE
                strUpdate = strUpdate + "@last_update_person='" + strUserName + "', "
                strUpdate = strUpdate + "@old_parent_skey=" + str(m_rec.Fields("parent_skey").OriginalValue) + ", "
                strUpdate = strUpdate + "@old_unit_cost_skey=" + str(m_rec.Fields("unit_cost_skey").OriginalValue) + ", "
                If IsNull(m_rec.Fields("sort_order").OriginalValue) Or m_rec.Fields("sort_order").OriginalValue = "" Then
                    strUpdate = strUpdate + "@old_sort_order=''"
                Else
                    strUpdate = strUpdate + "@old_sort_order='" + m_rec.Fields("sort_order").OriginalValue + "'"
                End If
                strUpdate = strUpdate + ", @skey_type='A'"
                blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                ' Reset on success
                If blnReturn = True Then
                    intSuccess = intSuccess + 1
                    clsRowInfo.State = STATE_NONE
                    m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                    m_rec.Fields("last_update_date").Value = Now
                    m_rec.Fields("last_update_person").Value = strUserName
'                    If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
                        SaveAssemblyID colAssemblyID, m_rec.Fields("parent_ID")
'                    End If
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
        'Update the unit cost for the modified materials.
        UpdateAssembly colAssemblyID
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
    m_objGridMap.m_colCloneCols.Add 1, "unit_cost_id"
    m_objGridMap.m_colCloneCols.Add 1, "parent_id"
    m_objGridMap.m_colCloneCols.Add 1, "alt_unit_cost_id"
    m_objGridMap.m_colCloneCols.Add 1, "unit_cost_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "unit_cost_metric_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "sort_order"
    m_objGridMap.m_colCloneCols.Add 1, "override_book_qty"
    m_objGridMap.m_colCloneCols.Add 1, "override_metric_book_qty"
    m_objGridMap.m_colCloneCols.Add 1, "assembly_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "assembly_metric_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "usage_unit"
    m_objGridMap.m_colCloneCols.Add 1, "usage_unit_qty"
    m_objGridMap.m_colCloneCols.Add 1, "adj_factor"
    m_objGridMap.m_colCloneCols.Add 1, "unit_cost_skey"
    m_objGridMap.m_colCloneCols.Add 1, "parent_skey"
    m_objGridMap.m_colCloneCols.Add 1, "usage_metric_unit"
    m_objGridMap.m_colCloneCols.Add 1, "usage_metric_unit_qty"
    m_objGridMap.m_colCloneCols.Add 1, "metric_adj_factor"
    m_objGridMap.m_colCloneCols.Add 1, "override_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "override_metric_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "call_out_id"
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    Dim strSELECT As String
    Dim rec As ADODB.RecordSet
    Dim blnReturn  As Boolean
    Dim i As Integer
    Dim bln_get_skey As Boolean
    On Error GoTo Exit_Sub
    
    m_rec.Bookmark = TDBGrid.Bookmark
    
    Select Case TDBGrid.Columns(ColIndex).Caption
        Case "Unit Cost ID"     'Unit Cost ID
            FillUnitCost Compress_String(TDBGrid.Columns(ColIndex).Value)
        Case "Assembly ID"
            FillAssembly Compress_String(TDBGrid.Columns(ColIndex).Value)
    End Select
    
Exit_Sub:

End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
'Validate Record
Dim strUnitCostId As String
Dim strAssemblyId As String
Dim vntMyBookmark As Variant
Dim strSELECT As String
Dim rec As ADODB.RecordSet
Dim blnReturn As Boolean
If TDBGrid.DataChanged = True Then
    If Screen.ActiveControl.Name <> "cmdDelete" Then 'delete not pending
        If TDBGrid.Columns("Assembly ID") = "" Then
            m_blnAssemblyID = True
            Cancel = True
        Else
            m_blnAssemblyID = False
        End If
        If Trim(TDBGrid.Columns("Unit Cost ID")) = "" Then
            m_blnUnitCostID = True
            Cancel = True
        Else
            m_blnUnitCostID = False
        End If
        If AssemblyUCSortRequired(Val(TDBGrid.Columns("Parent Skey"))) = True And Len(Trim(TDBGrid.Columns("Sort"))) = 0 Then
            m_blnSortOrder = True
            Cancel = True
        Else
            m_blnSortOrder = False
        End If
    End If
End If

End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)

    If m_blnAssemblyID = True Then
        MsgBox "Please enter a valid Assembly ID", vbExclamation
        m_blnAssemblyID = False
    End If
    If m_blnUnitCostID = True Then
        MsgBox "Please enter a valid Unit Cost ID", vbExclamation
        m_blnUnitCostID = False
    End If
    If m_blnSortOrder = True Then
        MsgBox "This assembly has an assembly book system line: The sort order is required.", vbExclamation
        m_blnSortOrder = False
    End If
    
    Response = 0
    TDBGrid.SetFocus

End Sub

Private Sub TDBGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    '27 = escape
    If KeyCode = 27 Then
        If TDBGrid.AddNewMode = dbgAddNewPending Then      'addnew pending
            TDBGrid.Delete
            TDBGrid.ReBind
        End If
    End If

End Sub

Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)
Dim rec As New ADODB.RecordSet ' Recordset to hold query results
Dim strSELECT As String
Dim blnReturn As Boolean
Dim i As Integer
Dim j As Integer

On Error Resume Next

Select Case MsgId
Case 1
    strSELECT = "Select  unit_cost_skey, tech_desc, metric_tech_desc from unit_cost_detail where unit_cost_id='" + Compress_String(TDBGrid.Columns("Unit Cost ID").Text) + "'"
     ' Use DAL to perform select

     blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
    m_rec.Bookmark = TDBGrid.Bookmark
    
    For i = 0 To rec.Fields.Count - 1
        m_rec.Fields(rec.Fields(i).Name).Value = rec.Fields(i).Value
    Next i
Case 2
    strSELECT = "Select  assembly_skey, tech_desc as parent_tech_description, metric_tech_desc as parent_met_tech_desc " + _
    "book_desc as parent_book_desc, metric_book_desc as parent_metric_book_desc from unit_cost_detail where unit_cost_id='" + Compress_String(TDBGrid.Columns("Unit Cost ID").Text) + "'"
     ' Use DAL to perform select
     blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
    m_rec.Bookmark = TDBGrid.Bookmark
    
    For i = 0 To rec.Fields.Count - 1
        m_rec.Fields(rec.Fields(i).Name).Value = rec.Fields(i).Value
    Next i
End Select
End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
    Dim colAssemblyID As New Collection
    Dim lngParentSkey As Long
    Dim strParentId As String
    ' Position to the desired record and execute delete
    m_rec.Bookmark = Bookmark
    
    Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
    If Not clsRowInfo Is Nothing Then
    If clsRowInfo.State = STATE_NEW Then
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
    ElseIf clsRowInfo.State <> STATE_NOT_SET _
        And clsRowInfo.State <> STATE_DELETED Then
            lngParentSkey = m_rec.Fields("parent_skey")
            strParentId = m_rec.Fields("parent_id")
            ' Build SQL statement
            strUpdate = "exec sp_delete_unit_cost_usage "
            strUpdate = strUpdate + "@parent_skey=" + str(lngParentSkey) + ","
            strUpdate = strUpdate + "@skey_type='A',"
            strUpdate = strUpdate + " @unit_cost_skey=" + str(m_rec.Fields("unit_cost_skey"))
            strUpdate = strUpdate + ", @sort_order='" + m_rec.Fields("sort_order") + "'"
            blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
            If Not blnReturn Then
                m_objGridMap.SetError m_rec.Bookmark, strError
                m_rec.Fields("Error").Value = "Error"
                Bookmark = Null
            Else
                TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
'                If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
                    SaveAssemblyID colAssemblyID, strParentId
'                End If
            End If
        End If
    End If
    
    TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
    m_rec.Delete ' Remove it from recordset so it doesn't show.
    MsgBox "Delete successful."
'    If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
        UpdateAssembly colAssemblyID
'    End If
    
    Exit Sub
   
NoDelete:
    MsgBox Err.Description
    Bookmark = Null
    
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim i As Integer
    Dim strError As String
    
    If TDBGrid.DataChanged = True Then
        strError = Empty
    ' Only worry about first two columns here
        If TDBGrid.Columns(ColIndex).Caption = "Assembly ID" Then
            strError = AsblyUCGridError_AssemblyID(Compress_String(TDBGrid.Text), Compress_String(TDBGrid.Columns("Unit Cost ID")), MasterFormat)
        ElseIf TDBGrid.Columns(ColIndex).Caption = "Unit Cost ID" Then
            strError = AsblyUCGridError_UnitCostID(Compress_String(TDBGrid.Text), Compress_String(TDBGrid.Columns("Assembly ID")), , MasterFormat)
            'TDBGrid.Columns(ColIndex + 1).Caption = TDBGrid.Text    'ccd 8.4 (rlh)
        ElseIf m_rec.Fields(TDBGrid.Columns(ColIndex).DataField).Type = adNumeric Then
            If Not IsNumeric(TDBGrid.Text) Then
                strError = TDBGrid.Columns(ColIndex).Caption + " must be numeric."
            End If
        End If
        
        If strError <> Empty Then
            MsgBox strError, vbInformation
            Cancel = True ' Cancel so old value is replaced
        End If
    End If

End Sub

Public Function CloneRow() As ADODB.RecordSet
    ' Pass to base class
    m_objGridMap.CloneRow (True)
    Set CloneRow = Nothing
'    TDBGrid.Columns("Parent Skey").Value = 0
'    TDBGrid.Columns("Unit Cost Skey").Value = 0
'    TDBGrid.Columns("Assembly ID").Value = TDBGrid.Columns("Assembly ID").Value
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

Public Sub SetError(intBookmark As Integer, strError As String)
    On Error Resume Next
    Dim objRowInfo As CRowInfo
    Set objRowInfo = Nothing
    Set objRowInfo = m_objGridMap.rowinfo(intBookmark)
    If Not objRowInfo Is Nothing Then
        objRowInfo.Error = True
        objRowInfo.ErrorText = strError
    End If
End Sub

Public Function IsPendingChange() As Boolean
    IsPendingChange = m_objGridMap.IsPendingChange
End Function

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub


