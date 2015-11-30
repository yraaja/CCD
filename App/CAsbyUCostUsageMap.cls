VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAsUCUsageMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Unit Cost Usage grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "AsblyUnitCostUsage" ' Identity of grid
' Names of some table this grid uses
Const ASSEMBLYDETAIL_TABLE = "Assembly_detail"
Const UNITCOSTUSAGE_TABLE = "Assembly_unit_cost_usage"
Const PUB_ASSEMBLY_COSTWORKS_TABLE = "published_assembly_costworks"
Const PUB_UC_COSTWORKS_TABLE = "published_unit_cost_costworks"
Const UNITCOSTDTL_TABLE = "unit_cost_detail"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private bln_AssemblyID As Boolean
Private bln_UnitCostID As Boolean
Private m_lngAssemblySKey As Long
Private m_strAssemblyID As String
Private m_strAssemblyType As String

Private m_blnRecords_Deleted As Boolean
Private m_blnSortOrder As Boolean

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Function SetRowStateNew() As Boolean
    If m_rec.RecordCount > 0 Then
        m_rec.MoveFirst
        While Not m_rec.EOF
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
            m_rec.MoveNext
        Wend
    End If

End Function

Public Function SetRowStateNone() As Boolean
    If m_rec.RecordCount > 0 Then
        m_rec.MoveFirst
        While Not m_rec.EOF
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
            m_rec.MoveNext
        Wend
    End If

End Function

Public Property Let AssemblySKey(intAssemblySKey As Long)
    m_lngAssemblySKey = intAssemblySKey
    
End Property

Public Property Let AssemblyID(strAssemblyId As String)
    m_strAssemblyID = strAssemblyId
End Property

Public Property Let AssemblyType(strTypeCode As String)
    m_strAssemblyType = strTypeCode
End Property

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
Dim bRet As Boolean
Dim rec As ADODB.RecordSet
'*** APEX Migration Utility Code Change ***
'Dim Item As New TrueOleDBGrid60.ValueItem
'*** APEX Migration Utility Code Change ***
'Dim Item As New TrueOleDBGrid70.ValueItem
Dim Item As New TrueOleDBGrid80.ValueItem

    m_objGridMap.GRIDTYPE = m_GRIDTYPE
    If MASTER_FORMAT_ASSEMBLIES = 1995 Then
        m_objGridMap.AddCol "Unit Cost ID", "unit_cost_id", vbNullString, 1500, , , , FORMAT_UNIT_COST
        m_objGridMap.AddCol "Unit Cost ID 04", "ext_unit_cost_id", vbNullString, 1500, True, , , FORMAT_UNIT_COST_04
    Else  '2004
        m_objGridMap.AddCol "Unit Cost ID", "unit_cost_id", vbNullString, 1500, True, , , FORMAT_UNIT_COST
        m_objGridMap.AddCol "Unit Cost ID 04", "ext_unit_cost_id", vbNullString, 1500, False, , , FORMAT_UNIT_COST_04
    End If
    m_objGridMap.AddCol "Sort", "sort_order", UNITCOSTUSAGE_TABLE, 450
    m_objGridMap.AddCol "UC Tech Desc", "tech_desc", vbNullString, 2000, True
    m_objGridMap.AddCol "UC Met Desc", "metric_tech_desc", vbNullString, 2000, True, False
    m_objGridMap.AddCol "Unit", "usage_unit", UNITCOSTUSAGE_TABLE, 1000, False, True, True, FORMAT_STRING_TRIM
    m_objGridMap.AddCol "Unit Qty", "usage_unit_qty", UNITCOSTUSAGE_TABLE, 1000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Adj Factor", "adj_factor", UNITCOSTUSAGE_TABLE, 1000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Metric Unit", "usage_metric_unit", UNITCOSTUSAGE_TABLE, 1000, True, False
    m_objGridMap.AddCol "Metric Unit Qty", "usage_metric_unit_qty", UNITCOSTUSAGE_TABLE, 1000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Metric Adj Factor", "metric_adj_factor", UNITCOSTUSAGE_TABLE, 1000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Override Book Description", "override_book_desc", UNITCOSTUSAGE_TABLE, 2000, False, True, True
    m_objGridMap.AddCol "Metric Override Book Description", "override_metric_book_desc", UNITCOSTUSAGE_TABLE, 2000, False, True, True
    m_objGridMap.AddCol "Override Book Qty", "override_book_qty", UNITCOSTUSAGE_TABLE, 2000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Metric Override Book Qty", "override_metric_book_qty", UNITCOSTUSAGE_TABLE, 2000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Callout ID", "call_out_id", UNITCOSTUSAGE_TABLE, 2000
    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
    m_objGridMap.AddCol "assembly_skey", "parent_skey", UNITCOSTUSAGE_TABLE, 800, True, False
    m_objGridMap.AddCol "skey_type", "skey_type", UNITCOSTUSAGE_TABLE, 800, True, False
    m_objGridMap.AddCol "last_update_id", "last_update_id", UNITCOSTUSAGE_TABLE, 800, True, False, False
    m_objGridMap.AddCol "unit_cost_skey", "unit_cost_skey", UNITCOSTUSAGE_TABLE, 800, False, False
    
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
    
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

End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    Dim colRowInfo As Collection
    Dim clsRowInfo As CRowInfo
    
    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0
    If m_objGridMap.IsPendingChange = True Then     'May just require deletes
        m_rec.MoveFirst
        While Not m_rec.EOF
            ' Skip the record if it didn't change
            Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
            If Not clsRowInfo Is Nothing Then
                If Not clsRowInfo.State = STATE_NONE Then
                ' If record was added, or if both assembly_id and unit_cost_id were changed
                If clsRowInfo.State = STATE_NEW Or (m_rec.Fields("parent_skey").Value <> m_rec.Fields("parent_skey").OriginalValue And m_rec.Fields("unit_cost_id").Value <> m_rec.Fields("unit_cost_id").OriginalValue) Then
                    strUpdate = "exec sp_insert_unit_cost_usage "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, UNITCOSTUSAGE_TABLE
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If blnReturn = True Then
                        intSuccess = intSuccess + 1
                        m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                        m_rec.Fields("last_update_date").Value = Now
                        m_rec.Fields("last_update_person").Value = strUserName
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
                    If IsNull(m_rec.Fields("sort_order").OriginalValue) Or Trim(m_rec.Fields("sort_order").OriginalValue) = "" Then
                        strUpdate = strUpdate + "@old_sort_order=''"
                    Else
                        strUpdate = strUpdate + "@old_sort_order='" + Trim(m_rec.Fields("sort_order").OriginalValue) + "'"
                    End If
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If blnReturn = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                        m_rec.Fields("last_update_date").Value = Now
                        m_rec.Fields("last_update_person").Value = strUserName
                        Reset_Orig_Values m_rec
                        m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
                    Else
                        intErrors = intErrors + 1
                        clsRowInfo.Error = True
                        clsRowInfo.ErrorText = strError
                    End If
                    
                    ' If the row was marked as changed but the user typed the same data
                    ' make sure the error condition is removed from the row
    '                If Not blnUpdateMatPrice And Not blnUpdateMat Then
    '                    intSuccess = intSuccess + 1
    '                    clsRowInfo.state = STATE_NONE
    '                End If
                End If
                End If
            End If
            m_rec.MoveNext
        Wend
        m_rec.UpdateBatch
    End If
Set colRowInfo = m_objGridMap.RowInfos
'Execute delete requests from rowinfo

For Each clsRowInfo In colRowInfo
    If clsRowInfo.State = STATE_DELETED And clsRowInfo.sKey > 0 Then
        ' Build SQL statement
        strUpdate = "exec sp_delete_unit_cost_usage " + _
            "@parent_skey=" + str(m_lngAssemblySKey) + _
            ", @skey_type = 'A'" + _
            ", @unit_cost_skey=" + str(clsRowInfo.sKey) + _
            ", @sort_order='" + clsRowInfo.SubKey + "'"
               blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
        If Not blnReturn Then
            intErrors = intErrors + 1
            m_objGridMap.SetError clsRowInfo.Bookmark, strError
        Else
            m_objGridMap.SetRowState clsRowInfo.Bookmark, STATE_NONE
        End If
    End If
Next
m_blnRecords_Deleted = False

Set colRowInfo = Nothing

    Dim strMsg As String
    strMsg = ""
'    If intSuccess > 0 Then
'        strMsg = str(intSuccess) + " Unit Cost Usage rows updated successfully." + Chr(10)
'    End If
    If intErrors > 0 Then
        strMsg = strMsg + str(intErrors) + " errors occurred."
        ' Return value will be False
        Update = False
    End If
    If Len(strMsg) > 0 Then
        MsgBox strMsg
    End If
    
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    Exit Function
    
CleanUp:
    MsgBox "Error in Update"
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

Public Function ValidSort() As Boolean

End Function

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    Dim strSELECT As String
    Dim rec As ADODB.RecordSet
    Dim blnReturn  As Boolean
    Dim i As Integer
    Dim strOrigBookmark As Variant
    
    On Error GoTo Errlbl        'rlh
    
    If Not (m_rec.BOF = True Or m_rec.EOF = True) Then
        strOrigBookmark = m_rec.Bookmark
        m_rec.Bookmark = TDBGrid.Bookmark
    End If
'    Stop 'rlh
    Select Case TDBGrid.Columns(ColIndex).Caption
        Case "Unit Cost ID"      'Unit Cost ID
            strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey " + _
                    "from unit_cost_detail where unit_cost_id='" + Compress_String(TDBGrid.Text) + "'"
                     
            
                ' Use DAL to perform select
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
            ' If found, copy that data into grid
            If rec.RecordCount > 0 Then
                ' Copy all values
                TDBGrid.Columns("UC Tech Desc") = rec.Fields("tech_desc").Value
                TDBGrid.Columns("UC Met Desc") = rec.Fields("metric_tech_desc").Value
                TDBGrid.Columns("Unit") = rec.Fields("unit").Value
                TDBGrid.Columns("unit_cost_skey") = rec.Fields("unit_cost_skey").Value
            End If
            
            If MASTER_FORMAT_ASSEMBLIES = 2004 Then
                '######################## rlh ########################
                '####               Get MF2004 ID
                '######################## rlh ########################
                
                '' added this block on 05/06/08
                'strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey from unit_cost_detail where unit_cost_skey = (select unit_cost_skey from unit_cost_detail_ext where unit_cost_id='" + Compress_String(TDBGrid.Text) + " ')"
                strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit_cost_skey from unit_cost_detail_ext where unit_cost_skey =" & TDBGrid.Columns("unit_cost_skey")
                ' Use DAL to perform select
                blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
                ' If found, copy that data into grid
                If rec.RecordCount > 0 Then
                    ' Copy all values
                    TDBGrid.Columns("UC Tech Desc") = rec.Fields("tech_desc").Value
                    TDBGrid.Columns("UC Met Desc") = rec.Fields("metric_tech_desc").Value
                    'TDBGrid.Columns("Unit") = rec.Fields("unit").Value
                    TDBGrid.Columns("unit_cost_skey") = rec.Fields("unit_cost_skey").Value
                End If
            End If
            '' down to here on 05/06/08
            
            If rec.State = adStateOpen Then
                rec.Close
            End If
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
        Case "Unit Cost ID 04"      'Ext Unit Cost ID  (mf04)
            strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey " + _
                    "from unit_cost_detail where unit_cost_id='" + TDBGrid.Columns("Unit Cost ID 04") + "'"
            
            'rlh 02/05/2009
            strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit_cost_skey " + _
                    "from unit_cost_detail_ext where unit_cost_id='" + Compress_String(TDBGrid.Columns("Unit Cost ID 04")) + "'"
            
            ' Use DAL to perform select
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
            ' If found, copy that data into grid
            If rec.RecordCount > 0 Then
                ' Copy all values
                TDBGrid.Columns("UC Tech Desc") = rec.Fields("tech_desc").Value
                TDBGrid.Columns("UC Met Desc") = rec.Fields("metric_tech_desc").Value
                'TDBGrid.Columns("Unit") = rec.Fields("unit").Value
                TDBGrid.Columns("unit_cost_skey") = rec.Fields("unit_cost_skey").Value
            End If
            
            If MASTER_FORMAT_ASSEMBLIES = 2004 Then
                '######################## rlh ########################
                '####               Get MF2004 ID
                '######################## rlh ########################
                
                '' added this block on 05/06/08
                'strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey from unit_cost_detail where unit_cost_skey = (select unit_cost_skey from unit_cost_detail_ext where unit_cost_id='" + Compress_String(TDBGrid.Text) + " ')"
                strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc,  unit_cost_skey from unit_cost_detail where unit_cost_skey =" & TDBGrid.Columns("unit_cost_skey")
                
                'rlh 02/05/2009 - Get "Unit" only!!!  (we already have the others...)
                strSELECT = "Select unit from unit_cost_detail where unit_cost_skey =" & TDBGrid.Columns("unit_cost_skey")
                
                ' Use DAL to perform select
                blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
                ' If found, copy that data into grid
                If rec.RecordCount > 0 Then
                    ' Copy "Unit" value
                    TDBGrid.Columns("Unit") = rec.Fields("unit").Value
                End If
            End If
            '' down to here on 05/06/08
            
            If rec.State = adStateOpen Then
                rec.Close
            End If
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)

    End Select
    If Not (m_rec.BOF = True Or m_rec.EOF = True) Then
        m_rec.Bookmark = strOrigBookmark
    End If
    Exit Sub
Errlbl: 'rlh 05/06/2008
    MsgBox ("(CAsUCUsageMap):TDBGRID_AfterColUpdate: " & Err.Description)
    Resume
    Stop 'rlh
    

End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
Dim strSELECT As String
Dim rec As ADODB.RecordSet
Dim blnReturn  As Boolean
Dim i As Integer
Dim strOrigBookmark As Variant
If TDBGrid.DataChanged = True Then
    If TDBGrid.AddNewMode > 0 Then
        If Screen.ActiveControl.Name <> "cmdDelete" Then 'delete not pending
            If TDBGrid.Columns("unit_cost_skey") = 0 Then
                bln_UnitCostID = True
                Cancel = True
            Else
                bln_UnitCostID = False
            End If
        End If
    End If
    If AssemblyUCSortRequired(m_lngAssemblySKey) = True And Len(Trim(TDBGrid.Columns("Sort"))) = 0 Then
            m_blnSortOrder = True
            Cancel = True
        Else
            m_blnSortOrder = False
    End If
End If
End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
If bln_UnitCostID = True Then
    MsgBox "Please enter a valid Unit Cost ID"
    bln_UnitCostID = False
End If
If m_blnSortOrder = True Then
    MsgBox "This assembly has an assembly book system line:  The sort order is required."
    m_blnSortOrder = False
End If
Response = 0
TDBGrid.SetFocus

End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim rowinfo As CRowInfo
    ' Position to the desired record and execute delete
    m_rec.Bookmark = Bookmark
    Set rowinfo = m_objGridMap.rowinfo(m_rec.Bookmark)

    TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
    If rowinfo.State = STATE_NEW Then
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
    ElseIf rowinfo.State <> STATE_NOT_SET _
        And rowinfo.State <> STATE_DELETED Then
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_objGridMap.SetRowSkey m_rec.Bookmark, m_rec.Fields("unit_cost_skey")
        m_objGridMap.SetRowSubKey m_rec.Bookmark, m_rec.Fields("sort_order")
        m_blnRecords_Deleted = True
    End If
    m_rec.Delete ' Remove it from recordset so it doesn't show.
    Set rowinfo = Nothing
    
    Exit Sub
    
NoDelete:
    MsgBox Err.Description
    Bookmark = Null
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim strUnitCostId As String
Dim lngAssemblySkey As Long
Dim rec As ADODB.RecordSet ' Recordset to hold query results
Dim vntMyBookmark As Variant
Dim strError As String

If TDBGrid.DataChanged = True Then
' Validate required fields:
'    Unit Cost ID
    If TDBGrid.Columns(ColIndex).Caption = "Unit Cost ID" Or TDBGrid.Columns(ColIndex).Caption = "Unit Cost ID 04" Then
    'Stop 'rlh
        If MASTER_FORMAT_ASSEMBLIES = 2004 Then    'rlh 05/13/08
            strError = AsblyUCGridError_UnitCostID(Compress_String(TDBGrid.Text), m_strAssemblyID, _
            m_strAssemblyType, EXT_MASTERFORMAT_VERSION) 'rlh 05/06/2008  changed masterformat forcing to 2004
        End If
        If MASTER_FORMAT_ASSEMBLIES = 1995 Then    'rlh 05/13/08
            strError = AsblyUCGridError_UnitCostID(Compress_String(TDBGrid.Text), m_strAssemblyID, _
            m_strAssemblyType, UCD_MASTERFORMAT_VERSION) 'rlh 05/06/2008  changed masterformat forcing to 2004
        End If
    End If
    If strError <> Empty Then
        MsgBox strError
        Cancel = True ' Cancel so old value is replaced
'    Else
'        strUnitCostId = Trim(TDBGrid.Text)
'        lngAssemblySkey = TDBGrid.Columns("assembly_skey").Value
'        'Check to see if the unit cost has been assigned for this assembly - may not have dups
'        vntMyBookmark = TDBGrid.Bookmark
'        If m_rec.RecordCount > 0 Then
'            m_rec.MoveFirst
'            Do Until m_rec.EOF
'                If Trim(m_rec.Fields("unit_cost_id")) = strUnitCostId Then
'                    If m_rec.Fields("parent_skey") = lngAssemblySkey Then
'                        MsgBox "The Unit Cost ID " + TDBGrid.Text + " has already been assigned to this assembly, and may not be assigned twice."
'                        Cancel = True
'                        TDBGrid.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                m_rec.MoveNext
'            Loop
'        End If
    End If
    If TDBGrid.Columns(ColIndex).Caption = "Unit Qty" _
    Or TDBGrid.Columns(ColIndex).Caption = "Adj Factor" _
    Or TDBGrid.Columns(ColIndex).Caption = "Metric Unit Qty" _
    Or TDBGrid.Columns(ColIndex).Caption = "Metric Adj Factor" _
    Or TDBGrid.Columns(ColIndex).Caption = "Override Book Qty" _
    Or TDBGrid.Columns(ColIndex).Caption = "Metric Override Book Qty" _
    Then
        If Not IsNumeric(TDBGrid.Text) Then
            MsgBox "Please enter a numeric value for " + TDBGrid.Columns(ColIndex).Caption
            Cancel = True ' Cancel so old value is replaced
        ElseIf TDBGrid.Text < 0 Then
            MsgBox "Please enter a valid number for " + TDBGrid.Columns(ColIndex).Caption
            Cancel = True ' Cancel so old value is replaced
        End If
    End If
End If
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("assembly_skey") = m_lngAssemblySKey
    If IsEmpty(TDBGrid.Columns("unit_cost_skey").Value) Then
        TDBGrid.Columns("unit_cost_skey").Value = 0
    End If
    TDBGrid.Columns("Adj Factor") = 1
    TDBGrid.Columns("Unit Qty") = 1
    TDBGrid.Columns("Metric Adj Factor") = 1
    TDBGrid.Columns("Metric Unit Qty") = 1
    TDBGrid.Columns("last_update_id") = 0
    TDBGrid.Columns("skey_type") = "A"
End Sub
Public Function CloneRow() As ADODB.RecordSet
    ' Pass to base class
    m_objGridMap.CloneRow (True)
    Set CloneRow = Nothing
End Function

Public Property Get GRIDTYPE() As String
    GRIDTYPE = m_GRIDTYPE
End Property

Public Property Get RecordsDeleted() As Boolean
    RecordsDeleted = m_blnRecords_Deleted
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
    If m_blnRecords_Deleted = True Then
        IsPendingChange = True
    End If
End Function

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub


