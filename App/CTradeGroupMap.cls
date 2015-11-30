VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTradeGroupMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Labor grid
Option Explicit

''' <modulename> CTradeGroupMap.cls</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Helper class module for frmTradeGroupGrid.frm
'''
'''  More important subs/functions:
'''"   InitGrid()
'''Where "AddCol"s are called to build datagrid columns from the recordset returned by call to stored procedure
'''"   Update()
'''Where stored procedures are prepared to do save of any new or changed data
'''"   RemoveTradeGroupMbr ()
'''User prompted with: This Trade ID/location will no longer participate in group
'''If the user says Yes/OK then the trade id is removed from the indicated group
'''"   TDBGrid_BeforeColUpdate
'''Called when a cell's data has changed as user moves to new row or presses Return
''' </summary>
'''
''' <seealso> frmTradeGroupGrid.frm </seealso>
''' <datastruct> TDBGRID</datastruct>
''' <datastruct> m_objGridMap</datastruct>
'''
''' <storedprocedurename> sp_update_trade_grp_mbr</storedprocedurename>
'''<storedprocedurename>sp_change_trade_group</storedprocedurename>
'''<storedprocedurename>sp_delete_labor_rate</storedprocedurename>
'''
'''
'''<returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
'''<code>
'''exec sp_change_trade_group @trade_skey = 8, @trade_group_code = '', @loc_id = 345, @start_date = '10/1/2010', @last_update_person='Hancockrl', @last_update_id=1
'''</code>
'''<code>
'''exec sp_replace_trade_groups  @trade_group_code = 'BOIL001', @new_trade_group_code = '', @last_update_person='Hancockrl'
'''</code>
'''<code>
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''"   frmTradeGroupGrid.frm
'''"   CGridMap.cls
'''"   CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''"   TDBGrid (ComponentOne DataGrid)
'''"   Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid  ' Access to Grid events
'''</dependson>



'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "Trade Group" ' Identity of grid
' Names of some table this grid uses
Const LAB_RATE_TABLE = "Labor_Rate"
Const LOCATION_TABLE = "Location"
Const LABOR_TRADE_TABLE = "Labor_Trade"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private m_blnWarnings As Boolean

Public Function RemoveTradeGroupMbr() As Boolean
Dim strMsg As String
Dim iResponse As Integer
Dim strUpdate As String
Dim strError As String
Dim blnRet As Boolean
Dim strBookmark As String
    
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Function
    End If
    strBookmark = TDBGrid.Bookmark
    m_rec.Bookmark = TDBGrid.Bookmark
    strMsg = "This Trade ID/location will no longer participate in group " + m_rec.Fields("trade_group_code") + ".  Are you sure you want to remove it?"
    iResponse = MsgBox(strMsg, vbYesNoCancel, "Remove Trade ID From Group")
    
    If iResponse = vbYes Then
        strUpdate = "exec sp_change_trade_group "
        strUpdate = strUpdate + "@trade_skey = " + CStr(m_rec.Fields("trade_skey"))
        strUpdate = strUpdate + ", @trade_group_code = ''"
        strUpdate = strUpdate + ", @loc_id = " + CStr(m_rec.Fields("loc_id"))
        strUpdate = strUpdate + ", @start_date = '" + CStr(m_rec.Fields("start_date")) + "'"
        strUpdate = strUpdate + ", @last_update_person='" + strUserName + "'"
        strUpdate = strUpdate + ", @last_update_id=" + CStr(m_rec.Fields("last_update_id").Value)
    
        blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
        If Not blnRet Then
            MsgBox strError
        Else
            m_blnWarnings = False
            TDBGrid.Columns("Trade Group Code").Value = ""
            TDBGrid.Update
            Dim clsRowInfo As CRowInfo
            Set clsRowInfo = m_objGridMap.rowinfo(CInt(strBookmark))
            If Not clsRowInfo Is Nothing Then
                clsRowInfo.State = STATE_NONE
            End If
            m_blnWarnings = True
            RemoveTradeGroupMbr = True
        End If
    End If

End Function

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub
Public Property Let Warnings(bln As Boolean)
    ' Display warnings before updating a grid row
    m_blnWarnings = bln
End Property


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
'*** APEX Migration Utility Code Change ***
'        Dim ActiveRow As TrueOleDBGrid60.Style
'*** APEX Migration Utility Code Change ***
'        Dim ActiveRow As TrueOleDBGrid70.Style
        Dim ActiveRow As TrueOleDBGrid80.Style

    ' Add the columns to the grid
    m_objGridMap.AddCol "Trade Group Code", "trade_group_code", LAB_RATE_TABLE, 1200, , True, True
    m_objGridMap.AddCol "Trade Skey", "trade_skey", LAB_RATE_TABLE, 1000, True, False
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Trade ID", "trade_id", vbNullString, 1000, True, True, True
    m_objGridMap.AddCol "Trade Desc", "trade_desc", vbNullString, 1000, True, True
    m_objGridMap.AddCol "Worker Type", "type_code", vbNullString, 500, True, False
    'm_objGridMap.AddCol "Traces Type Code", "traces_trade_type_code", vbNullString, 500, False, True, True
    m_objGridMap.AddCol "BIA Trade Desc", "bia_trade_desc", vbNullString, 500, False, True, True
    m_objGridMap.AddCol "City", "city", vbNullString, 1000, True, True, False
    m_objGridMap.AddCol "State", "state_code", vbNullString, 500, True, True, False
    m_objGridMap.AddCol "Start", "start_date", vbNullString, 1200, True, True, True, FORMAT_DATE, True
    m_objGridMap.AddCol "Term", "term_date", vbNullString, 1200, True, True, True, FORMAT_DATE
    m_objGridMap.AddCol "Union Base Hrly", "union_base_hrly", vbNullString, 1000, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Union Fringe Hrly", "union_fringe_hrly", vbNullString, 1000, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Total Union", "tot_union", vbNullString, 1000, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Estimated Ind", "estimated_ind", vbNullString, 1200, True, True, True
    m_objGridMap.AddCol "Local Union Num", "local_union_num", vbNullString, 1200, True, False, True
    m_objGridMap.AddCol "UOM", "uom", vbNullString, 500, False, False, True
    m_objGridMap.AddCol "Comment", "comment", vbNullString, 1200, True, False, True
    m_objGridMap.AddCol "Last Update Date", "last_update_date", vbNullString, 1000, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Last Update Person", "last_update_person", vbNullString, 1000, True, True, False
    m_objGridMap.AddCol "last_update_id", "last_update_id", LAB_RATE_TABLE, 800, True, False, False
    
    m_objGridMap.AddCol "MaxStartDate", "maxstartdate", vbNullString, 1000, True, False, True
    
    m_objGridMap.AddCol "Open Base Hrly", "opn_base_hrly", vbNullString, 1000, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Open Fringe Hrly", "opn_fringe_hrly", vbNullString, 1000, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Total Open", "tot_open", vbNullString, 1000, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "D-B Base hrly", "davis_bacon_base_hrly", vbNullString, 1000, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "D-B Fringe Hrly", "davis_bacon_fringe_hrly", vbNullString, 1000, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "D-B Overtime Hrly", "davis_bacon_overtime_hrly", vbNullString, 1000, True, True, , FORMAT_PRICE
    m_objGridMap.AddCol "D-B Travel Hrly", "davis_bacon_travel_hrly", vbNullString, 1000, True, True, , FORMAT_PRICE
    m_objGridMap.AddCol "Total D-B", "tot_db", vbNullString, 1000, True, True, , FORMAT_PRICE
    m_objGridMap.AddCol "Contact", "contact_id", vbNullString, 1200, True, False
'    m_objGridMap.AddCol "Worker Comp ID", "worker_comp_id", vbNullString, 1200, True, False, True
    m_objGridMap.AddCol "Location ID", "loc_id", LAB_RATE_TABLE, 1000, True, False, True
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    

    End If
End Sub
Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateLabor As Boolean
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
                blnReturn = False
                blnUpdateLabor = False
                ' Loop through the fields to see what has changed
                For Each fld In m_rec.Fields
                    ' If the value changed
                    If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
                        ' See what table the field is from
                        Set objColDef = m_objGridMap.GetCol(fld.Name)
                        ' flag as changed
                        blnUpdateLabor = True
                    End If
                Next
                If blnUpdateLabor Then
                ' If record was added
'                If clsRowInfo.State = STATE_NEW Or clsRowInfo.State = STATE_MODIFIED Then
                    If m_rec.Fields("trade_group_code") = "" Then   'trade group removed
                        strUpdate = "exec sp_change_trade_group @start_date='" + CStr(m_rec.Fields("start_date")) + "', "
                    
                    Else
                        strUpdate = "exec sp_update_trade_grp_mbr @prior_term_date='" + CStr(m_rec.Fields("term_date")) + "', "
                    End If
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, LAB_RATE_TABLE
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If blnReturn = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                        m_rec.Fields("last_update_date").Value = Now
                        m_rec.Fields("last_update_person").Value = strUserName
                        Reset_Orig_Values m_rec
                    Else
                        intErrors = intErrors + 1
                        clsRowInfo.Error = True
                        clsRowInfo.ErrorText = strError
                    End If
    '                ' If the row was marked as changed but the user typed the same data
    '                ' make sure the error condition is removed from the row
                    If Not blnUpdateLabor Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
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
End Function


Private Sub Class_Initialize()

    m_objGridMap.m_colCloneCols.Add 1, "trade_skey"
    m_objGridMap.m_colCloneCols.Add 1, "trade_id"
    m_objGridMap.m_colCloneCols.Add 1, "trade_desc"
    m_objGridMap.m_colCloneCols.Add 1, "type_code"
    m_objGridMap.m_colCloneCols.Add 1, "start_date"
    m_objGridMap.m_colCloneCols.Add 1, "term_date"
    m_objGridMap.m_colCloneCols.Add 1, "contact_id"
'    m_objGridMap.m_colCloneCols.Add 1, "worker_comp_id"
    m_objGridMap.m_colCloneCols.Add 1, "union_base_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "union_fringe_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "opn_base_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "opn_fringe_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "davis_bacon_base_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "davis_bacon_fringe_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "trade_group_code"
    m_objGridMap.m_colCloneCols.Add 1, "estimated_ind"
    m_objGridMap.m_colCloneCols.Add 1, "local_union_num"
    m_objGridMap.m_colCloneCols.Add 1, "comment"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_id"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_date"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_person"
    m_objGridMap.m_colCloneCols.Add 1, "City"
    m_objGridMap.m_colCloneCols.Add 1, "State_Code"
End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
Dim iResponse As Integer
    Dim strSelect As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim rsTemp As New ADODB.RecordSet
    Dim strMsg As String

If m_blnWarnings Then
    m_rec.Bookmark = TDBGrid.Bookmark
    If TDBGrid.Columns("Trade Group Code").Text <> m_rec.Fields("trade_group_code").OriginalValue Then
    'Check for existing labor trade records with the new trade group
        Screen.MousePointer = vbHourglass
        strSelect = "select '1' as Trades from labor_rate where trade_group_code = '" + TDBGrid.Columns("Trade Group Code").Text + "'"
        blnRet = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If Not blnRet Then
            MsgBox "An error occurred retrieving data."
        Else
            If rsTemp.RecordCount = 0 Then
                If Trim(m_rec.Fields("trade_group_code").OriginalValue) = "" Then
                    strMsg = "This trade ID/location will be added to the new group " + TDBGrid.Columns("Trade Group Code").Text + ". Do you wish to continue?"
                Else
                    strMsg = "This trade ID/location will be added to the new group " + TDBGrid.Columns("Trade Group Code").Text + " and removed from " + m_rec.Fields("trade_group_code").OriginalValue + ". Do you wish to continue?"
                End If
            Else
                strMsg = "The trade group will be changed from " + m_rec.Fields("trade_group_code").OriginalValue + " to " + TDBGrid.Columns("Trade Group Code").Text + " for this trade id/location.  Do you wish to continue?"
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing
        Screen.MousePointer = vbNormal
    
        iResponse = MsgBox(strMsg, vbOKCancel)
        If iResponse = vbCancel Then
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub TDBGrid_Change()
Dim iSelectStart As Integer
Dim iSelectLen As Integer

If TDBGrid.Col = TDBGrid.Columns("Trade Group Code").ColIndex Then
    iSelectStart = TDBGrid.SelStart
    iSelectLen = TDBGrid.SelLength
    TDBGrid.Text = UCase(TDBGrid.Text)
    TDBGrid.SelStart = iSelectStart
    TDBGrid.SelLength = iSelectLen
End If

End Sub

Private Sub TDBGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then KeyCode = 0    'Disable delete key
End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
    
    ' Position to the desired record and execute delete
    m_rec.Bookmark = Bookmark
    
    Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
    If Not clsRowInfo Is Nothing Then
        If clsRowInfo.State = STATE_NEW Then
            TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
            m_rec.Delete ' Remove it from recordset so it doesn't show.
            MsgBox "Delete successful."
            Exit Sub
        End If
    End If
    
    ' Build SQL statement
    strUpdate = "exec sp_delete_labor_rate "
    strUpdate = strUpdate + "@trade_skey=" + str(m_rec.Fields("trade_skey")) + ","
    strUpdate = strUpdate + " @loc_id=" + str(m_rec.Fields("loc_id")) + ","
    strUpdate = strUpdate + " @start_date='" + Format(m_rec.Fields("start_date"), "mm/dd/yyyy") + "',"
    strUpdate = strUpdate + " @trade_group_code='" + m_rec.Fields("trade_group_code") + "', "
    strUpdate = strUpdate + " @last_update_id=" + CStr(m_rec.Fields("last_update_id").Value)
    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetError m_rec.Bookmark, strError
        m_rec.Fields("Error").Value = "Error"
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
        MsgBox "Delete successful.  The Labor Rate Grid will now be refreshed."
    End If
'    TDBGrid.ReBind
    Exit Sub

NoDelete:
    MsgBox Err.Description
    Bookmark = Null
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error Resume Next
    Dim rec As ADODB.RecordSet
    Dim iResult As Integer
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim strGroup As String

    strGroup = TDBGrid.Columns("Trade Group Code").Text
    ' Only worry about first column here
    If ColIndex = 0 And Trim(strGroup) <> "" Then
        strSelect = "select max(start_date) as max_start_date from labor_rate where trade_group_code = '" _
        + strGroup + "' and loc_id <> " + CStr(TDBGrid.Columns("Location ID").Text)
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        ' If it does, copy that data into grid
        If rec.RecordCount = 0 Then     'New group
            iResult = MsgBox("The group " + strGroup + " does not exist - do you want to add it?", vbOKCancel)
            If iResult = vbCancel Then
                Cancel = True
            End If
        Else
            'Verify that the start date of the group matches - otherwise, error
            If rec.Fields("max_start_date") <> DateAdd("d", 1, CDate(TDBGrid.Columns("Term").Text)) Then
                MsgBox "The Term Date (" + Format(TDBGrid.Columns("Term").Text, "Short Date") + ") must match the day before the Start Date of the group the Trade ID is being added to (" + Format(rec.Fields("max_start_date"), "Short Date") + ")."
                Cancel = True
            End If
        End If
        rec.Close
        Set rec = Nothing
    End If

End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("Active").Value = True
    TDBGrid.Columns("Use").Value = True
    TDBGrid.Columns("West Use").Value = False
    TDBGrid.Columns("Estimated").Value = False
    TDBGrid.Columns("Factor").Value = False
    TDBGrid.Columns("TRACES").Value = False
    TDBGrid.Columns("Update").Value = False
    TDBGrid.Columns("last_update_id") = 0
End Sub

Public Function CloneRow() As ADODB.RecordSet
    ' Pass to base class
    Set CloneRow = m_objGridMap.CloneRowRecordset
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

Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub

