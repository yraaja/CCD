VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLaborRateMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Labor grid
Option Explicit

''' <modulename> CLaborRateMap.cls</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Helper class module for frmLaborRateGrid.frm
''' </summary>
''' <seealso> frmLaborRateGrid.frm </seealso>
''' <datastruct> TDBGRID</datastruct>
''' <datastruct> m_objGridMap</datastruct>
'''
''' <storedprocedurename> sp_update_labor_rate</storedprocedurename>
''' <storedprocedurename> sp_delete_labor_rate</storedprocedurename>
'''
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
'''<code>
'''</code>
'''<code>
'''exec sp_update_labor_rate @trade_skey= 8, @start_date='12/01/2005', @term_date='11/29/2011', @union_base_hrly= 64.93, @union_fringe_hrly= 9.99, @estimated_ind= 1, @trade_group_code='BOIL007', @last_update_id= 1, @opn_base_hrly= 0, @opn_fringe_hrly= 0, @davis_bacon_base_hrly= 0, @davis_bacon_fringe_hrly= 0, @davis_bacon_overtime_hrly= 0, @davis_bacon_travel_hrly= 0, @comment='', @local_union_num='146', @contact_id='999999', @loc_id= 95, @last_update_person='Hancockrl'</code>
'''<code>
'''exec sp_delete_labor_rate @trade_skey= 8, @loc_id= 18, @start_date='10/01/2010', @trade_group_code='BOIL001',  @last_update_id=1
'''</code>
'''<code>
'''</code>
''' <code>
'''</code>
'''
'''<code>
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
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

Const m_GRIDTYPE = "Labor" ' Identity of grid
' Names of some table this grid uses
Const LAB_RATE_TABLE = "Labor_Rate"
Const LOCATION_TABLE = "Location"
Const LABOR_TRADE_TABLE = "Labor_Trade"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
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
    m_objGridMap.AddCol "Trade ID", "trade_id", LABOR_TRADE_TABLE, 1000, True
    m_objGridMap.AddCol "Trade Skey", "trade_skey", ALL_TABLES, 1000, False, False
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "State", "state_code", LOCATION_TABLE, 500, True, True, False
    m_objGridMap.AddCol "City", "city", LOCATION_TABLE, 1000, True, True, False
    m_objGridMap.AddCol "Start", "start_date", LAB_RATE_TABLE, 1200, True, True, True, FORMAT_DATE
    m_objGridMap.AddCol "Term", "term_date", LAB_RATE_TABLE, 1200, False, True, True, FORMAT_DATE, True
    m_objGridMap.AddCol "Union Base Hrly", "union_base_hrly", LAB_RATE_TABLE, 1000, , , , FORMAT_PRICE
    m_objGridMap.AddCol "Union Fringe Hrly", "union_fringe_hrly", LAB_RATE_TABLE, 1000, , , , FORMAT_PRICE
    m_objGridMap.AddCol "Total Union", "tot_union", vbNullString, 1000, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Estimated Ind", "estimated_ind", LAB_RATE_TABLE, 1200, False, True, True
    m_objGridMap.AddCol "Trade Group Code", "trade_group_code", LAB_RATE_TABLE, 1200, True, True, True
    m_objGridMap.AddCol "Last Update Date", "last_update_date", vbNullString, 1000, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Last Update Person", "last_update_person", vbNullString, 1000, True, True, False
    m_objGridMap.AddCol "last_update_id", "last_update_id", LAB_RATE_TABLE, 800, True, False, False
    m_objGridMap.AddCol "lt_last_update_id", "lt_last_update_id", vbNullString, 800, True, False, False
    m_objGridMap.AddCol "MaxStartDate", "maxstartdate", "#labordate", 1000, True, False, True
    m_objGridMap.AddCol "Open Base Hrly", "opn_base_hrly", LAB_RATE_TABLE, 1000, , False, , FORMAT_PRICE
    m_objGridMap.AddCol "Open Fringe Hrly", "opn_fringe_hrly", LAB_RATE_TABLE, 1000, , False, , FORMAT_PRICE
    m_objGridMap.AddCol "Total Open", "tot_open", vbNullString, 1000, , False, , FORMAT_PRICE
    m_objGridMap.AddCol "D-B Base hrly", "davis_bacon_base_hrly", LAB_RATE_TABLE, 1000, , False, , FORMAT_PRICE
    m_objGridMap.AddCol "D-B Fringe Hrly", "davis_bacon_fringe_hrly", LAB_RATE_TABLE, 1000, , False, , FORMAT_PRICE
    m_objGridMap.AddCol "D-B Overtime Hrly", "davis_bacon_overtime_hrly", LAB_RATE_TABLE, 1000, False, True, , FORMAT_PRICE
    m_objGridMap.AddCol "D-B Travel Hrly", "davis_bacon_travel_hrly", LAB_RATE_TABLE, 1000, False, True, , FORMAT_PRICE
    m_objGridMap.AddCol "Total D-B", "tot_db", vbNullString, 1000, False, True, , FORMAT_PRICE
    m_objGridMap.AddCol "Comment", "comment", LAB_RATE_TABLE, 1200, False, False, True
    m_objGridMap.AddCol "Local Union Num", "local_union_num", LAB_RATE_TABLE, 1200, False, False, True
    m_objGridMap.AddCol "Trade Desc", "trade_desc", vbNullString, 1000, True, False
    m_objGridMap.AddCol "Worker Type", "worker_type_code", LABOR_TRADE_TABLE, 500, True, True
    m_objGridMap.AddCol "Contact", "contact_id", LAB_RATE_TABLE, 1200, , False
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
    Dim blnUpdateTrade As Boolean
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
                        If objColDef.Table = LAB_RATE_TABLE Then
                        ' flag as changed
                            blnUpdateLabor = True
                        ElseIf objColDef.Table = LABOR_TRADE_TABLE Then
                            blnUpdateTrade = True
                        End If
                     End If
                Next
                If blnUpdateLabor Or blnUpdateTrade Then
                ' If record was added
'                If clsRowInfo.State = STATE_NEW Or clsRowInfo.State = STATE_MODIFIED Then
                If blnUpdateLabor Then
                       strUpdate = "exec sp_update_labor_rate "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
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
                End If
                If blnUpdateTrade Then
                    strUpdate = "exec sp_update_labor_trade "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, LABOR_TRADE_TABLE
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    strUpdate = strUpdate + ", @last_update_id=" + CStr(TDBGrid.Columns("lt_last_update_id").Value)
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If blnReturn = True Then
                        m_rec.Fields("lt_last_update_id").Value = m_rec.Fields("lt_last_update_id").Value + 1
                        If blnUpdateLabor = False Then  'not required if done above
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("last_update_date").Value = Now
                            m_rec.Fields("last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        End If
                    Else
                        If blnUpdateLabor = False Then  'not required if done above
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    End If
                End If
    '                ' If the row was marked as changed but the user typed the same data
    '                ' make sure the error condition is removed from the row
                    If (Not blnUpdateLabor) And (Not blnUpdateTrade) Then
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
    m_objGridMap.m_colCloneCols.Add 1, "worker_type_code"
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
    m_objGridMap.m_colCloneCols.Add 1, "davis_bacon_overtime_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "davis_bacon_travel_hrly"
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
    
If m_rec.Fields("MaxStartDate") <> m_rec.Fields("start_date") Then
    m_objGridMap.SetError m_rec.Bookmark, "Only the active price for a labor rate record may be deleted."
    Exit Sub
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
    Exit Sub

NoDelete:
    MsgBox Err.Description
    Bookmark = Null
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim blnError As Boolean
Dim strErrorMessage As String

On Error Resume Next
    
    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Trade Group Code"     'Term Date
        If Len(TDBGrid.Text) > 0 And Len(TDBGrid.Text) < 7 Then
                strErrorMessage = "The Trade Group must be 7 characters long."
                blnError = True
        End If
    Case "Term"     'Term Date
        If IsDate(TDBGrid.Text) Then
            Dim intDateDiff As Integer
            intDateDiff = DateDiff("d", TDBGrid.Columns("Start"), TDBGrid.Text)
            If intDateDiff <= 0 Then
                strErrorMessage = "The Term Date must be after the Start Date."
                blnError = True
            End If
        Else
            blnError = True
            strErrorMessage = "Please enter a valid Term Date."
        End If
    Case "Union Base Hrly", "Union Fringe Hrly", "Open Base Hrly", "Open Fringe Hrly", "D-B Base hrly", "D-B Fringe Hrly", "D-B Overtime Hrly", "D-B Travel Hrly"
        If Len(Trim(TDBGrid.Text)) > 0 Then 'Only validate if data entered.
            If IsNumeric(TDBGrid.Text) = False Then
                blnError = True
                strErrorMessage = "Please enter a numeric amount for " + TDBGrid.Columns(ColIndex).Caption + "."
            ElseIf CDbl(TDBGrid.Text) < 0 Then
                blnError = True
                strErrorMessage = "The amount for " + TDBGrid.Columns(ColIndex).Caption + " must be 0 or greater."
            End If
        End If
    End Select

    If blnError = True Then
        MsgBox strErrorMessage
        Cancel = True
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
    TDBGrid.Columns("mat_last_update_id") = 0
    TDBGrid.Columns("matprice_last_update_id") = 0
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

Public Sub Factor(dblFactor As Double, strComment As String, intColumns As Integer)
    Dim vntBookmark As Variant
    
    For Each vntBookmark In TDBGrid.SelBookmarks
        m_rec.Bookmark = vntBookmark
        If intColumns And 1 Then
            m_rec.Fields("List_Price") = m_rec.Fields("List_Price") + m_rec.Fields("List_Price") * dblFactor / 100
            ' TRACES should always match list price
            m_rec.Fields("Traces_List_Price") = m_rec.Fields("Traces_List_Price") + m_rec.Fields("Traces_List_Price") * dblFactor / 100
            m_rec.Fields("Factor_ind") = -1 ' This is only for List Price, not TRACES
        ElseIf intColumns And 2 Then
            m_rec.Fields("Traces_List_Price") = m_rec.Fields("Traces_List_Price") + m_rec.Fields("Traces_List_Price") * dblFactor / 100
        End If
        m_rec.Fields("latest_price_update_comment") = strComment
        m_objGridMap.SetRowState Int(vntBookmark), STATE_MODIFIED
        'TDBGrid.Columns("Price").CellText(vntBookmark) = TDBGrid.Columns("Price").CellText(vntBookmark) * dblFactor
    Next
    vntBookmark = TDBGrid.SelBookmarks(0)
    TDBGrid.ReBind ' Reset grid contents
    TDBGrid.Bookmark = vntBookmark ' Set bookmark back again
End Sub

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub

Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub

