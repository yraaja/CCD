VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTradeHistMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Labor grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "Trade History" ' Identity of grid
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
'*** APEX Migration Utility Code Change ***
'        Dim ActiveRow As TrueOleDBGrid60.Style
'*** APEX Migration Utility Code Change ***
'        Dim ActiveRow As TrueOleDBGrid70.Style
        Dim ActiveRow As TrueOleDBGrid80.Style

    ' Add the columns to the grid
    m_objGridMap.AddCol "Trade Group Code", "trade_group_code", LAB_RATE_TABLE, 1200, True, True, True
    m_objGridMap.AddCol "Trade Skey", "trade_skey", LAB_RATE_TABLE, 1000, True, False
    m_objGridMap.AddSplit
    
    m_objGridMap.AddCol "Worker Type", "worker_type_code", vbNullString, 500, True, False
    m_objGridMap.AddCol "Start", "start_date", LAB_RATE_TABLE, 1200, True, True, True, FORMAT_DATE, True
    m_objGridMap.AddCol "Term", "term_date", vbNullString, 1200, True, True, True, FORMAT_DATE
    m_objGridMap.AddCol "Union Base Hrly", "union_base_hrly", vbNullString, 1000, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Union Fringe Hrly", "union_fringe_hrly", vbNullString, 1000, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Total Union", "tot_union", vbNullString, 1000, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Estimated Ind", "estimated_ind", vbNullString, 1200, True, True, True
    m_objGridMap.AddCol "Local Union Num", "local_union_num", vbNullString, 1200, True, False, True
    m_objGridMap.AddCol "Comment", "comment", vbNullString, 1200, True, False, True
    m_objGridMap.AddCol "Last Update Date", "last_update_date", vbNullString, 1000, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Last Update Person", "last_update_person", vbNullString, 1000, True, True, False
    m_objGridMap.AddCol "last_update_id", "last_update_id", LAB_RATE_TABLE, 800, True, False, False
    
    m_objGridMap.AddCol "MaxStartDate", "maxstartdate", vbNullString, 1000, True, False, True
    
    m_objGridMap.AddCol "Open Base Hrly", "opn_base_hrly", vbNullString, 1000, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Open Fringe Hrly", "opn_fringe_hrly", vbNullString, 1000, True, False
    m_objGridMap.AddCol "D-B Base hrly", "davis_bacon_base_hrly", vbNullString, 1000, True, False
    m_objGridMap.AddCol "D-B Fringe Hrly", "davis_bacon_fringe_hrly", vbNullString, 1000, True, False
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
                    strUpdate = "exec sp_change_trade_group "
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

