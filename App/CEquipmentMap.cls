VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEquipmentMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Equipment grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "Equipment" ' Identity of grid
' Names of table used by the grid
Const EQUIP_TABLE = "Equipment"
Const EQUIP_UPDATE = "Equipment update"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' recordset that populates grid
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
    TDBGrid.AllowAddNew = False
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
'    m_objGridMap.SetFrozenCols (2)
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "Type", "type_code", EQUIP_TABLE, 500, True
    m_objGridMap.AddCol "Equipment ID", "equip_id", EQUIP_TABLE, 1500, True
    m_objGridMap.AddCol "Alt Equip ID", "alt_equip_id", EQUIP_TABLE, 1500, False, False, True
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Tech Desc", "tech_desc", EQUIP_TABLE, 2000
    m_objGridMap.AddCol "Book Desc", "book_desc", EQUIP_TABLE, 2000
    m_objGridMap.AddCol "Crew Equip Desc", "crew_equip_desc", EQUIP_TABLE, 2000
    m_objGridMap.AddCol "Crew Equip Desc Pl", "crew_equip_desc_plural", EQUIP_TABLE, 2000
    m_objGridMap.AddCol "Metric Tech Desc", "metric_tech_desc", EQUIP_TABLE, 2000, , False
    m_objGridMap.AddCol "Metric Book Desc", "metric_book_desc", EQUIP_TABLE, 2000, , False
    m_objGridMap.AddCol "Metric Crew Equip Desc", "metric_crew_equip_desc", EQUIP_TABLE, 2000, , False
    m_objGridMap.AddCol "Metric Crew Equip Desc Pl", "metric_crew_equip_desc_plural", EQUIP_TABLE, 2000, , False
    m_objGridMap.AddCol "Index Desc", "index_desc", EQUIP_TABLE, 2000
    m_objGridMap.AddCol "Index Code", "index_code", EQUIP_TABLE, 1000
    m_objGridMap.AddCol "Unit", "unit", EQUIP_TABLE, 1000, False, True, True
    m_objGridMap.AddCol "Metric Unit", "metric_unit", EQUIP_TABLE, 1000, False, True, True
    m_objGridMap.AddCol "Model Name", "model_name", EQUIP_TABLE, 1500
    m_objGridMap.AddCol "TRACES", "traces_ind", EQUIP_TABLE, 1000
    ' Split here
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Indent Code", "indent_code", EQUIP_TABLE, 1000
    m_objGridMap.AddCol "Format Chars", "format_characters", EQUIP_TABLE, 1000
    m_objGridMap.AddCol "Format Code", "format_code", EQUIP_TABLE, 1000
    m_objGridMap.AddCol "Update Date", "equip_last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "equip_last_update_person", vbNullString, 1500, True, True, False
    m_objGridMap.AddCol "equip_skey", "equip_skey", EQUIP_TABLE, 800, True, False
    m_objGridMap.AddCol "equip_last_update_id", "equip_last_update_id", EQUIP_UPDATE, 800, True, False, False
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
               
    If Not m_blnPreferences Then
        Dim bRet As Boolean
        Dim rec As ADODB.RecordSet
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid60.ValueItem
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid70.ValueItem
        Dim Item As New TrueOleDBGrid80.ValueItem
        
        bRet = g_objDAL.GetRecordset(vbNullString, "select unit from unit_of_measure", rec)
        If Not bRet Then
            Exit Sub
        End If
        
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
        rec.Close
        Set rec = Nothing
    End If
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim bRet As Boolean
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    
    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0
    
    m_rec.MoveFirst
    While Not m_rec.EOF
        strUpdate = ""
        ' Skip the record if it didn't change
        Dim clsRowInfo As CRowInfo
        Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
        If Not clsRowInfo Is Nothing Then
            If Not clsRowInfo.State = STATE_NONE Then
                ' If the record was added
                If clsRowInfo.State = STATE_NEW Then
                    ' Could be new to recordset, but really an existing record
                    ' Happens when user types equip_id into new grid row
                    If m_rec.Fields("equip_skey") = "" Then
                        strUpdate = "exec sp_insert_equipment "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                    Else
                        strUpdate = "exec sp_update_equipment "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                    End If
                ' If the record was changed
                ElseIf clsRowInfo.State = STATE_MODIFIED Then
                    strUpdate = "exec sp_update_equipment "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                End If
                
                ' Make sure there is something to do
                If Not strUpdate = "" Then
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    
                    bRet = False
                    strError = ""
                    bRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If bRet = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                        m_rec.Fields("equip_last_update_date").Value = Now
                        m_rec.Fields("equip_last_update_person").Value = strUserName
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
    
'    m_rec.UpdateBatch
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
    MsgBox "Error in cmdUpdate_Click"
End Function

Private Sub Class_Initialize()
    m_objGridMap.m_colCloneCols.Add 1, "equip_skey"
    m_objGridMap.m_colCloneCols.Add 1, "equip_id"
    m_objGridMap.m_colCloneCols.Add 1, "alt_equip_id"
    m_objGridMap.m_colCloneCols.Add 1, "type_code"
    m_objGridMap.m_colCloneCols.Add 1, "book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "unit"
    m_objGridMap.m_colCloneCols.Add 1, "metric_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "crew_equip_desc"
    m_objGridMap.m_colCloneCols.Add 1, "crew_equip_desc_plural"
    m_objGridMap.m_colCloneCols.Add 1, "index_code"
    m_objGridMap.m_colCloneCols.Add 1, "index_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_crew_equip_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_crew_equip_desc_plural"
    m_objGridMap.m_colCloneCols.Add 1, "metric_unit"
    m_objGridMap.m_colCloneCols.Add 1, "model_name"
    m_objGridMap.m_colCloneCols.Add 1, "traces_ind"
    m_objGridMap.m_colCloneCols.Add 1, "indent_code"
    m_objGridMap.m_colCloneCols.Add 1, "format_characters"
    m_objGridMap.m_colCloneCols.Add 1, "format_code"
    m_objGridMap.m_colCloneCols.Add 1, "graphic_ref_id"
    m_objGridMap.m_colCloneCols.Add 1, "table_ref_id"
    m_objGridMap.m_colCloneCols.Add 1, "table_ref_col"
    m_objGridMap.m_colCloneCols.Add 1, "change_notice_code"
    m_objGridMap.m_colCloneCols.Add 1, "equip_last_update_id"
End Sub

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    ' Conversion Factor
'    If ColIndex = 6 Then
'        If IsNumeric(TDBGrid.Text) = False Then
'            MsgBox TDBGrid.Columns(ColIndex).Caption + " must be a number."
'            Cancel = True
'            Exit Sub
'        End If
'    End If

End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
  
    ' Position to the desired row and execute delete
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
    
    strUpdate = "exec sp_delete_equipment "
    strUpdate = strUpdate + "@equip_skey=" + str(m_rec.Fields("equip_skey")) + ", " 'Right(Str(m_rec.Fields("mat_skey")), Len(Str(m_rec.Fields("mat_skey"))) - 1) + ","
    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
    
    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetError m_rec.Bookmark, strError
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
        MsgBox "Delete successful."
    End If
    Exit Sub
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("Factor").Value = False
    TDBGrid.Columns("Estimated").Value = False
    TDBGrid.Columns("equip_last_update_id").Value = 0
End Sub

Public Sub CloneRow()
    ' Pass to base class
    m_objGridMap.CloneRow
End Sub

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


