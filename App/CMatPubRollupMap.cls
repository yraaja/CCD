VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMatRollupMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Material Publ. Rollup grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "MatPubRollup" ' Identity of grid
' Names of some table this grid uses
Const MAT_TABLE = "Material"
Const ALL_TABLES = "All"
Const MAT_PUB_ROLLUP = "Published_Material_Rollup"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private lng_mat_skey As Long
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
    m_objGridMap.AddCol "Material ID", "mat_id", MAT_PUB_ROLLUP, 1500, True, True, False
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "update_ind", "update_ind", MAT_PUB_ROLLUP, 1000, False, True, False
    m_objGridMap.AddCol "Process Date", "processed_date", MAT_PUB_ROLLUP, 1700, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "last_update_date", "last_update_date", MAT_PUB_ROLLUP, 1700, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "last_update_person", "last_update_person", MAT_PUB_ROLLUP, 1700, True, True, False
    m_objGridMap.AddCol "Mat_Skey", "mat_skey", MAT_PUB_ROLLUP, 860, True, True, True
    m_objGridMap.AddCol "last_update_id", "last_update_id", MAT_PUB_ROLLUP, 1200, True, True, True
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
End Sub

Public Function Update() As Boolean
    On Error Resume Next
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateMat As Boolean
    Dim blnUpdateMatRollup As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    
    m_rec.MoveFirst
    While Not m_rec.EOF
        ' Skip the record if it didn't change
        Dim clsRowInfo As CRowInfo
        Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
        If Not clsRowInfo Is Nothing Then
        
            If Not clsRowInfo.State = STATE_NONE Then
                If clsRowInfo.State = STATE_NEW Then             ' If record was added
                        strUpdate = "exec sp_insert_mat_pub_rollup "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MAT_PUB_ROLLUP
                        'm_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = RTrim(strUpdate)
                        If right(strUpdate, 1) = "," Then
                            strUpdate = left(strUpdate, Len(strUpdate) - 1)
                        End If
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
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
                End If  ' clsRowInfo.State = STATE_NEW
                
                If clsRowInfo.State = STATE_MODIFIED Then                 ' If the record was changed
                        strUpdate = "exec sp_update_mat_pub_rollup "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MAT_PUB_ROLLUP
                        'm_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        'strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        strUpdate = RTrim(strUpdate)
                        If right(strUpdate, 1) = "," Then
                            strUpdate = left(strUpdate, Len(strUpdate) - 1)
                        End If
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        clsRowInfo.State = STATE_NONE
                        If blnReturn = True Then
                        ' Reset on success
                            intSuccess = intSuccess + 1
                            m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                            m_rec.Fields("last_update_date").Value = Now
                            m_rec.Fields("last_update_person").Value = strUserName
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                End If  ' clsRowInfo.State = STATE_MODIFIED
                
            End If  'clsRowInfo.State = STATE_NONE
        
        End If  ' Not clsRowInfo Is Nothing
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
    m_objGridMap.m_colCloneCols.Add 1, "mat_skey"
    m_objGridMap.m_colCloneCols.Add 1, "mat_id"
   ' m_objGridMap.m_colCloneCols.Add 1, "alt_mat_id"
    m_objGridMap.m_colCloneCols.Add 1, "update_ind"
    m_objGridMap.m_colCloneCols.Add 1, "processed_date"
    m_objGridMap.m_colCloneCols.Add 1, "update_date"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_person"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_id"
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
    strUpdate = "exec sp_delete_mat_pub_rollup "
    strUpdate = strUpdate + "@mat_skey=" + str(m_rec.Fields("mat_skey"))
    
    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        MsgBox strError
        m_rec.Fields("Error").Value = "Error"
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
        MsgBox "Delete successful."
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
    'TDBGrid.Update
    Dim vntMyBookmark As Variant
    ' Only worry about first column here
    vntMyBookmark = m_rec.Bookmark
    
    If ColIndex = 0 Then
        ' Can only edit new rows and only once
        If OldValue = Empty Then
            Dim strSelect As String
            Dim blnReturn As Boolean
            Dim i As Integer
            ' Check to see if the mat_id entered exists already
            strSelect = "Select * from Published_Material_Rollup where mat_id='" + TDBGrid.Columns(ColIndex).Value + "'"
            ' Use DAL to perform select
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
            ' If it does, copy that data into grid
            If blnReturn = False Then
                MsgBox "An error occurred retrieving the  Published Material Rollup."
            Else
                If rec.RecordCount > 0 Then
                    MsgBox "The material id already exists.  Please enter a valid material ID."
                    Cancel = True ' Cancel this event so grid doesn't do any more processing
                Else
                    ' Check to see if the mat_id entered exists already
                    strSelect = "Select mat_skey from Material where mat_id='" + TDBGrid.Columns(ColIndex).Value + "'"
                    ' Use DAL to perform select
                    rec.Close
                    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
                    ' If it does, copy that data into grid
                    If blnReturn = False Then
                        MsgBox "An error occurred retrieving the Material record."
                    Else
                        If rec.BOF And rec.EOF Then
                            MsgBox "The material id does not exist.  Please enter a valid material ID."
                            Cancel = True ' Cancel this event so grid doesn't do any more processing
                        Else
                            If rec![mat_skey] > 0 Then
                                m_rec.AddNew ' Add new record to recordset that populates grid
                                m_rec![mat_skey] = rec![mat_skey]
                                m_rec![last_update_id] = 1  '?????
                                m_rec![processed_date] = Now()
                                m_rec![mat_id] = TDBGrid.Columns(ColIndex).Value
                                ' Defaults for new added row
                                m_rec![update_ind] = False
                                m_rec![last_update_id] = 0
                                m_rec![last_update_date] = Now()
                                m_rec![last_update_person] = strUserName
                                m_rec.Update
                                ' Make a new RowInfo
                                m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
                            End If
                        End If
                    End If
                End If
            End If
        Else
            MsgBox "You cannot edit this column for an existing row."
            Cancel = True
        End If
    End If
'    m_rec.Bookmark = vntMyBookmark
End Sub

'Public Function CloneRow() As ADODB.RecordSet
'    ' Pass to base class
'    Set CloneRow = m_objGridMap.CloneRowRecordset
'
'End Function

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

'Public Sub Factor(dblFactor As Double, strComment As String, intColumns As Integer)
'    Dim vntBookmark As Variant
'
'    For Each vntBookmark In TDBGrid.SelBookmarks
'        m_rec.Bookmark = vntBookmark
'        If intColumns And 1 Then
'            m_rec.Fields("List_Price") = m_rec.Fields("List_Price") + m_rec.Fields("List_Price") * dblFactor / 100
'            ' TRACES should always match list price
'            m_rec.Fields("Traces_List_Price") = m_rec.Fields("Traces_List_Price") + m_rec.Fields("Traces_List_Price") * dblFactor / 100
'            m_rec.Fields("Factor_ind") = -1 ' This is only for List Price, not TRACES
'        ElseIf intColumns And 2 Then
'            m_rec.Fields("Traces_List_Price") = m_rec.Fields("Traces_List_Price") + m_rec.Fields("Traces_List_Price") * dblFactor / 100
'        End If
'        m_rec.Fields("latest_price_update_comment") = strComment
'        m_objGridMap.SetRowState Int(vntBookmark), STATE_MODIFIED
'        'TDBGrid.Columns("Price").CellText(vntBookmark) = TDBGrid.Columns("Price").CellText(vntBookmark) * dblFactor
'    Next
'    vntBookmark = TDBGrid.SelBookmarks(0)
'    TDBGrid.ReBind ' Reset grid contents
'    TDBGrid.Bookmark = vntBookmark ' Set bookmark back again
'End Sub

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub

Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub

Private Function MaterialIdValidate(mRow As Long) As Boolean
    Dim rec As ADODB.RecordSet, oldRow As Long
    With TDBGrid
        oldRow = .Row
        .Row = mRow
        If Not Len(.Text) = 0 Then
            g_objDAL.GetRecordset vbNullString, "Select count(mat_id) from material where mat_id = '" + .Text + "'", rec
            If rec.Fields(0).Value = 0 Then
                MaterialIdValidate = False
            Else
                rec.Close
                g_objDAL.GetRecordset vbNullString, "Select mat_skey from material where mat_id = '" + .Text + "'", rec
                TDBGrid.Columns("Mat_Skey").Text = rec.Fields("mat_skey")
            End If
            rec.Close
            MaterialIdValidate = True
        End If
        .Row = oldRow
    End With
End Function

