VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUCostMatMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Material Price grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "UnitCostMaterialUsage" ' Identity of grid
' Names of some table this grid uses
Const MATUSAGE_TABLE = "Material_usage"
Const MATUSAGE_UPDATE = "Material_usage update"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private m_intUnitCostSKey As Long
Private m_blnErrorBlankMatID
Private m_intMat_SKey As Long
Private m_vntNewBookmark As Variant
Private m_strUnitCostID As String
Private m_blnRecords_Deleted As Boolean
Private m_blnGridUpdateErrs As Boolean


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

If m_intUnitCostSKey = 0 Then
    If Not (m_rec.EOF And m_rec.BOF) Then
        m_rec.MoveFirst
        Do Until m_rec.EOF
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
            m_rec.MoveNext
        Loop
        m_rec.MoveFirst
    End If
End If
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

Public Sub InitGrid()
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
'    m_objGridMap.SetFrozenCols (3)
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "Material ID", "mat_id", MATUSAGE_TABLE, 1500, , , , FORMAT_MATERIAL
    m_objGridMap.AddCol "Unit Qty", "unit_qty", MATUSAGE_TABLE, 1000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Input Factor", "input_factor", MATUSAGE_TABLE, 1000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Output Factor", "output_factor", MATUSAGE_TABLE, 1100, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Adj Factor", "adj_factor", MATUSAGE_TABLE, 1000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
    m_objGridMap.AddCol "mat_skey", "mat_skey", MATUSAGE_TABLE, 800, True, False
    m_objGridMap.AddCol "Comment", "comment", MATUSAGE_TABLE, 2000, True, False, False
    m_objGridMap.AddCol "last_update_id", "last_update_id", MATUSAGE_UPDATE, 800, True, False, False
    m_objGridMap.AddCol "unit_cost_skey", "unit_cost_skey", MATUSAGE_TABLE, 800, False, False
    m_objGridMap.BuildGrid
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

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    Dim colRowInfo As Collection
   
    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0
    If TDBGrid.EditActive = True Then
        TDBGrid.Update
        If m_blnErrorBlankMatID = True Then 'update cancelled, invalid
            Update = False
        End If
    End If
    If Update = True Then
        If m_objGridMap.IsPendingChange Then
        If m_rec.RecordCount > 0 Then
            m_rec.MoveFirst
            While Not m_rec.EOF
                ' Skip the record if it didn't change
                Dim clsRowInfo As CRowInfo
                Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
                If Not clsRowInfo Is Nothing Then
                    If Not clsRowInfo.State = STATE_NONE Then
                    ' If record was added or changed
                        If clsRowInfo.State = STATE_NEW Or clsRowInfo.State = STATE_MODIFIED Then
                            strUpdate = "exec sp_update_mat_usage_driver "
                            m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MATUSAGE_TABLE
                            strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                            strUpdate = strUpdate + ", @last_update_id=" + CStr(m_rec.Fields("last_update_id").Value)
                            strUpdate = strUpdate + ", @unit_cost_id='" + m_strUnitCostID + "'"
                            blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                            ' Reset on success
                            If blnReturn = True Then
                                intSuccess = intSuccess + 1
                                m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
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
                    End If
                End If
                m_rec.MoveNext
            Wend
            m_rec.UpdateBatch
            End If
        End If
If m_blnRecords_Deleted Then
    Set colRowInfo = m_objGridMap.RowInfos
    'Execute delete requests from rowinfo
    
    For Each clsRowInfo In colRowInfo
        If clsRowInfo.State = STATE_DELETED And clsRowInfo.sKey > 0 Then
            ' Build SQL statement
            strUpdate = "exec sp_delete_material_usage "
            strUpdate = strUpdate + "@mat_skey=" + str(clsRowInfo.sKey) + ","
            strUpdate = strUpdate + " @unit_cost_skey=" + str(m_intUnitCostSKey)
            strUpdate = strUpdate + ", @UseCommitControl=1"
            strUpdate = strUpdate + ", @verifylastmaterial=1"
            blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
            If Not blnReturn Then
                intErrors = intErrors + 1
                m_objGridMap.SetError clsRowInfo.Bookmark, strError
            Else
                m_objGridMap.SetRowState clsRowInfo.Bookmark, STATE_NONE
            End If
        End If
    Next
    
    Set colRowInfo = Nothing
End If
        
    
        Dim strMsg As String
        strMsg = ""
'        If intSuccess > 0 Then
'            strMsg = str(intSuccess) + " rows updated successfully." + Chr(10)
'        End If
        If intErrors > 0 Then
            strMsg = strMsg + str(intErrors) + " errors occurred."
            ' Return value will be False
            Update = False
        End If
        'Do not display update confirmation message, only errors
        If Len(strMsg) > 0 And intErrors > 0 Then
            MsgBox strMsg
        End If
        m_blnRecords_Deleted = False
        TDBGrid.ReBind
    End If
    Exit Function
    
CleanUp:
    MsgBox "Error in Update"
End Function

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 0 Then    'Mat ID
    TDBGrid.Columns("mat_skey").Value = m_intMat_SKey
End If
End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
'        If TDBGrid.AddNewMode > 0 Then
            If Screen.ActiveControl.Name <> "cmdMatUsageDelete" Then 'delete not pending
                If Trim(TDBGrid.Columns("mat_id").Text) = "" Then
                    Cancel = True
                    m_blnErrorBlankMatID = True
                    TDBGrid.PostMsg 1
                Else
                    m_blnErrorBlankMatID = False
                End If
            End If
'        End If
    m_blnGridUpdateErrs = Cancel

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
If m_blnErrorBlankMatID = True Then
    MsgBox "The Material ID may not be blank."
    Response = 0
    TDBGrid.SetFocus
    m_blnErrorBlankMatID = False
End If
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("Input Factor").Value = 1
    TDBGrid.Columns("Output Factor").Value = 1
    TDBGrid.Columns("Adj Factor").Value = 1
    TDBGrid.Columns("Unit Qty").Value = 1
    TDBGrid.Columns("last_update_id") = 0
    TDBGrid.Columns("unit_cost_skey") = str(m_intUnitCostSKey)

End Sub

Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)
If MsgId = 1 Then
    MsgBox "Please enter a valid Material ID"
End If
End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim rowinfo As CRowInfo
    
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton <> vbYes Then
        Exit Sub
    End If
    ' Position to the desired record and execute delete
    m_rec.Bookmark = Bookmark
    Set rowinfo = m_objGridMap.rowinfo(m_rec.Bookmark)
    
    If rowinfo.State = STATE_NEW Then
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
    ElseIf rowinfo.State <> STATE_NOT_SET _
        And rowinfo.State <> STATE_DELETED Then
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_objGridMap.SetRowSkey m_rec.Bookmark, m_rec.Fields("mat_skey")
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
    On Error Resume Next
    Dim strSELECT As String
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnReturn As Boolean
    Dim I As Integer
    Dim strError As String
    Dim int_mat_skey As Integer
    Dim strMaterialID As String
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

    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Material ID"
        
        TDBGrid.Text = Trim(TDBGrid.Text)     'rlh 03/13/2009 - Dave Drain issue for CCD 8.4
        
        If Trim(TDBGrid.Text) = "" Then
            MsgBox "The Material ID may not be blank."
            TDBGrid.Text = OldValue
            Cancel = True ' Cancel so old value is replaced
        Else
                    ' Check to see if the mat_id entered exists already
            Set rec = m_rec.Clone
            rec.MoveFirst
            rec.Find "mat_id = '" + Compress_String(TDBGrid.Text) + "'", , adSearchForward, 0
            If Not rec.EOF Then
                MsgBox "This material has already been assigned."
                TDBGrid.Text = OldValue
                Cancel = True ' Cancel so old value is replaced
            End If
            rec.Close
            Set rec = Nothing
            If Cancel = False Then
                strSELECT = "sp_validate_material_usage @key_type=2, @key = '" + Compress_String(TDBGrid.Text) + "'"
                ' Use DAL to perform select
                blnReturn = g_objDAL.ExecQuery(vbNullString, strSELECT, strError)
                If blnReturn = False Then
                    'rlh 03/05/09  Mel Mossman false alarm (bad material id) However the error handling
                    'here is ATROCIOUS!!!!
                    
                    If (InStr(Err.Description, "50030") > 0) Then  'rlh 03/05/09
                        MsgBox "The SKEY for the Material ID " + TDBGrid.Text + " does not exist."
                    Else
                        MsgBox "Error encountered while processing Material ID " + TDBGrid.Text
                    End If   'rlh 03/05/09 (end of block)
                    
                    TDBGrid.Text = OldValue
                    Cancel = True ' Cancel so old value is replaced
                Else
                    ' Check to see if the mat_id entered exists already
                    strSELECT = "Select mat_id, mat_skey from Material where mat_id='" + Compress_String(TDBGrid.Text) + "'"
                    ' Use DAL to perform select
                    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
                    If rec.RecordCount = 0 Then
                        MsgBox "The Material ID " + TDBGrid.Text + " does not exist."
                        TDBGrid.Text = OldValue
                        Cancel = True ' Cancel so old value is replaced
                    Else
                    ' If it does, copy that data into grid
                        ' If they entered into the add row
                        strMaterialID = Trim(TDBGrid.Text)
                        'Check to see if the material has been assigned for this unit cost - may not have dups
                        vntMyBookmark = TDBGrid.Bookmark
                        If m_rec.RecordCount > 0 Then
                            'Get existing mat_skey
                            If OldValue <> Empty Then    'existing record
                                m_rec.Bookmark = TDBGrid.Bookmark
                                m_intMat_SKey = m_rec.Fields("mat_skey")
                            Else
                                m_intMat_SKey = 0
                            End If
                            m_rec.MoveFirst
                            Do Until m_rec.EOF
                                If Trim(m_rec.Fields("mat_id")) = strMaterialID Then
                                    MsgBox "The Material ID " + TDBGrid.Text + " has already been assigned, and may not be assigned twice."
                                    TDBGrid.Text = OldValue
                                    Cancel = True
                                    Exit Do
                                End If
                                m_rec.MoveNext
                            Loop
                        End If
        '                If OldValue = Empty Then    'new record
                        If m_intMat_SKey = 0 Then
                            m_intMat_SKey = rec.Fields("mat_skey")
                        End If
                        'TDBGrid.PostMsg 1
                    End If
                End If
            End If
        End If
    Case "Unit Qty", "Input Factor", "Output Factor", "Adj Factor"
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
    m_blnGridUpdateErrs = Cancel
End Sub

Public Property Get GRIDTYPE() As String
    GRIDTYPE = m_GRIDTYPE
End Property

Public Property Get GridUpdateErrors() As Boolean
    GridUpdateErrors = m_blnGridUpdateErrs
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

Public Property Let UnitCostSKey(intUnitCostSKey As Long)
    m_intUnitCostSKey = intUnitCostSKey
End Property

Public Property Let UnitCostID(strUnitCostId As String)
    m_strUnitCostID = strUnitCostId
End Property
