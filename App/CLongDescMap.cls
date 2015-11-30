VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLongDescMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Long Description grid
Option Explicit

''' <modulename> CLongDescMap.cls</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Helper class module for frmLongDescriptionGrid.frm
'''
'''  More important subs/functions:
'''"   InitGrid()
'''"   Update()
''' </summary>
'''
''' <seealso> frmLongDescRpt </seealso>
''' <datastruct> TDBGRID</datastruct>
''' <datastruct> m_objGridMap</datastruct>
'''
''' <storedprocedurename> sp_select_attribute_value_grid_structure</storedprocedurename>
'''<storedprocedurename> sp_update_object_attribute_value</storedprocedurename>
'''<storedprocedurename> n/a</storedprocedurename>
'''
'''<returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
'''<code>
'''</code>
'''<code>
'''exec sp_select_attribute_value_grid_structure  'U', 'A'
'''</code>
'''<code>
'''exec sp_update_object_attribute_value @obj_id='681020300409', @row_meas_sys_cd='I', @obj_skey= 109032, @col_1='xxxx', @old_col_1='', @msc_1='', @old_msc_1='', @col_2='', @old_col_2='', @msc_2='', @old_msc_2='', @col_3='', @old_col_3='', @msc_3='', @old_msc_3='', @col_4='', @old_col_4='', @msc_4='', @old_msc_4='', @col_5='', @old_col_5='', @msc_5='', @old_msc_5='', @col_6='', @old_col_6='', @msc_6='', @old_msc_6='', @col_7='', @old_col_7='', @msc_7='', @old_msc_7='', @col_8='', @old_col_8='', @msc_8='', @old_msc_8=''</code>
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



'*** IF SET TO TRUE, USER IS ABLE TO EDIT METRIC ROWS
'*** SET TO FALSE FOR BUSINESS RULE CONFORMATION (BACKS OUT CHANGES)
'*** SATISFIES CR#1537/1547
'*** 6/1/2005 RTD
Private Const ALLOW_METRIC_EDITING As Boolean = True

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "LongDescription" ' Identity of grid
' Names of some table this grid uses
Const ALL_TABLES = "All"

Private m_UpdateStatusCode As Integer
Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean

Private m_objObjects As Collection
Dim rsColumnAttributeDefs As ADODB.RecordSet

Dim m_varSaveBookmark As Variant
Dim m_varCurrentBookmark As Variant
Dim m_iColIndex As Integer

Const UpdateMetricValues = 1
Const UpdateImperialDefault = 2
Const UpdateMetricDefault = 3

Private Function CurrentUpdateStatusCode() As Integer
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    Dim strSelect As String
    On Error Resume Next
    If m_UpdateStatusCode = 0 Then
        rec.Close
        strSelect = "SELECT update_status_code = domain_value from DOMAIN_TBL where domain_name = 'PAPER_CLIP'"
        g_objDAL.GetRecordset CONNECT, strSelect, rec
        If rec.EOF Then
            m_UpdateStatusCode = 0
        Else
            m_UpdateStatusCode = rec.Fields("update_status_code")
        End If
        rec.Close
        Set rec = Nothing
    End If
    CurrentUpdateStatusCode = m_UpdateStatusCode

End Function

Public Property Let RecordSet(RecSet As ADODB.RecordSet)
    Dim objObject As CObject
    Dim objObjectAttribute As CObjAttribute
    Dim i As Integer
    Dim sAttrFieldName As String
    Dim sFieldName As String
    
    On Error Resume Next
    
    ' Pass-through to base class
    m_objGridMap.RecordSet = RecSet
    Set m_rec = RecSet
    'Populate object collection - used for display in the grid
    If Not (m_rec.EOF And m_rec.BOF) Then
        m_rec.MoveFirst
        Set m_objObjects = Nothing
        Set m_objObjects = New Collection
        Do Until m_rec.EOF
            Set objObject = New CObject
            With objObject
                .Object_Key = Compress_String(m_rec.Fields("obj_skey")) + m_rec.Fields("row_meas_sys_cd")
                .Object_ID = Compress_String(m_rec.Fields("obj_id"))
                .RowMeasSys = m_rec.Fields("row_meas_sys_cd")
                rsColumnAttributeDefs.MoveFirst
                If Not rsColumnAttributeDefs.EOF Then
                    For i = 0 To (rsColumnAttributeDefs.RecordCount / 2) - 1
                        sAttrFieldName = rsColumnAttributeDefs.Fields("col_id")
                        rsColumnAttributeDefs.MoveNext
                        sFieldName = rsColumnAttributeDefs.Fields("col_id")
                        objObject.Add_Object_Attribute sAttrFieldName, m_rec.Fields(sFieldName)
                        rsColumnAttributeDefs.MoveNext
                    Next i
                End If
            End With
            m_objObjects.Add objObject, objObject.Object_Key
            m_rec.MoveNext
        Loop
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

Public Property Let Preferences(bln As Boolean)
    ' Pass-through to base class
    m_blnPreferences = bln
End Property

Public Sub InitGrid()
    Dim i As Integer
'*** APEX Migration Utility Code Change ***
'    Dim Col As TrueOleDBGrid70.Column
    Dim Col As TrueOleDBGrid80.Column
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid70.ValueItem
    Dim Item As New TrueOleDBGrid80.ValueItem
    Dim strSelect As String
    Dim blnReturn As Boolean

    strSelect = "exec sp_select_attribute_value_grid_structure  'U', 'A'"
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsColumnAttributeDefs)
    If blnReturn = False Then
        MsgBox "Error return description definitions."
    Else
        m_objGridMap.GRIDTYPE = m_GRIDTYPE
        ' Add the columns to the grid
        m_objGridMap.AddCol "Unit Cost ID", "obj_id", ALL_TABLES, 1500, True, , , FORMAT_UNIT_COST
        ' 8/23/2005 RTD - ADD MF2004 EXT UNIT COST ID COLUMN
        m_objGridMap.AddCol "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2), "ext_id", , 1500, True, , , FORMAT_UNIT_COST_04
        m_objGridMap.AddCol "MSys", "row_meas_sys_cd", ALL_TABLES, 200, True, True, False
        m_objGridMap.AddCol "obj_skey", "obj_skey", ALL_TABLES, 200, True, False, False
        m_objGridMap.AddSplit
        m_objGridMap.AddCol "Object Desc", "obj_desc", , 2000, True
        If Not rsColumnAttributeDefs.EOF Then
            For i = 0 To rsColumnAttributeDefs.RecordCount - 1
                m_objGridMap.AddCol rsColumnAttributeDefs.Fields("attribute_display_desc"), rsColumnAttributeDefs.Fields("col_id"), ALL_TABLES, IIf(rsColumnAttributeDefs.Fields("m_sys") = True, 400, 2500), , , , , True
                rsColumnAttributeDefs.MoveNext
            Next i
        End If
        
        ' Call base class
        If m_blnPreferences Then
            m_objGridMap.BuildPreferenceGrid
        Else
            m_objGridMap.BuildGrid
        End If
        'Add drop down lists for all measurement system fields
        If Not m_blnPreferences Then
            TDBGrid.Columns("Object Desc").WrapText = True
            If Not (rsColumnAttributeDefs.BOF And rsColumnAttributeDefs.EOF) Then
                rsColumnAttributeDefs.MoveFirst
                For i = 0 To rsColumnAttributeDefs.RecordCount - 1
                    If rsColumnAttributeDefs.Fields("m_sys") = True Then  'Need to set presentation to drop down list of values - I, M, A
                        Item.Value = "A"
                        TDBGrid.Columns(rsColumnAttributeDefs.Fields("attribute_display_desc")).ValueItems.Add Item
                        Item.Value = "I"
                        TDBGrid.Columns(rsColumnAttributeDefs.Fields("attribute_display_desc")).ValueItems.Add Item
                        Item.Value = "M"
                        TDBGrid.Columns(rsColumnAttributeDefs.Fields("attribute_display_desc")).ValueItems.Add Item
                        TDBGrid.Columns(rsColumnAttributeDefs.Fields("attribute_display_desc")).ValueItems.Presentation = dbgComboBox
                        TDBGrid.Columns(rsColumnAttributeDefs.Fields("attribute_display_desc")).AutoDropDown = True
                        TDBGrid.Columns(rsColumnAttributeDefs.Fields("attribute_display_desc")).ValueItems.Validate = True
                    End If
                    rsColumnAttributeDefs.MoveNext
                Next i
            End If
        End If
        m_UpdateStatusCode = 0  'Initialize domain table value - updated if referenced
    End If

End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnUpdateMat As Boolean
    Dim blnUpdateMatPrice As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim blnRollup As Boolean
    
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
                strUpdate = "exec sp_update_object_attribute_value "
                
                strUpdate = "exec sp_update_object_attribute_value "       'rlh 05/22/2007 - PRODUCTION
                'strUpdate = "exec sp_update_object_attribute_value_rlh2 "       'rlh 05/22/2007 - TEST/DEBUG
                
                m_objGridMap.BuildUpdate_OrigValues strUpdate, m_rec.Bookmark, ALL_TABLES
                If Right(RTrim(strUpdate), 1) = "," Then
                    strUpdate = Mid(RTrim(strUpdate), 1, Len(RTrim(strUpdate)) - 1)
                End If
                blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                ' Reset on success
                If blnReturn = True Then
                    intSuccess = intSuccess + 1
                    clsRowInfo.State = STATE_NONE
                Else
                    intErrors = intErrors + 1
                    clsRowInfo.Error = True
                    clsRowInfo.ErrorText = strError
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
    TDBGrid.Refresh
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

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub
Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub

Private Sub Class_Terminate()
    Set m_objObjects = Nothing
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    Dim iPostAction As Integer
    
    m_varSaveBookmark = TDBGrid.Bookmark
    m_iColIndex = ColIndex
    If Left(TDBGrid.Columns(ColIndex).DataField, 4) = "msc_" Then
        'System type was changed, need to change next record
        iPostAction = UpdateMetricValues
        
    Else
        'Not system type, need to fill default system type if I
        If TDBGrid.Columns("MSys") = "I" Then
            ' EDITED AN I-ROW; POST CHANGES TO M-ROW
            iPostAction = UpdateImperialDefault
        ElseIf TDBGrid.Columns("MSys") = "M" And ALLOW_METRIC_EDITING Then
            ' EDITED AN M-ROW; POST CHANGES TO I-ROW
            iPostAction = UpdateMetricDefault
        End If
    End If
    If iPostAction > 0 Then
        TDBGrid.PostMsg iPostAction
    End If

End Sub

Private Sub TDBGrid_Change()
    If Len(TDBGrid.Text) > 255 Then
        TDBGrid.Text = Left(TDBGrid.Text, 255)
    End If
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    Dim sFieldName As String
    Dim objObject As CObject
    Dim rs_m_rec_clone As ADODB.RecordSet
    Dim strObjKey As String
    Dim objAttribute As CObjAttribute
    Dim bLocked As Boolean
    
    sFieldName = TDBGrid.Columns(Col).DataField
    If IsNumeric(Bookmark) And m_rec.State = adStateOpen Then
        Set rs_m_rec_clone = m_rec.Clone
        rs_m_rec_clone.Bookmark = Bookmark
        If sFieldName <> "obj_id" And sFieldName <> "row_meas_sys_cd" And sFieldName <> "obj_skey" Then
            If Left(sFieldName, 4) = "msc_" Then
                If rs_m_rec_clone.Fields("row_meas_sys_cd") = "M" Then
                    'No updates allowed to meas_sys_code on Metric rows
                    ' *** ALLOW CHANGES TO M-RECORDS (CR#1537/1547)
                    If Not ALLOW_METRIC_EDITING Then
                        bLocked = True
                    End If
                End If
            Else
                strObjKey = Compress_String(rs_m_rec_clone.Fields("obj_skey")) + rs_m_rec_clone.Fields("row_meas_sys_cd")
                Set objObject = m_objObjects(strObjKey)
                Set objAttribute = m_objObjects(strObjKey).Object_Attribute(sFieldName)
                ' *** ALLOW CHANGES TO M-RECORDS (CR#1537/1547)
                If Not ALLOW_METRIC_EDITING Then
                    bLocked = objAttribute.Locked
                End If
            End If
        End If
        If bLocked Then
            CellStyle.Locked = True
            CellStyle.ForeColor = vbGrayText
        Else
            CellStyle.ForeColor = vbBlack
            CellStyle.Locked = False
        End If
    End If
    Set rs_m_rec_clone = Nothing
    Set objObject = Nothing
    Set objAttribute = Nothing
End Sub

Private Sub TDBGrid_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> vbKeyBack Then
'        If TDBGrid.Columns(TDBGrid.Col).DataField = "book_desc" then
         '   If Len(TDBGrid.Text) + 1 > 255 Then
         '       KeyAscii = 0
         '   End If
'        End If
'    End If

End Sub

Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)
    Dim sSystemType As String, sNewSystemType As String
    Dim sFieldCol As String
    Dim sTextValue As String
    Dim sCurrentObj_ID As String
    Dim oObject As CObject
    Dim oObjectAttribute As CObjAttribute
    Dim bLocked As Boolean
    Dim sKey As String
    Dim sDesc As String
    Dim varCurrentBookmark As Variant
    Dim i As Long
    
    On Error Resume Next
    'If system type was changed, need to change next records system type.
    'If it was changed to A, need to also change prev fld
    'The object attributes also need to be updated to lock/unlock the field
    If (MsgId < 1 Or MsgId > 2) And (ALLOW_METRIC_EDITING And MsgId <> UpdateMetricDefault) Then
        'Not one of the messages trapped
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    varCurrentBookmark = TDBGrid.Bookmark   'Save current location after mouseclick
    If TDBGrid.DataChanged = False Then         'Don't need to position if still editing current row.
        TDBGrid.Bookmark = m_varSaveBookmark    'Restore edited location
    End If
    
    Select Case MsgId
    Case UpdateMetricValues
        sSystemType = TDBGrid.Columns(m_iColIndex)
        If sSystemType = "I" Then
            sNewSystemType = "M"
            bLocked = False
        ElseIf sSystemType = "M" Then
            sNewSystemType = "I"
            bLocked = False
        Else
            'SYSTEM TYPE CHANGED TO 'A'
            'Need to retrieve previous field name and value
            bLocked = True
            If sSystemType = "A" Then sNewSystemType = "A"
            rsColumnAttributeDefs.MoveFirst
            Do Until rsColumnAttributeDefs.Fields("col_id") = TDBGrid.Columns(m_iColIndex).DataField Or rsColumnAttributeDefs.EOF
                rsColumnAttributeDefs.MoveNext
            Loop
            rsColumnAttributeDefs.MovePrevious
            sFieldCol = rsColumnAttributeDefs.Fields("col_id")
            sTextValue = TDBGrid.Columns(sFieldCol)
        End If
    
        '-- Remember key value of current row for comparison below.
        sCurrentObj_ID = TDBGrid.Columns("obj_id").Value
        
        TDBGrid.MoveNext
        If TDBGrid.Columns("obj_id") <> sCurrentObj_ID Then
            ' THE NEXT ROW DOES NOT MATCH BECAUSE USER
            ' EDITED AN M-ROW, SO BACK-UP TWO TO THE I-ROW
            TDBGrid.MovePrevious
            TDBGrid.MovePrevious
        End If
        
        '-- Compare key value of next row with original row.
        '-- If same, propragate data from original row to next row.
        If TDBGrid.Columns("obj_id") = sCurrentObj_ID Then
            TDBGrid.Columns(m_iColIndex) = sNewSystemType
            If sNewSystemType = "A" Then
                TDBGrid.Columns(sFieldCol) = sTextValue
            End If
            'Need to update next record locked state
            'sKey = Compress_String(TDBGrid.Columns("Unit Cost ID")) + TDBGrid.Columns("MSys")
            sKey = Compress_String(TDBGrid.Columns("obj_skey")) + TDBGrid.Columns("MSys")
            Set oObject = m_objObjects.Item(sKey)
            rsColumnAttributeDefs.MoveFirst
            Do Until rsColumnAttributeDefs.Fields("col_id") = TDBGrid.Columns(m_iColIndex).DataField Or rsColumnAttributeDefs.EOF
                rsColumnAttributeDefs.MoveNext
            Loop
            rsColumnAttributeDefs.MovePrevious
            sFieldCol = rsColumnAttributeDefs.Fields("col_id")
            Set oObjectAttribute = oObject.Object_Attribute(sFieldCol)
            oObjectAttribute.Locked = bLocked
            Set oObject = Nothing
            Set oObjectAttribute = Nothing
        End If
    Case UpdateImperialDefault
        sTextValue = TDBGrid.Columns(m_iColIndex)
        rsColumnAttributeDefs.MoveFirst
        Do Until rsColumnAttributeDefs.Fields("col_id") = TDBGrid.Columns(m_iColIndex).DataField Or rsColumnAttributeDefs.EOF
            rsColumnAttributeDefs.MoveNext
        Loop
        rsColumnAttributeDefs.MoveNext
        
        sFieldCol = rsColumnAttributeDefs.Fields("col_id")
        If Len(TDBGrid.Columns(sFieldCol)) = 0 Then 'No value yet
            TDBGrid.Columns(sFieldCol) = "A"
        End If
        sSystemType = TDBGrid.Columns(sFieldCol)
        
        '-- Remember key value of current row for comparison below.
        sCurrentObj_ID = TDBGrid.Columns("obj_id").Value
        
        TDBGrid.MoveNext
        
        '-- Compare key value of next row with original row.
        '-- If same, propragate data from original row to next row.
        If sCurrentObj_ID = TDBGrid.Columns("obj_id").Value Then
            If sSystemType = "I" Then
                sNewSystemType = "M"
            Else
                sNewSystemType = "A"
            End If
            TDBGrid.Columns(sFieldCol) = sNewSystemType
            If sNewSystemType = "A" Then
                TDBGrid.Columns(m_iColIndex) = sTextValue
                ' PROPOGATE LONG DESCRIPTION CHANGE TO 'M' ROW
                ' ADDED 5/25/2005 RTD
                ' UPDATED 8/24/2005 RTD - OBJECT DESC COLUMN NOW STARTS AT 5
                sDesc = ""
                For i = 5 To TDBGrid.Columns.Count - 1 Step 2
                    If TDBGrid.Columns(i).Text <> "" Then
                        If sDesc = "" Then
                            sDesc = UCase(Left(TDBGrid.Columns(i).Text, 1)) + Mid(TDBGrid.Columns(i).Text, 2)
                        Else
                            sDesc = sDesc + ", " + TDBGrid.Columns(i).Text
                        End If
                    End If
                Next
                If Len(sDesc) > 255 Then
                    sDesc = Left(sDesc, 255)
                End If
                TDBGrid.Columns("obj_desc") = sDesc
            End If
        End If
    Case UpdateMetricDefault
        ' REVERSE PROPOGATE CHANGES TO 'M' METRIC ROW TO 'I' ROW
        ' ADDED 6/1/2005 RTD
        sTextValue = TDBGrid.Columns(m_iColIndex)
        rsColumnAttributeDefs.MoveFirst
        Do Until rsColumnAttributeDefs.Fields("col_id") = TDBGrid.Columns(m_iColIndex).DataField Or rsColumnAttributeDefs.EOF
            rsColumnAttributeDefs.MoveNext
        Loop
        rsColumnAttributeDefs.MoveNext
        
        sFieldCol = rsColumnAttributeDefs.Fields("col_id")
        If Len(TDBGrid.Columns(sFieldCol)) = 0 Then 'No value yet
            TDBGrid.Columns(sFieldCol) = "A"
        End If
        sSystemType = TDBGrid.Columns(sFieldCol)
        
        '-- Remember key value of current row for comparison below.
        sCurrentObj_ID = TDBGrid.Columns("obj_id").Value
        
        TDBGrid.MovePrevious
        
        '-- Compare key value of next row with original row.
        '-- If same, propragate data from original row to next row.
        If sCurrentObj_ID = TDBGrid.Columns("obj_id").Value Then
            If sSystemType = "M" Then
                sNewSystemType = "I"
            Else
                sNewSystemType = "A"
            End If
            TDBGrid.Columns(sFieldCol) = sNewSystemType
            If sNewSystemType = "A" Then
                TDBGrid.Columns(m_iColIndex) = sTextValue
                ' PROPOGATE LONG DESCRIPTION CHANGE TO 'i' ROW
                ' ADDED 5/25/2005 RTD
                ' UPDATED 8/24/2005 RTD - OBJECT DESC COLUMN NOW STARTS AT 5
                sDesc = ""
                For i = 5 To TDBGrid.Columns.Count - 1 Step 2
                    If TDBGrid.Columns(i).Text <> "" Then
                        If sDesc = "" Then
                            sDesc = UCase(Left(TDBGrid.Columns(i).Text, 1)) + Mid(TDBGrid.Columns(i).Text, 2)
                        Else
                            sDesc = sDesc + ", " + TDBGrid.Columns(i).Text
                        End If
                    End If
                Next
                If Len(sDesc) > 255 Then
                    sDesc = Left(sDesc, 255)
                End If
                TDBGrid.Columns("obj_desc") = sDesc
            End If
        End If
    End Select

    TDBGrid.Bookmark = varCurrentBookmark
    Screen.MousePointer = vbNormal

End Sub
