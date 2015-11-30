VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAdminUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'   Access to grid events.
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid
Attribute TDBGrid.VB_VarHelpID = -1
'
'   Names of some of the tables this grid uses.
Const USER_TABLE = "USER_NAMES"
Const ALL_TABLES = "All"
Const m_strGridType = "UserAdmin"

'   "Parent" class
Private m_objGridMap As New CGridMap
'
'   Recordset that populates grid
Private m_rec                   As ADODB.RecordSet
Private m_blnPreferences        As Boolean
Private m_varSaveBookmark       As Variant
Private m_varCurrentBookmark    As Variant
Private m_colUpdatedRows        As New Collection
'
'   Used in the SuccessfulUpdates Property Get.
Private m_NumberOfSuccessfulUpdates As Integer
'
'   Used in the UpdateErrors Property Get.
Private m_NumberOfUpdateErrors As Integer

Private Sub Class_Initialize()
    With m_objGridMap.m_colCloneCols
        .Add 1, "user_id"
        .Add 1, "user_name"
        .Add 1, "user_first_name"
        .Add 1, "user_middle_initial"
        .Add 1, "user_last_name"
        .Add 1, "user_title"
        .Add 1, "user_extension"
        .Add 1, "user_fax"
        .Add 1, "user_role"
        .Add 1, "ccd_version"
        .Add 1, "last_logon"
        .Add 1, "last_update_date"
        .Add 1, "last_update_person"
    End With
End Sub

Public Sub InitGrid(Optional bPreferences As Boolean = False)
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid70.ValueItem
    Dim Item As New TrueOleDBGrid80.ValueItem
    Dim lRet As Long
    Dim hKey As Long
    Dim I As Integer
    Dim strTemp As String
    Dim aryFields
    Dim strKey As String

    With m_objGridMap
        .GRIDTYPE = m_strGridType
        '
        '   Add the columns to the grid, pointing to their
        '   corresponding db field names so that the grid can
        '   populate directly from the m_rec recordset.
        .AddCol "User ID", "user_id", USER_TABLE, 1200, True, True, False
        '
        '   Main Split, this grid only has 1.
        '
        .AddSplit
        .AddCol "Full Name", "user_name", USER_TABLE, 1800, False, True, False
        .AddCol "First Name", "user_first_name", USER_TABLE, 1000, False, True, False, FORMAT_STRING_TRIM
        .AddCol "M.I.", "user_middle_initial", USER_TABLE, 500, False, True, False, FORMAT_STRING_TRIM
        .AddCol "Last Name", "user_last_name", USER_TABLE, 1000, False, True, False, FORMAT_STRING_TRIM
        .AddCol "Title", "user_title", USER_TABLE, 1200, False, True, False, FORMAT_STRING_TRIM
        .AddCol "Extension", "user_extension", USER_TABLE, 800, False, True, False
        .AddCol "Fax", "user_fax", USER_TABLE, 1200, False, True, True
        .AddCol "Role", "user_role", USER_TABLE, 1000, False, True, False
        .AddCol "Last Logon", "last_logon", USER_TABLE, 1800, True, True, False
        .AddCol "CCD Version", "ccd_version", USER_TABLE, 1000, True, True, False
        .AddCol "Last Updated", "last_update_date", ALL_TABLES, 1800, True, True, False
        .AddCol "Updated By", "last_update_person", ALL_TABLES, 1000, True, True, False
        '
        '   Call base class.
        If m_blnPreferences Then
            .BuildPreferenceGrid
        Else
            .BuildGrid True, False
        End If
    End With
    '
    '   Give headings for the splits.
    With TDBGrid
        .Splits(0).Caption = "User"
        .Splits(1).Caption = "Attributes"
        .Splits(1).Size = 5
        .AllowAddNew = False
        
        ' Format columns
        .Columns("Full Name").AutoSize
        .Columns("Extension").AutoSize
        .Columns("Title").AutoSize
        
        ' User Role needs a pop-up
        .Columns("Role").Button = True
        
    End With
    
End Sub

Public Property Get GRIDTYPE() As String
    GRIDTYPE = m_strGridType
End Property

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub

Public Sub SetGrid(grd As TrueOleDBGrid80.TDBGrid)
    ' Pass-through to base class
    m_objGridMap.SetGrid grd
    Set TDBGrid = grd
End Sub
'
'   Sets whether or not the user wants the row to wrap in the grid.
Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub

Public Function IsPendingChange() As Boolean
    IsPendingChange = m_objGridMap.IsPendingChange
End Function

Public Function GetError(intBookmark As Integer) As String
    Dim objRowInfo As CRowInfo

    On Error Resume Next
    Set objRowInfo = Nothing
    Set objRowInfo = m_objGridMap.rowinfo(intBookmark)
    If Not objRowInfo Is Nothing Then
        GetError = objRowInfo.ErrorText
    Else
        GetError = ""
    End If
End Function

Public Property Let RecordSet(RecSet As ADODB.RecordSet)
    '
    '   Pass-through to base class
    m_objGridMap.RecordSet = RecSet
    Set m_rec = RecSet
End Property

Public Property Let Preferences(bln As Boolean)
    '
    '   Pass-through to base class
    m_blnPreferences = bln
End Property

Public Property Get SuccessfulUpdates() As Integer
    SuccessfulUpdates = m_NumberOfSuccessfulUpdates
End Property

Public Property Get UpdateErrors() As Integer
    UpdateErrors = m_NumberOfUpdateErrors
End Property

Public Function CloneRowRecordset() As ADODB.RecordSet
    Set CloneRowRecordset = m_objGridMap.CloneRowRecordset
End Function

Public Function Update() As Boolean
    Dim strUpdate As String     ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateUnitCost As Boolean
    Dim blnUpdateUnitCostEx As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String      ' Error string returned from DAL
    Dim intErrors As Integer    ' Tracks if any errors have occurred
    Dim intSuccess As Integer   ' Tracks successful updates

    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0

    m_rec.MoveFirst

    ' Loop through all grid records
    While Not m_rec.EOF
        Dim clsRowInfo As CRowInfo
        Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
        ' Skip the record if it didn't change
        If Not clsRowInfo Is Nothing Then
            blnReturn = False
            blnUpdateUnitCost = False
            blnUpdateUnitCostEx = False
            ' Loop through the fields to see what has changed
            For Each fld In m_rec.Fields
                ' If the value changed
                If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
                    ' See what table the field is from
                    Set objColDef = m_objGridMap.GetCol(fld.Name)
                    ' Mark the table we should update
                    blnUpdateUnitCost = True
                End If
            Next
            
            ' If the row is new
            If clsRowInfo.State = STATE_NEW Then
            ' NOTE: Cannot do NEWs in the grid, launches single record view!
            ' If this row was modified
            ElseIf clsRowInfo.State = STATE_MODIFIED Then
                If blnUpdateUnitCost Then
                'Update all types of rows.  The sp_update_unit_cost_driver will update the appropriate records.
                    strUpdate = "exec usp_update_user_names "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, USER_TABLE
                    'm_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                    strUpdate = strUpdate + " @last_update_person='" + strUserName + "',"
                    strUpdate = strUpdate + " @last_update_date='" + Format(Now(), "General Date") + "'"
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
        End If
        m_rec.MoveNext
    Wend
    m_rec.UpdateBatch

    Dim strMsg As String
    strMsg = ""
    If intSuccess > 0 Then
'        For Each m_varCurrentBookmark In m_colUpdatedRows
'            m_objGridMap.RefreshRecord CInt(m_varCurrentBookmark)
'            m_colUpdatedRows.Remove 1
'        Next
        strMsg = str(intSuccess) + " rows updated successfully." + vbCrLf
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

' 9/13/2005 RTD
' VERIFY THAT THE RECORD IS OK FOR DELETION
Private Function check_error(varBookmark As Variant, strSkey As String) As Boolean

    check_error = False

End Function

' 9/13/2005 RTD
' DELETE RECORD(S)
Public Function Delete() As Boolean
    Dim varSelectedBookmarks() As Variant
    Dim varCurrentSelectedBookmark As Variant
    Dim varButton
    Dim blnResult As Boolean
    Dim I As Integer
    Dim i1 As Integer
    
    'On Error Resume Next

    If TDBGrid.SelBookmarks.Count > 0 Then
        ReDim varSelectedBookmarks(0 To TDBGrid.SelBookmarks.Count)
        I = 0
        For Each varCurrentSelectedBookmark In TDBGrid.SelBookmarks
            varSelectedBookmarks(I) = varCurrentSelectedBookmark
            I = I + 1
        Next varCurrentSelectedBookmark
        For I = 0 To UBound(varSelectedBookmarks) - 1
            TDBGrid.Bookmark = varSelectedBookmarks(I)
            If check_error(TDBGrid.Bookmark, TDBGrid.Columns("User ID")) = True Then
                For i1 = 0 To TDBGrid.SelBookmarks.Count
                    If TDBGrid.SelBookmarks(i1) = varSelectedBookmarks(I) Then
                        TDBGrid.SelBookmarks.Remove i1
                        Exit For
                    End If
                Next i1
            End If
        Next I
        If TDBGrid.SelBookmarks.Count > 0 Then
            Delete = m_objGridMap.Delete
        Else
            Delete = False
        End If
    Else
        If check_error(TDBGrid.Bookmark, TDBGrid.Columns("User ID")) = True Then
            Delete = False
        Else
            Delete = m_objGridMap.Delete
        End If
    End If
    
End Function

' 9/13/2005 RTD
' PROCESS ROWS DELETED FROM THE GRID
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
   
    ' Position to the desired record and execute delete
    If check_error(Bookmark, TDBGrid.Columns("User ID")) = True Then Exit Sub
    
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

    ' Build SQL statement
    strUpdate = "DELETE FROM user_names "
    strUpdate = strUpdate & " WHERE user_id = '" & m_rec.Fields("user_id") & "'"

    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetError m_rec.Bookmark, strError
        m_rec.Fields("Error").Value = "Error"
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
    End If
'    TDBGrid.ReBind
    Exit Sub

NoDelete:
    MsgBox Err.Description, vbCritical
    Bookmark = Null

End Sub

