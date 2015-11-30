VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProjectMap"
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
Const PROJECT_TABLE = "Project_Building_Detail"
Const ALL_TABLES = "All"
Const m_strGridType = "PROJECT"
'
'   Grid Messages.
Const CLEAR_BLDG_CATEGORY = 1
'
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
        .Add 1, "proj_bldg_skey"
        .Add 1, "bid_date"
        .Add 1, "exterior_material_desc"
        .Add 1, "facility1_class_id"
        .Add 1, "proj_bldg_project_tot_cost"
        .Add 1, "gross_floor_area"
        .Add 1, "state_code"
        .Add 1, "last_update_person"
        .Add 1, "last_update_date"
    End With
End Sub

Public Sub InitGrid(aryList() As String, aryExterior() As String, aryState() As String, Optional bPreferences As Boolean = False)
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid70.ValueItem
    Dim Item As New TrueOleDBGrid80.ValueItem
    Dim lRet As Long
    Dim hKey As Long
    Dim I As Integer
    Dim strTemp As String
    Dim aryFields
    Dim strKey As String

    '
    '   Do not clear the registry if calling from the Preferences screen!
    If bPreferences = False Then
        '
        '   Do not clear PROJECT registry for main fields because it erases the
        '   preferences the user set on the Preferences screen.
        aryFields = Split("proj_bldg_id,bid_date,use_ind,exterior_material_desc,facility1_class_id,proj_bldg_project_tot_cost,gross_floor_area,state_code,class_uom", ",")
'        For i = LBound(aryFields) To UBound(aryFields)
'            strKey = CCD_KEY + "\" + m_strGridType + "\" + aryFields(i)
'            lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
'            lRet = RegSetValueExLong(hKey, "Order", 0&, REG_DWORD, CLng(i), 4)
'        Next
        For I = LBound(aryList()) To UBound(aryList())
            strKey = CCD_KEY + "\" + m_strGridType + "\" + aryList(I)
            lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
            lRet = RegSetValueExLong(hKey, "Order", 0&, REG_DWORD, CLng(I + UBound(aryFields) + 1), 4)
        Next
    
'        strKey = CCD_KEY + "\" + m_strGridType + "\last_update_person"
'        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
'        lRet = RegSetValueExLong(hKey, "Order", 0&, REG_DWORD, CLng(UBound(aryList()) + UBound(aryFields) + 2), 4)
'        strKey = CCD_KEY + "\" + m_strGridType + "\last_update_date"
'        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
'        lRet = RegSetValueExLong(hKey, "Order", 0&, REG_DWORD, CLng(UBound(aryList()) + UBound(aryFields) + 3), 4)
    End If
    '
    '   Identifies the Grid.
    With m_objGridMap
        .GRIDTYPE = m_strGridType
        '
        '   Add the columns to the grid, pointing to their
        '   corresponding db field names so that the grid can
        '   populate directly from the m_rec recordset.
        .AddCol "ID", "proj_bldg_id", PROJECT_TABLE, 800, True, True
        '
        '   Main Split, this grid only has 1.
        '
        .AddSplit
        .AddCol "Year", "bid_date", PROJECT_TABLE, 500, False, True
        .AddCol "Use", "use_ind", PROJECT_TABLE, 500, False, True, False
        .AddCol "Material", "exterior_material_desc", PROJECT_TABLE, 2000, False, True
        .AddCol "Class", "facility1_class_id", PROJECT_TABLE, 500, False, True
        .AddCol "Tot. Cost", "proj_bldg_project_tot_cost", PROJECT_TABLE, 1000, False, True, True, 1
        ' UPDATED 9/7/2005 RTD
        ' CR#1513 - CHANGE AREA COLUMN HEADER TO "TOT. AREA"
        .AddCol "Tot. Area", "gross_floor_area", PROJECT_TABLE, 1000, False, True, True, 1
        .AddCol "Location", "state_code", PROJECT_TABLE, 1000, False, True
        .AddCol "UOM", "class_uom", PROJECT_TABLE, 500, True, True
        For I = LBound(aryList()) To UBound(aryList())
            .AddCol aryList(I), aryList(I), PROJECT_TABLE, 1000, False, True, True, 1
        Next
        .AddCol "Updated By", "last_update_person", PROJECT_TABLE, 1000, True, True
        .AddCol "Updated", "last_update_date", PROJECT_TABLE, 1000, True, True
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
        .Splits(0).Caption = "Project"
        .Splits(1).Caption = "Attributes"
        .Splits(1).Size = 5
        
        '   Allow new record to be added
        .AllowAddNew = True
        '   Format the number columns
        .Columns("proj_bldg_project_tot_cost").NumberFormat = "#,###,###,###"
        .Columns("gross_floor_area").NumberFormat = "#,###,###,###"
        For I = LBound(aryList()) To UBound(aryList())
            .Columns(aryList(I)).NumberFormat = "#,###,###,###"
        Next
        '
        '   create exterior material drop down list
        For I = LBound(aryExterior()) To UBound(aryExterior())
            Item.Value = aryExterior(I)
            .Columns("exterior_material_desc").ValueItems.Add Item
        Next
        .Columns("exterior_material_desc").ValueItems.Presentation = dbgComboBox
        .Columns("exterior_material_desc").ValueItems.Validate = True
        .Columns("exterior_material_desc").AutoDropDown = True
        '
        '   create state code drop down list
        For I = LBound(aryState()) To UBound(aryState())
            Item.Value = aryState(I)
            .Columns("state_code").ValueItems.Add Item
        Next
        .Columns("state_code").ValueItems.Presentation = dbgComboBox
        .Columns("state_code").ValueItems.Validate = True
        .Columns("state_code").AutoDropDown = True
    End With
End Sub

Public Property Get GRIDTYPE() As String
    GRIDTYPE = "PROJECT"
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

'*** APEX Migration Utility Code Change ***
'Public Sub SetGrid(grd As TrueOleDBGrid70.TDBGrid)
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

Public Sub Update(aryList() As String, classid As String)
    Dim strSQL As String
    Dim strError As String
    Dim I As Integer
    'On Error Resume Next
    m_NumberOfUpdateErrors = 0
    m_NumberOfSuccessfulUpdates = 0

    With m_rec

        If .State <> adStateClosed Then
            .MoveFirst
            While Not .EOF
                '
                '   Skip the record if it didn't change
                Dim clsRowInfo As CRowInfo
                Set clsRowInfo = m_objGridMap.rowinfo(.Bookmark)
                

                If Not clsRowInfo Is Nothing Then
                    If clsRowInfo.State <> STATE_NONE Then
                        strSQL = "EXEC sp_update_project_list "
                        If Trim(.Fields("proj_bldg_skey")) = "" Or IsNull(.Fields("proj_bldg_skey")) Then
                            strSQL = strSQL & " @projkey = null, "
                        Else
                            strSQL = strSQL & " @projkey = " & Trim(.Fields("proj_bldg_skey")) & ", "
                        End If
                        If Trim(.Fields("proj_bldg_project_tot_cost")) = "" Or IsNull(.Fields("proj_bldg_project_tot_cost")) Then
                            strSQL = strSQL & " @totalvalue = 0, "
                        Else
                            strSQL = strSQL & " @totalvalue = " & Trim(.Fields("proj_bldg_project_tot_cost")) & ", "
                        End If
                        If Trim(.Fields("gross_floor_area")) = "" Or IsNull(.Fields("gross_floor_area")) Then
                            strSQL = strSQL & " @grossarea = 0, "
                        Else
                            strSQL = strSQL & " @grossarea = " & Trim(.Fields("gross_floor_area")) & ", "
                        End If
                        strSQL = strSQL & " @useind = '" & .Fields("use_ind") & "', "
                        strSQL = strSQL & " @exterior = '" & .Fields("exterior_material_desc") & "', "
                        strSQL = strSQL & " @list = '"
                        For I = LBound(aryList()) To UBound(aryList())
                            strSQL = strSQL & aryList(I) & "=" & .Fields(aryList(I)) & ","
                        Next
                        strSQL = Left(strSQL, Len(strSQL) - 1) & "',"
                        strSQL = strSQL & " @classid = '" & .Fields("facility1_class_id") & "', "
                        If Trim(.Fields("bid_date")) = "" Or IsNull(.Fields("bid_date")) Then
                            strSQL = strSQL & " @biddate = null, "
                        Else
                            strSQL = strSQL & " @biddate = '1/1/" & Trim(.Fields("bid_date")) & "', "
                        End If
                        strSQL = strSQL & " @state = '" & .Fields("state_code") & "'"
                        ' exec store proc
                        If g_objDAL.ExecQuery(vbNullString, strSQL, strError) Then
                            If strError = "" Then
                                m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                            Else
                                m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
                                m_objGridMap.SetError .Bookmark, strError
                            End If
                        End If
                        clsRowInfo.State = STATE_NONE
                    End If
                End If
                .MoveNext
            Wend
            
            .UpdateBatch
        End If
    End With
End Sub

' 9/9/2005 RTD
' VERIFY THAT THE RECORD IS OK FOR DELETION
Private Function check_error(varBookmark As Variant, lngSkey As Long) As Boolean

    check_error = False

End Function

' 9/9/2005 RTD
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
            If check_error(TDBGrid.Bookmark, TDBGrid.Columns("ID")) = True Then
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
        If check_error(TDBGrid.Bookmark, TDBGrid.Columns("ID")) = True Then
            Delete = False
        Else
            Delete = m_objGridMap.Delete
        End If
    End If
    
End Function

' 9/9/2005 RTD
' PROCESS ROWS DELETED FROM THE GRID
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
   
    ' Position to the desired record and execute delete
    If check_error(Bookmark, TDBGrid.Columns("ID")) = True Then Exit Sub
    
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
    strUpdate = "exec usp_delete_project "
    strUpdate = strUpdate + "@proj_bldg_skey=" + str(m_rec.Fields("proj_bldg_skey"))

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
    MsgBox Err.Description
    Bookmark = Null

End Sub
