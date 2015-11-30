VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBldgComAdds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'   This class is "derived" from CGridMap
'   It serves the Building grid.
'
'   Access to grid events.
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid
Attribute TDBGrid.VB_VarHelpID = -1
'
'   Names of some of the tables this grid uses.
Const COMMON_ADD_BOOK_DETAIL = "common_add_book_detail"
Const ALL_TABLES = "All"
'
'   "Parent" class
Private m_objGridMap As New CGridMap
'
'   Recordset that populates grid
Private m_rec As ADODB.RecordSet
Private m_blnPreferences As Boolean
Private m_varSaveBookmark As Variant
Private m_varCurrentBookmark As Variant
Private m_colUpdatedRows As New Collection
Private bInserting As Boolean
Private m_nMaxSortOrder As Integer
Private m_NumberOfSuccessfulUpdates     As Integer
Private bMultiDeleteInProcess   As Boolean
Private m_colSortOrderItems As New Collection

Private Sub Class_Initialize()
   
    With m_objGridMap.m_colCloneCols
        .Add 1, "bldg_id"
        .Add 1, "type_code"
        .Add 1, "common_add_skey"
        .Add 1, "common_add_skey_type"
        .Add 1, "sort_order"

        .Add 1, "common_add_child_skey"
        .Add 1, "common_add_child_skey_max"
        .Add 1, "id"
        .Add 1, "id_max"
        .Add 1, "unit"
        .Add 1, "book_desc"
        .Add 1, "tech_desc"
        .Add 1, "pct_ind"
        .Add 1, "mat_cost_op"
        .Add 1, "inst_cost_op"
        .Add 1, "total_cost_op"

        .Add 1, "last_update_date"
        .Add 1, "last_update_person"
        .Add 1, "last_update_id"
    End With
End Sub

Public Sub InitGrid(bResidential As Boolean)
    '
    '   Identifies the Grid.
    
    
    With m_objGridMap
        '
        'Allow the op columns to be updated if the model is residential.
        If bResidential Then
            .GRIDTYPE = "Common Additives"
            
            .AddCol "Bldg Skey", "bldg_skey", COMMON_ADD_BOOK_DETAIL, 750, True, False, False
            .AddCol "Type Code", "type_code", COMMON_ADD_BOOK_DETAIL, 750, True, False, False
            
            .AddCol "Common Add Skey", "common_add_skey", COMMON_ADD_BOOK_DETAIL, 2000, True, False, False
            
            '.AddCol "Skey Type", "common_add_skey_type", COMMON_ADD_BOOK_DETAIL, 700, True, True, False
            .AddCol "Skey Type", "common_add_skey_type", COMMON_ADD_BOOK_DETAIL, 700, False, True, False
            
            .AddCol "Sort Order", "sort_order", COMMON_ADD_BOOK_DETAIL, 1000, False, True, False
            .AddCol "Child Skey", "common_add_child_skey", COMMON_ADD_BOOK_DETAIL, 1000, True, False, False
            .AddCol "Child Skey Max", "common_add_child_skey_max", COMMON_ADD_BOOK_DETAIL, 1000, True, False, False
             'Legacy
'            .AddCol "ID", "ID", COMMON_ADD_BOOK_DETAIL, 1300, False, True, False, FORMAT_UNIT_COST
'            .AddCol "ID Max", "ID_Max", COMMON_ADD_BOOK_DETAIL, 1300, False, True, False, FORMAT_UNIT_COST
           
            'CCD 8.4 (rlh)  Change format from MF95 to MF04 unit cost id formatting
            '04/15/2009
            .AddCol "ID", "ID", COMMON_ADD_BOOK_DETAIL, 1300, False, True, False, FORMAT_UNIT_COST_04
            .AddCol "ID Max", "ID_Max", COMMON_ADD_BOOK_DETAIL, 1300, False, True, False, FORMAT_UNIT_COST_04
            
            .AddCol "Unit", "unit", COMMON_ADD_BOOK_DETAIL, 525, True, True, False
            .AddCol "Book Desc", "book_desc", COMMON_ADD_BOOK_DETAIL, 4000, False, True, False, FORMAT_STRING_TRIM
            .AddCol "Tech Desc", "tech_desc", COMMON_ADD_BOOK_DETAIL, 4600, True, True, False, FORMAT_STRING_TRIM
            
            .AddCol "Pct Ind", "pct_ind", COMMON_ADD_BOOK_DETAIL, 700, False, False, False
            .AddCol "Mat Cost OP", "mat_cost_op", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            .AddCol "Inst Cost OP", "inst_cost_op", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            .AddCol "Total Cost OP", "total_cost_op", COMMON_ADD_BOOK_DETAIL, 1100, False, True, False, FORMAT_PRICE
            
            .AddCol "Max Mat Cost OP", "mat_cost_op_max", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            .AddCol "Max Inst Cost OP", "inst_cost_op_max", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            .AddCol "Max Total Cost OP", "total_cost_op_max", COMMON_ADD_BOOK_DETAIL, 1100, False, True, False, FORMAT_PRICE
            
            .AddCol "Format Code", "format_code", COMMON_ADD_BOOK_DETAIL, 650, False, True, True
            .AddCol "Indent Code", "indent_code", COMMON_ADD_BOOK_DETAIL, 650, False, True, True
            
            .AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
            .AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
            .AddCol "last_update_id", "last_update_id", COMMON_ADD_BOOK_DETAIL, 800, True, False, False
        Else
            .GRIDTYPE = "Common Additives"
            
            .AddCol "Bldg Skey", "bldg_skey", COMMON_ADD_BOOK_DETAIL, 750, True, False, False
            .AddCol "Type Code", "type_code", COMMON_ADD_BOOK_DETAIL, 750, True, False, False
            
            .AddCol "Common Add Skey", "common_add_skey", COMMON_ADD_BOOK_DETAIL, 2000, True, False, False
            .AddCol "Skey Type", "common_add_skey_type", COMMON_ADD_BOOK_DETAIL, 700, True, True, False
            .AddCol "Sort Order", "sort_order", COMMON_ADD_BOOK_DETAIL, 1000, False, True, False
            .AddCol "Child Skey", "common_add_child_skey", COMMON_ADD_BOOK_DETAIL, 1000, True, False, False
            .AddCol "Child Skey Max", "common_add_child_skey_max", COMMON_ADD_BOOK_DETAIL, 1000, True, False, False
            
            'Legacy MF95 unit cost id formatting
            '.AddCol "ID", "ID", COMMON_ADD_BOOK_DETAIL, 1300, False, True, False, FORMAT_UNIT_COST
            '.AddCol "ID Max", "ID_Max", COMMON_ADD_BOOK_DETAIL, 1300, False, True, False, FORMAT_UNIT_COST
            
            'CCD 8.4 (rlh)  Change format from MF95 to MF04 unit cost id formatting
            '04/15/2009
            .AddCol "ID", "ID", COMMON_ADD_BOOK_DETAIL, 1300, False, True, False, FORMAT_UNIT_COST_04
            .AddCol "ID Max", "ID_Max", COMMON_ADD_BOOK_DETAIL, 1300, False, True, False, FORMAT_UNIT_COST_04
            
            .AddCol "Unit", "unit", COMMON_ADD_BOOK_DETAIL, 525, True, True, False
            .AddCol "Book Desc", "book_desc", COMMON_ADD_BOOK_DETAIL, 4000, False, True, False, FORMAT_STRING_TRIM
            .AddCol "Tech Desc", "tech_desc", COMMON_ADD_BOOK_DETAIL, 4600, True, True, False, FORMAT_STRING_TRIM
            
            .AddCol "Pct Ind", "pct_ind", COMMON_ADD_BOOK_DETAIL, 700, False, False, False
            .AddCol "Mat Cost OP", "mat_cost_op", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            .AddCol "Inst Cost OP", "inst_cost_op", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            .AddCol "Total Cost OP", "total_cost_op", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            
            .AddCol "Max Mat Cost OP", "mat_cost_op_max", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            .AddCol "Max Inst Cost OP", "inst_cost_op_max", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            .AddCol "Max Total Cost OP", "total_cost_op_max", COMMON_ADD_BOOK_DETAIL, 1100, True, True, False, FORMAT_PRICE
            
            .AddCol "Format Code", "format_code", COMMON_ADD_BOOK_DETAIL, 650, False, True, True
            .AddCol "Indent Code", "indent_code", COMMON_ADD_BOOK_DETAIL, 650, False, True, True
            
            .AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
            .AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
            .AddCol "last_update_id", "last_update_id", COMMON_ADD_BOOK_DETAIL, 800, True, False, False
        End If
        '
        '   Call base class.
        If m_blnPreferences Then
            .BuildPreferenceGrid
        Else
            .BuildGrid True, False
        End If
    End With
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

'*** APEX Migration Utility Code Change ***
'Public Sub SetGrid(grd As TrueOleDBGrid70.TDBGrid)
Public Sub SetGrid(grd As TrueOleDBGrid80.TDBGrid)
    ' Pass-through to base class
    m_objGridMap.SetGrid grd
    Set TDBGrid = grd
End Sub

Public Function IsPendingChange() As Boolean
    IsPendingChange = m_objGridMap.IsPendingChange
End Function
'
'   Sets whether or not the user wants the row to wrap in the grid.
Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub

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

Public Property Get bInsertInProcess()
    bInsertInProcess = bInserting
End Property

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

Public Property Let MaxSortOrder(nMaxSortOrder As Integer)
    m_nMaxSortOrder = nMaxSortOrder + 5
End Property

Public Property Get GRIDTYPE() As String
    GRIDTYPE = "Common Additives"
End Property

Public Sub CloneRow()
    Dim vntBookmark
    vntBookmark = TDBGrid.Bookmark
    '
    '   Pass to base class
    m_objGridMap.CloneRow
End Sub

Public Function CloneRowRecordset() As ADODB.RecordSet
    Set CloneRowRecordset = m_objGridMap.CloneRowRecordset
End Function

Public Function Update(sbldg_skey As String, stype_code As String, nbldg_form As Integer) As Boolean
    Dim strUpdate As String ' SQL string
    Dim sBookDesc   As String
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates

    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0
    With m_rec
        If .State <> adStateClosed Then
            .MoveFirst
            While Not .EOF
                ' Skip the record if it didn't change
                Dim clsRowInfo As CRowInfo
                Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
                '   If record was added meaning clsRowInfo.State = STATE_NEW
                '   Or clsRowInfo.State = STATE_MODIFIED Then
                If Not clsRowInfo Is Nothing Then
                    If clsRowInfo.State <> STATE_NONE Then
                        
                        strUpdate = "exec sp_update_common_add_book_detail_holding_table @common_add_skey = '"
                         
                        If Len(Trim(.Fields("Common_Add_Skey").Value)) = 0 Then
                           strUpdate = strUpdate & "0',"
                        Else
                            strUpdate = strUpdate & .Fields("Common_Add_Skey").Value & "',"
                        End If
                        '
                        '   If the user is inserting we wouldn't have the bldg_skey & type_code
                        '   so need it passed in.
                        strUpdate = strUpdate & "@bldg_skey  = '" & sbldg_skey & "',"
                        strUpdate = strUpdate & "@type_code  = '" & stype_code & "',"
                        strUpdate = strUpdate & "@sort_order  = '" & .Fields("Sort_Order").Value & "',"
                        '
                        '   If we're updating Residential the skey_type must be E for exception.
                        If stype_code = "R" Then
                            If (.Fields("common_add_Skey_Type").Value) = "" Then
                                strUpdate = strUpdate & "@common_add_skey_type  = 'E',"
                            ElseIf (.Fields("common_add_Skey_Type").Value) <> "E" And (.Fields("ID").Value) = "" Then
                                strUpdate = strUpdate & "@common_add_skey_type  = 'E',"
                            Else
                                strUpdate = strUpdate & "@common_add_skey_type  = '" & .Fields("common_add_Skey_Type").Value & "',"
                            End If
                        Else
                            strUpdate = strUpdate & "@common_add_skey_type  = '" & .Fields("common_add_Skey_Type").Value & "',"
                        End If
                        strUpdate = strUpdate & "@common_add_child_skey  = '" & .Fields("common_add_Child_Skey").Value & "',"
                        strUpdate = strUpdate & "@common_add_child_skey_max  = '" & .Fields("common_add_Child_Skey_Max").Value & "',"
                        strUpdate = strUpdate & "@format_code  = '" & Trim(.Fields("Format_Code").Value) & "',"
                        strUpdate = strUpdate & "@indent_code   = '" & Trim(.Fields("Indent_Code").Value) & "',"
                        '
                        '   Since we might have ' marks in our book desc, need to replace with '' for SQL.
                        sBookDesc = Trim(.Fields("Book_Desc").Value)
                        sBookDesc = Replace(sBookDesc, "'", "''", 1)
                        strUpdate = strUpdate & "@book_desc  = '" & sBookDesc & "',"
                        
                        strUpdate = strUpdate & "@unit  = '" & Trim(.Fields("Unit").Value) & "',"
                        strUpdate = strUpdate & "@pct_ind  = '0'," ' & .Fields("pct_ind").Value & "',"
    
                        strUpdate = strUpdate & "@mat_cost_op = '" & IIf(Trim(.Fields("mat_cost_op").Value) = "", "0", Trim(.Fields("mat_cost_op").Value)) & "',"
                        strUpdate = strUpdate & "@inst_cost_op = '" & IIf(Trim(.Fields("inst_cost_op").Value) = "", "0", Trim(.Fields("inst_cost_op").Value)) & "',"
                        strUpdate = strUpdate & "@total_cost_op = '" & IIf(Trim(.Fields("total_cost_op").Value) = "", "0", Trim(.Fields("total_cost_op").Value)) & "',"
                        strUpdate = strUpdate & "@mat_cost_op_max = '" & IIf(Trim(IsNull(.Fields("mat_cost_op_max").Value)), "0", Trim(.Fields("mat_cost_op_max").Value)) & "',"
                        strUpdate = strUpdate & "@inst_cost_op_max = '" & IIf(Trim(IsNull(.Fields("inst_cost_op_max").Value)), "0", Trim(.Fields("inst_cost_op_max").Value)) & "',"
                        strUpdate = strUpdate & "@total_cost_op_max = '" & IIf(Trim(IsNull(.Fields("total_cost_op_max").Value)), "0", Trim(.Fields("total_cost_op_max").Value)) & "',"
                        '
                        '   Set the last_update_id to 0 for new's in common_add_book_detail_holding_table
                        If Len(Trim(.Fields("last_update_id").Value)) = 0 Then
                            strUpdate = strUpdate & "@last_update_id = '0',"
                        Else
                            strUpdate = strUpdate & "@last_update_id = '" & Trim(.Fields("last_update_id").Value) & "',"
                        End If
                        strUpdate = strUpdate & "@bldg_form = '" & nbldg_form & "'"
                        '
                        ' Reset on success
                        If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
                            If strError = "" Then
                                intSuccess = intSuccess + 1
                                clsRowInfo.State = STATE_NONE
                                .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                                .Fields("last_update_date").Value = Now
                                .Fields("last_update_person").Value = strUserName
                                Reset_Orig_Values m_rec
                            Else
                                m_objGridMap.SetError m_rec.Bookmark, strError
                                m_rec.Fields("Error").Value = "Error"
                                Update = False
                                intErrors = intErrors + 1
                            End If
                        Else
                            m_objGridMap.SetError m_rec.Bookmark, strError
                            m_rec.Fields("Error").Value = "Error"
                            Update = False
                            intErrors = intErrors + 1
                        End If
                    ' If the row was marked as changed but the user typed the same data
                    ' make sure the error condition is removed from the row
                    End If
                End If
                .MoveNext
            Wend
            .UpdateBatch
        End If
    End With
    TDBGrid.ReBind
End Function

Public Sub SetupInitialSortOrder()
    Dim i       As Integer
    
    On Error Resume Next
    
    With TDBGrid
        i = 1
        Do Until i > .ApproxCount
            m_colSortOrderItems.Add Trim(.Columns("Sort Order").CellText(i)), Trim(.Columns("Sort Order").CellText(i))
            i = i + 1
            Dim duh As String
            duh = .Bookmark
        Loop
        If .ApproxCount = 0 Then
            m_nMaxSortOrder = "5"
        Else
            m_nMaxSortOrder = IIf(Trim(IsNull(.Columns("Sort Order").CellText(i - 1))), "5", Trim(.Columns("Sort Order").CellText(i - 1)) + 5)
        End If
    End With
End Sub

Private Sub SetSortOrder()
    Dim i           As Integer
    Dim varBookmark As Variant
    
    On Error Resume Next
    Set m_colSortOrderItems = Nothing
    Set m_colSortOrderItems = New Collection
    
    With TDBGrid
        varBookmark = .Bookmark
        .MoveLast
        m_nMaxSortOrder = IIf(Trim(IsNull(.Columns("Sort Order").Value)), "5", Trim(.Columns("Sort Order").Value) + 5)

        i = 1
        Do Until i > .ApproxCount
            m_colSortOrderItems.Add Trim(.Columns("Sort Order").CellText(i)), Trim(.Columns("Sort Order").CellText(i))
            If m_nMaxSortOrder < Trim(.Columns("Sort Order").CellText(i)) Then
                m_nMaxSortOrder = Trim(.Columns("Sort Order").CellText(i))
            ElseIf m_nMaxSortOrder = Trim(.Columns("Sort Order").CellText(i)) Then
                m_nMaxSortOrder = m_nMaxSortOrder + 5
            End If
            i = i + 1
        Loop
        .Bookmark = varBookmark
    End With
End Sub

'**********GRID STUFF***************
Private Sub TDBGrid_AfterDelete()
    SetSortOrder
End Sub

Private Sub TDBGrid_AfterInsert()
    bInserting = False
    SetSortOrder
End Sub

Private Sub TDBGrid_AfterUpdate()
    bInserting = False
    SetSortOrder
End Sub

Private Sub TDBGrid_OnAddNew()
    bInserting = True
End Sub

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim strSELECT           As String
    Dim ID                  As String
    Dim recAssemblyUnitCost As New ADODB.RecordSet
    Dim i                   As Integer
    Dim sSortOrder          As String
    Dim nFound              As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With TDBGrid
        If .Columns(ColIndex).Caption = "ID" And Trim(.Columns(ColIndex).Value) <> "" Then
            If .Columns("ID").Value <> OldValue Then
                If Len(Trim(.Columns("ID").Value)) = 14 Then
                    ID = Left$(.Columns("ID").Value, 5) & Right$(Left$(.Columns("ID").Value, 9), 3) & Right$(.Columns("ID").Value, 4)
                Else
                    ID = Trim(.Columns("ID").Value)
                End If
                '
                '   We need to search for a new assembly/unit cost.
                strSELECT = "exec sp_select_model_assemblies @assembly_id = '" & ID & "'"
                
                If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recAssemblyUnitCost) Then
                    Screen.MousePointer = vbNormal
                    MsgBox "An error occurred while searching for assembly/unit cost.", vbCritical
                Else
                    If Not recAssemblyUnitCost.EOF Then
                        .Columns("Skey Type").Value = "A"
                        .Columns("Child Skey").Value = Trim(recAssemblyUnitCost.Fields("assembly_skey").Value)
                        .Columns("ID").Value = ID
                        
                        If .Columns("Sort Order").Value = "" Then
                            .Columns("Sort Order").Value = m_nMaxSortOrder
                        Else
                            .Columns("Sort Order").Value = .Columns("Sort Order").Value
                        End If
                        
                        'rlh (new lines)
                        .Columns("type_code") = recAssemblyUnitCost.Fields("type_code")             'rlh
                        .Columns("common_add_skey") = recAssemblyUnitCost.Fields("common_add_skey") 'rlh
                        .Columns("child_skey_max") = 0                                              'rlh
                        'rlh (end of new lines block
                        
                        .Columns("book desc").Value = Trim(recAssemblyUnitCost.Fields("book_desc").Value)
                        .Columns("tech desc").Value = Trim(recAssemblyUnitCost.Fields("tech_desc").Value)
                        .Columns("unit").Value = Trim(recAssemblyUnitCost.Fields("unit").Value)
                        '.Columns("pct ind").Value = Trim(recAssemblyUnitCost.Fields("pct_ind").Value)
                        .Columns("mat cost op").Value = Trim(recAssemblyUnitCost.Fields("std_mat_cost_op").Value)
                        .Columns("inst cost op").Value = Trim(recAssemblyUnitCost.Fields("std_inst_cost_op").Value)
                        .Columns("total cost op").Value = Trim(recAssemblyUnitCost.Fields("std_total_cost_op").Value)
                        
                        '(rlh) 10/10/2006   (new lines)
                        '       These setting were failing in the sp at line "0"
                        '       blanks were being assigned to DECIMAL (12,2) fields
                        '       resulting in immediate failure!
                        'Max Mat Cost  (rlh)
                        .Columns("Max Mat Cost OP").Value = CDec("0.00")    'rlh
                        'Max Inst Cost  (rlh)
                        .Columns("Max Inst Cost OP").Value = CDec("0.00")   'rlh
                        'Max Total  (rlh)
                        .Columns("Max Total Cost OP").Value = CDec("0.00")  'rlh
                        'rlh  (end of new lines block)
                    Else
                        recAssemblyUnitCost.Close
                        strSELECT = "exec sp_select_unit_cost @start_unit_cost_id = '"
                        strSELECT = strSELECT & ID & "', @end_unit_cost_id = '" & ID
                        strSELECT = strSELECT & "', @alt_unit_cost_id = '%', @tech_desc = '%'"

                        If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recAssemblyUnitCost) Then
                            Screen.MousePointer = vbNormal
                            MsgBox "An error occurred while searching for assembly/unit cost.", vbCritical
                        Else
                            If Not recAssemblyUnitCost.EOF Then
                                .Columns("Skey Type").Value = "U"
                                .Columns("Child Skey").Value = Trim(recAssemblyUnitCost.Fields("unit_cost_skey").Value)
                                
                                If .Columns("Sort Order").Value = "" Then
                                    .Columns("Sort Order").Value = m_nMaxSortOrder
                                Else
                                    .Columns("Sort Order").Value = .Columns("Sort Order").Value
                                End If
                                
                                .Columns("ID").Value = ID
                                .Columns("book desc").Value = Trim(recAssemblyUnitCost.Fields("book_desc").Value)
                                .Columns("tech desc").Value = Trim(recAssemblyUnitCost.Fields("tech_desc").Value)
                                .Columns("unit").Value = Trim(recAssemblyUnitCost.Fields("unit").Value)
                                '.Columns("pct ind").Value = Trim(recAssemblyUnitCost.Fields("pct_ind").Value)
                                .Columns("mat cost op").Value = Trim(recAssemblyUnitCost.Fields("std_mat_cost_op").Value)
                                .Columns("inst cost op").Value = Val(Trim(recAssemblyUnitCost.Fields("std_labor_cost_op").Value)) + Val(Trim(recAssemblyUnitCost.Fields("std_equip_cost_op").Value))
                                .Columns("total cost op").Value = Trim(recAssemblyUnitCost.Fields("std_total_cost_op").Value)
                            Else
                                '
                                '   Appears that the user has entered an invalid ID,
                                '   raise message.
                                MsgBox "ID not found in either Assemblies or Unit Costs.  Please provide a valid ID.", vbCritical
                                Cancel = True
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf .Columns(ColIndex).Caption = "ID Max" And Trim(.Columns(ColIndex).Value) <> "" Then
            If .Columns("ID Max").Value <> OldValue Then
                If Len(Trim(.Columns("ID Max").Value)) = 14 Then
                    ID = Left$(.Columns("ID Max").Value, 5) & Right$(Left$(.Columns("ID Max").Value, 9), 3) & Right$(.Columns("ID Max").Value, 4)
                Else
                    ID = Trim(.Columns("ID Max").Value)
                End If

                If .Columns("Skey Type").Value = "A" Then
                    '
                    '   We need to search for Child Skey Max for assembly/unit cost.
                    strSELECT = "exec sp_select_model_assemblies @assembly_id = '" & ID & "'"

                    If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recAssemblyUnitCost) Then
                        Screen.MousePointer = vbNormal
                        MsgBox "An error occurred while searching for assembly/unit cost.", vbCritical
                    Else
                        If Not recAssemblyUnitCost.EOF Then
                            .Columns("Child Skey Max").Value = Trim(recAssemblyUnitCost.Fields("assembly_skey").Value)
                            .Columns("ID Max").Value = ID
                            
                            If .Columns("Sort Order").Value = "" Then
                                .Columns("Sort Order").Value = m_nMaxSortOrder
                            Else
                                .Columns("Sort Order").Value = .Columns("Sort Order").Value
                            End If
                            
                            .Columns("max mat cost op").Value = Trim(recAssemblyUnitCost.Fields("std_mat_cost_op").Value)
                            .Columns("max inst cost op").Value = Trim(recAssemblyUnitCost.Fields("std_inst_cost_op").Value)
                            .Columns("max total cost op").Value = Trim(recAssemblyUnitCost.Fields("std_total_cost_op").Value)
                        Else
                            '
                            '   Appears that the user has entered an invalid ID,
                            '   raise message.
                            MsgBox "ID not found in Assemblies.  Please provide a valid ID.", vbCritical
                            Cancel = True
                        End If
                    End If
                ElseIf .Columns("Skey Type").Value = "U" Then
                    '
                    '   We need to search for Child Skey Max for assembly/unit cost.
                    'strSelect = "SELECT unit_cost_skey FROM unit_cost_detail WHERE unit_cost_id = '"
                    'strSelect = strSelect & ID & "'"
                    strSELECT = "exec sp_select_unit_cost @start_unit_cost_id = '"
                    strSELECT = strSELECT & ID & "', @end_unit_cost_id = '" & ID
                    strSELECT = strSELECT & "', @alt_unit_cost_id = '%', @tech_desc = '%'"

                    If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recAssemblyUnitCost) Then
                        Screen.MousePointer = vbNormal
                        MsgBox "An error occurred while searching for assembly/unit cost.", vbCritical
                    Else
                        If Not recAssemblyUnitCost.EOF Then
                            .Columns("Child Skey Max").Value = Trim(recAssemblyUnitCost.Fields("unit_cost_skey").Value)
                            .Columns("ID Max").Value = ID

                            If .Columns("Sort Order").Value = "" Then
                                .Columns("Sort Order").Value = m_nMaxSortOrder
                            Else
                                .Columns("Sort Order").Value = .Columns("Sort Order").Value
                            End If

                            .Columns("max mat cost op").Value = Trim(recAssemblyUnitCost.Fields("std_mat_cost_op").Value)
                            .Columns("max inst cost op").Value = Val(Trim(recAssemblyUnitCost.Fields("std_labor_cost_op").Value)) + Val(Trim(recAssemblyUnitCost.Fields("std_equip_cost_op").Value))
                            .Columns("max total cost op").Value = Trim(recAssemblyUnitCost.Fields("std_total_cost_op").Value)
                        Else
                            '
                            '   Appears that the user has entered an invalid ID,
                            '   raise message.
                            MsgBox "ID not found in Unit Costs.  Please provide a valid ID.", vbCritical
                            Cancel = True
                        End If
                    End If
                Else
                    '
                    '   We need to search for Child Skey Max until we find it.
                    'strSelect = "SELECT assembly_Skey FROM assembly_detail WHERE assembly_id = '"
                    'strSelect = strSelect & ID & "'"
                    strSELECT = "exec sp_select_model_assemblies @assembly_id = '" & ID & "'"
                
                    If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recAssemblyUnitCost) Then
                        Screen.MousePointer = vbNormal
                        MsgBox "An error occurred while searching for assembly/unit cost.", vbCritical
                    Else
                        If Not recAssemblyUnitCost.EOF Then
                            .Columns("Skey Type").Value = "A"
                            .Columns("Child Skey Max").Value = Trim(recAssemblyUnitCost.Fields("assembly_skey").Value)
                            .Columns("ID Max").Value = ID

                            If .Columns("Sort Order").Value = "" Then
                                .Columns("Sort Order").Value = m_nMaxSortOrder
                            Else
                                .Columns("Sort Order").Value = .Columns("Sort Order").Value
                            End If

                            .Columns("max mat cost op").Value = Trim(recAssemblyUnitCost.Fields("std_mat_cost_op").Value)
                            .Columns("max inst cost op").Value = Trim(recAssemblyUnitCost.Fields("std_inst_cost_op").Value)
                            .Columns("max total cost op").Value = Trim(recAssemblyUnitCost.Fields("std_total_cost_op").Value)
                        Else
                            recAssemblyUnitCost.Close
                            strSELECT = "exec sp_select_unit_cost @start_unit_cost_id = '"
                            strSELECT = strSELECT & ID & "', @end_unit_cost_id = '" & ID
                            strSELECT = strSELECT & "', @alt_unit_cost_id = '%', @tech_desc = '%'"

                            If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recAssemblyUnitCost) Then
                                Screen.MousePointer = vbNormal
                                MsgBox "An error occurred while searching for assembly/unit cost.", vbCritical
                            Else
                                If Not recAssemblyUnitCost.EOF Then
                                    .Columns("Skey Type").Value = "U"
                                    .Columns("Child Skey Max").Value = Trim(recAssemblyUnitCost.Fields("unit_cost_skey").Value)
                                    .Columns("ID Max").Value = ID

                                    If .Columns("Sort Order").Value = "" Then
                                        .Columns("Sort Order").Value = m_nMaxSortOrder
                                    Else
                                        .Columns("Sort Order").Value = .Columns("Sort Order").Value
                                    End If

                                    .Columns("max mat cost op").Value = Trim(recAssemblyUnitCost.Fields("std_mat_cost_op").Value)
                                    .Columns("max inst cost op").Value = Val(Trim(recAssemblyUnitCost.Fields("std_labor_cost_op").Value)) + Val(Trim(recAssemblyUnitCost.Fields("std_equip_cost_op").Value))
                                    .Columns("max total cost op").Value = Trim(recAssemblyUnitCost.Fields("std_total_cost_op").Value)
                                Else
                                    '
                                    '   Appears that the user has entered an invalid ID,
                                    '   raise message.
                                    MsgBox "ID not found in either Assemblies or Unit Costs.  Please provide a valid ID.", vbCritical
                                    Cancel = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf .Columns(ColIndex).Caption = "Book Desc" And Trim(.Columns("Sort Order").Value) = "" Then
            If .Columns("Sort Order").Value = "" Then
                .Columns("Sort Order").Value = m_nMaxSortOrder
            Else
                .Columns("Sort Order").Value = .Columns("Sort Order").Value
            End If
        End If
        '
        '   If we have changed the Sort Order we need to make sure that the
        '   sort order is not already in use.
        If .Columns(ColIndex).Caption = "Sort Order" And Trim(.Columns(ColIndex).Value) <> "" Then
            sSortOrder = Trim(.Columns(ColIndex).Value)
            
            Err.Clear
            m_colSortOrderItems.Add sSortOrder, sSortOrder
            If Err.Number <> 0 Then
                MsgBox "Please provide a numeric sort order that is not already in use.", vbCritical
                Cancel = True
            End If
        End If
    End With
    
'    Debug.Print "* * * (CBldgComAdds.cls) TDBGRID GRID COLUMNS * * *"      'rlh
'    Debug.Print "   "                                   'rlh
'    For I = 0 To TDBGrid.Columns.Count - 1              'rlh
'        Debug.Print I & "  TdbGrid.Column.Caption: " & TDBGrid.Columns(I).Caption & vbTab & " Value: " & TDBGrid.Columns(I).Value
'    Next    'rlh
    
    'Me.SetRowStateNew       'rlh
    recAssemblyUnitCost.Close
    Screen.MousePointer = vbNormal
End Sub

Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    Dim strUpdate       As String
    Dim blnReturn       As Boolean
    Dim strError        As String
    Dim clsRowInfo      As CRowInfo
    
    On Error Resume Next
    '
    '   Position to the desired record and execute delete
    m_rec.Bookmark = Bookmark
    
    With m_objGridMap
        Set clsRowInfo = .rowinfo(m_rec.Bookmark)
            
        If Not clsRowInfo Is Nothing Then
            If clsRowInfo.State = STATE_NEW Then
                TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
                .SetRowState m_rec.Bookmark, STATE_DELETED
                '
                '   Remove it from recordset so it doesn't show.
                m_rec.Delete
                If bMultiDeleteInProcess = False Then
                    MsgBox "Common Add deleted successfully", vbInformation
                Else
                    m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                End If
            End If
        Else
            '
            '   Build SQL statement.
            '
            strUpdate = "exec sp_delete_common_add_book_detail @common_add_skey = '" _
                    & Trim(m_rec.Fields("common_add_skey").Value) & "'"
            
            If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
                If strError = "" Then
                    TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
                    .SetRowState m_rec.Bookmark, STATE_DELETED
                    '
                    '   Remove it from recordset so it doesn't show.
                    m_rec.Delete
                    If bMultiDeleteInProcess = False Then
                        MsgBox "Common Add deleted successfully", vbInformation
                    Else
                        m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                    End If
                Else
                    .SetError m_rec.Bookmark, strError
                    m_rec.Fields("Error").Value = "Error"
                End If
            Else
                .SetError m_rec.Bookmark, strError
                m_rec.Fields("Error").Value = "Error"
            End If
        End If
    End With
End Sub

Private Sub TDBGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim varButton
    Dim i                           As Integer
    
    On Error Resume Next
    With TDBGrid
        '
        '   Intercept the delete key - prompt user.
        If KeyCode = 46 Then
            If .SelBookmarks.Count = 1 Then
                varButton = MsgBox(CStr(.SelBookmarks.Count) + " records will be deleted.  Are you sure you want to delete this row(s) permanently?", vbYesNo + vbCritical)
                If varButton <> vbYes Then
                    KeyCode = 0
                End If
            '
            '   If multiple records are selected.
            ElseIf .SelBookmarks.Count > 1 Then
                varButton = MsgBox(CStr(.SelBookmarks.Count) + " records will be deleted.  Are you sure you want to delete this row(s) permanently?", vbYesNo + vbCritical)
                If varButton <> vbYes Then
                    KeyCode = 0
                Else
                    Delete
                End If
            End If
        End If
    End With
End Sub

Public Function Delete() As Boolean
    Dim varSelectedBookmarks()      As Variant
    Dim varCurrentSelectedBookmark  As Variant
    Dim varButton
    Dim i                           As Integer
    
    On Error Resume Next
    With TDBGrid
        If .SelBookmarks.Count > 0 Then
            bMultiDeleteInProcess = True
            m_NumberOfSuccessfulUpdates = 0

            ReDim varSelectedBookmarks(0 To .SelBookmarks.Count)
            i = 0
            For Each varCurrentSelectedBookmark In .SelBookmarks
                varSelectedBookmarks(i) = varCurrentSelectedBookmark
                i = i + 1
            Next varCurrentSelectedBookmark
            For i = 0 To UBound(varSelectedBookmarks) - 1
                .Bookmark = varSelectedBookmarks(i)
                .Delete
            Next i
            MsgBox m_NumberOfSuccessfulUpdates & " Common Adds were deleted successfully."
            bMultiDeleteInProcess = False
            m_NumberOfSuccessfulUpdates = 0
        End If
    End With
End Function
Public Function SetRowStateNew() As Boolean     'rlh
    If m_rec.RecordCount > 0 Then
        m_rec.MoveFirst
        While Not m_rec.EOF
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
            m_rec.MoveNext
        Wend
    End If
    
End Function
