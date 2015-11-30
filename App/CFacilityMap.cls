VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBuildingMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' <modulename> CBuildingMap.cls</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Helper class module for frmBuildingGrid.frm (frmFacilityGrid.frm)
''' </summary>
''' <seealso>frmAssemblyGrid</seealso>
''' <datastruct> TDBGRID</datastruct>
''' <datastruct> m_objGridMap</datastruct>
'''<datastruct> m_rec</datastruct>
'''
'''<storedprocedurename> sp_update_commercial_building</storedprocedurename>
'''<storedprocedurename> sp_update_residential_building</storedprocedurename>
'''<storedprocedurename> sp_update_bldg_model</storedprocedurename>
'''
'''<storedprocedurename> sp_update_bldg_model_resi</storedprocedurename>
'''
'''
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''exec sp_update_commercial_building @bldg_skey = '274',@bldg_id = '670',@bldg_category = 'Commercial',@bldg_desc = 'Apartment, 1-3 Story',@bldg_stories = 3,@bldg_stories_hgt = 10,@bldg_part_density = 9,@bldg_part_hgt = 8,@bldg_door_density = 80,
'''@bldg_type = '0',@bldg_area_std = 22500,@bldg_perimeter_std = 400,@bldg_wall_factor = 0.88,@bldg_elev_no = 9,@bldg_fixture_area = 200,@window_area = 15,@op_factor = 0.25,@architect_fee = 0.08,@row_to_bold = 2,@col_to_bold = 5,@graphic_ref_id = 'b200-1f',@graphic_ref_id2 = '',
'''@last_update_id_bldg = '1',@bldg_area_1 = '0',@bldg_perimeter_1 = '0',@bldg_orig_area_1 = '0',@area_ind_1 = '0',@last_update_id_area_1 = '0',@bldg_area_2 = '0',@bldg_perimeter_2 = '0',@bldg_orig_area_2 = '0',@area_ind_2 = '0',@last_update_id_area_2 = '0',@bldg_area_3 = '0',@bldg_perimeter_3 = '0',
'''@bldg_orig_area_3 = '0',@area_ind_3 = '0',@last_update_id_area_3 = '0',@bldg_area_4 = '0',@bldg_perimeter_4 = '0',@bldg_orig_area_4 = '0',@area_ind_4 = '0',@last_update_id_area_4 = '0',@bldg_area_5 = '
'''@bldg_perimeter_5 = '0',@bldg_orig_area_5 = '0',@area_ind_5 = '0',@last_update_id_area_5 = '0',@bldg_area_6 = '0',@bldg_perimeter_6 = '0',@bldg_orig_area_6 = '0',@area_ind_6 = '0',@last_update_id_area_6 = '0',@bldg_area_7 = '0',@bldg_perimeter_7 = '0',@bldg_orig_area_7 = '0',@area_ind_7 = '0',
'''@last_update_id_area_7 = '0',@bldg_area_8 = '0',@bldg_perimeter_8 = '0',@bldg_orig_area_8 = '0',@area_ind_8 = '0',@last_update_id_area_8 = '0',@bldg_area_9 = '0',@bldg_perimeter_9 = '0',@bldg_orig_area_9 = '0',@area_ind_9 = '0',@last_update_id_area_9 = '0',@bldg_form = '0', @last_update_person = 'Hancockrl'
'''</code>
''' <code>
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  frmBuildingGrid.frm
'''2.  CGridMap.cls
'''3.  CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''</dependson>
'''


'
'   Access to grid events.
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid
Attribute TDBGrid.VB_VarHelpID = -1
'
'   Names of some of the tables this grid uses.
Const BUILDING_TABLE = "bldg_detail"
Const ALL_TABLES = "All"
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
'
'   Indicates a field that will affect overall cost rollups
'   has been changed so we must RefreshCosts if they update.
Dim bRefreshCosts As Boolean

Private Sub Class_Initialize()
    With m_objGridMap.m_colCloneCols
        .Add 1, "bldg_skey"
        .Add 1, "bldg_id"
        .Add 1, "type_code"
        .Add 1, "bldg_category"
        .Add 1, "bldg_desc"
        .Add 1, "bldg_stories"
        .Add 1, "bldg_stories_hgt"
        .Add 1, "bldg_part_density"
        .Add 1, "bldg_part_hgt"
        .Add 1, "bldg_door_density"
        .Add 1, "bldg_type"
        .Add 1, "bldg_area_std"
        .Add 1, "bldg_perimeter_std"
        .Add 1, "bldg_wall_factor"
        .Add 1, "bldg_elev_no"
        .Add 1, "bldg_fixture_area"
        .Add 1, "window_area"
        .Add 1, "op_factor"
        .Add 1, "architect_fee"
        .Add 1, "row_to_bold"
        .Add 1, "col_to_bold"
        .Add 1, "graphic_ref_id"
        .Add 1, "graphic_ref_id2"
        .Add 1, "last_update_date"
        .Add 1, "last_update_person"
        .Add 1, "last_update_id"
    End With
End Sub

Public Sub InitGrid()
    '
    '   Identifies the Grid.
    With m_objGridMap
        .GRIDTYPE = "Building"
        '
        '   Add the columns to the grid, pointing to their
        '   corresponding db field names so that the grid can
        '   populate directly from the m_rec recordset.
        .AddCol "SKEY", "bldg_skey", BUILDING_TABLE, 800, True, False, True
        .AddCol "Bldg ID", "bldg_id", BUILDING_TABLE, 800, True, True
        .AddCol "Type", "type_code", BUILDING_TABLE, 500, True, True
        .AddCol "Category", "bldg_category", BUILDING_TABLE, 1250, True, True
        .AddCol "Bldg Desc", "bldg_desc", BUILDING_TABLE, 2000, , , , FORMAT_STRING_TRIM
        '
        '   Main Split, this grid only has 1.
        '
        .AddSplit
        '
        '   Note: the architect-fee col is
        '   not displayed since it is a final total paid to the architect.
        .AddCol "Stories", "bldg_stories", BUILDING_TABLE, 1000
        .AddCol "Stories Hgt", "bldg_stories_hgt", BUILDING_TABLE, 1000
        .AddCol "Part Density", "bldg_part_density", BUILDING_TABLE, 1000
        .AddCol "Part Hgt", "bldg_part_hgt", BUILDING_TABLE, 900, , True, True
        .AddCol "Dr Density", "bldg_door_density", BUILDING_TABLE, 900
        .AddCol "Elev No", "bldg_elev_no", BUILDING_TABLE, 1000
        .AddCol "Wall Fact", "bldg_wall_factor", BUILDING_TABLE, 1000
        .AddCol "OP Fact", "op_factor", BUILDING_TABLE, 1000
        .AddCol "Arch Fees", "architect_fee", BUILDING_TABLE, 900
        
        .AddCol "Fixture Area", "bldg_fixture_area", BUILDING_TABLE, 1000
        .AddCol "Win Area", "window_area", BUILDING_TABLE, 1300
        
        .AddCol "Std Area", "bldg_area_std", BUILDING_TABLE, 900, True
        .AddCol "Std Perim.", "bldg_perimeter_std", BUILDING_TABLE, 900, True
        .AddCol "Row Bold", "row_to_bold", BUILDING_TABLE, 1000, True
        .AddCol "Col Bold", "col_to_bold", BUILDING_TABLE, 1000, True
        .AddCol "Graphic ID", "graphic_ref_id", BUILDING_TABLE, 2000
        .AddCol "Graphic ID 2", "graphic_ref_id2", BUILDING_TABLE, 2000
        .AddCol "Bldg Type", "bldg_type", BUILDING_TABLE, 900, True
        .AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
        .AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
        .AddCol "last_update_id", "last_update_id", BUILDING_TABLE, 800, True, False, False
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
        .Splits(0).Caption = "Building"
        .Splits(1).Caption = "Attributes"
        .Splits(1).Size = 5
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
'
'   Sets whether or not the user wants the row to wrap in the grid.
Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub

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

Public Sub Update()

    On Error Resume Next
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
                '
                '   If record was added meaning clsRowInfo.State = STATE_NEW
                '   Or clsRowInfo.State = STATE_MODIFIED Then
                If Not clsRowInfo Is Nothing Then
                    If clsRowInfo.State <> STATE_NONE Then
                        If Trim(.Fields("type_code").Value) = "C" Then
                            If UpdateCommercial Then
                                clsRowInfo.State = STATE_NONE
                            End If
                        Else
                            If UpdateResidential Then
                                clsRowInfo.State = STATE_NONE
                            End If
                        End If
                    End If
                End If
                .MoveNext
            Wend
            .UpdateBatch
        End If
    End With
    TDBGrid.ReBind
End Sub

Private Function UpdateCommercial() As Boolean
    Dim strUpdate As String ' SQL string
    Dim sBldgDesc   As String
    Dim strError As String ' Error string returned from DAL
    
    On Error Resume Next
    With m_rec
        strUpdate = "exec sp_update_commercial_building @bldg_skey = '" & Trim(.Fields("bldg_skey").Value) & "',"
        strUpdate = strUpdate & "@bldg_id = '" & Trim(.Fields("bldg_id").Value) & "',"
        strUpdate = strUpdate & "@bldg_category = '" & Trim(.Fields("bldg_category").Value) & "',"
        '
        '   Since we might have ' marks in our book desc, need to replace with '' for SQL.
        sBldgDesc = Trim(.Fields("bldg_desc").Value)
        sBldgDesc = Replace(sBldgDesc, "'", "''", 1)
        strUpdate = strUpdate & "@bldg_desc = '" & sBldgDesc & "',"
        strUpdate = strUpdate & "@bldg_stories = " & Trim(.Fields("bldg_stories").Value) & ","
        strUpdate = strUpdate & "@bldg_stories_hgt = " & Trim(.Fields("bldg_stories_hgt").Value) & ","
        strUpdate = strUpdate & "@bldg_part_density = " & Trim(.Fields("bldg_part_density").Value) & ","
        strUpdate = strUpdate & "@bldg_part_hgt = " & Trim(.Fields("bldg_part_hgt").Value) & ","
        strUpdate = strUpdate & "@bldg_door_density = " & Trim(.Fields("bldg_door_density").Value) & ","
        strUpdate = strUpdate & "@bldg_type = '" & Trim(.Fields("bldg_type").Value) & "',"
        strUpdate = strUpdate & "@bldg_area_std = " & Trim(.Fields("bldg_area_std").Value) & ","
        strUpdate = strUpdate & "@bldg_perimeter_std = " & Trim(.Fields("bldg_perimeter_std").Value) & ","
        strUpdate = strUpdate & "@bldg_wall_factor = " & Trim(.Fields("bldg_wall_factor").Value) & ","
        strUpdate = strUpdate & "@bldg_elev_no = " & Trim(.Fields("bldg_elev_no").Value) & ","
        strUpdate = strUpdate & "@bldg_fixture_area = " & Trim(.Fields("bldg_fixture_area").Value) & ","
        strUpdate = strUpdate & "@window_area = " & Trim(.Fields("window_area").Value) & ","
        strUpdate = strUpdate & "@op_factor = " & Trim(.Fields("op_factor").Value) & ","
        strUpdate = strUpdate & "@architect_fee = " & Trim(.Fields("architect_fee").Value) & ","
        strUpdate = strUpdate & "@row_to_bold = " & Trim(.Fields("row_to_bold").Value) & ","
        strUpdate = strUpdate & "@col_to_bold = " & Trim(.Fields("col_to_bold").Value) & ","
        strUpdate = strUpdate & "@graphic_ref_id = '" & Trim(.Fields("graphic_ref_id").Value) & "',"
        strUpdate = strUpdate & "@graphic_ref_id2 = '" & Trim(.Fields("graphic_ref_id2").Value) & "',"
        strUpdate = strUpdate & " @last_update_id_bldg = '" & Trim(m_rec.Fields("last_update_id").Value) & "',"
        
        strUpdate = strUpdate & "@bldg_area_1 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_1 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_1 = '0',"
        strUpdate = strUpdate & "@area_ind_1 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_1 = '0',"
        
        strUpdate = strUpdate & "@bldg_area_2 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_2 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_2 = '0',"
        strUpdate = strUpdate & "@area_ind_2 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_2 = '0',"
        
        strUpdate = strUpdate & "@bldg_area_3 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_3 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_3 = '0',"
        strUpdate = strUpdate & "@area_ind_3 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_3 = '0',"
        
        strUpdate = strUpdate & "@bldg_area_4 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_4 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_4 = '0',"
        strUpdate = strUpdate & "@area_ind_4 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_4 = '0',"
        strUpdate = strUpdate & "@bldg_area_5 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_5 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_5 = '0',"
        strUpdate = strUpdate & "@area_ind_5 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_5 = '0',"
        strUpdate = strUpdate & "@bldg_area_6 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_6 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_6 = '0',"
        strUpdate = strUpdate & "@area_ind_6 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_6 = '0',"
        strUpdate = strUpdate & "@bldg_area_7 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_7 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_7 = '0',"
        strUpdate = strUpdate & "@area_ind_7 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_7 = '0',"
        strUpdate = strUpdate & "@bldg_area_8 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_8 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_8 = '0',"
        strUpdate = strUpdate & "@area_ind_8 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_8 = '0',"
        strUpdate = strUpdate & "@bldg_area_9 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_9 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_9 = '0',"
        strUpdate = strUpdate & "@area_ind_9 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_9 = '0',"
        '
        '   common_add_book_detail -ensures that we only add records that our form inserted.
        strUpdate = strUpdate & "@bldg_form = '0',"
        strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"
        '
        ' Reset on success
        If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
            If strError = "" Then
                If bRefreshCosts Then
                    If RefreshCostsCommercial = "" Then
                        m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                        .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                        .Fields("last_update_date").Value = Now
                        .Fields("last_update_person").Value = strUserName
                        Reset_Orig_Values m_rec
                        UpdateCommercial = True
                    Else
                        m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
                        m_objGridMap.SetError m_rec.Bookmark, RefreshCostsCommercial
                        .Fields("Error").Value = "Error"
                    End If
                Else
                    m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                    .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                    .Fields("last_update_date").Value = Now
                    .Fields("last_update_person").Value = strUserName
                    Reset_Orig_Values m_rec
                    UpdateCommercial = True
                End If
            Else
                m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                .Fields("last_update_date").Value = Now
                .Fields("last_update_person").Value = strUserName
                Reset_Orig_Values m_rec
                UpdateCommercial = True
            End If
        Else
            m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
            m_objGridMap.SetError .Bookmark, strError
            .Fields("Error").Value = "Error"
        End If
    End With

End Function

Private Function UpdateResidential() As Boolean
    Dim strUpdate As String ' SQL string
    Dim sBldgDesc   As String
    Dim strError As String ' Error string returned from DAL
    
    With m_rec
        strUpdate = "exec sp_update_residential_building @bldg_skey = '" & Trim(.Fields("bldg_skey").Value) & "',"
        strUpdate = strUpdate & "@bldg_id = '" & Trim(.Fields("bldg_id").Value) & "',"
        strUpdate = strUpdate & "@bldg_category = '" & Trim(.Fields("bldg_category").Value) & "',"
        '
        '   Since we might have ' marks in our book desc, need to replace with '' for SQL.
        sBldgDesc = Trim(.Fields("bldg_desc").Value)
        sBldgDesc = Replace(sBldgDesc, "'", "''", 1)
        strUpdate = strUpdate & "@bldg_desc = '" & sBldgDesc & "',"
        strUpdate = strUpdate & "@bldg_stories = " & Trim(.Fields("bldg_stories").Value) & ","
        strUpdate = strUpdate & "@bldg_stories_hgt = " & Trim(.Fields("bldg_stories_hgt").Value) & ","
        strUpdate = strUpdate & "@bldg_part_density = " & Trim(.Fields("bldg_part_density").Value) & ","
        strUpdate = strUpdate & "@bldg_part_hgt = " & Trim(.Fields("bldg_part_hgt").Value) & ","
        strUpdate = strUpdate & "@bldg_door_density = " & Trim(.Fields("bldg_door_density").Value) & ","
        strUpdate = strUpdate & "@bldg_type = '" & Trim(.Fields("bldg_type").Value) & "',"
        strUpdate = strUpdate & "@bldg_area_std = " & Trim(.Fields("bldg_area_std").Value) & ","
        strUpdate = strUpdate & "@bldg_perimeter_std = " & Trim(.Fields("bldg_perimeter_std").Value) & ","
        strUpdate = strUpdate & "@bldg_wall_factor = " & Trim(.Fields("bldg_wall_factor").Value) & ","
        strUpdate = strUpdate & "@bldg_elev_no = " & Trim(.Fields("bldg_elev_no").Value) & ","
        strUpdate = strUpdate & "@bldg_fixture_area = " & Trim(.Fields("bldg_fixture_area").Value) & ","
        strUpdate = strUpdate & "@window_area = " & Trim(.Fields("window_area").Value) & ","
        strUpdate = strUpdate & "@op_factor = " & Trim(.Fields("op_factor").Value) & ","
        strUpdate = strUpdate & "@architect_fee = " & Trim(.Fields("architect_fee").Value) & ","
        strUpdate = strUpdate & "@row_to_bold = " & Trim(.Fields("row_to_bold").Value) & ","
        strUpdate = strUpdate & "@col_to_bold = " & Trim(.Fields("col_to_bold").Value) & ","
        strUpdate = strUpdate & "@graphic_ref_id = '" & Trim(.Fields("graphic_ref_id").Value) & "',"
        strUpdate = strUpdate & "@graphic_ref_id2 = '" & Trim(.Fields("graphic_ref_id2").Value) & "',"
        strUpdate = strUpdate & " @last_update_id_bldg = '" & Trim(m_rec.Fields("last_update_id").Value) & "',"
        
        strUpdate = strUpdate & "@bldg_area_1 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_1 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_1 = '0',"
        strUpdate = strUpdate & "@area_ind_1 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_1 = '0',"
        strUpdate = strUpdate & "@bldg_area_2 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_2 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_2 = '0',"
        strUpdate = strUpdate & "@area_ind_2 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_2 = '0',"
        strUpdate = strUpdate & "@bldg_area_3 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_3 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_3 = '0',"
        strUpdate = strUpdate & "@area_ind_3 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_3 = '0',"
        strUpdate = strUpdate & "@bldg_area_4 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_4 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_4 = '0',"
        strUpdate = strUpdate & "@area_ind_4 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_4 = '0',"
        strUpdate = strUpdate & "@bldg_area_5 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_5 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_5 = '0',"
        strUpdate = strUpdate & "@area_ind_5 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_5 = '0',"
        strUpdate = strUpdate & "@bldg_area_6 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_6 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_6 = '0',"
        strUpdate = strUpdate & "@area_ind_6 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_6 = '0',"
        strUpdate = strUpdate & "@bldg_area_7 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_7 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_7 = '0',"
        strUpdate = strUpdate & "@area_ind_7 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_7 = '0',"
        strUpdate = strUpdate & "@bldg_area_8 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_8 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_8 = '0',"
        strUpdate = strUpdate & "@area_ind_8 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_8 = '0',"
        strUpdate = strUpdate & "@bldg_area_9 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_9 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_9 = '0',"
        strUpdate = strUpdate & "@area_ind_9 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_9 = '0',"
        strUpdate = strUpdate & "@bldg_area_10 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_10 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_10 = '0',"
        strUpdate = strUpdate & "@area_ind_10 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_10 = '0',"
        strUpdate = strUpdate & "@bldg_area_11 = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_11 = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_11 = '0',"
        strUpdate = strUpdate & "@area_ind_11 = '0',"
        strUpdate = strUpdate & "@last_update_id_area_11 = '0',"
        '
        '   common_add_book_detail -ensures that we only add records that our form inserted.
        strUpdate = strUpdate & "@bldg_form = '0',"
        strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"
        '
        ' Reset on success
        If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
            If strError = "" Then
                If bRefreshCosts Then
                    If RefreshCostsResidential = "" Then
                        m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                        .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                        .Fields("last_update_date").Value = Now
                        .Fields("last_update_person").Value = strUserName
                        Reset_Orig_Values m_rec
                        UpdateResidential = True
                    Else
                        m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
                        m_objGridMap.SetError m_rec.Bookmark, RefreshCostsCommercial
                        .Fields("Error").Value = "Error"
                    End If
                Else
                    m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                    .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                    .Fields("last_update_date").Value = Now
                    .Fields("last_update_person").Value = strUserName
                    Reset_Orig_Values m_rec
                    UpdateResidential = True
                End If
            Else
                m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                .Fields("last_update_date").Value = Now
                .Fields("last_update_person").Value = strUserName
                Reset_Orig_Values m_rec
                UpdateResidential = True
            End If
        Else
            m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
            m_objGridMap.SetError .Bookmark, strError
            .Fields("Error").Value = "Error"
        End If
    End With

End Function

Private Function RefreshCostsCommercial() As String
    Dim strUpdate           As String
    Dim strSelect           As String
    Dim sbldg_model_skey    As String
    Dim i                   As Integer
    Dim recTemp             As New ADODB.RecordSet
    Dim cnTemp              As New ADODB.Connection
    Dim cmdTemp             As New ADODB.Command
    
    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    RefreshCostsCommercial = ""
    Status ("Updating Building Cost Information ...")
    With cnTemp
        .ConnectionTimeout = 50000
        .CommandTimeout = 50000
        '.Open "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
        .Open strConnect
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp

        strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & Trim(m_rec.Fields("bldg_skey").Value) _
                    & "' AND model_code != '7' AND model_code != '8'"
        If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                RefreshCostsCommercial = "Errors in the RefreshCostsCommercial routine searching for bldg_model_skey's."
        Else
            With recTemp
                If .RecordCount > 0 Then
                    Do Until .EOF
                        Status ("Updating Building Cost Information For Model " & Trim(.Fields("bldg_model_skey").Value) & " ...")
                        strUpdate = "exec sp_update_bldg_model @bldg_model_skey = '"
                        strUpdate = strUpdate & Trim(.Fields("bldg_model_skey").Value) & "',"
                        strUpdate = strUpdate & "@op_code = 'STD',"
                        strUpdate = strUpdate & "@country_code = 'USA',"
                        strUpdate = strUpdate & "@region_code = 'NAT'"
                        With cmdTemp
                            .CommandTimeout = 50000
                            .CommandType = adCmdText
                            .CommandText = strUpdate
                            .Execute adExecuteNoRecords
                        End With
                        DoEvents
                        If cnTemp.Errors.Count = 0 Then
                            strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
                            With cmdTemp
                                .CommandTimeout = 50000
                                .CommandType = adCmdText
                                .CommandText = strUpdate
                                .Execute adExecuteNoRecords
                            End With
                            DoEvents
                            If cnTemp.Errors.Count <> 0 Then
                                RefreshCostsCommercial = "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                                    & Trim(.Fields("bldg_model_skey").Value) & " " & vbCrLf & cnTemp.Errors(0).Description _
                                    & vbCrLf & "RefreshCostsCommercial routine will continue for other models."
                            End If
                        Else
                            RefreshCostsCommercial = "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                                & Trim(.Fields("bldg_model_skey").Value) & " " & vbCrLf & cnTemp.Errors(0).Description _
                                & vbCrLf & "RefreshCostsCommercial routine will continue for other models."
                        End If
                        .MoveNext
                    Loop
                Else
                    RefreshCostsCommercial = "Errors in the RefreshCostsCommercial routine, unable to locate bldg_model_skey's " _
                            & "associated with the building"
                End If
            End With
        End If
        .Close
    End With
    Exit Function

errorHandler:
    RefreshCostsCommercial = "Errors in RefreshCosts routine: " & Err.Description
    Status ("")
End Function

Private Function RefreshCostsResidential() As String
    Dim strUpdate           As String
    Dim strSelect           As String
    Dim sbldg_model_skey    As String
    Dim i                   As Integer
    Dim recTemp             As New ADODB.RecordSet
    Dim cnTemp              As New ADODB.Connection
    Dim cmdTemp             As New ADODB.Command
    
    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    RefreshCostsResidential = ""
    Status ("Updating Building Cost Information ...")
    With cnTemp
        .ConnectionTimeout = 50000
        .CommandTimeout = 50000
        '.Open "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
        .Open strConnect
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
        
        strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & Trim(m_rec.Fields("bldg_skey").Value) _
            & "' AND model_code != '7' AND model_code != '8'"
        If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                RefreshCostsResidential = "Errors in the RefreshCostsResidential routine searching for bldg_model_skey's."
        Else
            With recTemp
                If .RecordCount > 0 Then
                    Do Until .EOF
                        Status ("Updating Building Cost Information For Model " & Trim(.Fields("bldg_model_skey").Value) & " ...")
                        '
                        '   Updates cost for STD & OPN op_codes.
                        strUpdate = "exec sp_update_bldg_model_resi @bldg_model_skey = '"
                        strUpdate = strUpdate & Trim(.Fields("bldg_model_skey").Value) & "',"
                        strUpdate = strUpdate & "@op_code = 'STD',"
                        '
                        'allow to update & change order of models?
                        strUpdate = strUpdate & "@country_code = 'USA',"
                        strUpdate = strUpdate & "@region_code = 'NAT'"
                        With cmdTemp
                            .CommandTimeout = 50000
                            .CommandType = adCmdText
                            .CommandText = strUpdate
                            .Execute adExecuteNoRecords
                        End With
                        DoEvents
                        If cnTemp.Errors.Count = 0 Then
                            strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
                            With cmdTemp
                                .CommandTimeout = 50000
                                .CommandType = adCmdText
                                .CommandText = strUpdate
                                .Execute adExecuteNoRecords
                            End With
                
                            DoEvents
                            If cnTemp.Errors.Count <> 0 Then
                                Screen.MousePointer = vbNormal
                                RefreshCostsResidential = "Errors in the RefreshCostsResidential routine for Building Model skey: " _
                                    & Trim(.Fields("bldg_model_skey").Value) & " " & vbCrLf & cnTemp.Errors(0).Description _
                                    & vbCrLf & "RefreshCostsResidential routine will continue for other models."
                            End If
                        Else
                            Screen.MousePointer = vbNormal
                            RefreshCostsResidential = "Errors in the RefreshCostsResidential routine for Building Model skey: " _
                                & Trim(.Fields("bldg_model_skey").Value) & " " & vbCrLf & cnTemp.Errors(0).Description _
                                & vbCrLf & "RefreshCostsResidential routine will continue for other models."
                        End If
                        .MoveNext
                    Loop
                Else
                    RefreshCostsResidential = "Errors in the RefreshCostsResidential routine, unable to locate bldg_model_skey's " _
                            & "associated with the building"
                End If
            End With
        End If
        .Close
    End With
    Exit Function

errorHandler:
    RefreshCostsResidential = "Errors in the RefreshCosts routine: " & Err.Description
    Status ("")
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

Public Property Get GRIDTYPE() As String
    GRIDTYPE = "Building"
End Property

Public Property Get SuccessfulUpdates() As Integer
    SuccessfulUpdates = m_NumberOfSuccessfulUpdates
End Property

Public Property Get UpdateErrors() As Integer
    UpdateErrors = m_NumberOfUpdateErrors
End Property

'**********GRID STUFF***************
Private Sub TDBGrid_Change()
    With TDBGrid
        If .Columns(.Col).Caption = "Bldg Desc" Then
            '
            '   Truncate the bldg desc to only 75 chars.
            If Len(.Text) > 75 Then .Text = Left(Trim(.Text), 75)
        End If
    End With
End Sub
'
'   Called when a cell's data has changed as user moves to new row or presses Return
'   Used to validate values entered.
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, Cancel As Integer)

    Dim recTemp     As New ADODB.RecordSet
    Dim strSelect   As String

    On Error Resume Next
    With TDBGrid
        '
        '   They cannot change any fields for Residential Quality Series buildings.
        If .Columns("Bldg ID").Value = "100" Or .Columns("Bldg ID").Value = "200" _
        Or .Columns("Bldg ID").Value = "300" Or .Columns("Bldg ID").Value = "400" Then
            MsgBox "Building values cannot be changed for Residential Quality Series buildings.", vbCritical
            Cancel = True
        Else
            Select Case .Columns(ColIndex).Caption
                Case "Bldg Desc"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a building description.", vbCritical
                        .Columns("Bldg Desc") = OldValue
                        Cancel = True
                    ElseIf Len(Trim(.Text)) > 75 Then
                        MsgBox "Please provide a building description that is less than 75 characters.", vbCritical
                        .Columns("Bldg Desc") = Left$(Trim(.Columns("Bldg Desc")), 75)
                    End If
                               
                Case "Stories"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide the number of building stories.", vbCritical
                        .Columns("Stories") = OldValue
                        Cancel = True
                        
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please enter a whole number for the total stories."
                        .Columns("Stories") = OldValue
                        Cancel = True
                    
                    ElseIf Trim(.Text) < 1 Then
                        MsgBox "Please provide a number of stories that is greater than or equal to 1.", vbCritical
                        .Columns("Stories") = OldValue
                        Cancel = True
      
                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) > 3 Then
                        MsgBox "Maximum number of stories for residential building is 3.", vbCritical
                        .Columns("Stories") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) > 50 Then
                        MsgBox "Maximum number of stories for commercial building is 50.", vbCritical
                        .Columns("Stories") = OldValue
                        Cancel = True
                    End If
                    
                Case "Stories Hgt"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide the building story height.", vbCritical
                        .Columns("Flr-Hgt") = OldValue
                        Cancel = True
                    
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric building story height.", vbCritical
                        .Columns("Flr-Hgt") = OldValue
                        Cancel = True
                
                    ElseIf Trim(.Text) < 1 Then
                        MsgBox "Please provide a building story height that is greater than or equal to 1.", vbCritical
                        .Columns("Flr-Hgt") = OldValue
                        Cancel = True
                        
                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) > 12 Then
                        MsgBox "Maximum story height for residential building is 12.", vbCritical
                        .Columns("Flr-Hgt") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) > 30 Then
                        MsgBox "Maximum story height for commercial building is 30.", vbCritical
                        .Columns("Flr-Hgt") = OldValue
                        Cancel = True
                    End If

                Case "Part Density"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a building partition density.", vbCritical
                        .Columns("Part Density") = OldValue
                        Cancel = True
                
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric building partition density.", vbCritical
                        .Columns("Part Density") = OldValue
                        Cancel = True
                    
                    ElseIf Trim(.Text) < 1 Then
                        MsgBox "Please provide a building partition density that is greater than or equal to 1.", vbCritical
                        .Columns("Part Density") = OldValue
                        Cancel = True
                        
                    ElseIf Trim(.Text) > 999 Then
                        MsgBox "Maximum building partition density is 999.", vbCritical
                        .Columns("Part Density") = OldValue
                        Cancel = True
                    End If

                Case "Part Hgt"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a building partition height.", vbCritical
                        .Columns("Part Hgt") = OldValue
                        Cancel = True
                    
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric building partition height.", vbCritical
                        .Columns("Part Hgt") = OldValue
                        Cancel = True
                   
                    ElseIf Trim(.Text) < 1 Then
                        MsgBox "Please provide a building partition height that is greater than or equal to 1.", vbCritical
                        .Columns("Part Hgt") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) > 12 Then
                        MsgBox "Maximum building partition height for residential building is 12.", vbCritical
                        .Columns("Part Hgt") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) > 30 Then
                        MsgBox "Maximum building partition height for commercial building is 30.", vbCritical
                        .Columns("Part Hgt") = OldValue
                        Cancel = True
                    End If
                    
                Case "Dr Density"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a building door density.", vbCritical
                        .Columns("Dr Density") = OldValue
                        Cancel = True
                
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric building door density.", vbCritical
                        .Columns("Dr Density") = OldValue
                        Cancel = True
                        
                    ElseIf Trim(.Text) < 1 Then
                        MsgBox "Please provide a building door density that is greater than or equal to 1.", vbCritical
                        .Columns("Dr Density") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) > 999 Then
                        MsgBox "Maximum building door density for residential building is 999.", vbCritical
                        .Columns("Dr Density") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) > 9999 Then
                        MsgBox "Maximum building door density for commercial building is 9999.", vbCritical
                        .Columns("Dr Density") = OldValue
                        Cancel = True
                    End If

                Case "Elev No"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide the number of building elevators.", vbCritical
                        .Columns("Elev No") = OldValue
                        Cancel = True
                
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric number of building elevators.", vbCritical
                        .Columns("Elev No") = OldValue
                        Cancel = True
                    
                    'REVISED 6/16/2005 RTD FOR VERSION 7.4.0 CR#1312
                    'ElseIf Trim(.Text) < 1 Then
                    '    MsgBox "Please provide a number of elevators that is greater than or equal to 1.", vbCritical
                    '    .Columns("Elev No") = OldValue
                    '    Cancel = True
                        
                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) > 5 Then
                        MsgBox "Maximum number of elevators for a residential building is 5.", vbCritical
                        .Columns("Elev No") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) > 50 Then
                        MsgBox "Maximum number of elevators for a commercial building is 50.", vbCritical
                        .Columns("Elev No") = OldValue
                        Cancel = True
                    End If

                Case "Fixture Area"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a building fixture area.", vbCritical
                        .Columns("Fixture Area") = OldValue
                        Cancel = True
                    
                    ElseIf IsNumeric(Trim(.Text)) = False Then
                        MsgBox "Please provide a numeric building fixture area.", vbCritical
                        .Columns("Fixture Area") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) < 0 Then
                        MsgBox "Please provide a building fixture area that is greater than or equal to zero.", vbCritical
                        .Columns("Fixture Area") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) > 99 Then
                        MsgBox "Maximum building fixture area for a residential building is 99.", vbCritical
                        .Columns("Fixture Area") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) < 1 Then
                        MsgBox "Please provide a building fixture area that is greater than or equal to 1.", vbCritical
                        .Columns("Fixture Area") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) > 99999 Then
                        MsgBox "Maximum building fixture area for a commercial building is 99999.", vbCritical
                        .Columns("Fixture Area") = OldValue
                        Cancel = True
                    End If
       
                Case "Wall Fact"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a building wall factor.", vbCritical
                        .Columns("Wall Fact") = OldValue
                        Cancel = True
                
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric building wall factor.", vbCritical
                        .Columns("Wall Fact") = OldValue
                        Cancel = True
                    
                    ElseIf Trim(.Text) > 1 Then
                        MsgBox "Please provide a building wall factor that is less than or equal to 1.", vbCritical
                        .Columns("Wall Fact") = OldValue
                        Cancel = True
                
                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) < 0 Then
                        MsgBox "Please provide a building wall factor that is greater than or equal to 0.", vbCritical
                        .Columns("Wall Fact") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) < 0.01 Then
                        MsgBox "Please provide a building wall factor that is greater than or equal to .01.", vbCritical
                        .Columns("Wall Fact") = OldValue
                        Cancel = True
                    End If

               Case "Win Area"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a building window area.", vbCritical
                        .Columns("Win Area") = OldValue
                        Cancel = True
                    
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric building window area.", vbCritical
                        .Columns("Win Area") = OldValue
                        Cancel = True

                    ElseIf Trim(.Text) > 99 Then
                        MsgBox "Maximum building window area is 99.", vbCritical
                        .Columns("Win Area") = OldValue
                        Cancel = True
                        
                    ElseIf .Columns("Type").Value = "R" And Trim(.Text) < 1 Then
                        MsgBox "Please provide a building window area that is greater than or equal to 1.", vbCritical
                        .Columns("Win Area") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) < 0 Then
                        MsgBox "Please provide a building window area that is greater than or equal to zero.", vbCritical
                        .Columns("Win Area") = OldValue
                        Cancel = True
                    
                    ElseIf Trim(.Text) = 0 Then
                        strSelect = "SELECT formula_code FROM assembly_usage JOIN bldg_model bm ON bm.bldg_skey = '" _
                            & Trim(.Columns("Skey").Value) & "' WHERE parent_skey = bldg_model_skey" _
                            & " AND formula_code = 'WW'"

                        If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                            MsgBox "An error occurred while searching to validate that algorithm 'WW' did not exist with window area of zero.", vbCritical
                        ElseIf recTemp.RecordCount <> 0 Then
                            MsgBox "Please provide a different Algorithm or change the building window area.  " _
                                    & "Algorithm 'WW' is not allowed when the building window area is zero. ", vbCritical
                            .Columns("Win Area") = OldValue
                            Cancel = True
                        End If
                        recTemp.Close
                    End If
 
                Case "OP Fact"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a building overhead & profit factor.", vbCritical
                        .Columns("OP Fact") = OldValue
                        Cancel = True
                
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric building overhead & profit factor.", vbCritical
                        .Columns("OP Fact") = OldValue
                        Cancel = True
                        
                    ElseIf Trim(.Text) > 1 Then
                        MsgBox "Please provide a building overhead & profit factor that is less than or equal to 1.", vbCritical
                        .Columns("OP Fact") = OldValue
                        Cancel = True
                        
                     ElseIf .Columns("Type").Value = "R" And Trim(.Text) < 0 Then
                        MsgBox "Please provide a building overhead & profit factor that is greater than or equal to zero.", vbCritical
                        .Columns("OP Fact") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) < 0.01 Then
                        MsgBox "Please provide a building overhead & profit factor that is greater than or equal to .01.", vbCritical
                        .Columns("OP Fact") = OldValue
                        Cancel = True
                    End If
                                           
                Case "Arch Fees"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide the building architect fees.", vbCritical
                        .Columns("Arch Fees") = OldValue
                        Cancel = True
                
                    ElseIf Not IsNumeric(Trim(.Text)) Then
                        MsgBox "Please provide a numeric building architect fees.", vbCritical
                        .Columns("Arch Fees") = OldValue
                        Cancel = True
                    
                    ElseIf Trim(.Text) > 1 Then
                        MsgBox "Please provide a building architect fee that is less than or equal to 1.", vbCritical
                        .Columns("Arch Fees") = OldValue
                        Cancel = True
                
                     ElseIf .Columns("Type").Value = "R" And Trim(.Text) < 0 Then
                        MsgBox "Please provide a building architect fee that is greater than or equal to zero.", vbCritical
                        .Columns("Arch Fees") = OldValue
                        Cancel = True

                    ElseIf .Columns("Type").Value = "C" And Trim(.Text) < 0.01 Then
                        MsgBox "Please provide a building architect fee that is greater than or equal to .01", vbCritical
                        .Columns("Arch Fees") = OldValue
                        Cancel = True
                    End If
                    
                Case "Graphic ID"
                    If Len(Trim(.Text)) > 12 Then
                        MsgBox "Please alter the file name to be 12 characters or less.", vbCritical
                        .Columns("Graphic ID") = Left$(Trim(.Columns("Graphic ID")), 12)
                        Cancel = True
                    End If
    
                Case "Graphic ID 2"
                    If Len(Trim(.Text)) > 12 Then
                        MsgBox "Please alter the file name to be 12 characters or less.", vbCritical
                        .Columns("Graphic ID 2") = Left$(Trim(.Columns("Graphic ID 2")), 12)
                        Cancel = True
                    End If
            End Select
        End If
    End With
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    On Error Resume Next
    
    With TDBGrid
        If .Columns(ColIndex).Caption <> "Bldg Desc" And .Columns(ColIndex).Caption <> "Graphic ID" _
            And .Columns(ColIndex).Caption <> "Graphic ID 2" Then
            
            bRefreshCosts = True
        End If
    End With
End Sub



