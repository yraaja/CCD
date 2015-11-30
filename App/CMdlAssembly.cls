VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMdlAssembly"
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
Const BLDGCOMPONENT_TABLE = "published_bldg_component_cost"
Const ASSEMBLY_USAGE = "assembly_usage"
Const ALL_TABLES = "All"
'
'   Grid Messages.
Const CLEAR_BLDG_CATEGORY = 1
'
'   "Parent" class
Private m_objGridMap As New CGridMap
'
'   Recordset that populates grid
Private m_rec As ADODB.RecordSet
'
'   From bldg Recordset, used to compute values.
Private bldg_stories        As String
Private bldg_stories_hgt    As String
Private bldg_area_std       As String
Private bldg_perimeter_std  As String
Private bldg_part_hgt       As String
Private bldg_part_density   As String
Private bldg_door_density   As String
Private op_code             As String
Private window_area         As String

Private perim_hold          As String
Private bldg_stair_factor   As String
Private ground_area         As String
Private bldg_f_height       As String
Private bldg_factor         As String
Private bldg_door_exterior  As String
Private bldg_door_interior  As String
Private bldg_partition      As String
Private bldg_type           As String
Private type_code           As String

Private m_blnPreferences                As Boolean
Private m_varSaveBookmark               As Variant
Private m_varCurrentBookmark            As Variant
Private m_colUpdatedRows                As New Collection
'
'   Used in the SuccessfulUpdates Property Get.
Private m_NumberOfSuccessfulUpdates     As Integer
'
'   Used in the UpdateErrors Property Get.
Private m_NumberOfUpdateErrors          As Integer
Private bInserting                      As Boolean
Private bCloningAssemblies              As Boolean
Private m_colAssembliesFormulasItems    As New Collection
Private bMultiDeleteInProcess           As Boolean
Private sPriorAlgorithm                 As String
Public Event RefreshCostsAssemblyDeleted()

Private Sub Class_Initialize()
    With m_objGridMap.m_colCloneCols
        .Add 1, "assembly_id"
        .Add 1, "alt_assembly_id"
        .Add 1, "skey_type"
        .Add 1, "parent_skey"
        .Add 1, "assembly_skey"
        .Add 1, "formula_factor"
        .Add 1, "override_bldg_book_desc"
        .Add 1, "override_bldg_book_qty"
    End With
End Sub

Public Sub InitGrid(Optional bViewOnly As Boolean)
   
    On Error Resume Next
    '
    '   Identifies the Grid.
    
    TDBGrid.AllowAddNew = True  'rlh CCD 8.4  3/04/2009
    
    With m_objGridMap
        .GRIDTYPE = "Model Assemblies"
        
    
        If bViewOnly Then
            .AddCol "Model Skey", "bldg_model_skey", ASSEMBLY_USAGE, 1000, True, False, False, True
            .AddCol "Assembly Skey", "assembly_skey", ASSEMBLY_USAGE, 1000, True, False, False, True
            'We need the original assembly_skey in case the modify an existing row with a new assembly
            'thus we need to know to delete and insert the new. Same for formula_code it's part of the key.
            .AddCol "Prior Assembly Skey", "prior_assembly_skey", ASSEMBLY_USAGE, 1000, True, False
            .AddCol "Assembly ID", "assembly_id", vbNullString, 1350, True, True, True, FORMAT_UNIT_COST
            .AddCol "Alt Assembly ID", "alt_assembly_id", vbNullString, 1350, True, True, True, FORMAT_UNIT_COST
    
            .AddCol "Tech Desc", "assembly_desc", vbNullString, 2000, True, True
            
            .AddCol "Algorithm", "formula_code", ASSEMBLY_USAGE, 1050, True, True
            .AddCol "Prior Algorithm", "prior_formula_code", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Formula Factor", "formula_factor", ASSEMBLY_USAGE, 1205, True, True
            .AddCol "Unit", "unit", vbNullString, 750, True, True

            .AddCol "Total OP", "total_cost_op", ALL_TABLES, 1200, True, , , FORMAT_PRICE, True
            .AddCol "Total Per SF", "total_cost_op_per_sf", ALL_TABLES, 1200, True, , , FORMAT_PRICE, True
            
            .AddCol "Mat OP", "mat_cost_op", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Mat Per SF", "mat_cost_op_per_sf", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Labor OP", "labor_cost_op", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Labor Per SF", "labor_cost_op_per_sf", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Equip OP", "equip_cost_op", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Equip Per SF", "equip_cost_op_per_sf", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Instl OP", "inst_cost_op", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Instl Per SF", "inst_cost_op_per_sf", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
                       
            .AddCol "Stories", "bldg_stories", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Stories Hgt", "bldg_stories_hgt", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Part Hgt", "bldg_part_hgt", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Part Density", "bldg_part_density", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Door Density", "bldg_door_density", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Arch Fees", "architect_fee", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Area Std", "bldg_area_std", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Perm Std", "bldg_perimeter_std", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Wall Factor", "bldg_wall_factor", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Elev No", "bldg_elev_no", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Fixture Area", "bldg_fixture_area", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "OP Factor", "op_factor", ASSEMBLY_USAGE, 1050, True, False
                          
            .AddCol "Override Book Desc", "override_bldg_book_desc", ASSEMBLY_USAGE, 1200, True, True
            .AddCol "Override Book Qty", "override_bldg_book_qty", ASSEMBLY_USAGE, 900, True, True
                          
            .AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
            .AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
            .AddCol "last_update_id", "last_update_id", BLDGCOMPONENT_TABLE, 800, True, False, False
        Else
            .AddCol "Model Skey", "bldg_model_skey", ASSEMBLY_USAGE, 1000, True, False, False, True
            .AddCol "Assembly Skey", "assembly_skey", ASSEMBLY_USAGE, 1000, True, False, False, True
            'We need the original assembly_skey in case the modify an existing row with a new assembly
            'thus we need to know to delete and insert the new. Same for formula_code it's part of the key.
            .AddCol "Prior Assembly Skey", "prior_assembly_skey", ASSEMBLY_USAGE, 1000, True, False
            .AddCol "Assembly ID", "assembly_id", vbNullString, 1350, False, True, True, FORMAT_UNIT_COST
            .AddCol "Alt Assembly ID", "alt_assembly_id", vbNullString, 1350, False, True, True, FORMAT_UNIT_COST
    
            .AddCol "Tech Desc", "assembly_desc", vbNullString, 2000, True, True
            
            .AddCol "Algorithm", "formula_code", ASSEMBLY_USAGE, 1050, False, True
            .AddCol "Prior Algorithm", "prior_formula_code", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Formula Factor", "formula_factor", ASSEMBLY_USAGE, 1205, False, True
            .AddCol "Unit", "unit", vbNullString, 750, True, True

            .AddCol "Total OP", "total_cost_op", ALL_TABLES, 1200, True, True, , FORMAT_PRICE, True
            .AddCol "Total Per SF", "total_cost_op_per_sf", ALL_TABLES, 1200, True, True, , FORMAT_PRICE, True

            .AddCol "Mat OP", "mat_cost_op", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Mat Per SF", "mat_cost_op_per_sf", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Labor OP", "labor_cost_op", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Labor Per SF", "labor_cost_op_per_sf", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Equip OP", "equip_cost_op", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Equip Per SF", "equip_cost_op_per_sf", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Instl OP", "inst_cost_op", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True
            .AddCol "Instl Per SF", "inst_cost_op_per_sf", ALL_TABLES, 1200, True, False, , FORMAT_PRICE, True

            .AddCol "Stories", "bldg_stories", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Stories Hgt", "bldg_stories_hgt", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Part Hgt", "bldg_part_hgt", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Part Density", "bldg_part_density", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Door Density", "bldg_door_density", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Arch Fees", "architect_fee", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Area Std", "bldg_area_std", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Perm Std", "bldg_perimeter_std", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Wall Factor", "bldg_wall_factor", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Elev No", "bldg_elev_no", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "Fixture Area", "bldg_fixture_area", ASSEMBLY_USAGE, 1050, True, False
            .AddCol "OP Factor", "op_factor", ASSEMBLY_USAGE, 1050, True, False

            .AddCol "Override Book Desc", "override_bldg_book_desc", ASSEMBLY_USAGE, 1200, False, True
            .AddCol "Override Book Qty", "override_bldg_book_qty", ASSEMBLY_USAGE, 900, False, True

            .AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
            .AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
            .AddCol "last_update_id", "last_update_id", BLDGCOMPONENT_TABLE, 800, True, False, False
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
Public Property Get bCloneAssembliesInProcess() As Boolean
    bCloneAssembliesInProcess = bCloningAssemblies
End Property
Public Property Let bCloneAssembliesInProcess(bNewValue As Boolean)
    bCloningAssemblies = bNewValue
End Property
Public Property Let bldgStories(sNewValue As String)
    bldg_stories = sNewValue
End Property

Public Property Let bldgStoriesHgt(sNewValue As String)
    bldg_stories_hgt = sNewValue
End Property

Public Property Let bldgAreaStd(sNewValue As String)
    bldg_area_std = sNewValue
End Property

Public Property Let bldgPerimeterStd(sNewValue As String)
    bldg_perimeter_std = sNewValue
End Property

Public Property Let bldgPartHgt(sNewValue As String)
    bldg_part_hgt = sNewValue
End Property

Public Property Let bldgPartDensity(sNewValue As String)
    bldg_part_density = sNewValue
End Property

Public Property Let bldgDoorDensity(sNewValue As String)
    bldg_door_density = sNewValue
End Property

Public Property Let opCode(sNewValue As String)
    op_code = sNewValue
End Property

Public Property Let WindowArea(sNewValue As String)
    window_area = sNewValue
End Property

Public Property Let BldgStairFactor(sNewValue As String)
    bldg_stair_factor = sNewValue
End Property

Public Property Let BldgFHeight(sNewValue As String)
    bldg_f_height = sNewValue
End Property

Public Property Let BldgFactor(sNewValue As String)
    bldg_factor = sNewValue
End Property

Public Property Let BldgType(sNewValue As String)
    bldg_type = sNewValue
End Property

Public Property Let TypeCode(sNewValue As String)
    type_code = sNewValue
End Property

Public Property Let RecordSet(RecSet As ADODB.RecordSet)
    '
    '   Pass-through to base class
    m_objGridMap.RecordSet = RecSet
    Set m_rec = RecSet
End Property

Public Sub ComputeValues(Area As String, Perm As String)
    Dim FormulaFactor   As String
    Dim I               As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With TDBGrid
        FormulaFactor = .Columns("Formula Factor").Value
        
        Select Case Trim(UCase(.Columns("Algorithm").Value))
            Case "G"
                '
                '   G = (assembly_cost * Factor) / bldg_stories
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor) / bldg_stories
                End If
            Case "L"
                '
                '   L = (assembly_cost * bldg_perimeter * Factor) / bldg_area
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * Perm * FormulaFactor) / Area
                End If
            
            Case "EF"
                '
                '   EF = ((bldg_stories - 1) * assembly_cost * factor) / bldg_stories
                 If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = ((bldg_stories - 1) * .Columns("Total OP").Value * FormulaFactor) / bldg_stories
                End If

            Case "F"
                '
                '   F = assembly_cost * Factor
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = .Columns("Total OP").Value * FormulaFactor
                End If
                
            Case "W"
                '
                '   W = (assembly_cost * bldg_perimeter * bldg_stories * bldg_stories_hgt * factor) / bldg_area"
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * Perm * bldg_stories * bldg_stories_hgt * FormulaFactor) / Area
                End If

            Case "WW"
                '
                '   WW = (assembly_cost * bldg_perimeter * bldg_stories * bldg_stories_hgt * Factor) / (bldg_area * window_area)"
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * Perm * bldg_stories * bldg_stories_hgt * FormulaFactor) / (Area * window_area)
                End If
            
            Case "E", "DE"
                '
                '   = (assembly_cost * factor) / bldg_area_std"
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor) / bldg_area_std
                End If

            Case "EA"
                '
                '   = (assembly_cost * factor) / bldg_area"
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor) / Area
                End If

            Case "P"
                '
                '   P = (assembly_cost * bldg_part_hgt * factor) / bldg_part_density"
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * bldg_part_hgt * FormulaFactor) / bldg_part_density
                End If

            Case "DI"
                '
                '   DI = (assembly_cost * factor) / bldg_door_density"
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor) / bldg_door_density
                End If

            Case "DW"
                '
                '   DW = ((assembly_cost * 144) * factor * bldg_stories_hgt * bldg_perimeter * bldg_stories) / bldg_area"
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = ((.Columns("Total OP").Value * 144) * FormulaFactor * bldg_stories_hgt * Perm * bldg_stories) / Area
                End If

            Case "S"
                '
                '   "S = (assembly_cost * 2 * bldg_part_hgt * factor) / bldg_part_density"
                If .Columns("Total OP").Value <> "" Then
                    .Columns("Total Per SF").Value = (.Columns("Total OP").Value * 2 * bldg_part_hgt * FormulaFactor) / bldg_part_density
                End If
        End Select
    End With
    Screen.MousePointer = vbNormal
End Sub

Public Sub ComputeValuesResi(Area As String, Perm As String)
    Dim FormulaFactor   As String
    Dim I               As Integer
    Dim recTemp     As New ADODB.RecordSet
    Dim strSELECT   As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With TDBGrid
        '
        '   Compute the following values.
        ground_area = Area / bldg_factor
        FormulaFactor = .Columns("Formula Factor").Value
        '
        '   Query to get the following values from the database based upon the area.
        'perim_hold based on bldg_type
        If bldg_type = "H" Or bldg_type = "I" Or bldg_type = "J" Then
            strSELECT = "SELECT bldg_door_exterior, bldg_door_interior, bldg_partition, bldg_perim_hold " _
                        & "FROM bldg_wing_resi WHERE " & ground_area & " BETWEEN ground_area_min AND ground_area_max"
        Else
            strSELECT = "SELECT bldg_door_exterior, bldg_door_interior, bldg_partition, '0' AS bldg_perim_hold " _
                        & "FROM bldg_ratio_resi WHERE " & ground_area & " BETWEEN ground_area_min AND ground_area_max"
        End If
        If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recTemp) Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching for formula values to use in the calculation.", vbCritical
        Else
            With recTemp
                bldg_door_exterior = .Fields("bldg_door_exterior").Value
                bldg_door_interior = .Fields("bldg_door_interior").Value
                bldg_partition = .Fields("bldg_partition").Value
                perim_hold = .Fields("bldg_perim_hold").Value
                .Close
            End With
            
            Select Case Trim(UCase(.Columns("Algorithm").Value))
                Case "G"
                    '
                    '   G = (assembly_cost * ground_area * Factor) / bldg_area
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * ground_area * FormulaFactor) / Area
                    End If
                Case "L"
                    '
                    '   L = (assembly_cost * perim_hold * Factor) / bldg_area
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * perim_hold * FormulaFactor) / Area
                    End If
                Case "EF"
                    '
                    '   EF = (assembly_cost * (bldg_area - ground_area)) / bldg_area
                     If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * (Area - ground_area)) / Area
                    End If
                Case "F"
                    '
                    '   F = assembly_cost * Factor
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = .Columns("Total OP").Value * FormulaFactor
                        '.Columns("Total Per SF").Value = .Columns("Total OP").Value * FormulaFactor
                    End If
                Case "W"
                    '
                    '   W = (assembly_cost * perim_hold * bldg_f_height * factor) / bldg_area
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * perim_hold * bldg_f_height * FormulaFactor) / Area
                    End If
                Case "WW"
                    '
                    '   WW = (assembly_cost * perim_hold * bldg_f_height * Factor) / (bldg_area * window_area)
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * perim_hold * bldg_f_height * FormulaFactor) / (Area * window_area)
                    End If
                Case "E"
                    '
                    '   E = (assembly_cost * factor) / 1200
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor) / 1200
                    End If
                Case "DE"
                    '
                    '   DE = (assembly_cost * factor * bldg_door_exterior) / bldg_area
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor * bldg_door_exterior) / Area
                    End If
                Case "EA", "BS"
                    '
                    '   EA = (assembly_cost * factor) / bldg_area
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor) / Area
                    End If
                Case "P"
                    '
                    '   P = (assembly_cost * factor * bldg_part_hgt) / bldg_partition
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor * bldg_part_hgt) / bldg_partition
                    End If
                Case "DI"
                    '
                    '   DI = (assembly_cost * factor) / bldg_door_interior
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * FormulaFactor) / bldg_door_interior
                    End If
                Case "DW"
                    '
                    '   DW = ((assembly_cost /144) * Factor * bldg_stories_hgt * perim_hold * bldg_stories) / bldg_area
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = ((.Columns("Total OP").Value * 144) * FormulaFactor * bldg_stories_hgt * perim_hold * bldg_stories) / Area
                    End If
                Case "S"
                    '
                    '   S = (assembly_cost * 2 * bldg_part_hgt * factor) / bldg_partition
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * 2 * bldg_part_hgt * FormulaFactor) / bldg_partition
                    End If
                Case "ST"
                    '
                    '   ST = (assembly_cost * bldg_stair_factor) / bldg_area
                    '   AND bldg_type not in ('H','I','J') Enforced by calling application.
                    If .Columns("Total OP").Value <> "" Then
                        .Columns("Total Per SF").Value = (.Columns("Total OP").Value * bldg_stair_factor) / Area
                    End If
            End Select
        End If
    End With
    Screen.MousePointer = vbNormal
End Sub

Public Function Update(sSkeyType As String, sParentSkey As String, nmodel_form As Integer) As Boolean
    Dim strUpdate As String ' SQL string
    Dim sBookDesc   As String
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    Dim iRecCount As Long
    Dim iRec As Long
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Update = True
    intErrors = 0
    intSuccess = 0
    With m_rec
        If .State <> adStateClosed Then
            .MoveFirst
            iRecCount = m_rec.RecordCount
            iRec = 0
            While Not .EOF
                iRec = iRec + 1
                Status ("Updating Temporary Assemblies Table " & iRec & " of " & iRecCount & "...")
                DoEvents
                ' Skip the record if it didn't change
                Dim clsRowInfo As CRowInfo
                Set clsRowInfo = m_objGridMap.rowinfo(.Bookmark)
                '   If record was added meaning clsRowInfo.State = STATE_NEW
                '   Or clsRowInfo.State = STATE_MODIFIED Then
                If Not clsRowInfo Is Nothing Then
                    If clsRowInfo.State <> STATE_NONE Then
                        strUpdate = "exec  sp_update_assembly_usage_holding_table @parent_skey = '"
                        '
                        '   We can't depend upon the grid row having the model/parent_skey and
                        '   skey_type populated ie-they inserted rows.  So have it passed in.
                        strUpdate = strUpdate & sParentSkey & "',"
                        strUpdate = strUpdate & "@skey_type = '" & sSkeyType & "',"
                        
                        strUpdate = strUpdate & "@assembly_skey  = '" & Trim(.Fields("assembly_skey").Value) & "',"
                        strUpdate = strUpdate & "@formula_code  = '" & Trim(.Fields("formula_code").Value) & "',"
                        
                        strUpdate = strUpdate & "@prior_assembly_skey = '" _
                            & IIf(Trim(.Fields("prior_assembly_skey").Value) = "", _
                            Trim(.Fields("assembly_skey").Value), Trim(.Fields("prior_assembly_skey").Value)) & "',"
                            
                        strUpdate = strUpdate & "@prior_formula_code  = '" _
                            & IIf(Trim(.Fields("prior_formula_code").Value) = "0", _
                            Trim(.Fields("formula_code").Value), Trim(.Fields("prior_formula_code").Value)) & "',"
                        
                        strUpdate = strUpdate & "@formula_factor  = " & IIf(Trim(.Fields("formula_factor").Value) = "", 0, .Fields("formula_factor").Value) & ","
                        '
                        '   Since we might have ' marks in our book desc, need to replace with '' for SQL.
                        sBookDesc = Trim(.Fields("override_bldg_book_desc").Value)
                        sBookDesc = Replace(sBookDesc, "'", "''", 1)
        
                        strUpdate = strUpdate & "@override_bldg_book_desc  = '" & sBookDesc & "',"
                        If Len(Trim(.Fields("override_bldg_book_qty").Value)) = 0 Or IsNull(.Fields("override_bldg_book_qty").Value) Then
                            strUpdate = strUpdate & "@override_bldg_book_qty  = 0,"
                        Else
                            strUpdate = strUpdate & "@override_bldg_book_qty = " & .Fields("override_bldg_book_qty").Value & ","
                        End If
                        '
                        '   Set the last_update_id to 0 for new's in assembly_usage_holding_table
                        If Len(Trim(.Fields("last_update_id").Value)) = 0 Then
                            strUpdate = strUpdate & "@last_update_id = '0',"
                        Else
                            strUpdate = strUpdate & "@last_update_id = '" & Trim(.Fields("last_update_id").Value) & "',"
                        End If
                        strUpdate = strUpdate & "@model_form = '" & nmodel_form & "'"
                        '
                        '   Reset on success
                        ' rlh
                        If DEBUGON Then
                            Debug.Print "CMdlAssembly:Update: " & strUpdate
                        End If
                        
                        If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
                            If strError = "" Then
                                intSuccess = intSuccess + 1
                                clsRowInfo.State = STATE_NONE
                                .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                                .Fields("last_update_date").Value = Now
                                .Fields("last_update_person").Value = strUserName
                                Reset_Orig_Values m_rec
                            Else
                                intErrors = intErrors + 1
                                m_objGridMap.SetError .Bookmark, strError
                                .Fields("Error").Value = "Error"
                                Update = False
                            End If
                        Else
                            intErrors = intErrors + 1
                            m_objGridMap.SetError .Bookmark, strError
                            .Fields("Error").Value = "Error"
                            Update = False
                        End If
                    End If
                End If
                .MoveNext
            Wend
            ' We do not want to update the underlying recordset when cloning because we read
            ' the values from the standard model.
            If (Not bCloningAssemblies) Then
                .UpdateBatch
            End If
        End If
    End With
    TDBGrid.ReBind
    Screen.MousePointer = vbDefault
    
End Function

Public Property Let Preferences(bln As Boolean)
    '
    '   Pass-through to base class
    m_blnPreferences = bln
End Property

Public Property Get GRIDTYPE() As String
    GRIDTYPE = "Model Assemblies"
End Property

Public Property Get SuccessfulUpdates() As Integer
    SuccessfulUpdates = m_NumberOfSuccessfulUpdates
End Property

Public Property Get UpdateErrors() As Integer
    UpdateErrors = m_NumberOfUpdateErrors
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

Public Sub SetupAssembliesFormulas(bDontSetBookmark As Boolean)
    Dim I           As Integer
    Dim varBookmark As Variant
    Dim sAlgorithm  As String
    Dim sSkey       As String
    
    On Error Resume Next
    Set m_colAssembliesFormulasItems = Nothing
    Set m_colAssembliesFormulasItems = New Collection
    
    With TDBGrid
        varBookmark = .Bookmark
        
        I = 1
        Do Until I > TDBGrid.ApproxCount
            sAlgorithm = Trim(.Columns("Algorithm").CellText(I))
            sSkey = Trim(.Columns("Assembly Skey").CellText(I))
            
            m_colAssembliesFormulasItems.Add sAlgorithm & sSkey, sAlgorithm & sSkey
            I = I + 1
        Loop
        If Not bDontSetBookmark And Not IsNull(varBookmark) Then
            .Bookmark = varBookmark
        End If
    End With
End Sub

'**********GRID STUFF***************
Private Sub TDBGrid_AfterDelete()
    SetupAssembliesFormulas False
    sPriorAlgorithm = ""
End Sub

Private Sub TDBGrid_AfterInsert()
    bInserting = False
    SetupAssembliesFormulas False
    sPriorAlgorithm = ""
End Sub

Private Sub TDBGrid_AfterUpdate()
    bInserting = False
    SetupAssembliesFormulas False
    sPriorAlgorithm = ""
End Sub

Private Sub TDBGrid_OnAddNew()
    bInserting = True
End Sub

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim strSELECT           As String
    Dim ID                  As String
    Dim recAssemblyUnitCost As New ADODB.RecordSet
    Dim I                   As Integer
    Dim sAlgorithm          As String
    Dim sAssemblySkey       As String
    Dim nFound              As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With TDBGrid
        If .Columns(ColIndex).Caption = "Assembly ID" And Trim(.Columns(ColIndex).Value) <> "" Then
            If Trim(.Columns(ColIndex).Value) <> Trim(OldValue) Then
                If Len(Trim(.Columns("Assembly ID").Value)) = 14 Then
                    ID = Left$(.Columns("Assembly ID").Value, 5) & Right$(Left$(.Columns("Assembly ID").Value, 9), 3) & Right$(.Columns("Assembly ID").Value, 4)
                ElseIf Len(Trim(.Columns("Assembly ID").Value)) = 12 And InStr(1, Trim(.Columns("Assembly ID").Value), " ") <> 0 Then
                    ID = Left$(.Columns("Assembly ID").Value, 3) & Right$(Left$(.Columns("Assembly ID").Value, 7), 3) & Right$(.Columns("Assembly ID").Value, 4)
                Else
                    ID = Trim(.Columns("Assembly ID").Value)
                End If
                '
                '   We need to search for a new assembly/unit cost.
                strSELECT = "exec sp_select_model_assemblies @assembly_id = '" & ID & "'"
            End If
            
        ElseIf .Columns(ColIndex).Caption = "Alt Assembly ID" And Trim(.Columns(ColIndex).Value) <> "" Then
            If Trim(.Columns(ColIndex).Value) <> Trim(OldValue) Then
                If Len(Trim(.Columns("Alt Assembly ID").Value)) = 12 And InStr(1, Trim(.Columns("Alt Assembly ID").Value), " ") <> 0 Then
                    ID = Left$(Trim(.Columns("Alt Assembly ID").Value), 3) & Right$(Left$(Trim(.Columns("Alt Assembly ID").Value), 7), 3) & Right$(Trim(.Columns("Alt Assembly ID").Value), 4)
                ElseIf Len(Trim(.Columns("Alt Assembly ID").Value)) = 14 Then
                    ID = Left$(.Columns("Alt Assembly ID").Value, 5) & Right$(Left$(.Columns("Alt Assembly ID").Value, 9), 3) & Right$(.Columns("Alt Assembly ID").Value, 4)
                Else
                    ID = Trim(.Columns("Alt Assembly ID").Value)
                End If
                '
                '   We need to search for a new assembly/unit cost.
                strSELECT = "exec sp_select_model_assemblies @assembly_id = '" & ID & "'"
            End If
        '
        '   If we have changed the Algorithm we need to make sure that the assembly_skey
        '   algorithm combination is not already in use.
        ElseIf .Columns(ColIndex).Caption = "Algorithm" And Trim(.Columns(ColIndex).Value) <> "" Then
            '
            '   If we're on a Resi bldg with bldg_type 'H','I', or 'J' they can't
            '   use formula 'ST' or 'BS'
            If type_code = "R" Then
                If bldg_type = "H" Or bldg_type = "I" Or bldg_type = "J" Then
                    If Trim(.Columns(ColIndex).Value) = "ST" Then
                        MsgBox "Please provide a different Algorithm, 'ST' is not allowed for Residential building type: " & bldg_type, vbCritical
                        Cancel = True
                    ElseIf Trim(.Columns(ColIndex).Value) = "BS" Then
                        MsgBox "Please provide a different Algorithm, 'BS' is not allowed for Residential building type: " & bldg_type, vbCritical
                        Cancel = True
                    End If
                End If
            End If
            If Cancel = False And window_area = "0" Then
                If Trim(.Columns(ColIndex).Value) = "WW" Then
                    MsgBox "Please provide a different Algorithm or change the building window area.  " _
                        & vbCrLf & "Algorithm 'WW' is not allowed when the building window area is zero. ", vbCritical
                    Cancel = True
                End If
            End If
            
            If Cancel = False Then  'And Trim(.Columns(ColIndex).Value) <> Trim(OldValue)
                If .Columns("Assembly ID").Value <> "" Then
                    '   If you are inserting you have to make sure that all of the
                    '   other combinations are in the collection before trying to insert
                    '   the new one.  If editing an existing row you can't do this because
                    '   the new value gets added to the collection and then when you try
                    '   to "re-insert" here it fails.
                    If bInserting Then
                        SetupAssembliesFormulas True
                    Else
                        If sPriorAlgorithm = "" Then
                            sPriorAlgorithm = Trim(OldValue)
                        End If
                        For I = 1 To m_colAssembliesFormulasItems.Count
                            If m_colAssembliesFormulasItems.Item(I) = Trim(sPriorAlgorithm) & Trim(.Columns("Assembly Skey").Value) Then
                                m_colAssembliesFormulasItems.Remove I
                                sPriorAlgorithm = Trim(.Columns(ColIndex).Value)
                                Exit For
                            End If
                        Next I
                    End If
                    
                    sAlgorithm = Trim(.Columns(ColIndex).Value)
                    sAssemblySkey = Trim(.Columns("Assembly Skey").Value)
                    
                    Err.Clear
                    m_colAssembliesFormulasItems.Add sAlgorithm & sAssemblySkey, sAlgorithm & sAssemblySkey
                    If Err.Number <> 0 Then
                        MsgBox "Please provide an Assembly ID/Formula Code combination that is not already in use.", vbCritical
                        Cancel = True
                    End If
                End If
            End If
        End If

        If ID <> "" And Cancel = False Then
            If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recAssemblyUnitCost) Then
                Screen.MousePointer = vbNormal
                MsgBox "An error occurred while searching for assemblies.", vbCritical
            Else
                If recAssemblyUnitCost.EOF Then
                    '
                    '   Appears that the user has entered an invalid ID,
                    '   raise message.
                    MsgBox "ID not found in Assemblies.  Please provide a valid ID.", vbCritical
                    Cancel = True
                Else
                    If .Columns("Algorithm").Value <> "" Then
                        '   If you are inserting you have to make sure that all of the
                        '   other combinations are in the collection before trying to insert
                        '   the new one.  If editing an existing row you can't do this because
                        '   the new value gets added to the collection and then when you try
                        '   to "re-insert" here it fails.
                        If bInserting Then
                            SetupAssembliesFormulas True
                        End If
        
                        sAlgorithm = Trim(.Columns("Algorithm").Value)
                        sAssemblySkey = Trim(recAssemblyUnitCost.Fields("assembly_skey").Value)
                        
                        Err.Clear
                        m_colAssembliesFormulasItems.Add sAlgorithm & sAssemblySkey, sAlgorithm & sAssemblySkey
                        If Err.Number <> 0 Then
                            MsgBox "Please provide an Assembly ID/Formula Code combination that is not already in use.", vbCritical
                            .Columns("Assembly ID").Value = OldValue
                            Cancel = True
                        End If
                    End If
                    
                    If Cancel = False Then
                        .Columns("Assembly ID").Value = Trim(recAssemblyUnitCost.Fields("assembly_id").Value)
                        .Columns("Alt Assembly ID").Value = Trim(recAssemblyUnitCost.Fields("alt_assembly_id").Value)
                        .Columns("Assembly_skey").Value = Trim(recAssemblyUnitCost.Fields("assembly_skey").Value)
                        .Columns("Tech Desc").Value = Trim(recAssemblyUnitCost.Fields("tech_desc").Value)
                                     
                        .Columns("Unit").Value = Trim(recAssemblyUnitCost.Fields("unit").Value)
                        .Columns("Metric Unit").Value = Trim(recAssemblyUnitCost.Fields("metric_unit").Value)
                        
                        If .Columns("Prior Algorithm").Value = "" Then
                            .Columns("Prior Algorithm").Value = "0"
                        End If
                        
                        If op_code = "STD" Then
                            .Columns("Total OP").Value = Trim(recAssemblyUnitCost.Fields("std_total_cost_op").Value)
                            .Columns("Mat OP").Value = Trim(recAssemblyUnitCost.Fields("std_mat_cost_op").Value)
                            .Columns("Labor OP").Value = Trim(recAssemblyUnitCost.Fields("std_labor_cost_op").Value)
                            .Columns("Equip OP").Value = Trim(recAssemblyUnitCost.Fields("std_equip_cost_op").Value)
                            .Columns("Instl OP").Value = Trim(recAssemblyUnitCost.Fields("std_inst_cost_op").Value)
                        Else
                            .Columns("Total OP").Value = Trim(recAssemblyUnitCost.Fields("opn_total_cost_op").Value)
                            .Columns("Mat OP").Value = Trim(recAssemblyUnitCost.Fields("opn_mat_cost_op").Value)
                            .Columns("Labor OP").Value = Trim(recAssemblyUnitCost.Fields("opn_labor_cost_op").Value)
                            .Columns("Equip OP").Value = Trim(recAssemblyUnitCost.Fields("opn_equip_cost_op").Value)
                            .Columns("Instl OP").Value = Trim(recAssemblyUnitCost.Fields("opn_inst_cost_op").Value)
                        End If
                    End If
                End If
            End If
        End If
    End With
    
    recAssemblyUnitCost.Close
    Screen.MousePointer = vbNormal
End Sub

Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    Dim strUpdate       As String
    Dim blnReturn       As Boolean
    Dim strError        As String
    Dim clsRowInfo      As CRowInfo
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
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
                    MsgBox "Assembly deleted successfully.", vbInformation
                Else
                    m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                End If
            Else
                '
                '   Build SQL statement.
                With TDBGrid
                    '
                    '   The skey_type is always 'SF' for Square Foot models from here.
                    strUpdate = "exec sp_delete_assembly_usage @skey_type = 'SF',"
                    strUpdate = strUpdate & "@parent_skey = '"
                    strUpdate = strUpdate & m_rec.Fields("bldg_model_skey").Value & "', @assembly_skey = '"
                    strUpdate = strUpdate & m_rec.Fields("assembly_skey").Value & "', @formula_code = '"
                    strUpdate = strUpdate & Trim(m_rec.Fields("prior_formula_code").Value) & "'"
                End With
                
                If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
                    If strError = "" Then
                        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
                        .SetRowState m_rec.Bookmark, STATE_DELETED
                        '
                        '   Remove it from recordset so it doesn't show.
                        m_rec.Delete
                        If bMultiDeleteInProcess = False Then
                            MsgBox "Assembly deleted successfully.", vbInformation
                            '
                            '   Now refresh costs
                            RaiseEvent RefreshCostsAssemblyDeleted
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
        Else
            '
            '   Build SQL statement.
            With TDBGrid
                '
                '   The skey_type is always 'SF' for Square Foot models from here.
                strUpdate = "exec sp_delete_assembly_usage @skey_type = 'SF',"
                strUpdate = strUpdate & "@parent_skey = '"
                strUpdate = strUpdate & m_rec.Fields("bldg_model_skey").Value & "', @assembly_skey = '"
                strUpdate = strUpdate & m_rec.Fields("assembly_skey").Value & "', @formula_code = '"
                strUpdate = strUpdate & Trim(m_rec.Fields("prior_formula_code").Value) & "'"
            End With
            
            If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
                If strError = "" Then
                    TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
                    .SetRowState m_rec.Bookmark, STATE_DELETED
                    '
                    '   Remove it from recordset so it doesn't show.
                    m_rec.Delete
                    'MsgBox "Assembly deleted successfully", vbInformation
                    If bMultiDeleteInProcess = False Then
                        MsgBox "Assembly deleted successfully.", vbInformation
                        '
                        '   Now refresh costs
                        RaiseEvent RefreshCostsAssemblyDeleted
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
    Screen.MousePointer = vbNormal
End Sub

Private Sub TDBGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim varButton
    Dim I                           As Integer
    
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
    Dim I                           As Integer
    
    On Error Resume Next
    With TDBGrid
        If .SelBookmarks.Count > 0 Then
            bMultiDeleteInProcess = True
            m_NumberOfSuccessfulUpdates = 0
            
            ReDim varSelectedBookmarks(0 To .SelBookmarks.Count)
            I = 0
            For Each varCurrentSelectedBookmark In .SelBookmarks
                varSelectedBookmarks(I) = varCurrentSelectedBookmark
                I = I + 1
            Next varCurrentSelectedBookmark
            For I = 0 To UBound(varSelectedBookmarks) - 1
                .Bookmark = varSelectedBookmarks(I)
                .Delete
            Next I
            MsgBox m_NumberOfSuccessfulUpdates & " Assemblies were deleted successfully."
            bMultiDeleteInProcess = False
            m_NumberOfSuccessfulUpdates = 0
            '
            '   Now refresh costs
            RaiseEvent RefreshCostsAssemblyDeleted
        End If
    End With
End Function

' We call this routine when we clone the assemblies because otherwise, if the row info structure
' isn't set up, the update code will think the row hasn't changed and therefore skip the actual update.
' This routine sets the row state to new.  The Set Row State
' routine will create a new row info rec if necessary.
Public Sub UpdateRowInfo(dbRec As ADODB.RecordSet)
   
    dbRec.MoveFirst
    Do While Not dbRec.EOF
    
        m_objGridMap.SetRowState dbRec.Bookmark, STATE_NEW
    
        dbRec.MoveNext
    Loop
    
End Sub


