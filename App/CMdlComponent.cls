VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMdlComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
Private m_blnPreferences As Boolean
Private m_varSaveBookmark As Variant
Private m_varCurrentBookmark As Variant
Private m_colUpdatedRows As New Collection
'
'   Used in the SuccessfulUpdates Property Get.
Private m_NumberOfSuccessfulUpdates As Integer
'
'   Used in the UpdateErrors Property Get.
Private m_NumberOfUpdateErrors As Integer

Private Sub Class_Initialize()
    With m_objGridMap.m_colCloneCols
    
        .Add "bldg_model_skey"
        .Add "bldg_skey"
        .Add 1, "bldg_id"
        .Add 1, "bldg_desc"
        .Add 1, "model_code"
        .Add 1, "frame_type"
        .Add 1, "wall_type"

        .Add 1, "class_id"
        .Add 1, "system_component"
        .Add 1, "specifications"
        .Add 1, "unit"
        .Add 1, "unit_cost"
        .Add 1, "cost_per_sf_inst"
        .Add 1, "cost_per_sf_mat"
        .Add 1, "cost_per_sf"
        .Add 1, "pct_of_total"

        .Add 1, "last_update_date"
        .Add 1, "last_update_person"
        .Add 1, "last_update_id"
    End With
End Sub

Public Sub InitGrid()
    '
    '   Identifies the Grid.
    
    TDBGrid.AllowAddNew = True   'rlh 03/04/2009  - CCD 8.4
    
    With m_objGridMap
        .GRIDTYPE = "Summary Estimate"
        
        .AddCol "Sort Order", "sort_order", BLDGCOMPONENT_TABLE, 750, True, False, False
        .AddCol "Class ID", "ref_num", BLDGCOMPONENT_TABLE, 750, True, True, False
       
        .AddCol "System Component", "system_component", BLDGCOMPONENT_TABLE, 2000, True, True, False
        .AddCol "Specifications", "specifications", BLDGCOMPONENT_TABLE, 2000, False, True, False
        .AddCol "Unit", "unit", BLDGCOMPONENT_TABLE, 1000, False, True, False
        .AddCol "Unit Cost", "unit_cost", BLDGCOMPONENT_TABLE, 1000, True, True, False, FORMAT_PRICE
        .AddCol "Cost Per SF", "cost_per_sf", BLDGCOMPONENT_TABLE, 1000, True, True, False, FORMAT_PRICE
        .AddCol "Pct Of Total", "pct_of_total", BLDGCOMPONENT_TABLE, 750, True, True, False
        
        .AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
        .AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
        .AddCol "last_update_id", "last_update_id", BLDGCOMPONENT_TABLE, 800, True, False, False
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

Public Function Update(bAllAreas As Boolean, nmodel_form As Integer) As Boolean
    Dim strUpdate As String ' SQL string
    Dim sDesc   As String
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
                Set clsRowInfo = m_objGridMap.rowinfo(.Bookmark)
                '   If record was added meaning clsRowInfo.State = STATE_NEW
                '   Or clsRowInfo.State = STATE_MODIFIED Then
                If Not clsRowInfo Is Nothing Then
                    If clsRowInfo.State <> STATE_NONE Then
                        strUpdate = "exec sp_update_published_bldg_component_cost_holding_table @bldg_model_skey = '"
                        strUpdate = strUpdate & Trim(.Fields("bldg_model_skey").Value) & "',"
                       
                        strUpdate = strUpdate & "@bldg_area  = '" & Trim(.Fields("bldg_area").Value) & "',"
                        strUpdate = strUpdate & "@class_system_id  = '" & Trim(.Fields("class_system_id").Value) & "',"
                        strUpdate = strUpdate & "@ref_num  = '" & Trim(.Fields("ref_num").Value) & "',"
                        strUpdate = strUpdate & "@sort_order  = '" & Trim(.Fields("sort_order").Value) & "',"
                        strUpdate = strUpdate & "@op_code  = '" & Trim(.Fields("op_code").Value) & "',"
                        strUpdate = strUpdate & "@country_code  = '" & Trim(.Fields("country_code").Value) & "',"
                        strUpdate = strUpdate & "@region_code  = '" & Trim(.Fields("region_code").Value) & "',"
                        '
                        '   Since we might have ' marks in our desc, need to replace with '' for SQL.
                        sDesc = Trim(.Fields("specifications").Value)
                        sDesc = Replace(sDesc, "'", "''", 1)
        
                        strUpdate = strUpdate & "@specifications  = '" & sDesc & "',"
                        '
                        '   Should we apply the changes to all of the areas within the bldg.
                        strUpdate = strUpdate & "@bAllAreas = '" & IIf(bAllAreas = True, "1", "0") & "',"
                        strUpdate = strUpdate & "@unit  = '" & Trim(.Fields("unit").Value) & "',"
    
                        strUpdate = strUpdate & "@last_update_id = '" & Trim(.Fields("last_update_id").Value) & "',"
                        strUpdate = strUpdate & "@model_form = '" & nmodel_form & "'"
                        '
                        '   Reset on success
                        
                        ' rlh
                        If DEBUGON Then
                            Debug.Print "CMdlComponent:Update: " & strUpdate
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
            .UpdateBatch
        End If
    End With
    TDBGrid.ReBind
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
    GRIDTYPE = "Summary Estimate"
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

Private Sub TDBGrid_Change()
    With TDBGrid
        If .Columns(.col).Caption = "Bldg Desc" Then
            '
            '   Truncate the bldg desc to only 75 chars.
            If Len(.Text) > 75 Then .Text = Left(Trim(.Text), 75)
        End If
    End With
End Sub
