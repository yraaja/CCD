VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModelMap"
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
Const MODEL_TABLE = "bldg_model"
Const BUILDING_TABLE = "bldg_detail"
Const PUBLISHED_BLDG_MATRIX_COST = "published_bldg_matrix_cost"
Const ALL_TABLES = "All"
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
        .Add 1, "bldg_model_skey"
        .Add 1, "bldg_skey"
        .Add 1, "bldg_id"
        .Add 1, "type_code"
        .Add 1, "bldg_category"
        .Add 1, "bldg_type"
        .Add 1, "bldg_desc"
        .Add 1, "costworks_desc"
        .Add 1, "model_code"
        .Add 1, "format_code"
        .Add 1, "wall_type"
        .Add 1, "frame_type"
        .Add 1, "format_row"
        .Add 1, "last_update_date"
        .Add 1, "last_update_person"
        .Add 1, "last_update_id"
    End With
End Sub

Public Sub InitGrid()
    Dim strSELECT   As String
    Dim recTemp     As ADODB.RecordSet
'*** APEX Migration Utility Code Change ***
'    Dim Item        As New TrueOleDBGrid70.ValueItem
    Dim Item        As New TrueOleDBGrid80.ValueItem
    '
    '   Identifies the Grid.
    With m_objGridMap
        .GRIDTYPE = "Model"
        '
        '   Add the columns to the grid, pointing to their
        '   corresponding db field names so that the grid can
        '   populate directly from the m_rec recordset.
        .AddCol "Mdl SKEY", "bldg_model_skey", MODEL_TABLE, 600, True, False, True
        .AddCol "Bldg SKEY", "bldg_skey", MODEL_TABLE, 600, True, False, True
        .AddCol "Bldg ID", "bldg_id", MODEL_TABLE, 770, True, True, False
        .AddCol "Type", "type_code", MODEL_TABLE, 500, True, True, False
        .AddCol "Category", "bldg_category", MODEL_TABLE, 1250, True, True, False
        .AddCol "Bldg Desc", "bldg_desc", MODEL_TABLE, 2000, True, True, False, FORMAT_STRING_TRIM
        '
        '   Main Split, this grid only has 1.
        '
        .AddSplit

        '.AddCol "Prior Model Code", "prior_model_code", MODEL_TABLE, 950, True, False, True
        .AddCol "Model Code", "model_code", MODEL_TABLE, 950, True, True, True
        .AddCol "Format Code", "format_code", MODEL_TABLE, 1000, False, True, True
        .AddCol "Wall Type", "wall_type", MODEL_TABLE, 3000, False, True, True, FORMAT_STRING_TRIM
        .AddCol "Frame Type", "frame_type", MODEL_TABLE, 2000, False, True, True, FORMAT_STRING_TRIM
        
        'ADDED COSTWORKS DESCRIPTION FIELD SO THAT IT IS AVAILABLE TO FORMS
        '7/5/2005 RTD (CR#1530)
        .AddCol "CostWorks Desc", "costworks_desc", MODEL_TABLE, 2000, False, False, True, FORMAT_STRING_TRIM
        
        .AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
        .AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
        .AddCol "last_update_id", "last_update_id", MODEL_TABLE, 800, True, False, False
        '
        '   Call base class.
        If m_blnPreferences Then
            .BuildPreferenceGrid
        Else
            .BuildGrid True, False
        End If
        '
        '   Give headings for the splits.
        With TDBGrid
            .Splits(0).Caption = "Building"
            .Splits(1).Caption = "Model"
            .Splits(1).Size = 5
            ' MODIFIED 6/16/05 RTD - MERGE PROPERTY IS NOW INTEGER [0|1|2], NOT BOOLEAN
            .Columns("Bldg Desc").Merge = 1
            .Columns("Bldg Desc").Style.Alignment = dbgLeft
            .Columns("Bldg Desc").Style.VerticalAlignment = dbgVertCenter
            .Columns("Bldg Desc").WrapText = True
        End With
        '
        '   Add items for frame combo drop down.
        strSELECT = "SELECT DISTINCT frame_type FROM bldg_model WHERE model_code != '0' AND " _
                        & "model_code != '7' AND model_code != '8' AND frame_type != '' ORDER BY frame_type"
        '
        '   Use DAL to perform select.
        If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recTemp) Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching.", vbCritical
        Else
            With TDBGrid
                With recTemp
                    While Not .EOF
                        Item.Value = Trim(.Fields("frame_type"))
                        TDBGrid.Columns("Frame Type").ValueItems.Add Item
                        .MoveNext
                    Wend
                End With
                .Columns("Frame Type").ValueItems.Presentation = dbgComboBox
                recTemp.Close
                Set recTemp = Nothing
            End With
        End If
        '
        '   Add items for wall combo drop down.
        strSELECT = "SELECT DISTINCT wall_type FROM bldg_model WHERE model_code != '0' AND " _
                        & "model_code != '7' AND model_code != '8' AND wall_type != '' ORDER BY wall_type"
        '
        '   Use DAL to perform select.
        If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recTemp) Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching.", vbCritical
        Else
            With TDBGrid
                With recTemp
                    While Not .EOF
                        Item.Value = Trim(.Fields("wall_type"))
                        TDBGrid.Columns("Wall Type").ValueItems.Add Item
                        .MoveNext
                    Wend
                End With
                .Columns("Wall Type").ValueItems.Presentation = dbgComboBox
                recTemp.Close
                Set recTemp = Nothing
            End With
        End If
        '
        '   Now add valid format codes to choose from
        With TDBGrid
            Item.Value = "A3"
            .Columns("Format Code").ValueItems.Add Item
            Item.Value = "A4"
            .Columns("Format Code").ValueItems.Add Item
            .Columns("Format Code").ValueItems.Presentation = dbgComboBox
            .Columns("Format Code").ValueItems.Validate = True
        End With
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
    Dim strUpdate           As String
    Dim strError            As String
    
    On Error Resume Next
    With m_rec
        '
        '   Note they cannot change the format_code or model_code once updated.
        strUpdate = "exec sp_update_commercial_model @bldg_model_skey = '" & Trim(.Fields("bldg_model_skey").Value) & "',"
        strUpdate = strUpdate & "@frame_type = '" & SQLFixString(Trim(.Fields("frame_type").Value)) & "',"
        strUpdate = strUpdate & "@wall_type = '" & SQLFixString(Trim(.Fields("wall_type").Value)) & "',"
        '
        '   assembly_usage & published_bldg_component_cost
        '   -ensures that we only add records that our form inserted.
        strUpdate = strUpdate & "@format_code = '" & Trim(.Fields("format_code").Value) & "',"
        strUpdate = strUpdate & "@model_form = '0',"
        ' ADDED @COSTWORKS_DESC 7/6/2005 RTD - TO SUPPORT CR#1530
        strUpdate = strUpdate & "@costworks_desc = '" & SQLFixString(Trim(.Fields("costworks_desc").Value)) & "',"
        strUpdate = strUpdate & "@last_update_id = '" & .Fields("last_update_id").Value & "',"
        strUpdate = strUpdate & " @last_update_person='" & strUserName & "'"
        '
        ' Reset on success
        If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
            If strError = "" Then
                m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                .Fields("last_update_date").Value = Now
                .Fields("last_update_person").Value = strUserName
                Reset_Orig_Values m_rec
                UpdateCommercial = True
            Else
                m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
                m_objGridMap.SetError .Bookmark, strError
                .Fields("Error").Value = "Error"
            End If
        Else
            m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
            m_objGridMap.SetError .Bookmark, strError
            .Fields("Error").Value = "Error"
        End If
    End With
End Function

Private Function UpdateResidential() As Boolean
    Dim strUpdate           As String
    Dim strError            As String
    
    On Error Resume Next
    With m_rec
        '
        '   Note they cannot change the format_code or model_code once updated.
        strUpdate = "exec sp_update_residential_model @bldg_model_skey = '" & Trim(.Fields("bldg_model_skey").Value) & "',"
        strUpdate = strUpdate & "@frame_type = '" & SQLFixString(Trim(.Fields("frame_type").Value)) & "',"
        strUpdate = strUpdate & "@wall_type = '" & SQLFixString(Trim(.Fields("wall_type").Value)) & "',"
        '   ADDED @FORMAT_CODE 6/29/2005 RTD TO CORRECT ERROR IN MODELS GRID UPDATE
        strUpdate = strUpdate & "@format_code = '" & Trim(.Fields("format_code").Value) & "',"
        '   ADDED @COSTWORKS_DESC 7/6/2005 RTD TO SUPPORT CR#1530
        strUpdate = strUpdate & "@costworks_desc = '" & SQLFixString(Trim(.Fields("costworks_desc").Value)) & "',"
        '
        '   assembly_usage & published_bldg_component_cost
        '   -ensures that we only add records that our form inserted.
        strUpdate = strUpdate & "@model_form = '0',"
        strUpdate = strUpdate & "@last_update_id = '" & .Fields("last_update_id").Value & "',"
        strUpdate = strUpdate & " @last_update_person='" & strUserName & "'"
        '
        ' Reset on success
        If g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
            If strError = "" Then
                m_NumberOfSuccessfulUpdates = m_NumberOfSuccessfulUpdates + 1
                .Fields("last_update_id").Value = .Fields("last_update_id").Value + 1
                .Fields("last_update_date").Value = Now
                .Fields("last_update_person").Value = strUserName
                Reset_Orig_Values m_rec
                UpdateResidential = True
            Else
                m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
                m_objGridMap.SetError .Bookmark, strError
                .Fields("Error").Value = "Error"
            End If
        Else
            m_NumberOfUpdateErrors = m_NumberOfUpdateErrors + 1
            m_objGridMap.SetError .Bookmark, strError
            .Fields("Error").Value = "Error"
        End If
    End With
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
    GRIDTYPE = "Model"
End Property

Public Property Get SuccessfulUpdates() As Integer
    SuccessfulUpdates = m_NumberOfSuccessfulUpdates
End Property

Public Property Get UpdateErrors() As Integer
    UpdateErrors = m_NumberOfUpdateErrors
End Property

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With TDBGrid
        '
        '   They cannot change the order of model codes for Residential Quality Series buildings.
        If .Columns("Bldg ID").Value = "100" Or .Columns("Bldg ID").Value = "200" _
        Or .Columns("Bldg ID").Value = "300" Or .Columns("Bldg ID").Value = "400" Then
            MsgBox "Model values cannot be changed for Residential Quality Series buildings.", vbCritical
            Cancel = True
        Else
            Select Case .Columns(ColIndex).Caption
                Case "Wall Type"
                    If Trim(.Text) = "" Then
                        MsgBox "Please provide a wall type description.", vbCritical
                        .Columns("Wall Type") = OldValue
                        Cancel = True
                    ElseIf Len(Trim(.Text)) > 30 Then
                        MsgBox "Please provide a wall type description that is less than 30 characters.", vbCritical
                        .Columns("Wall Type") = Left$(Trim(.Columns("Wall Type")), 30)
                    End If
                Case "Frame Type"
                    If Trim(.Columns("Type").Value) = "R" Then
                        MsgBox "Frame type descriptions are included within the wall type description for Residential buildings." _
                            & vbCrLf & "Please add the frame type to the wall type description.", vbCritical
                        Cancel = True
                    ElseIf Trim(.Text) = "" Then
                        MsgBox "Please select a frame type description.", vbCritical
                        Cancel = True
                    ElseIf Len(Trim(.Text)) > 30 Then
                        MsgBox "Please provide a frame type description that is less than 30 characters.", vbCritical
                        .Columns("Frame Type") = Left$(Trim(.Columns("Frame Type")), 30)
                    End If
                Case "Format Code"
                    If Trim(.Columns("Model Code").Value) = "7" Or Trim(.Columns("Model Code").Value) = "8" Then
                        MsgBox "Format Codes cannot be changed for Models 7 & 8 which are format rows.", vbCritical
                        Cancel = True
                    End If
            End Select
        End If
    
    End With
    Screen.MousePointer = vbNormal
End Sub


