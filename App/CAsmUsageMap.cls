VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAsmUsageMap"
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

Const m_GRIDTYPE = "Assembly_Usage" ' Identity of grid (Seen as REGISTRY folder)
'Const m_GRIDTYPE = "Assembly_Usage_1" ' Identity of grid (Seen as REGISTRY folder)
'############################### IMPORTANT ######################################
'
'IMPORTANT!!!
'
'To date, 6/11/2010, use "Assembly_Usage"
'What was happening that for a DataGrid with no existing "Preference Grid" in place
'the VISIBLE property (as seen in the registry) was NEVER being set to true (was always = 0)
'Upon hard coding each of the registry VISIBLE values to ffffff, the columns showed
'up in the grid!
'
'###############################################################################

' Names of some table this grid uses
'Const ASMUSAGE_TABLE = "Material_usage"
'Const ASMUSAGE_UPDATE = "Material_usage update"

Const ASMUSAGE_TABLE = "Assembly_DETAIL"
Const ASMUSAGE_UPDATE = "Assembly_usage update"

Const MATERIAL = "Material"     'Source validation
Const UNIT_COST = "Unit Cost"   'Source validation - see where we came from

'Const MATERIAL = "Model"     'Source validation
'Const UNIT_COST = "Assembly"   'Source validation - see where we came from

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private m_blnMat_ID_Error As Boolean
Private m_blnUC_ID_Error As Boolean
Private m_strMaterialID As String
Private m_strUnitCostID As String
Private m_blnUseMatDft As Boolean
Private m_blnUseUCDft As Boolean
Private m_iPostMsgID As Integer

Public strSource As String   'Source of form - Unit Cost, Material, or None

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Sub FillMaterial(strSelMatID As String, Optional m_recBookmark As Variant = Empty)
    Dim strSQL As String
    
    If Not IsEmpty(m_recBookmark) Then
        
        m_rec.Bookmark = m_recBookmark
    End If
    
    FillFields ("Select mat_id,  mat_skey, usage_unit, tech_desc as material_tech_desc, metric_tech_desc as material_metric_tech_desc from Material where mat_id='" + strSelMatID + "'")
    
    strSQL = "Select ave_mat_price, country_code as material_country_code, region_code as material_region_code" + _
        " from published_Material_price " + _
        " where mat_skey=" + CStr(m_rec.Fields("mat_skey").Value) + _
        " and start_date <= getdate()" + _
        " and term_date >= getdate()" + _
        " and country_code = 'USA'" + _
        " and region_code = 'NAT'"
    
    FillFields (strSQL)

End Sub

Public Sub FillUnitCost(strSelUCID As String, Optional iMasterFormatVersion As Long = UCD_MASTERFORMAT_VERSION)
'MODIFIED 8/25/2005 RTD - CORRECTS PROBLEM REPORTED BY G. MEDEIROS
    Dim sSQL As String
    
    If iMasterFormatVersion = UCD_MASTERFORMAT_VERSION Then
        sSQL = "SELECT unit_cost_id, ext_unit_cost_id, unit_cost_skey, tech_desc as unit_tech_desc, " + _
            " metric_tech_desc as unit_metric_tech_desc, ucd.unit " + _
            " FROM vw_unit_cost_detail AS ucd " + _
            " WHERE unit_cost_id = '" + strSelUCID + "'"
    Else
        sSQL = "SELECT unit_cost_id, ext_unit_cost_id, unit_cost_skey, ext_tech_desc as unit_tech_desc," & _
            " ext_metric_tech_desc as unit_metric_tech_desc, ucd.unit " & _
            " FROM vw_UNIT_COST_DETAIL AS UCD" & _
            " WHERE ext_unit_cost_id = '" + strSelUCID + "'"
    End If
    FillFields (sSQL)
    'TDBGrid.ReBind

End Sub

Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
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

Public Property Let Material_id(strMatID As String)
    m_strMaterialID = Compress_String(strMatID)
End Property

Public Property Let UnitCost_ID(strUCID As String)
    m_strUnitCostID = strUCID
End Property

Public Sub InitGrid()
    
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
    
    'm_objGridMap.SetFrozenCols (3)
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "Bldg Desc", "bldg_desc", ASMUSAGE_TABLE, 2000, True, False
    m_objGridMap.AddCol "Assembly ID", "Assembly_id", ASMUSAGE_TABLE, 1500, True, False
'    m_objGridMap.AddCol "Tech Desc", "tech_desc", ASMUSAGE_TABLE, 10000, False, True ', True ', FORMAT_DECIMAL '--vbNullString, 1500, True, False

    ' 8/19/2005 RTD - ADD COLUMN FOR EXT-UNIT-COST-ID.
    'm_objGridMap.AddCol "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2), "ext_unit_cost_id", "", 1500, , , , FORMAT_UNIT_COST_04
    'm_objGridMap.AddCol "Alt Mat ID", "alt_mat_id", vbNullString, 1500, True, False, True
    
    '################################################################################
    '
    ' HAVING TO COMMENT OUT THE "AddSplit" as it prevents pulling out the remaining
    ' grid columns
    '
    '################################################################################
    m_objGridMap.AddSplit
    
    m_objGridMap.AddCol "Tech Desc", "tech_desc", ASMUSAGE_TABLE, 4000, False, True, True ', FORMAT_DECIMAL '--vbNullString, 1500, True, False
    m_objGridMap.AddCol "Bldg Model Skey", "bldg_model_skey", ASMUSAGE_TABLE, 1500, True, False
    m_objGridMap.AddCol "Bldg ID", "bldg_id", ASMUSAGE_TABLE, 1500, True, False
    m_objGridMap.AddCol "Model Code", "model_code", ASMUSAGE_TABLE, 1100, True, False
    m_objGridMap.AddCol "Frame Type", "frame_type", ASMUSAGE_TABLE, 1500, True, False
    m_objGridMap.AddCol "Wall Type", "wall_type", ASMUSAGE_TABLE, 1500, True, False



'
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
'''        If strSource = MATERIAL Then
'''            If TDBGrid.Columns("Material ID").Order > TDBGrid.Columns("Unit Cost ID").Order Then
'''                TDBGrid.Columns("Material ID").Order = TDBGrid.Columns("Unit Cost ID").Order
'''                TDBGrid.Columns("Unit Cost ID").Order = TDBGrid.Columns("Material ID").Order + 1
'''                TDBGrid.Columns("Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)).Order = TDBGrid.Columns("Unit Cost ID").Order + 1
'''                TDBGrid.ReBind
'''            End If
'''        ElseIf strSource = UNIT_COST Then
'''            If TDBGrid.Columns("Unit Cost ID").Order > TDBGrid.Columns("Material ID").Order Then
'''                TDBGrid.Columns("Unit Cost ID").Order = TDBGrid.Columns("Material ID").Order
'''                TDBGrid.Columns("Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)).Order = TDBGrid.Columns("Unit Cost ID").Order + 1
'''                TDBGrid.Columns("Material ID").Order = TDBGrid.Columns("Unit Cost ID").Order + 2
'''                TDBGrid.ReBind
'''            End If
'''        End If
        
        '#######################################################################################################
        '
        'SOLUTION to missing columns:
        '
        'Converted all of the column names, as seen in the registry under the mGridType folder name,
        'from visible = 0  to ffffffff
        '
        '#######################################################################################################

        '#############################################
            TDBGrid.ReBind
        'UPON ADDING THE TDBGrid.Rebind here, ensuring that it will always be called,
        'the columns still don't show (mGridType="Assembly_Usage") but could be pulled out!!!
        '
        '#############################################
    End If
    
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    Dim colUnitCostID As New Collection
    
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
            ' If record was added, or if both mat_id and unit_cost_id were changed
                If clsRowInfo.State = STATE_NEW Or clsRowInfo.State = STATE_MODIFIED Then
                    strUpdate = "exec sp_update_mat_usage_driver "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ASMUSAGE_TABLE
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    strUpdate = strUpdate + ", @last_update_id=" + str(m_rec.Fields("last_update_id").Value)
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If strError = "" Then
                        intSuccess = intSuccess + 1
'                        If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
                        '8/25/2005 RTD - SAVE EXT_UCID IF UCID IS EMPTY
                        If m_rec.Fields("unit_cost_ID") <> "" Then
                            SaveUnitCostID colUnitCostID, m_rec.Fields("unit_cost_ID")
                        Else
                            SaveUnitCostID colUnitCostID, m_rec.Fields("ext_unit_cost_ID")
                        End If
'                        End If
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                        m_rec.Fields("last_update_date").Value = Now
                        m_rec.Fields("last_update_person").Value = strUserName
'                        Reset_Orig_Values m_rec
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
        strMsg = str(intSuccess) + " rows updated successfully." + Chr(10)
        'Update the unit cost for the modified materials.
        UpdateUnitCost colUnitCostID
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
    Exit Function
    
CleanUp:
    MsgBox "Error in Update"
End Function

Private Sub Class_Initialize()
'''    m_objGridMap.m_colCloneCols.Add 1, "mat_id"
'''    m_objGridMap.m_colCloneCols.Add 1, "unit_cost_id"
'''    m_objGridMap.m_colCloneCols.Add 1, "alt_mat_id"
'''    m_objGridMap.m_colCloneCols.Add 1, "material_tech_desc"
'''    m_objGridMap.m_colCloneCols.Add 1, "material_metric_tech_desc"
'''    m_objGridMap.m_colCloneCols.Add 1, "unit_tech_desc"
'''    m_objGridMap.m_colCloneCols.Add 1, "unit_metric_tech_desc"
'''    m_objGridMap.m_colCloneCols.Add 1, "usage_unit"
'''    m_objGridMap.m_colCloneCols.Add 1, "material_country_code"
'''    m_objGridMap.m_colCloneCols.Add 1, "material_region_code"
'''    m_objGridMap.m_colCloneCols.Add 1, "unit_qty"
'''    m_objGridMap.m_colCloneCols.Add 1, "input_factor"
'''    m_objGridMap.m_colCloneCols.Add 1, "output_factor"
'''    m_objGridMap.m_colCloneCols.Add 1, "adj_factor"
'''    m_objGridMap.m_colCloneCols.Add 1, "unit"
'''    m_objGridMap.m_colCloneCols.Add 1, "comment"
'    m_objGridMap.m_colCloneCols.Add 1, "mat_skey"
'    m_objGridMap.m_colCloneCols.Add 1, "unit_cost_skey"

'rlh 06/09/2010
'    m_objGridMap.m_colCloneCols.Add 1, "bldg_desc"
'    m_objGridMap.m_colCloneCols.Add 1, "Assembly_id"
'    'm_objGridMap.m_colCloneCols.Add 1, "alt_mat_id"
'    m_objGridMap.m_colCloneCols.Add 1, "tech_desc"
'    m_objGridMap.m_colCloneCols.Add 1, "bldg_model_skey"
'    m_objGridMap.m_colCloneCols.Add 1, "model_code"
'    m_objGridMap.m_colCloneCols.Add 1, "frame_type"
'    m_objGridMap.m_colCloneCols.Add 1, "wall_type"
'
'    m_objGridMap.m_colCloneCols.Add 1, "material_country_code"
'    m_objGridMap.m_colCloneCols.Add 1, "material_region_code"
'    m_objGridMap.m_colCloneCols.Add 1, "unit_qty"
'    m_objGridMap.m_colCloneCols.Add 1, "input_factor"
'    m_objGridMap.m_colCloneCols.Add 1, "output_factor"
'    m_objGridMap.m_colCloneCols.Add 1, "adj_factor"
'    m_objGridMap.m_colCloneCols.Add 1, "unit"
'    m_objGridMap.m_colCloneCols.Add 1, "comment"
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Material ID"
        If TDBGrid.AddNewMode = 2 Then
            TDBGrid.PostMsg 1
        Else
            FillMaterial TDBGrid.Text
        End If
    Case "Unit Cost ID"
        If TDBGrid.AddNewMode = 2 Then
            TDBGrid.PostMsg 2
        Else
            FillUnitCost TDBGrid.Text, UCD_MASTERFORMAT_VERSION
        End If
    ' 8/25/2005 RTD - ADD SUPPORT FOR MASTERFORMAT 2004 (EXT ID)
    Case "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)
        If TDBGrid.AddNewMode = 2 Then
            TDBGrid.PostMsg 3
        Else
            FillUnitCost TDBGrid.Text, EXT_MASTERFORMAT_VERSION
        End If
    End Select
End Sub

Private Sub TDBGrid_BeforeDelete(Cancel As Integer)
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbQuestion)
    If varButton = vbNo Then
        Cancel = True
    End If

End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Error_Processing
    If Trim(TDBGrid.Columns("Material Id").Value) = "" Then
        m_blnMat_ID_Error = True
        Cancel = True
    End If
    ' UPDATED 8/25/2005 RTD - UNIT COST OR EXT UNIT COST CAN BE BLANK, BUT NOT BOTH
    If Trim(TDBGrid.Columns("Unit Cost Id").Value) = "" And Trim(TDBGrid.Columns("Unit Cost Id " & Right(EXT_MASTERFORMAT_VERSION, 2)).Value) = "" Then
        m_blnUC_ID_Error = True
        Cancel = True
    End If
    
Exit_Sub:
    Exit Sub
    
Error_Processing:
    Cancel = True
    Resume Exit_Sub
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
    If m_blnMat_ID_Error = True Then
        MsgBox "Please enter a valid Material ID.", vbExclamation
        m_blnMat_ID_Error = False
    End If
    If m_blnUC_ID_Error = True Then
        MsgBox "Please enter a valid Unit Cost ID.", vbExclamation
        m_blnUC_ID_Error = False
    End If
    Response = 0
    TDBGrid.SetFocus

End Sub

Private Sub TDBGrid_OnAddNew()
''Fill defaults from the grid values
'TDBGrid.Columns("Input Factor").Value = 1
'TDBGrid.Columns("Output Factor").Value = 1
'TDBGrid.Columns("Adj Factor").Value = 1
'TDBGrid.Columns("Unit Qty").Value = 1
'TDBGrid.Columns("last_update_id") = 0

End Sub

Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)
Dim blnPositionToUC As Boolean
Dim blnPositionToMat As Boolean
Dim blnUseDftMat As Boolean
Dim blnSetDft As Boolean
Dim strSelMatID As String
Dim strSelUCID As String
Dim strSaveBookMark As String

Dim i As Integer
Dim j As Integer

On Error Resume Next

If MsgId = 1 Then       'Material changed
    strSelMatID = TDBGrid.Columns("Material ID").Text
    If m_blnUseUCDft = True Then
        strSelUCID = m_strUnitCostID
        blnPositionToUC = False         'Fill in UC, position to qty
    Else
        blnPositionToUC = True          'UC not filled in, position to it
    End If
    blnSetDft = True
End If

If MsgId = 2 Then       'Unit Cost ID changed
    strSelUCID = TDBGrid.Columns("Unit Cost ID").Text
    If m_blnUseMatDft = True Then
        strSelMatID = "M" + m_strMaterialID
        blnPositionToMat = False
    Else
        blnPositionToMat = True
    End If
    blnSetDft = True
End If

If MsgId = 3 Then       'EXT Unit Cost ID changed
    strSelUCID = TDBGrid.Columns("Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)).Text
    If m_blnUseMatDft = True Then
        strSelMatID = "M" + m_strMaterialID
        blnPositionToMat = False
    Else
        blnPositionToMat = True
    End If
    blnSetDft = True
End If

If blnSetDft = True Then
    m_rec.AddNew
    FillMaterial strSelMatID
    FillUnitCost strSelUCID

    m_rec.Fields("input_factor").Value = 1
    m_rec.Fields("output_factor").Value = 1
    m_rec.Fields("adj_factor").Value = 1
    m_rec.Fields("unit_qty").Value = 1
    m_rec.Fields("last_update_id").Value = 0

    m_rec.Update
    TDBGrid.ReBind
    TDBGrid.SetFocus
    TDBGrid.MoveLast
    m_rec.MoveLast
    Set TDBGrid.EditActive = True

    If blnPositionToMat = True Then
        TDBGrid.Split = 0
        TDBGrid.Col = TDBGrid.Columns("Material ID").ColIndex
    ElseIf blnPositionToUC = True Then
        TDBGrid.Split = 0
        TDBGrid.Col = TDBGrid.Columns("Unit Cost ID").ColIndex
    Else
        TDBGrid.Split = 1
        TDBGrid.Col = TDBGrid.Columns("Unit Qty").ColIndex
    End If
    m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW

End If

End Sub

Private Sub FillFields(strSQL As String)
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnReturn As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim varCurBookmark As Variant

    On Error Resume Next
    If m_rec.EOF = True Then
        'On Error GoTo Err_Handler
        'varCurBookmark = m_rec.Bookmark
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSQL, rec)
        If Not rec.EOF Then
            'On Error Resume Next
            For i = 0 To rec.Fields.Count - 1
               m_rec.Fields(rec.Fields(i).Name).Value = rec.Fields(i).Value
               For j = 0 To TDBGrid.Columns.Count
                   If TDBGrid.Columns(j).DataField = rec.Fields(i).Name Then
                       TDBGrid.Columns(j).Value = rec.Fields(i).Value
                       Exit For
                   End If
               Next j
            Next i
        End If
        rec.Close
        Set rec = Nothing
        'm_rec.Bookmark = varCurBookmark
    End If
    Exit Sub

Err_Handler:
    Debug.Print "CMatUsageMap.FillFields() error: " & Err.Description
    
End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
    Dim colUnitCostID As New Collection
   
    ' Position to the desired record and execute delete
    m_rec.Bookmark = Bookmark
    
    Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
    If Not clsRowInfo Is Nothing Then
        If clsRowInfo.State = STATE_NEW Then
            TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
            m_rec.Delete ' Remove it from recordset so it doesn't show.
            MsgBox "Delete successful.", vbInformation
            Exit Sub
        End If
    End If
    
    ' Build SQL statement
    strUpdate = "exec sp_delete_material_usage "
    strUpdate = strUpdate + "@mat_skey=" + str(m_rec.Fields("mat_skey")) + ","
    strUpdate = strUpdate + " @unit_cost_skey=" + str(m_rec.Fields("unit_cost_skey"))
    strUpdate = strUpdate + ", @UseCommitControl=1"
    strUpdate = strUpdate + ", @verifylastmaterial=1"
    
    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetError m_rec.Bookmark, strError
        m_rec.Fields("Error").Value = "Error"
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        '8/25/2005 RTD - PASS EXT_UCID IF UNIT_COST_ID IS EMPTY
        If m_rec.Fields("unit_cost_id") <> "" Then
            SaveUnitCostID colUnitCostID, m_rec.Fields("unit_cost_id")
        Else
            SaveUnitCostID colUnitCostID, m_rec.Fields("ext_unit_cost_id")
        End If
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
        MsgBox "Delete successful.", vbInformation
    End If
'    If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
        UpdateUnitCost colUnitCostID
'    End If
    Exit Sub

NoDelete:
    MsgBox Err.Description, vbExclamation
    Bookmark = Null
    
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

    On Error Resume Next
    Dim strSelect As String
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnReturn As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim int_mat_skey As Integer
    Dim strMaterialID As String
    Dim strUnitCostId As String
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
    Dim vntSaveM_RecBookmark As Variant
    Dim vntMyBookmark As Variant
    Dim strError As String

    vntSaveM_RecBookmark = m_rec.Bookmark
    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Material ID"
        If (Len(m_strMaterialID) = 0 And Len(m_strUnitCostID) = 12) And Right(m_strUnitCostID, 1) <> "*" Then
            m_blnUseUCDft = True
        Else
            m_blnUseUCDft = False
        End If
        m_iPostMsgID = 1 'Used for addnew - post defaults

        If Trim(TDBGrid.Text) = "" Then
            MsgBox "The Material ID may not be blank."
            TDBGrid.Text = OldValue
            Cancel = True ' Cancel so old value is replaced
        Else
            strSelect = "sp_validate_material_usage @key_type=2, @key = '" + Compress_String(TDBGrid.Text) + "'"
            ' Use DAL to perform select
            blnReturn = g_objDAL.ExecQuery(vbNullString, strSelect, strError)
            If blnReturn = False Then
            'rlh
            'rlh 03/05/09  Mel Mossman false alarm (bad material id) However the error handling
            'here is ATROCIOUS!!!!
                    
                If (InStr(Err.Description, "50030") > 0) Then  'rlh 03/05/09
                    MsgBox "The SKEY for the Material ID " + TDBGrid.Text + " does not exist."
                Else
                    MsgBox "Error encountered while processing Material ID " + TDBGrid.Text
                End If   'rlh 03/05/09 (end of block)
                
                'MsgBox strError
                TDBGrid.Text = OldValue
                Cancel = True ' Cancel so old value is replaced
            Else
                ' Check to see if the mat_id entered exists already
               strSelect = "Select mat_skey from Material where mat_id='" + Compress_String(TDBGrid.Text) + "'"
                ' Use DAL to perform select
                blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                If rec.RecordCount = 0 Then
                    MsgBox "The Material ID " + TDBGrid.Text + " does not exist."
                    TDBGrid.Text = OldValue
                    Cancel = True ' Cancel so old value is replaced
                Else
                    strMaterialID = Trim(Compress_String(TDBGrid.Text))
                    strUnitCostId = Compress_String(TDBGrid.Columns("Unit Cost ID").Value)
                    'Check to see if the material has been assigned for this unit cost - may not have dups
                    vntMyBookmark = TDBGrid.Bookmark
                    If m_rec.RecordCount > 0 Then
'                        If TDBGrid.Columns("Unit Cost ID").Value > "" Then
                            m_rec.MoveFirst
                            Do Until m_rec.EOF
                                If Trim(m_rec.Fields("mat_id")) = strMaterialID Then
                                    If Trim(m_rec.Fields("unit_cost_id")) = strUnitCostId Then
                                        MsgBox "The Material ID " + TDBGrid.Text + " has already been assigned to this unit cost, and may not be assigned twice."
                                        Cancel = True
                                        TDBGrid.Text = OldValue
                                        TDBGrid.SetFocus
                                        Exit Sub
'                                    Else
'                                        If m_blnUseUCDft = True And Trim(m_rec.Fields("unit_cost_id")) = m_strUnitCostID Then
'                                            m_blnUseUCDft = False
'                                        End If
                                    End If
                                End If
                                m_rec.MoveNext
                            Loop
'                        End If
                    End If
                End If
                rec.Close
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
    Case "Unit Cost ID"
        If Left(m_strMaterialID, 1) = "M" Then m_strMaterialID = Right(m_strMaterialID, Len(m_strMaterialID) - 1)
        If (Len(m_strUnitCostID) = 0 And Len(m_strMaterialID) = 10) And Right(m_strMaterialID, 1) <> "*" Then
            m_blnUseMatDft = True
        Else
            m_blnUseMatDft = False
        End If
        m_iPostMsgID = 2    'Used for addnew - post defaults
        ' Check to see if the unit_cost_id entered exists already
        strSelect = "Select unit_cost_skey, type_code from unit_cost_detail where unit_cost_id='" + Compress_String(TDBGrid.Text) + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        If rec.RecordCount = 0 Then
            MsgBox "The Unit Cost ID " + TDBGrid.Text + " does not exist."
            Cancel = True
        ElseIf rec.Fields("type_code") = "E" Or rec.Fields("type_code") = "H" Or rec.Fields("type_code") = "B" Then
            MsgBox "E, H, and B Type Unit Costs are not valid for Material Usage."
            Cancel = True
        Else
            strUnitCostId = Trim(TDBGrid.Text)
            strMaterialID = TDBGrid.Columns("Material ID").Value
            vntMyBookmark = TDBGrid.Bookmark
            If m_rec.RecordCount > 0 Then
    '            If TDBGrid.Columns("Unit Cost ID").Value > "" Then
                    m_rec.MoveFirst
                    Do Until m_rec.EOF
                        If Trim(m_rec.Fields("unit_cost_id")) = strUnitCostId Then
                            If Trim(m_rec.Fields("mat_id")) = strMaterialID Then
                                MsgBox "The Material ID " + TDBGrid.Text + " has already been assigned to this unit cost, and may not be assigned twice."
                                Cancel = True
                                TDBGrid.SetFocus
                                Exit Sub
                            Else
                                If m_blnUseMatDft = True And Trim(m_rec.Fields("mat_id")) = "M" + m_strMaterialID Then
                                    m_blnUseMatDft = False
                                End If
                            End If
                        End If
                        m_rec.MoveNext
                    Loop
    '            End If
            End If
            rec.Close
            Set rec = Nothing
        End If
    '8/25/2005 RTD - CHECK THE EXT UNIT COST ID
    Case "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)
        If Left(m_strMaterialID, 1) = "M" Then m_strMaterialID = Right(m_strMaterialID, Len(m_strMaterialID) - 1)
        If (Len(m_strUnitCostID) = 0 And Len(m_strMaterialID) = 10) And Right(m_strMaterialID, 1) <> "*" Then
            m_blnUseMatDft = True
        Else
            m_blnUseMatDft = False
        End If
        m_iPostMsgID = 3    'Used for addnew - post defaults
        ' Check to see if the unit_cost_id entered exists already
        strSelect = "Select unit_cost_skey, type_code from vw_unit_cost_detail where ext_unit_cost_id='" + Compress_String(TDBGrid.Text) + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        If rec.RecordCount = 0 Then
            MsgBox "The Ext Unit Cost ID " + TDBGrid.Text + " does not exist.", vbCritical
            Cancel = True
        ElseIf rec.Fields("type_code") = "E" Or rec.Fields("type_code") = "H" Or rec.Fields("type_code") = "B" Then
            MsgBox "E, H, and B Type Unit Costs are not valid for Material Usage.", vbCritical
            Cancel = True
        Else
            strUnitCostId = Compress_String(Trim(TDBGrid.Text))
            strMaterialID = TDBGrid.Columns("Material ID").Value
            vntMyBookmark = TDBGrid.Bookmark
            If m_rec.RecordCount > 0 Then
    '            If TDBGrid.Columns("Unit Cost ID").Value > "" Then
                    m_rec.MoveFirst
                    Do Until m_rec.EOF
                        If Trim(m_rec.Fields("ext_unit_cost_id")) = strUnitCostId Then
                            If Trim(m_rec.Fields("mat_id")) = strMaterialID Then
                                MsgBox "The Material ID " + TDBGrid.Text + " has already been assigned to this unit cost, and may not be assigned twice.", vbCritical
                                Cancel = True
                                TDBGrid.SetFocus
                                Exit Sub
                            Else
                                If m_blnUseMatDft = True And Trim(m_rec.Fields("mat_id")) = "M" + m_strMaterialID Then
                                    m_blnUseMatDft = False
                                End If
                            End If
                        End If
                        m_rec.MoveNext
                    Loop
    '            End If
            End If
            rec.Close
            Set rec = Nothing
        End If
    
    End Select
m_rec.Bookmark = vntSaveM_RecBookmark

End Sub

Public Function CloneRow() As ADODB.RecordSet
    ' Pass to base class
    m_objGridMap.CloneRow (True)
    Set CloneRow = Nothing
End Function

Public Property Get GRIDTYPE() As String
    GRIDTYPE = m_GRIDTYPE
End Property
Public Function CloneRowRecordset() As ADODB.RecordSet
    Set CloneRowRecordset = m_objGridMap.CloneRowRecordset
End Function
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
Public Sub SetRowState(intBookmark As Integer, intState As Integer)
    m_objGridMap.SetRowState intBookmark, intState
End Sub


