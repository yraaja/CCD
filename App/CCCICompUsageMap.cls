VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCCICompUseMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Material grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "published_CCI_Equipment_Rate" ' Identity of grid
Const UPDATE_ADD = "Update/Add"
Const ALL_TABLES = "All Tables"
Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' recordset that populates grid
Private m_blnPreferences As Boolean ' Set to True if on Grid Preferences screen

Const EQUIPMENT_TYPE = "E"
Const LABOR_TYPE = "L"
Const MATERIAL_TYPE = "M"

Const MATERIAL_ERROR = 1
Const EQUIPMENT_ERROR = 2
Const LABOR_ERROR = 3

Dim m_iError As Integer
Public Function Delete() As Boolean
        Delete = m_objGridMap.Delete
End Function

Private Sub FillID(strSelID As String, Optional m_recBookmark As Variant = Empty)
Dim strSQL As String
Dim sType  As String
sType = TDBGrid.Columns("Type")

Select Case sType
    Case EQUIPMENT_TYPE
        strSQL = "SELECT  cci_equip_desc description, cci_equip_skey cci_skey FROM cci_equipment Where cci_equip_id = '" & m_rec.Fields("cci_id") & "'"
    Case LABOR_TYPE
        strSQL = "SELECT  cci_trade_desc description, trade_skey cci_skey FROM cci_labor Where trade_id = '" & m_rec.Fields("cci_id") & "'"
    Case MATERIAL_TYPE
        strSQL = "SELECT  cci_mat_desc description, usage_unit, cci_mat_skey cci_skey FROM cci_material Where cci_mat_id = '" & strSelID & "'"
End Select

FillFields (strSQL)

End Sub
Private Sub FillFields(strSQL As String)
Dim rec As New ADODB.RecordSet ' Recordset to hold query results
Dim blnReturn As Boolean
Dim i As Integer
Dim j As Integer
Dim varCurBookmark As Variant

On Error Resume Next
varCurBookmark = m_rec.Bookmark
m_rec.Bookmark = TDBGrid.Bookmark
   
'On Error GoTo Err
    
     ' Use DAL to perform select
     blnReturn = g_objDAL.GetRecordset(vbNullString, strSQL, rec)
     
     
    For i = 0 To rec.Fields.Count - 1
   
   On Error Resume Next
  
        m_rec.Bookmark = TDBGrid.Bookmark
        m_rec.Fields(rec.Fields(i).Name).Value = rec.Fields(i).Value
    
        For j = 0 To TDBGrid.Columns.Count - 1
        
            If TDBGrid.Columns(j).DataField = rec.Fields(i).Name Then
            
                TDBGrid.Columns(j).Value = CStr(rec.Fields(i).Value)
                
                Exit For
            End If
'          End If
        Next j
        
    Next i
   
    rec.Close
   
    Set rec = Nothing
  
     m_rec.Bookmark = varCurBookmark
 
Err:
Debug.Print Err.Description
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

Public Sub SetRowState(intBookmark As Integer, intState As Integer)
    m_objGridMap.SetRowState intBookmark, intState
End Sub
Public Sub InitGrid()
    TDBGrid.AllowAddNew = False
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
'    m_objGridMap.SetFrozenCols (2)

    If DEBUGON Then Stop
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "Sys", "class_system_id", ALL_TABLES, 300, True
    m_objGridMap.AddCol "Cls ID", "class_id", ALL_TABLES, 500, , , , , True
    m_objGridMap.AddCol "Type", "rec_type", ALL_TABLES, 390, , , , , True
    m_objGridMap.AddCol "CCI ID", "cci_id", vbNullString, 1000, , , , , True
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Quarter", "quarter_id", vbNullString, 625, True
    m_objGridMap.AddCol "State", "state_code", vbNullString, 300, True
    m_objGridMap.AddCol "City", "city", vbNullString, 1200, True
    m_objGridMap.AddCol "Description", "description", vbNullString, 3300, True
    m_objGridMap.AddCol "Unit", "usage_unit", vbNullString, 800, True
    m_objGridMap.AddCol "Qty", "usage_qty", UPDATE_ADD, 1000, , , , FORMAT_DECIMAL
    m_objGridMap.AddCol "Base Cost", "base_cost", vbNullString, 1300, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Mat'l Ext Cost", "mat_extended_cost", vbNullString, 1400, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Labor Ext Cost", "labor_extended_cost", vbNullString, 1400, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Equip Ext Cost", "equip_extended_cost", vbNullString, 1400, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Total Cost", "tot_ext_cost", vbNullString, 1400, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "last Update Date", "last_update_date", vbNullString, 1000, True, , , FORMAT_DATE
    m_objGridMap.AddCol "Last Update Person", "last_update_person", vbNullString, 1000, True
    m_objGridMap.AddCol "Last Update ID", "last_update_id", vbNullString, 1000, True, False
    m_objGridMap.AddCol "cci_skey", "cci_skey", ALL_TABLES, 10, True, False
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
    
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid70.ValueItem
    Dim Item As New TrueOleDBGrid80.ValueItem
    ' Type Code drop down
    Item.Value = "M"
    TDBGrid.Columns("Type").ValueItems.Add Item
    Item.Value = "E"
    TDBGrid.Columns("Type").ValueItems.Add Item
    Item.Value = "L"
    TDBGrid.Columns("Type").ValueItems.Add Item

    TDBGrid.Columns("Type").ValueItems.Presentation = dbgComboBox
    TDBGrid.Columns("Type").ValueItems.Validate = True
    TDBGrid.Columns("Type").AutoDropDown = True
    
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim bRet As Boolean
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    
    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0
    
    m_rec.MoveFirst
    While Not m_rec.EOF
        strUpdate = ""
        ' Skip the record if it didn't change
        Dim clsRowInfo As CRowInfo
        Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
        If Not clsRowInfo Is Nothing Then
            If clsRowInfo.State = STATE_MODIFIED Then
                strUpdate = "exec sp_update_cci_comp_usage "
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, UPDATE_ADD
                ' Make sure there is something to do
                If Not strUpdate = "" Then
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    strUpdate = strUpdate + ", @last_update_id=" + CStr(m_rec.Fields("last_update_id").Value)
                    bRet = False
                    strError = ""
                    bRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If bRet = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                    Else
                        intErrors = intErrors + 1
                        clsRowInfo.Error = True
                        clsRowInfo.ErrorText = strError
                    End If
                End If
            ElseIf clsRowInfo.State = STATE_NEW Then
                strUpdate = "exec sp_insert_cci_comp_usage "
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, UPDATE_ADD
                ' Make sure there is something to do
                If Not strUpdate = "" Then
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    bRet = False
                    strError = ""
                    bRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If bRet = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
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
    MsgBox "Error in cmdUpdate_Click"
End Function

Private Sub Class_Initialize()
    m_objGridMap.m_colCloneCols.Add 1, "state_code"
    m_objGridMap.m_colCloneCols.Add 1, "city"
    m_objGridMap.m_colCloneCols.Add 1, "quarter_id"
    m_objGridMap.m_colCloneCols.Add 1, "cci_equip_desc"
    m_objGridMap.m_colCloneCols.Add 1, "rent_per_week"
    m_objGridMap.m_colCloneCols.Add 1, "estimated_ind"
    m_objGridMap.m_colCloneCols.Add 1, "loc_id"
    m_objGridMap.m_colCloneCols.Add 1, "qtr_dt_skey"
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "CCI ID"
        If TDBGrid.AddNewMode = 2 Then
            TDBGrid.PostMsg 1
        Else
            FillID TDBGrid.Text
        End If
    End Select
End Sub

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim strSELECT As String
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnRet As Boolean
    Dim sSQL As String
    Dim sType As String
    Dim i As Integer
    On Error Resume Next
    sType = TDBGrid.Columns("Type").Text
    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "CCI ID"
        TDBGrid.Text = UCase(TDBGrid.Text)
        m_rec.Bookmark = TDBGrid.Bookmark
        Select Case sType
        Case "M"    'Material
            sSQL = "select count(cci_mat_id) from cci_material where cci_mat_id = '" + TDBGrid.Text + "'"
        Case "E"    'Equipment
            sSQL = "select count(cci_equip_id) from cci_equipment where cci_equip_id = '" + TDBGrid.Text + "'"
        Case "L"    'Labor
            sSQL = "select count(trade_id) from labor_trade where trade_id = '" + TDBGrid.Text + "'"
        End Select
        g_objDAL.GetRecordset vbNullString, sSQL, rec
        If rec.Fields(0).Value = 0 Then
            MsgBox "You must enter a valid CCI ID."
            Cancel = True
        End If
    End Select

End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
Dim sMsg As String
If m_iError > 0 Then
    Select Case m_iError
    Case MATERIAL_ERROR
        sMsg = "Please enter a valid Material Id."
    Case EQUIPMENT_ERROR
        sMsg = "Please enter a valid Equipment Id."
    Case LABOR_ERROR
        sMsg = "Please enter a valid Labor Id."
    End Select
    
    MsgBox sMsg
    m_iError = 0
    Response = 0
    TDBGrid.SetFocus
End If

End Sub


Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)
'Dim blnPositionToUC As Boolean
'Dim blnPositionToMat As Boolean
'Dim blnUseDftMat As Boolean
'Dim blnSetDft As Boolean
'Dim strSelMatID As String
'Dim strSelUCID As String
'Dim strSaveBookMark As String
'
'Dim i As Integer
'Dim J As Integer
'
'On Error Resume Next
'
'If MsgId = 1 Then       'Material changed
'    strSelMatID = TDBGrid.Columns("Material ID").Text
'    If m_blnUseUCDft = True Then
'        strSelUCID = m_strUnitCostID
'        blnPositionToUC = False         'Fill in UC, position to qty
'    Else
'        blnPositionToUC = True          'UC not filled in, position to it
'    End If
'    blnSetDft = True
'End If
'
'If MsgId = 2 Then       'Unit Cost ID changed
'    strSelUCID = TDBGrid.Columns("Unit Cost ID").Text
'    If m_blnUseMatDft = True Then
'        strSelMatID = "M" + m_strMaterialID
'        blnPositionToMat = False
'    Else
'        blnPositionToMat = True
'    End If
'    blnSetDft = True
'End If
'
'If blnSetDft = True Then
'    m_rec.AddNew
'    FillMaterial strSelMatID
'    FillUnitCost strSelUCID
'
'    m_rec.Fields("input_factor").Value = 1
'    m_rec.Fields("output_factor").Value = 1
'    m_rec.Fields("adj_factor").Value = 1
'    m_rec.Fields("unit_qty").Value = 1
'    m_rec.Fields("last_update_id").Value = 0
'
'    m_rec.Update
'    TDBGrid.ReBind
'    TDBGrid.SetFocus
'    TDBGrid.MoveLast
'    m_rec.MoveLast
'    Set TDBGrid.EditActive = True
'
'    If blnPositionToMat = True Then
'        TDBGrid.Split = 0
'        TDBGrid.Col = TDBGrid.Columns("Material ID").ColIndex
'    ElseIf blnPositionToUC = True Then
'        TDBGrid.Split = 0
'        TDBGrid.Col = TDBGrid.Columns("Unit Cost ID").ColIndex
'    Else
'        TDBGrid.Split = 1
'        TDBGrid.Col = TDBGrid.Columns("Unit Qty").ColIndex
'    End If
'    m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW
'
'End If

End Sub


' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
    Dim strSELECT As String
    Dim rec As New ADODB.RecordSet
    Dim strcontact_id As String
    Dim strmat_skey As Integer
    Dim strmanufacturer_id As String
    Dim strStart_date As String
    
  
    ' Position to the desired row and execute delete
    m_rec.Bookmark = Bookmark
    
    Set clsRowInfo = m_objGridMap.rowinfo(CStr(m_rec.Bookmark))
    If Not clsRowInfo Is Nothing Then
        If clsRowInfo.State = STATE_NEW Then
            TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
            m_rec.Delete ' Remove it from recordset so it doesn't show.
            Exit Sub
        End If
    End If
    clsRowInfo = Nothing
    
    strUpdate = "exec sp_delete_cci_comp_usage "
    
    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
    strUpdate = Mid(Trim(strUpdate), 1, Len(Trim(strUpdate)) - 1)   'Remove trailing comma
    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_NONE
        m_objGridMap.SetError m_rec.Bookmark, strError
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
    End If
    Exit Sub
End Sub

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


