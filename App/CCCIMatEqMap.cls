VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCCIMatEqMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Material grid
Option Explicit

' <modulename> CCCIMatEqMap.cls</modulename>
' <functionname>General (Main) </functionname>
'
' <summary>
'Helper class for the window/form, frmCCIMatEquRptGrid:
'
'Key subs / functions:
'"   InitGrid()
'       Calls CgridMap "AddCols" to build datagrid columns based upon resultset/recordset returned by stored procedure executed in cmdSearch_Click()
'"   Update()
'       N/A
'       No updating enabled.  Read Only!!!'
'</summary>
'
' <seealso> frmCCIMatEquRptGrid </seealso>
'<seealso> </seealso>
'
' <datastruct>m_rec</datastruct>
'<datastruct>m_objGridMap</datastruct>
'
' <storedprocedurename> N/A
'</storedprocedurename>
'
' <returns>N/A</returns>
' <exception> Always trap with an accompanying message box</exception>
' <example>
' <code>  * * * MAKE A CHANGE TO A VALUE ON THE GRID ROW * * *
'N/A
'</code>
'
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  frmCCIMatEquRptGrid
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>



'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "CCI Quarterly Estimated Report" ' Identity of grid
Const PUB_CCI_INDEX = "published_cci_index"
Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' recordset that populates grid
Private m_blnPreferences As Boolean ' Set to True if on Grid Preferences screen

Public Function Delete() As Boolean
        Delete = m_objGridMap.Delete
End Function

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

Public Sub InitGrid()
    TDBGrid.AllowAddNew = False
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
'    m_objGridMap.SetFrozenCols (2)
    
    If DEBUGON Then Stop
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "ID", "cci_mat_equ_id", vbNullString, 775, True
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Country", "country_code", vbNullString, 390, True
    m_objGridMap.AddCol "State", "state_code", vbNullString, 390, True
    m_objGridMap.AddCol "Zip Code", "zip_3", vbNullString, 850, True
    m_objGridMap.AddCol "City", "city", vbNullString, 1850, True
    m_objGridMap.AddCol "Type", "type_code", vbNullString, 625, True
    
    m_objGridMap.AddCol "Q1", "Q1", vbNullString, 1500, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Q1 Est", "Q1_est", vbNullString, 625, True
    m_objGridMap.AddCol "Q1 %", "Q1_pct", vbNullString, 1000, True, , , FORMAT_DECIMAL3
    m_objGridMap.AddCol "Q2", "Q2", vbNullString, 1500, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Q2 Est", "Q2_est", vbNullString, 625, True
    m_objGridMap.AddCol "Q2 %", "Q2_pct", vbNullString, 1000, True, , , FORMAT_DECIMAL3
    m_objGridMap.AddCol "Q3", "Q3", vbNullString, 1500, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Q3 Est", "Q3_est", vbNullString, 625, True
    m_objGridMap.AddCol "Q3 %", "Q3_pct", vbNullString, 1000, True, , , FORMAT_DECIMAL3
    m_objGridMap.AddCol "Q4", "Q4", vbNullString, 1500, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Q4 Est", "Q4_est", vbNullString, 625, True
    m_objGridMap.AddCol "Q4 %", "Q4_pct", vbNullString, 1000, True, , , FORMAT_DECIMAL3
    m_objGridMap.AddCol "YTD Pct", "YTD_pct", vbNullString, 800, True, , , FORMAT_DECIMAL3
    m_objGridMap.AddCol "30-City", "cci_nat_avg_ind", vbNullString, 625, True
    m_objGridMap.AddCol "316-City", "cci_primary_city_ind", vbNullString, 625, True
    m_objGridMap.AddCol "719-City", "cci_ind", vbNullString, 625, True
    m_objGridMap.AddCol "loc_factor_ind", "loc_factor_ind", vbNullString, 800, True
    
    m_objGridMap.AddCol "cci_mat_equ_skey", "cci_mat_equ_skey", vbNullString, 3300, True, False
    m_objGridMap.AddCol "loc_id", "loc_id", vbNullString, 800, True, False
    m_objGridMap.AddCol "qtr_dt_skey", "qtr_dt_skey", vbNullString, 800, True, False
    m_objGridMap.AddCol "last Update Date", "last_update_date", vbNullString, 1000, , , , FORMAT_DATE
    m_objGridMap.AddCol "Last Update Person", "last_update_person", vbNullString, 1000, True
    m_objGridMap.AddCol "Last Update ID", "last_update_id", vbNullString, 1000, True, False
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid70.ValueItem
    Dim Item As New TrueOleDBGrid80.ValueItem
    Item.Value = 0
    Item.DisplayValue = " "
    TDBGrid.Columns("Q1_est").ValueItems.Add Item
    TDBGrid.Columns("Q2_est").ValueItems.Add Item
    TDBGrid.Columns("Q3_est").ValueItems.Add Item
    TDBGrid.Columns("Q4_est").ValueItems.Add Item
    Item.Value = -1
    Item.DisplayValue = "E"
    TDBGrid.Columns("Q1_est").ValueItems.Add Item
    TDBGrid.Columns("Q2_est").ValueItems.Add Item
    TDBGrid.Columns("Q3_est").ValueItems.Add Item
    TDBGrid.Columns("Q4_est").ValueItems.Add Item
    
    TDBGrid.Columns("Q1_est").ValueItems.Translate = True
    TDBGrid.Columns("Q2_est").ValueItems.Translate = True
    TDBGrid.Columns("Q3_est").ValueItems.Translate = True
    TDBGrid.Columns("Q4_est").ValueItems.Translate = True
               
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
                strUpdate = "exec sp_update_published_cci_material_price "
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUB_CCI_INDEX
                
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
                        m_rec.Fields("mat_last_update_date").Value = Now
                        m_rec.Fields("mat_last_update_person").Value = strUserName
                        Reset_Orig_Values m_rec
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


