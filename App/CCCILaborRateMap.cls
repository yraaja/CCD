VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCCILabRtMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Material grid
Option Explicit


''' <modulename> CCCILabRtMap.cls</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
'''Helper class for the window/form, frmCCILaborRateGrid,in support of:
'''
'''Key subs / functions:
'''"   InitGrid()
'''"   TDBGrid_UnboundDeleteRow(Bookmark As Variant)
'''    ' Fired when a row is deleted
'''o   Keeps recordset in sync w/datagrid
'''o   Sets "str" values for
'''(eg. strContact_id, strStart_Date)
'''o   Executes sp, sp_delete_material, that actually removes the material from the database
'''
'''</summary>
'''
''' <seealso>frmCCILaborRateGrid</seealso>
'''<seealso> </seealso>
'''
''' <datastruct>m_rec</datastruct>
'''<datastruct>m_objGridMap</datastruct>
'''
''' <storedprocedurename> sp_select_cci_labor_rate_static_ksr
'''</storedprocedurename>
'''<storedprocedurename> sp_update_published_cci_material_price
'''</storedprocedurename>
'''
'''
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''exec dbo.sp_select_cci_labor_rate_static_ksr   @trade_id = '', @quarter_id = '2006Q4', @zip = '', @loc_id = 0, @state_code = '', @select_type = 2, @country_code = '%' </code>
'''
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  frmCCILaborRateGrid.cls
'''2.  CGridMap.cls
'''3.  CCDdal.CRSMDataAccess (
'''4.  Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''</dependson>


'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "published_CCI_Labor_Rate" ' Identity of grid
Const PUB_CCI_MAT_PRICE = "published_cci_labor_rate"
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
    m_objGridMap.AddCol "CCI Trade ID", "trade_id", vbNullString, 1500, True, , , FORMAT_MATERIAL
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "State", "state_code", vbNullString, 400, True
    m_objGridMap.AddCol "City", "city", vbNullString, 1850, True
    m_objGridMap.AddCol "Zip", "zip_code", vbNullString, 1850, True, False
    m_objGridMap.AddCol "Quarter", "quarter_id", vbNullString, 625, True
    m_objGridMap.AddCol "Description", "cci_trade_desc", vbNullString, 3300, True
    
    m_objGridMap.AddCol "Prev Base", "prev_base_hrly", vbNullString, 775, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Prev Fringe", "prev_fringe_hrly", vbNullString, 1000, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Prev Hourly", "prev_rate_hrly", vbNullString, 615, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Prev Worker Comp", "prev_worker_comp_rate", vbNullString, 615, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Prev WC Hourly", "prev_rate_wc_hrly", vbNullString, 615, True, , , FORMAT_PRICE
    
    m_objGridMap.AddCol "Base", "base_hrly", vbNullString, 775, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Fringe", "fringe_hrly", vbNullString, 1000, True, False, , FORMAT_PRICE
    m_objGridMap.AddCol "Hourly", "rate_hrly", vbNullString, 615, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "QTR-1", "rate_prev_1", vbNullString, 615, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "QTR-2", "rate_prev_2", vbNullString, 615, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "QTR-3", "rate_prev_3", vbNullString, 615, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "QTR-4", "rate_prev_4", vbNullString, 615, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "Worker Comp", "worker_comp_rate", vbNullString, 615, True, , , FORMAT_PRICE
    m_objGridMap.AddCol "WC Hourly", "rate_wc_hrly", vbNullString, 615, True, , , FORMAT_PRICE
'    m_objGridMap.AddCol "Est Ind", "estimated_ind", vbNullString, 615, True
'    m_objGridMap.AddCol "Contact", "contact_id", vbNullString, 800, True
'    m_objGridMap.AddCol "Last Update Date", "last_update_date", vbNullString, 1000, True, , , FORMAT_DATE
'    m_objGridMap.AddCol "Last Update Person", "last_update_person", vbNullString, 1000, True
    m_objGridMap.AddCol "Last Update ID", "last_update_id", vbNullString, 1000, True, False
    m_objGridMap.AddCol "trade_skey", "trade_skey", vbNullString, 500, True, False
    m_objGridMap.AddCol "loc_id", "loc_id", vbNullString, 500, True, False
    m_objGridMap.AddCol "qtr_dt_skey", "qtr_dt_skey", vbNullString, 500, True, False
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
               

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
                m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUB_CCI_MAT_PRICE
                
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

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim strSelect As String
    Dim rec As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnRet As Boolean
    Dim i As Integer
    On Error Resume Next

    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Contact"
            TDBGrid.Text = UCase(TDBGrid.Text)
            g_objDAL.GetRecordset vbNullString, "select count(contact_id) from information_source where contact_id = '" + TDBGrid.Text + "'", rec
            If rec.Fields(0).Value = 0 Then
                MsgBox "You must enter a valid Contact ID."
                Cancel = True
            End If
    End Select

End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
    Dim strSelect As String
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
    strmat_skey = m_rec.Fields("mat_skey").Value
    strSelect = "select distinct contact_id from material_price inner join material on material_price_skey = material_skey where material_price.mat_skey = '" + m_rec.Fields("mat_skey").Value + "'"
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
    strcontact_id = strSelect

    
    strStart_date = "select start_date from material inner join material on material_price.mat_skey=material.mat_skey where material_price.mat_skey = '" + m_rec.Fields("mat_skey").Value + "'"
    blnReturn = g_objDAL.GetRecordset(CONNECT, strStart_date, rec)
    strStart_date = strStart_date

    strUpdate = "exec sp_delete_material "
    strUpdate = strUpdate + "@mat_skey='" + str(m_rec.Fields("mat_skey").Value) + "', "
    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
    
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

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
'    TDBGrid.Columns("Active").Value = True
    TDBGrid.Columns("mat_last_update_id").Value = 0
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


