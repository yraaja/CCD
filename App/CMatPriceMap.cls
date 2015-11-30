VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMatPriceMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Material Price grid
Option Explicit

''' <modulename> CMatPriceMap.cls</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' Helper class module for frmMatPriceGrid.frm
''' </summary>
''' <seealso> frmMatPriceGrid.frm </seealso>
''' <datastruct> TDBGRID</datastruct>
''' <datastruct> m_objGridMap</datastruct>
'''
''' <storedprocedurename> sp_publish_mat_price</storedprocedurename>
'''<storedprocedurename> sp_insert_material_and_price</storedprocedurename>
''' <storedprocedurename> sp_update_material_and_price </storedprocedurename>
'''<storedprocedurename> sp_update_material_price </storedprocedurename>
'''<storedprocedurename> sp_update_material </storedprocedurename>
'''<storedprocedurename> sp_select_mp_count</storedprocedurename>
'''<storedprocedurename> sp_delete_material_price</storedprocedurename>
'''<storedprocedurename> sp_insert_material_price</storedprocedurename>
'''<storedprocedurename> sp_insert_material_and_price</storedprocedurename>
'''<storedprocedurename> sp_update_material_and_price</storedprocedurename>
'''<storedprocedurename> sp_update_material_price</storedprocedurename>
'''<storedprocedurename> sp_update_material</storedprocedurename>
'''
'''<param name="data"> ???a dataset containing all the data for updating ?
'''</param>
''' <param name="someParameter">
'''???? Description of someParameter goes here  updating
''' </param>
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
'''<code>
'''exec sp_publish_mat_price  @mat_skey = 117050, @country_code = 'USA', @region_code = 'NAT', @last_update_person='Hancockrl'
''' </code>
'''<code>
'''exec sp_insert_material_and_price  @alt_mat_id='M0331680199', @usage_unit='S.F.', @purchase_unit='Bag', @purchase_usage_conv_factor=1., @metric_tech_desc='Cementitious repair mortar, 23 kg bag yields 0.012 m3', @tech_desc='Cementitious repair mortar, 50 Lb bag yields 0.42 CF', @active_status_ind= 1, @mat_id='M039206000199', @traces_ind= 0, @use_ind= 1, @wst_use_ind= 1, @estimated_ind= 0, @update_status_code='8', @list_price=20.45, @traces_list_price=20.45, @pct_multiplier=100., @contact_id='030PHI', @commodity_code='', @manufacturer_id='WEB030', @catalog_num='SpeedPatch', @page_num='', @model_name='', @item_num='0048000', @latest_price_update_comment='', @comment='', @last_update_person='Hancockrl'
'''</code>
'''<code>
'''exec sp_update_material_and_price @mat_skey=117050, @matprice_last_update_id=9, @mat_last_update_id=4,  @alt_mat_id='M0331680150', @usage_unit='S.F.', @purchase_unit='Unit', @purchase_usage_conv_factor=1., @metric_tech_desc='Concrete floor patch, 3-part epoxy mortar, .014 m3 unit', @tech_desc='Concrete floor patch, 3-part epoxy mortar, .50 CF unit rlh', @active_status_ind= 1, @mat_id='M039206000150', @traces_ind= 0, @use_ind= 1, @wst_use_ind= 1, @estimated_ind= 0, @update_status_code='7', @list_price=188., @traces_list_price=188., @pct_multiplier=100., @contact_id='030PHI', @commodity_code='', @manufacturer_id='WEB030', @catalog_num='43Patch-Pak', @page_num='', @model_name='', @item_num='0400300', @latest_price_update_comment='', @comment='', @old_contact_id='030PHI',  @old_manufacturer_id='WEB030',  @start_date='7/16/2010',  @factor_ind= 0,  @last_update_person='Hancockrl'
'''</code>
'''<code>
'''exec sp_update_material_price @mat_skey=117050, @matprice_last_update_id=7,  @traces_ind= 0, @use_ind= 1, @wst_use_ind= 1, @estimated_ind= 0, @update_status_code='7', @list_price=188., @traces_list_price=188., @pct_multiplier=100., @contact_id='030PHI', @commodity_code='', @manufacturer_id='WEB030', @catalog_num='43Patch-Pak', @page_num='', @model_name='', @item_num='0400300', @latest_price_update_comment='', @comment='', @old_contact_id='030PHI',  @old_manufacturer_id='WEB030',  @start_date='7/11/2006',  @factor_ind= 0,  @last_update_person='Hancockrl'
'''</code>
''' <code>
'''exec sp_delete_material_price @mat_skey= 117051, @contact_id='030PHI', @manufacturer_id='WEB030', @start_date='07/11/2006', @last_update_person='Hancockrl'
''' </code>
'''
'''<code>
'''exec sp_update_material  @alt_mat_id='M0331680150', @purchase_usage_conv_factor=1., @purchase_unit='Unit', @usage_unit='S.F.', @mat_skey=117050., @mat_id='M039206000150', @active_status_ind= 1, @tech_desc='Concrete floor patch, 3-part epoxy mortar, .50 CF unit rlhzzz', @metric_tech_desc='Concrete floor patch, 3-part epoxy mortar, .014 m3 unit',@last_update_person='Hancockrl', @last_update_id=6
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''    CGridMap.Cls
'''    CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''TDBGrid (ComponentOne DataGrid)
'''Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid  ' Access to Grid events
'''</dependson>


'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "MaterialPrice" ' Identity of grid
' Names of some table this grid uses
Const MAT_TABLE = "Material"
Const MATPRICE_TABLE = "Material_price"
Const MAT_UPDATE = "Material update"
Const MATMAN_TABLE = "Material_manufacturer"
Const MATPRICE_UPDATE = "Material_price update"
Const ALL_TABLES = "All"

Private m_UpdateStatusCode As Integer
Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean

Private colMatSkey As New Collection
Private Function CurrentUpdateStatusCode() As Integer
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    Dim strSelect As String
    On Error Resume Next
    If m_UpdateStatusCode = 0 Then
        rec.Close
        strSelect = "SELECT update_status_code = domain_value from DOMAIN_TBL where domain_name = 'PAPER_CLIP'"
        g_objDAL.GetRecordset CONNECT, strSelect, rec
        If rec.EOF Then
            m_UpdateStatusCode = 0
        Else
            m_UpdateStatusCode = rec.Fields("update_status_code")
        End If
        rec.Close
        Set rec = Nothing
    End If
    CurrentUpdateStatusCode = m_UpdateStatusCode

End Function

Public Function Delete() As Boolean
        Delete = m_objGridMap.Delete
End Function
Private Sub RollupUCPrice(colMatSkey As Collection)
Dim varMatSkey As Variant
Dim rec As ADODB.RecordSet
Dim strSelect As String
Dim colUnitCost As New Collection
Dim blnReturn As Boolean

On Error Resume Next
'Rollup the unit cost
'The material price update triggers an update to published material price

For Each varMatSkey In colMatSkey
    strSelect = "select distinct unit_cost_id from material_usage inner join " + _
        "unit_cost_detail on material_usage.unit_cost_skey = unit_cost_detail.unit_cost_skey where mat_skey = " + CStr(varMatSkey)
    rec.Close
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
    If blnReturn = True Then
        If rec.RecordCount > 0 Then
            Do Until rec.EOF
                SaveUnitCostID colUnitCost, rec.Fields("unit_cost_id")
                rec.MoveNext
            Loop
        End If
    End If
Next varMatSkey

If colUnitCost.Count > 0 Then
    UpdateUnitCost colUnitCost
End If

End Sub

Private Sub RollupAssemblyPrice(colUnitCostSkey As Collection)
Dim varUCSkey As Variant
Dim rec As ADODB.RecordSet
Dim strSelect As String
Dim colAssemblyID As New Collection
Dim blnReturn As Boolean

On Error Resume Next
'Rollup the unit cost
'The material price update triggers an update to published material price

For Each varUCSkey In colUnitCostSkey
    strSelect = "select distinct assembly_id from unit_cost_usage inner join " + _
        "assembly_detail on unit_cost_usage.parent_skey = assembly_detail.assembly_skey where " + _
        "skey_type = 'A' and unit_cost_skey = " + CStr(varUCSkey)
    rec.Close
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
    If blnReturn = True Then
        If rec.RecordCount > 0 Then
            Do Until rec.EOF
                SaveAssemblyID colAssemblyID, rec.Fields("assembly_id")
            Loop
        End If
    End If
Next varUCSkey

If colUnitCostSkey.Count > 0 Then
    UpdateAssembly colAssemblyID
End If

End Sub
Public Sub Publish()
    Dim vntBookmark As Variant
    Dim vntSaveBookmark As Variant

    Dim strUpdate As String
    Dim strError As String
    Dim blnReturn As Boolean
    Dim intErrors As Integer
    Dim intSuccess As Integer
    Dim i As Integer
    
    On Error Resume Next
    
    
    If TDBGrid.SelBookmarks.Count = 0 Then
        MsgBox "Please select rows to publish."
    Else
        vntSaveBookmark = TDBGrid.SelBookmarks(0)
        

        For Each vntBookmark In TDBGrid.SelBookmarks
            m_rec.Bookmark = vntBookmark
            m_objGridMap.SetRowState Int(m_rec.Bookmark), STATE_PUBLISHED
            Dim clsRowInfo As CRowInfo
            Set clsRowInfo = m_objGridMap.rowinfo(CInt(m_rec.Bookmark))
            clsRowInfo.State = STATE_PUBLISHED
           If m_rec.Fields("use_ind") = True Then
                strUpdate = "exec sp_publish_mat_price" + _
                    "  @mat_skey = " + CStr(m_rec.Fields("mat_skey")) + _
                    ", @country_code = 'USA'" + _
                    ", @region_code = 'NAT'" + _
                    ", @last_update_person='" + strUserName + "' "
                
                blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnReturn = True Then
                    intSuccess = intSuccess + 1
                Else
                    intErrors = intErrors + 1
                    clsRowInfo.Error = True
                    clsRowInfo.ErrorText = strError
                End If
            End If
        Next
        For i = 0 To TDBGrid.SelBookmarks.Count - 1
            TDBGrid.SelBookmarks.Remove 0
        Next i

        TDBGrid.Bookmark = vntSaveBookmark ' Set bookmark back again
        Dim strMsg As String
        strMsg = ""
        If intSuccess > 0 Then
            strMsg = str(intSuccess) + " material prices published successfully." + Chr(10)
        End If
        If intErrors > 0 Then
            strMsg = strMsg + str(intErrors) + " errors occurred."
        End If
        If Len(strMsg) > 0 Then
            MsgBox strMsg
        End If
        TDBGrid.ReBind
    End If
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
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "Material ID", "mat_id", MAT_TABLE, 1500, True, , , FORMAT_MATERIAL
    m_objGridMap.AddCol "Alt Mat ID", "alt_mat_id", MAT_TABLE, 1500, False, False, True
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Tech Desc", "tech_desc", MAT_TABLE, 2000
    m_objGridMap.AddCol "Metric Tech Desc", "metric_tech_desc", MAT_TABLE, 2000, , False
    m_objGridMap.AddCol "Purch Unit", "purchase_unit", MAT_TABLE, 1000, False, True, True
    m_objGridMap.AddCol "Price", "list_price", MATPRICE_TABLE, 1000, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "TRACES Price", "traces_list_price", MATPRICE_TABLE, 1200, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "Contact", "contact_id", MATPRICE_TABLE, 1200
    m_objGridMap.AddCol "Manufacturer", "manufacturer_id", MATPRICE_TABLE, 1200
    m_objGridMap.AddCol "Manufac. Desc", "manufacturer_desc", MATMAN_TABLE, 2000, True
    m_objGridMap.AddCol "Use", "use_ind", MATPRICE_TABLE, 1000
    m_objGridMap.AddCol "West Use", "wst_use_ind", MATPRICE_TABLE, 1000
    m_objGridMap.AddCol "Estimated", "estimated_ind", MATPRICE_TABLE, 1000
    m_objGridMap.AddCol "TRACES", "traces_ind", MATPRICE_TABLE, 1000
    m_objGridMap.AddCol "Active", "active_status_ind", MAT_TABLE, 1000
    m_objGridMap.AddCol "Material Price", "final_price", vbNullString, 1200, True, True, False, FORMAT_PRICE
    m_objGridMap.AddCol "Start Date", "start_date", MATPRICE_UPDATE, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Term Date", "term_date", vbNullString, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Conv", "purchase_usage_conv_factor", MAT_TABLE, 800, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Pct Mult", "pct_multiplier", MATPRICE_TABLE, 1000, False, True, True, FORMAT_DECIMAL
    m_objGridMap.AddCol "Factor", "factor_ind", MATPRICE_UPDATE, 1000, True, True, False
    m_objGridMap.AddCol "Use Unit", "usage_unit", MAT_TABLE, 1000, False, True, True
    m_objGridMap.AddCol "Cat Num", "catalog_num", MATPRICE_TABLE, 1500
    m_objGridMap.AddCol "Item Num", "item_num", MATPRICE_TABLE, 1500
    m_objGridMap.AddCol "Page Num", "page_num", MATPRICE_TABLE, 1500
    m_objGridMap.AddCol "Model", "model_name", MATPRICE_TABLE, 1500
    m_objGridMap.AddCol "Comm Code", "commodity_code", MATPRICE_TABLE, 1500
    m_objGridMap.AddCol "Update", "update_ind", vbNullString, 1000, True, True, False
    m_objGridMap.AddCol "Updt Stat", "update_status_code", MATPRICE_TABLE, 1000, False, True, False
    m_objGridMap.AddCol "Latest Price Comment", "latest_price_update_comment", MATPRICE_TABLE, 2000
    m_objGridMap.AddCol "Comment", "comment", MATPRICE_TABLE, 2000
    m_objGridMap.AddCol "Update Date", "matprice_last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "matprice_last_update_person", vbNullString, 1500, True, True, False
    m_objGridMap.AddCol "mat_skey", "mat_skey", ALL_TABLES, 800, True, False, True
    m_objGridMap.AddCol "matprice_last_update_id", "matprice_last_update_id", MATPRICE_UPDATE, 800, True, False, False
    m_objGridMap.AddCol "mat_last_update_id", "mat_last_update_id", vbNullString, 800, True, False, True
   
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
        
    If Not m_blnPreferences Then
        Dim bRet As Boolean
        Dim rec As ADODB.RecordSet
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid60.ValueItem
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid70.ValueItem
        Dim Item As New TrueOleDBGrid80.ValueItem
        
        bRet = g_objDAL.GetRecordset(vbNullString, "select unit from unit_of_measure", rec)
        If Not bRet Then
            Exit Sub
        End If
        
        While Not rec.EOF
            Item.Value = rec.Fields(0)
            TDBGrid.Columns("Purch Unit").ValueItems.Add Item
            TDBGrid.Columns("Use Unit").ValueItems.Add Item
            rec.MoveNext
        Wend
        TDBGrid.Columns("Purch Unit").ValueItems.Presentation = dbgComboBox
        TDBGrid.Columns("Purch Unit").ValueItems.Validate = True
        TDBGrid.Columns("Use Unit").ValueItems.Presentation = dbgComboBox
        TDBGrid.Columns("Use Unit").ValueItems.Validate = True
    End If
    m_UpdateStatusCode = 0  'Initialize domain table value - updated if referenced
End Sub
Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnUpdateMat As Boolean
    Dim blnUpdateMatPrice As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    Dim strSelect As String
    
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
            ' If record was added
                If clsRowInfo.State = STATE_NEW Then
                    ' If mat_skey doesn't exist, then insert material and price
                    If m_rec.Fields("mat_skey") = 0 Or m_rec.Fields("mat_skey") = "" Then
                        strUpdate = "exec sp_insert_material_and_price "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MAT_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MATPRICE_TABLE
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"

                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                             
                            m_rec.Fields("matprice_last_update_id").Value = m_rec.Fields("matprice_last_update_id").Value + 1
                            m_rec.Fields("mat_last_update_id").Value = m_rec.Fields("mat_last_update_id").Value + 1
                            m_rec.Fields("mat_last_update_date").Value = Now
                            m_rec.Fields("mat_last_update_person").Value = strUserName
                            m_rec.Fields("matprice_last_update_date").Value = Now
                            m_rec.Fields("matprice_last_update_person").Value = strUserName
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' mat_skey does exist, so insert material price only
                    Else
                        strUpdate = "exec sp_insert_material_price "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MATPRICE_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("matprice_last_update_id").Value = m_rec.Fields("matprice_last_update_id").Value + 1
                            m_rec.Fields("matprice_last_update_date").Value = Now
                            m_rec.Fields("matprice_last_update_person").Value = strUserName
                            SaveMatID colMatSkey, m_rec.Fields("mat_skey")
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    End If
                ' If the record was changed
                ElseIf clsRowInfo.State = STATE_MODIFIED Then
                    blnUpdateMat = False
                    blnUpdateMatPrice = False
                    ' Loop through the fields
                    For Each fld In m_rec.Fields
                        ' If the value changed
                        If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
                            ' See what table the field is from
                            Set objColDef = m_objGridMap.GetCol(fld.Name)
                            ' Mark the table we should update
                            If objColDef.Table = MAT_TABLE Then
                                blnUpdateMat = True
                            ElseIf objColDef.Table = MATPRICE_TABLE Then
                                blnUpdateMatPrice = True
                            End If
                        End If
                    Next
                    blnReturn = False
                    strError = ""
                    If blnUpdateMatPrice And blnUpdateMat Then
                        strUpdate = "exec sp_update_material_and_price "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MATPRICE_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MAT_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MATPRICE_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MAT_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@old_contact_id='" + m_rec.Fields("contact_id").OriginalValue + "', "
                        strUpdate = strUpdate + "@Old_pct_multiplier='" + m_rec.Fields("pct_multiplier").OriginalValue + "', "
                        strUpdate = strUpdate + "@old_list_price='" + m_rec.Fields("list_price").OriginalValue + "', "
                        strUpdate = strUpdate + "@old_manufacturer_id='" + m_rec.Fields("manufacturer_id").OriginalValue + "', "
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            SaveMatID colMatSkey, m_rec.Fields("Mat_skey")
                            m_rec.Fields("matprice_last_update_id").Value = m_rec.Fields("matprice_last_update_id").Value + 1
                            m_rec.Fields("mat_last_update_id").Value = m_rec.Fields("mat_last_update_id").Value + 1
                            m_rec.Fields("mat_last_update_date").Value = Now
                            m_rec.Fields("mat_last_update_person").Value = strUserName
                            m_rec.Fields("matprice_last_update_date").Value = Now
                            m_rec.Fields("matprice_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ElseIf blnUpdateMatPrice Then
                        strUpdate = "exec sp_update_material_price "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MATPRICE_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MATPRICE_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@old_contact_id='" + m_rec.Fields("contact_id").OriginalValue + "', "
                        strUpdate = strUpdate + "@old_manufacturer_id='" + m_rec.Fields("manufacturer_id").OriginalValue + "', "
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            SaveMatID colMatSkey, m_rec.Fields("Mat_skey")
                            m_rec.Fields("matprice_last_update_id").Value = m_rec.Fields("matprice_last_update_id").Value + 1
                            m_rec.Fields("matprice_last_update_date").Value = Now
                            m_rec.Fields("matprice_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ElseIf blnUpdateMat Then
                        strUpdate = "exec sp_update_material "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MAT_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, MAT_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        strUpdate = strUpdate + ", @last_update_id=" + CStr(m_rec.Fields("mat_last_update_id"))
                        
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            SaveMatID colMatSkey, m_rec.Fields("Mat_skey")
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("mat_last_update_id").Value = m_rec.Fields("mat_last_update_id").Value + 1
                            m_rec.Fields("mat_last_update_date").Value = Now
                            m_rec.Fields("mat_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    End If
                    
                    ' If the row was marked as changed but the user typed the same data
                    ' make sure the error condition is removed from the row
'                    If Not blnUpdateMatPrice And Not blnUpdateMat Then
'                        intSuccess = intSuccess + 1
'                        clsRowInfo.State = STATE_NONE
'                    End If
                End If
            End If
        End If
        If blnReturn = True Then    'Retrieve the latest update_status_code
        'synchronize the update_status_code
            rec.Close
            strSelect = "select update_status_code from material_price " + _
                "where mat_skey = " & str(m_rec.Fields("mat_skey")) + _
                " and contact_id = '" & m_rec.Fields("contact_id") + _
                "' and manufacturer_id = '" & m_rec.Fields("manufacturer_id") + _
                "' and start_date = '" & str(m_rec.Fields("start_date")) + "'"
            g_objDAL.GetRecordset CONNECT, strSelect, rec
            If Not rec.EOF Then
                m_rec.Fields("update_status_code") = rec.Fields("update_status_code")
            End If
            rec.Close
            Set rec = Nothing
            blnReturn = False
        End If
        m_rec.MoveNext
    Wend
'Check the rollup flag in the domain table
    rec.Close
    Set rec = Nothing
    
    ' Only roll up if the flag in the db says to.
    Dim strDomainValue As String
    strDomainValue = GetDomainTableValue("PUB_MAT_OPTION")

    If strDomainValue = "1" Then
        RollupUCPrice colMatSkey
    End If
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
    
    TDBGrid.Refresh
End Function
Public Sub SaveMatID(colMatSkey As Collection, sKey As Long)
    'Add the specified ID to the collection if it does not exist
Dim blnfound As Boolean
Dim varMatSkey As Variant

    'Add the skey to the collection if it does not exist in it.
    blnfound = False
    'Add to the collection of unit_cost_skey entries
    For Each varMatSkey In colMatSkey
        If varMatSkey = sKey Then
            blnfound = True
            Exit For
        End If
    Next varMatSkey
    If blnfound = False Then
        colMatSkey.Add sKey
    End If

End Sub

Private Sub Class_Initialize()
    m_objGridMap.m_colCloneCols.Add 1, "mat_skey"
    m_objGridMap.m_colCloneCols.Add 1, "mat_id"
    m_objGridMap.m_colCloneCols.Add 1, "alt_mat_id"
    m_objGridMap.m_colCloneCols.Add 1, "tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "active_status_code"
    m_objGridMap.m_colCloneCols.Add 1, "update_status_code"
    m_objGridMap.m_colCloneCols.Add 1, "active_status_ind"
    m_objGridMap.m_colCloneCols.Add 1, "purchase_unit"
    m_objGridMap.m_colCloneCols.Add 1, "purchase_usage_conv_factor"
    m_objGridMap.m_colCloneCols.Add 1, "usage_unit"
    m_objGridMap.m_colCloneCols.Add 1, "contact_id"
    m_objGridMap.m_colCloneCols.Add 1, "manufacturer_id"
    m_objGridMap.m_colCloneCols.Add 1, "list_price"
    m_objGridMap.m_colCloneCols.Add 1, "traces_list_price"
    m_objGridMap.m_colCloneCols.Add 1, "pct_multiplier"
    m_objGridMap.m_colCloneCols.Add 1, "use_ind"
    m_objGridMap.m_colCloneCols.Add 1, "wst_use_ind"
    m_objGridMap.m_colCloneCols.Add 1, "estimated_ind"
    m_objGridMap.m_colCloneCols.Add 1, "catalog_num"
    m_objGridMap.m_colCloneCols.Add 1, "traces_ind"
    m_objGridMap.m_colCloneCols.Add 1, "model_name"
    m_objGridMap.m_colCloneCols.Add 1, "page_num"
    m_objGridMap.m_colCloneCols.Add 1, "item_num"
    m_objGridMap.m_colCloneCols.Add 1, "commodity_code"
    m_objGridMap.m_colCloneCols.Add 1, "comment"
    m_objGridMap.m_colCloneCols.Add 1, "latest_price_update_comment"
    m_objGridMap.m_colCloneCols.Add 1, "mat_last_update_id"
End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
    Dim rec As ADODB.RecordSet
    Dim strSelect As String
    Dim intNbrMatsUsed As Integer
    Dim intPriceCount As Integer
    ' Position to the desired record and execute delete
    m_rec.Bookmark = Bookmark
    strSelect = "exec sp_select_mp_count @mat_skey = " + CStr(m_rec.Fields("mat_skey"))
    g_objDAL.GetRecordset CONNECT, strSelect, rec
    If Not rec.EOF Then
        intPriceCount = rec.Fields("PriceCount")
    Else
        intPriceCount = 0
    End If
    rec.Close
    If intPriceCount = 1 Then    'Only 1 use
        strSelect = "select count(*) as NbrMatsUsed from material_usage as mu where mat_skey = " + CStr(m_rec.Fields("mat_skey"))
        g_objDAL.GetRecordset CONNECT, strSelect, rec
        If Not rec.EOF Then
            intNbrMatsUsed = rec.Fields("NbrMatsUsed")
        Else
            intNbrMatsUsed = 0
        End If
        rec.Close
        If intNbrMatsUsed > 0 Then
            m_objGridMap.SetError m_rec.Bookmark, str(intNbrMatsUsed) + " material usage record(s) exist.  The material may not be deleted."
            Exit Sub
        End If
    End If
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
    strUpdate = "exec sp_delete_material_price "
    strUpdate = strUpdate + "@mat_skey=" + str(m_rec.Fields("mat_skey")) + ","
    strUpdate = strUpdate + " @contact_id='" + m_rec.Fields("contact_id") + "',"
    strUpdate = strUpdate + " @manufacturer_id='" + m_rec.Fields("manufacturer_id") + "',"
    strUpdate = strUpdate + " @start_date='" + Format(m_rec.Fields("start_date"), "mm/dd/yyyy") + "',"
    strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    
    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetError m_rec.Bookmark, strError
        m_rec.Fields("Error").Value = "Error"
        Bookmark = Null
    Else
        SaveMatID colMatSkey, m_rec.Fields("Mat_ID")
        
        ' Only roll up if the flag in the db says to.
        Dim strDomainValue As String
        strDomainValue = GetDomainTableValue("PUB_MAT_OPTION")
        If (strDomainValue = "1") Then
            RollupUCPrice colMatSkey
        End If
        
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
    End If
    rec.Close
    Exit Sub

NoDelete:
    MsgBox Err.Description
    Bookmark = Null
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error Resume Next
    Dim rec As ADODB.RecordSet
    Dim strSelect As String

    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Active"   'Validate active flag - may not be set off if usage records exist
        If TDBGrid.Columns(ColIndex).Value = 0 Then
            rec.Close
            strSelect = "select count(*) as NbrMatsUsed from material_usage as mu where mat_skey = " + CStr(TDBGrid.Columns("mat_skey").Value)
            g_objDAL.GetRecordset CONNECT, strSelect, rec
            If Not rec.EOF Then
                If rec.Fields("NbrMatsUsed") > 0 Then
                    MsgBox str(rec.Fields("NbrMatsUsed")) + " material usage record(s) exist.  The material may not be inactivated."
                    Cancel = True
                    TDBGrid.Columns(ColIndex).RefetchCell
                End If
            End If
            rec.Close
        End If
    Case "Material ID"
        ' Can only edit new rows and only once
        If OldValue <> Empty Then
            MsgBox "You cannot edit this column for an existing row."
            Cancel = True
        End If
    Case "Price"
            If CDbl(TDBGrid.Text) < 0 Then
                MsgBox TDBGrid.Columns(ColIndex).Caption + " must be greater than or equal to zero."
                Cancel = True
            End If
            TDBGrid.Columns("factor_ind").Value = False ' Price was changed, so no longer factored
            TDBGrid.Columns("traces_list_price").Value = TDBGrid.Text ' Always make TRACES = list when list changes
        ' TRACES List Price
    Case "TRACES Price"
        If CDbl(TDBGrid.Text) < 0 Then
            MsgBox TDBGrid.Columns(ColIndex).Caption + " must be greater than or equal to zero."
            Cancel = True
        End If
        ' Conversion Factor
    Case "Conv"
        If IsNumeric(TDBGrid.Text) = False Then
            MsgBox TDBGrid.Columns(ColIndex).Caption + " must be a number."
            Cancel = True
        End If
        If CDbl(TDBGrid.Text) <= 0 Then
            MsgBox TDBGrid.Columns(ColIndex).Caption + " must be greater than zero."
            Cancel = True
        End If
    Case "Contact"
            TDBGrid.Text = UCase(TDBGrid.Text)
            g_objDAL.GetRecordset vbNullString, "select count(contact_id) from information_source where contact_id = '" + TDBGrid.Text + "'", rec
            If rec.Fields(0).Value = 0 Then
                MsgBox "You must enter a valid Contact ID."
                Cancel = True
            End If
    Case "Manufacturer"
        TDBGrid.Text = UCase(TDBGrid.Text)
        g_objDAL.GetRecordset vbNullString, "select manufacturer_desc from material_manufacturer where manufacturer_id = '" + TDBGrid.Text + "'", rec
        If rec.RecordCount = 0 Then
            MsgBox "You must enter a valid Manufacturer ID."
            Cancel = True
            rec.Close
        Else
            TDBGrid.Columns("manufacturer_desc").Text = rec.Fields(0).Value
            rec.Close
        End If
    Case Else
    ' list_price
        If m_rec.Fields(TDBGrid.Columns(ColIndex).DataField).Type = adCurrency Then
            If IsNumeric(TDBGrid.Text) = False Then
                MsgBox TDBGrid.Columns(ColIndex).Caption + " must be a dollar value."
                Cancel = True
            Else
                TDBGrid.Text = Format(TDBGrid.Text, "#,###,##0.00")
            End If
        End If
    End Select
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("Active").Value = True
    TDBGrid.Columns("Use").Value = True
    TDBGrid.Columns("West Use").Value = False
    TDBGrid.Columns("Estimated").Value = False
    TDBGrid.Columns("Factor").Value = False
    TDBGrid.Columns("TRACES").Value = False
    TDBGrid.Columns("Update").Value = False
    TDBGrid.Columns("mat_last_update_id") = 0
    TDBGrid.Columns("matprice_last_update_id") = 0
End Sub

Public Function CloneRow() As ADODB.RecordSet
    ' Pass to base class
    Set CloneRow = m_objGridMap.CloneRowRecordset
    
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

Public Sub Factor(dblFactor As Double, strComment As String, intColumns As Integer)
    Dim vntBookmark As Variant
    
    For Each vntBookmark In TDBGrid.SelBookmarks
        m_rec.Bookmark = vntBookmark
        If intColumns And 1 Then
            m_rec.Fields("List_Price") = Format(m_rec.Fields("List_Price") + m_rec.Fields("List_Price") * dblFactor / 100, "#.00")
            ' TRACES should always match list price
            m_rec.Fields("Traces_List_Price") = Format(m_rec.Fields("Traces_List_Price") + m_rec.Fields("Traces_List_Price") * dblFactor / 100, "#.00")
            m_rec.Fields("Factor_ind") = -1 ' This is only for List Price, not TRACES
        ElseIf intColumns And 2 Then
            m_rec.Fields("Traces_List_Price") = Format(m_rec.Fields("Traces_List_Price") + m_rec.Fields("Traces_List_Price") * dblFactor / 100, "#.00")
        End If
        m_rec.Fields("update_status_code") = CurrentUpdateStatusCode()
        m_rec.Fields("latest_price_update_comment") = strComment
        m_objGridMap.SetRowState Int(vntBookmark), STATE_MODIFIED
        'TDBGrid.Columns("Price").CellText(vntBookmark) = TDBGrid.Columns("Price").CellText(vntBookmark) * dblFactor
    Next
    vntBookmark = TDBGrid.SelBookmarks(0)
    TDBGrid.ReBind ' Reset grid contents
    TDBGrid.Bookmark = vntBookmark ' Set bookmark back again
End Sub
Public Sub SetMultiplier(dblMultiplier As Double)
    Dim vntBookmark As Variant
    For Each vntBookmark In TDBGrid.SelBookmarks
        m_rec.Bookmark = vntBookmark
        m_rec.Fields("Pct_Multiplier") = dblMultiplier
        m_rec.Fields("update_status_code") = CurrentUpdateStatusCode()
        m_objGridMap.SetRowState Int(vntBookmark), STATE_MODIFIED
    Next
    vntBookmark = TDBGrid.SelBookmarks(0)
    TDBGrid.ReBind ' Reset grid contents
    TDBGrid.Bookmark = vntBookmark ' Set bookmark back again
End Sub
Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
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

