VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUnitCostMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Unit Cost grid
Option Explicit

'<modulename> CUnitCostMap.cls</modulename>
'<functionname>General (Main) </functionname>
'
'<summary>  Helper class module for frmUnitCostGrid.frm
'</summary>
'
'<seealso>frmUnitCostGrid</seealso>
'
'<datastruct> TDBGRID</datastruct>
'<datastruct> m_objGridMap</datastruct>
'
'<storedprocedurename> USP_UPDATE_UNIT_COST_DRIVER_EXT
'</storedprocedurename>
'<storedprocedurename> USP_UPDATE_UNIT_COST_DRIVER
'</storedprocedurename>
'<storedprocedurename> USP_SELECT_UNIT_COST_EXT</storedprocedurename>
'<storedprocedurename> SP_DELETE_UNIT_COST</storedprocedurename>
'
'STORED PROCEDURES:
'
'Refer to documentation in "stored procedures" section of frmUnitCostGrid.frm
'
'<returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'exec usp_update_unit_cost_driver_ext_rlh  @ext_unit_cost_id='031104023006', @alt_unit_cost_id='031104023003', @unit_cost_id='030110203004', @type_code='M', @index_code='', @index_desc='', @format_characters=0, @format_code='F1', @indent_code=1, @metric_daily_output='23.22500', @daily_output='1.00000', @crew_id='C2', @unit='SFCA', @crew_qty='1', @assembly_book_desc='Forms (place & strip), beams, plywood, 1 use', @tech_desc='Forms (place & strip), beams, plywood, 1 use', @book_desc='Forms (place & strip), beams, plywood, 1 use', @metric_tech_desc='Forms (place & strip), beams, plywood, 1 use', @metric_book_desc='Forms (place & strip), beams, plywood, 1 use', @metric_assembly_book_desc='Forms (place & strip), beams, plywood, 1 use', @metric_unit='m2CA    ', @comment='', @unit_cost_skey=109103, @res_labor_hour='0.19200', @res_total_cost_op='7.35', @res_equip_cost_op='', @res_labor_cost_op='7.35', @res_mat_cost_op='', @res_total_cost='4.33', @res_equip_cost='', @res_labor_cost='4.33', @res_mat_cost='', @metric_l
'bor_hour='2.06674', @rr_labor_hour='0.19200', @opn_labor_hour='0.19200', @std_labor_hour='0.19200', @metric_equip_cost_op='', @metric_labor_cost_op='115.00', @metric_mat_cost_op='42.00', @rr_equip_cost_op='', @rr_labor_cost_op='11.35', @rr_mat_cost_op='3.92', @opn_equip_cost_op='', @opn_labor_cost_op='8.30', @opn_mat_cost_op='3.92', @std_equip_cost_op='', @std_labor_cost_op='10.70', @std_mat_cost_op='3.92', @metric_total_cost_op='157.00', @rr_total_cost_op='15.25', @opn_total_cost_op='12.20', @std_total_cost_op='14.60', @metric_total_cost='112.50', @metric_equip_cost='', @metric_labor_cost='74.00', @metric_mat_cost='38.50', @rr_total_cost='10.41', @rr_equip_cost='', @rr_labor_cost='6.85', @rr_mat_cost='3.56', @opn_total_cost='8.44', @opn_equip_cost='', @opn_labor_cost='4.88', @opn_mat_cost='3.56', @std_total_cost='10.41', @std_equip_cost='', @std_labor_cost='6.85', @std_mat_cost='3.56', @percent_flag='',  @last_update_person='Hancockrl', @bypass_ucd_ind = 0, @ucd_last_update_id=6,  @cstw_last_update_id=6,
'@update_material_usage_ind=0, @cost_change_ind=1
' </code>
' <code>
'</code>
'</example>
'<permission>CCDUSER</Permission>
'<dependson>This component depends on the following
'"   CGridMap.cls
'"   CCDdal.CRSMDataAccess (
'Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>


'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "UnitCost" ' Identity of grid
' Names of some table this grid uses
Const UNIT_COST_TABLE = "published_unit_cost_costworks"
Const UNITCOSTDTL_TABLE = "unit_cost_detail"
Const UNITCOSTDTLEXT_TABLE = "unit_cost_detail_ext"
Const UNITCOST_UPDATE = "published_unit_cost_costworks update"
Const PUBUNITCOSTEX_TABLE = "published_unit_cost_exception"
Const PUBUNITCOSTEX_UPDATE = "published_unit_cost_exception update"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap    ' "Parent" class
Private m_rec As ADODB.RecordSet        ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private m_varSaveBookmark As Variant
Private m_varCurrentBookmark As Variant
Private m_blnCrewQtyErr As Boolean
Private m_blnUnitErr As Boolean
Private m_intMasterFormat As Long

Private m_colUpdatedRows As New Collection




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

'ExtUnitCostIDColumnName Read-Only Property
'RETURNS THE GRID COLUMN NAME THAT HOLDS THE DATABASE FIELD "ext_unit_cost_id"
Public Property Get ExtUnitCostIDColumnName() As String
    Dim objColumnDef As CColumnDef
    
    Set objColumnDef = m_objGridMap.GetCol("ext_unit_cost_id")
    If Not (objColumnDef Is Nothing) Then
        ExtUnitCostIDColumnName = objColumnDef.Name
    End If
    Set objColumnDef = Nothing
    
End Property

Public Property Get MasterFormat() As Long
    MasterFormat = m_intMasterFormat
End Property
Public Property Let MasterFormat(NewValue As Long)
    Dim objColumnDef As CColumnDef
    Dim Col As TrueOleDBGrid80.Column
    
    m_intMasterFormat = NewValue
    If m_intMasterFormat = UCD_MASTERFORMAT_VERSION Then
        Set objColumnDef = m_objGridMap.GetCol("unit_cost_id")
        If Not (objColumnDef Is Nothing) Then
            objColumnDef.Format = FORMAT_UNIT_COST
        End If
        Set objColumnDef = m_objGridMap.GetCol("ext_unit_cost_id")
        If Not (objColumnDef Is Nothing) Then
            objColumnDef.Format = FORMAT_UNIT_COST_04
            objColumnDef.Name = "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2)
            For Each Col In TDBGrid.Columns
                If Col.DataField = "ext_unit_cost_id" Then
                    Col.Caption = objColumnDef.Name
                    Exit For
                End If
            Next
        End If
    Else
        Set objColumnDef = m_objGridMap.GetCol("unit_cost_id")
        If Not (objColumnDef Is Nothing) Then
            objColumnDef.Format = FORMAT_UNIT_COST_04
        End If
        Set objColumnDef = m_objGridMap.GetCol("ext_unit_cost_id")
        If Not (objColumnDef Is Nothing) Then
            objColumnDef.Format = FORMAT_UNIT_COST
            objColumnDef.Name = "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2)
            For Each Col In TDBGrid.Columns
                If Col.DataField = "ext_unit_cost_id" Then
                    Col.Caption = objColumnDef.Name
                    Exit For
                End If
            Next
        End If
    End If

    Set objColumnDef = Nothing
    
End Property

Public Sub InitGrid()
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "Type", "type_code", UNITCOSTDTL_TABLE, 500, True
    If m_intMasterFormat = UCD_MASTERFORMAT_VERSION Then
        m_objGridMap.AddCol "Unit Cost ID", "unit_cost_id", UNITCOSTDTL_TABLE, 1450, True, True, , FORMAT_UNIT_COST
        m_objGridMap.AddCol "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2), "ext_unit_cost_id", UNITCOSTDTL_TABLE, 1450, False, False, , FORMAT_UNIT_COST_04
    Else
        m_objGridMap.AddCol "Unit Cost ID", "unit_cost_id", UNITCOSTDTL_TABLE, 1450, True, True, , FORMAT_UNIT_COST_04
        'rlh CCD 8.4
        'm_objGridMap.AddCol "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2), "ext_unit_cost_id", UNITCOSTDTL_TABLE, 1450, False, False, , FORMAT_UNIT_COST
        
        'rlh CCD 8.4 - Set parm 5 to True to BLOCK!!!
         m_objGridMap.AddCol "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2), "ext_unit_cost_id", UNITCOSTDTL_TABLE, 1450, True, False, , FORMAT_UNIT_COST
    End If
    m_objGridMap.AddCol "Alt Unit Cost ID", "alt_unit_cost_id", UNITCOSTDTL_TABLE, 1450, False, False
    ' Main Split
    m_objGridMap.AddSplit
    'rlh 10/08/2010
    'vicm m_objGridMap.AddCol "Excp Chk", "Excp_chk", UNIT_COST_TABLE, 1000, , , , , True
    
    m_objGridMap.AddCol "Book Desc", "book_desc", UNITCOSTDTL_TABLE, 2000
    m_objGridMap.AddCol "Metric Book Desc", "metric_book_desc", UNITCOSTDTL_TABLE, 2000, , False
    m_objGridMap.AddCol "Tech Desc", "tech_desc", UNITCOSTDTL_TABLE, 2000, False, False
    m_objGridMap.AddCol "Metric Tech Desc", "metric_tech_desc", UNITCOSTDTL_TABLE, 2000, False, False
    m_objGridMap.AddCol "Index Code", "index_code", UNITCOSTDTL_TABLE, 1000, False, False
    m_objGridMap.AddCol "Index Desc", "index_desc", UNITCOSTDTL_TABLE, 2000, False, False
    m_objGridMap.AddCol "Crew Qty", "crew_qty", UNIT_COST_TABLE, 900, False, True, True, FORMAT_DECIMAL3, True
    m_objGridMap.AddCol "Crew ID", "crew_id", UNIT_COST_TABLE, 900, False, True, True, FORMAT_STRING_TRIM, True
    
    m_objGridMap.AddCol "Crew Type", "crew_type_code", vbNullString, 900, True, False, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Daily Out", "daily_output", UNIT_COST_TABLE, 900, False, True, True, FORMAT_DECIMAL, True
    m_objGridMap.AddCol "Metric Daily Out", "metric_daily_output", UNIT_COST_TABLE, 900, True, False, True, FORMAT_DECIMAL, True
    m_objGridMap.AddCol "Unit", "unit", UNIT_COST_TABLE, 1000, , , True, FORMAT_STRING_TRIM, True
    m_objGridMap.AddCol "Metric Unit", "metric_unit", UNIT_COST_TABLE, 1000, False, False, True, FORMAT_STRING_TRIM, True
    m_objGridMap.AddCol "Assm Book Desc", "assembly_book_desc", UNITCOSTDTL_TABLE, 2000, False, False
    m_objGridMap.AddCol "Metric Assm Bk Dsc", "metric_assembly_book_desc", UNITCOSTDTL_TABLE, 2000, False, False
'    m_objGridMap.AddCol "Traces Book Desc", "traces_book_desc", UNIT_COST_TABLE, 2000, False, False
    m_objGridMap.AddCol "Comment", "comment", UNITCOSTDTL_TABLE, 2000, False, False
    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
    m_objGridMap.AddCol "unit_cost_skey", "unit_cost_skey", ALL_TABLES, 800, True, False, True
    m_objGridMap.AddCol "last_update_id", "ucd_last_update_id", UNITCOST_UPDATE, 800, True, False, False
    m_objGridMap.AddCol "last_update_id", "cstw_last_update_id", vbNullString, 800, True, False, False
    ' Published Price Data
    m_objGridMap.AddCol "Percent", "percent_flag", UNIT_COST_TABLE, 1000, , , , , True
    m_objGridMap.AddCol "Std L Hr", "std_labor_hour", UNIT_COST_TABLE, 900, False, True, False, , True
    m_objGridMap.AddCol "Std Mat", "std_mat_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Mat O&P", "std_mat_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Std L $", "std_labor_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Labor O&P", "std_labor_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Eqpt", "std_equip_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Eqpt O&P", "std_equip_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Total", "std_total_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Std Total O&P", "std_total_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn L Hr", "opn_labor_hour", UNIT_COST_TABLE, 900, False, True, False, , True
    m_objGridMap.AddCol "Opn Mat", "opn_mat_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Mat O&P", "opn_mat_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn L $", "opn_labor_cost", UNIT_COST_TABLE, 900, True, False, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Labor O&P", "opn_labor_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Eqpt", "opn_equip_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Eqpt O&P", "opn_equip_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Total", "opn_total_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Opn Total O&P", "opn_total_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "RR L Hr", "rr_labor_hour", UNIT_COST_TABLE, 900, False, True, False, , True
    m_objGridMap.AddCol "RR Mat", "rr_mat_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "RR Mat O&P", "rr_mat_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "RR L $", "rr_labor_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "RR Labor O&P", "rr_labor_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "RR Eqpt", "rr_equip_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "RR Eqpt O&P", "rr_equip_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "RR Total", "rr_total_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "RR Total O&P", "rr_total_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Met L Hr", "metric_labor_hour", UNIT_COST_TABLE, 900, False, True, False, , True
    m_objGridMap.AddCol "Met Mat", "metric_mat_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Mat O&P", "metric_mat_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Met L $", "metric_labor_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Labor O&P", "metric_labor_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Eqpt", "metric_equip_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Eqpt O&P", "metric_equip_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Total", "metric_total_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Met Total O&P", "metric_total_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    
    'RESI (rlh) 05/7/2010
    m_objGridMap.AddCol "Res L Hr", "res_labor_hour", UNIT_COST_TABLE, 900, False, True, False, , True
    m_objGridMap.AddCol "Res Mat", "res_mat_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Res Mat O&P", "res_mat_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Res L $", "res_labor_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Res Labor O&P", "res_labor_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Res Eqpt", "res_equip_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Res Eqpt O&P", "res_equip_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Res Total", "res_total_cost", UNIT_COST_TABLE, 900, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "Res Total O&P", "res_total_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    
    'FMR (In-House) (rlh) 05/14/2010
    m_objGridMap.AddCol "FMR Mat O&P", "inhouse_mat_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "FMR Labor O&P", "inhouse_labor_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "FMR Eqpt O&P", "inhouse_equip_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True
    m_objGridMap.AddCol "FMR Total O&P", "inhouse_total_cost_op", UNIT_COST_TABLE, 1200, False, True, False, FORMAT_PRICE, True

'    ' Format split
'    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Format Code", "format_code", UNITCOSTDTL_TABLE, 1000, False, True, True
    m_objGridMap.AddCol "Indent Code", "indent_code", UNIT_COST_TABLE, 1000, False, True, True
    m_objGridMap.AddCol "Format Chars", "format_characters", UNIT_COST_TABLE, 1000, False, True, True
    
    
    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If
    
    ' Give headings for the splits
    TDBGrid.Splits(0).Caption = "Heading"
    TDBGrid.Splits(1).Caption = "Unit Cost"
    TDBGrid.Splits(1).Size = 5
'    TDBGrid.Splits(2).Caption = "Formatting"

    If Not m_blnPreferences Then
        Dim bRet As Boolean
        Dim rec As ADODB.RecordSet
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid60.ValueItem
'*** APEX Migration Utility Code Change ***
'        Dim Item As New TrueOleDBGrid70.ValueItem
        Dim Item As New TrueOleDBGrid80.ValueItem
        
        ' Unit drop downs
        bRet = g_objDAL.GetRecordset(vbNullString, "select unit from unit_of_measure order by unit", rec)
        If bRet Then
            While Not rec.EOF
                If Len(Trim(rec.Fields(0))) > 0 Then
                    Item.Value = rec.Fields(0)
                    TDBGrid.Columns("Unit").ValueItems.Add Item
                    TDBGrid.Columns("Metric Unit").ValueItems.Add Item
                End If
                rec.MoveNext
            Wend
            TDBGrid.Columns("Unit").ValueItems.Presentation = dbgComboBox
            TDBGrid.Columns("Unit").ValueItems.Validate = True
            TDBGrid.Columns("Metric Unit").ValueItems.Presentation = dbgComboBox
            TDBGrid.Columns("Metric Unit").ValueItems.Validate = True
        End If
        rec.Close

        ' Type Code drop down
        Item.Value = "M"
        TDBGrid.Columns("Type").ValueItems.Add Item
        Item.Value = "E"
        TDBGrid.Columns("Type").ValueItems.Add Item
        Item.Value = "H"
        TDBGrid.Columns("Type").ValueItems.Add Item
        Item.Value = "B"
        TDBGrid.Columns("Type").ValueItems.Add Item

        TDBGrid.Columns("Type").ValueItems.Presentation = dbgComboBox
        TDBGrid.Columns("Type").ValueItems.Validate = True
        TDBGrid.Columns("Type").AutoDropDown = True

        ' Percent Flag drop down
        Item.Value = "Y"
        Item.DisplayValue = True
        TDBGrid.Columns("Percent").ValueItems.Add Item
        Item.Value = " "
        Item.DisplayValue = False
        TDBGrid.Columns("Percent").ValueItems.Add Item
        TDBGrid.Columns("Percent").ValueItems.Translate = True
        TDBGrid.Columns("Percent").ValueItems.Presentation = dbgCheckBox
        TDBGrid.Columns("Percent").ValueItems.Validate = True
        
        
        ' Exception Check Flag '(rlh 10/08/2010 - week of the move to Norwell)
        'vicm Item.Value = "Y"
        'Item.DisplayValue = True
        'TDBGrid.Columns("Excp_chk").ValueItems.Add Item
        'Item.Value = " "
        'Item.DisplayValue = False
        'TDBGrid.Columns("Excp_chk").ValueItems.Add Item
        'TDBGrid.Columns("Excp_chk").ValueItems.Translate = True
        'TDBGrid.Columns("Excp_chk").ValueItems.Presentation = dbgCheckBox
        'TDBGrid.Columns("Excp_chk").ValueItems.Validate = True
        
        TDBGrid.Columns("Book Desc").FetchStyle = True
        
'*** APEX Migration Utility Code Change ***
'        Dim ShadedRow As TrueOleDBGrid60.Style
'*** APEX Migration Utility Code Change ***
'        Dim ShadedRow As TrueOleDBGrid70.Style
        Dim ShadedRow As TrueOleDBGrid80.Style
        ' Define a New Style that will be used within this Application
        Set ShadedRow = TDBGrid.Styles.Add("ShadedRow")
        ShadedRow.BackColor = vbBlack
    End If
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String     ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateUnitCost As Boolean
    Dim blnUpdateUnitCostEx As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String      ' Error string returned from DAL
    Dim intErrors As Integer    ' Tracks if any errors have occurred
    Dim intSuccess As Integer   ' Tracks successful updates

    On Error Resume Next
    Update = True
    intErrors = 0
    intSuccess = 0

    m_rec.MoveFirst

    ' Loop through all grid records
    While Not m_rec.EOF
        Dim clsRowInfo As CRowInfo
        Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
        ' Skip the record if it didn't change
        If Not clsRowInfo Is Nothing Then
            blnReturn = False
            blnUpdateUnitCost = False
            blnUpdateUnitCostEx = False
            ' Loop through the fields to see what has changed
            For Each fld In m_rec.Fields
                ' If the value changed
                If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
                    ' See what table the field is from
                    Set objColDef = m_objGridMap.GetCol(fld.Name)
                    ' Mark the table we should update
                    blnUpdateUnitCost = True
                End If
            Next
            
            ' If the row is new
            If clsRowInfo.State = STATE_NEW Then
            ' NOTE: Cannot do NEWs in the grid, launches single record view!
            ' If this row was modified
            ElseIf clsRowInfo.State = STATE_MODIFIED Then
                If blnUpdateUnitCost Then
                
                    Dim excludeList As New Collection
                    
                    'Update all types of rows.  The sp_update_unit_cost_driver will update the appropriate records.
                    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
                        strUpdate = "exec usp_update_unit_cost_driver_ext_rlh "         'rlh 05/22/2007  PRODUCTION
                        'strUpdate = "exec usp_update_unit_cost_driver_ext_rlh2 "       'rlh 05/22/2007  TEST/DEBUG
                    Else
                        strUpdate = "exec sp_update_unit_cost_driver_res "
                        
                        ' We need to exclude these columns from the stored proc.
                        excludeList.Add "ext_unit_cost_id", "ext_unit_cost_id"
                        excludeList.Add "inhouse_total_cost_op", "inhouse_total_cost_op"
                        excludeList.Add "inhouse_equip_cost_op", "inhouse_equip_cost_op"
                        excludeList.Add "inhouse_mat_cost_op", "inhouse_mat_cost_op"
                        excludeList.Add "inhouse_labor_cost_op", "inhouse_labor_cost_op"
                    End If
                    
                    'rlh 03/06/09 - Swap the values of the MF95 and MF04 unit cost ids
                    'from the recordset!!!
'                    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
'                        Dim tmp_unit_cost_id As String
'                        tmp_unit_cost_id = m_rec("unit_cost_id")
'                        m_rec("unit_cost_id") = m_rec("ext_unit_cost_id")
'                        m_rec("ext_unit_cost_id") = tmp_unit_cost_id
'                    End If
                    
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, UNIT_COST_TABLE, excludeList
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, UNITCOSTDTL_TABLE, excludeList
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, UNITCOST_UPDATE, excludeList
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES, excludeList
                    strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                    If IsNull(m_rec![cstw_last_update_id]) Then
                        strUpdate = strUpdate + ", @cstw_last_update_id=1"
                    Else
                        strUpdate = strUpdate + ", @cstw_last_update_id=" + str(m_rec![cstw_last_update_id].Value)
                    End If
                    strUpdate = strUpdate + ", @update_material_usage_ind=0"
                    strUpdate = strUpdate + ", @cost_change_ind=" + CStr(SetCostChange())
                    strUpdate = strUpdate + ", @bypass_ucd_ind = 0"
                    
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If blnReturn = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("ucd_last_update_id").Value = m_rec.Fields("ucd_last_update_id").Value + 1
                        m_rec.Fields("cstw_last_update_id").Value = m_rec.Fields("cstw_last_update_id").Value + 1
                        m_rec.Fields("unit_cost_last_update_date").Value = Now
                        m_rec.Fields("unit_cost_last_update_person").Value = strUserName
'                        m_colUpdatedRows.Add clsRowInfo.Bookmark
                        clsRowInfo.RowSQL = "exec usp_select_unit_cost_ext_rlh2 @start_unit_cost_id = '" + m_rec.Fields("unit_cost_id") + "%', @end_unit_cost_id = '%', @alt_unit_cost_id = '%', @tech_desc = '%', @master_format=" & MasterFormat
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
'        For Each m_varCurrentBookmark In m_colUpdatedRows
'            m_objGridMap.RefreshRecord CInt(m_varCurrentBookmark)
'            m_colUpdatedRows.Remove 1
'        Next
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

End Function

Private Function SetCostChange() As Integer
    SetCostChange = 0
    'Set the unit cost change flag to determine if a new cost or exception record is to be generated
    If m_rec.Fields("percent_flag").Value <> m_rec.Fields("percent_flag").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("type_code").Value <> m_rec.Fields("type_code").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("crew_qty").Value <> m_rec.Fields("crew_qty").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("crew_id").Value <> m_rec.Fields("crew_id").OriginalValue Then SetCostChange = 1
    
    If m_rec.Fields("daily_output").Value <> m_rec.Fields("daily_output").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_daily_output").Value <> m_rec.Fields("metric_daily_output").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("unit").Value <> m_rec.Fields("unit").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_unit").Value <> m_rec.Fields("metric_unit").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_mat_cost").Value <> m_rec.Fields("std_mat_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_labor_cost").Value <> m_rec.Fields("std_labor_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_labor_hour").Value <> m_rec.Fields("std_labor_hour").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_mat_cost_op").Value <> m_rec.Fields("std_mat_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_labor_cost").Value <> m_rec.Fields("std_labor_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_labor_cost_op").Value <> m_rec.Fields("std_labor_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_equip_cost").Value <> m_rec.Fields("std_equip_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_equip_cost_op").Value <> m_rec.Fields("std_equip_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_total_cost").Value <> m_rec.Fields("std_total_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("std_total_cost_op").Value <> m_rec.Fields("std_total_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_labor_hour").Value <> m_rec.Fields("opn_labor_hour").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_mat_cost").Value <> m_rec.Fields("opn_mat_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_mat_cost_op").Value <> m_rec.Fields("opn_mat_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_labor_cost").Value <> m_rec.Fields("opn_labor_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_labor_cost_op").Value <> m_rec.Fields("opn_labor_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_equip_cost").Value <> m_rec.Fields("opn_equip_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_equip_cost_op").Value <> m_rec.Fields("opn_equip_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_total_cost").Value <> m_rec.Fields("opn_total_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("opn_total_cost_op").Value <> m_rec.Fields("opn_total_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_labor_hour").Value <> m_rec.Fields("rr_labor_hour").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_mat_cost").Value <> m_rec.Fields("rr_mat_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_mat_cost_op").Value <> m_rec.Fields("rr_mat_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_labor_cost").Value <> m_rec.Fields("rr_labor_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_labor_cost_op").Value <> m_rec.Fields("rr_labor_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_equip_cost").Value <> m_rec.Fields("rr_equip_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_equip_cost_op").Value <> m_rec.Fields("rr_equip_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_total_cost").Value <> m_rec.Fields("rr_total_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("rr_total_cost_op").Value <> m_rec.Fields("rr_total_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_labor_hour").Value <> m_rec.Fields("metric_labor_hour").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_mat_cost").Value <> m_rec.Fields("metric_mat_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_mat_cost_op").Value <> m_rec.Fields("metric_mat_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_labor_cost").Value <> m_rec.Fields("metric_labor_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_labor_cost_op").Value <> m_rec.Fields("metric_labor_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_equip_cost").Value <> m_rec.Fields("metric_equip_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_equip_cost_op").Value <> m_rec.Fields("metric_equip_cost_op").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_total_cost").Value <> m_rec.Fields("metric_total_cost").OriginalValue Then SetCostChange = 1
    If m_rec.Fields("metric_total_cost_op").Value <> m_rec.Fields("metric_total_cost_op").OriginalValue Then SetCostChange = 1
End Function

Private Sub Class_Initialize()

    m_intMasterFormat = UCD_MASTERFORMAT_VERSION

    m_objGridMap.m_colCloneCols.Add 1, "unit_cost_skey"
    m_objGridMap.m_colCloneCols.Add 1, "unit_cost_id"
    m_objGridMap.m_colCloneCols.Add 1, "ext_unit_cost_id"
    m_objGridMap.m_colCloneCols.Add 1, "alt_unit_cost_id"
    m_objGridMap.m_colCloneCols.Add 1, "type_code"
    m_objGridMap.m_colCloneCols.Add 1, "crew_qty"
    m_objGridMap.m_colCloneCols.Add 1, "crew_id"
  
    m_objGridMap.m_colCloneCols.Add 1, "crew_type_code"
    m_objGridMap.m_colCloneCols.Add 1, "unit"
    m_objGridMap.m_colCloneCols.Add 1, "metric_unit"
    m_objGridMap.m_colCloneCols.Add 1, "daily_output"
    m_objGridMap.m_colCloneCols.Add 1, "metric_daily_output"
    'vicm m_objGridMap.m_colCloneCols.Add 1, "Excp_chk"       'rlh 11/08/2010 (week of the move)
    m_objGridMap.m_colCloneCols.Add 1, "book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "assembly_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_assembly_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "index_code"
    m_objGridMap.m_colCloneCols.Add 1, "index_desc"
    m_objGridMap.m_colCloneCols.Add 1, "indent_code"
    m_objGridMap.m_colCloneCols.Add 1, "format_characters"
    m_objGridMap.m_colCloneCols.Add 1, "format_code"
    m_objGridMap.m_colCloneCols.Add 1, "graphic_ref_id"
    m_objGridMap.m_colCloneCols.Add 1, "table_ref_id"
    m_objGridMap.m_colCloneCols.Add 1, "table_ref_col"
    m_objGridMap.m_colCloneCols.Add 1, "change_notice_code"
    m_objGridMap.m_colCloneCols.Add 1, "comment"
    m_objGridMap.m_colCloneCols.Add 1, "last_update_id"
    ' Exception
    m_objGridMap.m_colCloneCols.Add 1, "percent_flag"
    m_objGridMap.m_colCloneCols.Add 1, "std_mat_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_mat_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "std_labor_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_labor_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "std_equip_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_equip_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "std_total_cost"
    m_objGridMap.m_colCloneCols.Add 1, "std_total_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "std_labor_hour"
    m_objGridMap.m_colCloneCols.Add 1, "opn_mat_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_mat_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_equip_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_equip_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_total_cost"
    m_objGridMap.m_colCloneCols.Add 1, "opn_total_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "opn_labor_hour"
    m_objGridMap.m_colCloneCols.Add 1, "rr_mat_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_mat_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_labor_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_labor_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_equip_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_equip_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_total_cost"
    m_objGridMap.m_colCloneCols.Add 1, "rr_total_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "rr_labor_hour"
    m_objGridMap.m_colCloneCols.Add 1, "metric_mat_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_mat_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_labor_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_labor_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_equip_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_equip_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_total_cost"
    m_objGridMap.m_colCloneCols.Add 1, "metric_total_cost_op"
    m_objGridMap.m_colCloneCols.Add 1, "metric_labor_hour"
    
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    Dim intMsg As Integer
    m_varSaveBookmark = TDBGrid.Bookmark
    If TDBGrid.Columns(ColIndex).Caption = "Type" Or TDBGrid.Columns(ColIndex).Caption = "Crew Id" Then    'If record type or crew id change, update appropriate values
        Select Case TDBGrid.Columns(TDBGrid.Col).Caption
            Case "Type"
                intMsg = 1
            Case "Crew Id"
                intMsg = 2
                TDBGrid.RefreshCol (TDBGrid.Columns("Daily Out").ColIndex)
        End Select
    End If
    If intMsg > 0 Then
        TDBGrid.PostMsg intMsg
    End If
End Sub

Private Sub TDBGrid_AfterUpdate()
    If TDBGrid.Columns("Crew ID") = "" Then
                TDBGrid.Columns("Crew Qty") = ""
                TDBGrid.Columns("Crew Type") = ""
                TDBGrid.Columns("Daily Out") = ""
                TDBGrid.Columns("Metric Daily Out") = ""
    End If
    If m_blnCrewQtyErr = True Then
        TDBGrid.PostMsg 3
    End If
End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
    'Validate the Crew Qty based on Crew.type_code
    m_blnUnitErr = False
    m_blnCrewQtyErr = False
    If TDBGrid.Columns("Crew Type").Value = "L" Then
        If Not IsNumeric(TDBGrid.Columns("Crew Qty")) Then
            m_blnCrewQtyErr = True
            Cancel = True
        ElseIf Val(CInt(TDBGrid.Columns("Crew Qty"))) <> Val(TDBGrid.Columns("Crew Qty")) Then
            m_blnCrewQtyErr = True
            Cancel = True
            Screen.MousePointer = vbNormal
        ElseIf Val(TDBGrid.Columns("Crew Qty")) < 1 Then    'Quantity is required
            m_blnCrewQtyErr = True
            Cancel = True
            Screen.MousePointer = vbNormal
        End If
    Else
        If TDBGrid.Columns("Type") <> "H" And Trim(TDBGrid.Columns("Crew ID")) <> "" And Trim(TDBGrid.Columns("Crew Qty")) = "" Then
            m_blnCrewQtyErr = True
            Cancel = True
        End If
    End If
    If TDBGrid.Columns("Type") = "M" Then
        If Len(Trim(TDBGrid.Columns("Unit"))) = 0 Then    'unit required
            m_blnUnitErr = True
            Cancel = True
        End If
    End If

End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
    If m_blnCrewQtyErr = True Then
        MsgBox "Please enter a valid Crew Qty."
        Response = 0
        TDBGrid.SetFocus
        m_blnCrewQtyErr = False
    End If
    If m_blnUnitErr = True Then
        MsgBox "Please enter a valid Unit of Measure."
        Response = 0
        TDBGrid.SetFocus
        m_blnCrewQtyErr = False
    End If

End Sub

'*** APEX Migration Utility Code Change ***
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
'Validate the Crew Qty based on Crew.type_code
    On Error Resume Next

    Dim strType As String
    Dim objRowInfo As CRowInfo
    Dim bLocked As Boolean

    objRowInfo = Nothing
    Set objRowInfo = m_objGridMap.rowinfo(str(Bookmark))
    If Split = 1 Then   'Price/Exception split
        If IsNumeric(Bookmark) Then
            m_rec.Bookmark = Bookmark
            strType = m_rec.Fields("type_code").Value
            With TDBGrid.Columns(Col)
                'Process Exception record fields
                If strType = "E" Then
                    Select Case .Caption
                        Case "Crew Qty"
                            bLocked = True
                        Case "Crew Id"
                            bLocked = True
                        Case "Daily Out"
                            bLocked = True
                        Case "Metric Daily Out"
                            bLocked = True
                        Case "Std L Hr"
                            bLocked = True
                        Case "Opn L Hr"
                            bLocked = True
                        Case "RR L Hr"
                            bLocked = True
                        Case "Met L Hr"
                            bLocked = True
                    End Select
                ElseIf strType = "B" Or strType = "M" Or strType = "H" Then
                    Select Case .Caption
                        Case "Crew Qty"
                            If TDBGrid.DataChanged = True Then
                                If Not IsNull(TDBGrid.Columns("Crew Type").Value) Then
                                    If TDBGrid.Columns("Crew Type").Value = "C" Or Trim(TDBGrid.Columns("Crew ID").Value) = "" Then
                                        bLocked = True
                                    End If
                                End If
                            Else
                                If Not IsNull(m_rec.Fields("crew_type_code").Value) Then
                                    If m_rec.Fields("crew_type_code").Value = "C" Or Trim(m_rec.Fields("crew_id").Value) = "" Then
                                        bLocked = True
                                    End If
                                End If
                            End If
                        Case "Daily Out"
                            If TDBGrid.DataChanged = True Then
                                If Trim(TDBGrid.Columns("Crew ID").Value) = "" Then
                                    bLocked = True
                                End If
                            Else
                                If IsNull(m_rec.Fields("crew_id").Value) Or Trim(m_rec.Fields("crew_id").Value) = "" Then
                                    bLocked = True
                                End If
                            End If
                        Case "Metric Daily Out"
                            If TDBGrid.DataChanged = True Then
                                If Trim(TDBGrid.Columns("Crew ID").Value) = "" Then
                                    bLocked = True
                                End If
                            Else
                                If IsNull(m_rec.Fields("crew_id").Value) Or Trim(m_rec.Fields("crew_id").Value) = "" Then
                                    bLocked = True
                                End If
                            End If
                        Case "Metric Unit"
                            bLocked = True
'                        Case "Book Desc"
'                            bLocked = False
                        Case "Percent"
                            bLocked = True
                        Case "Std L Hr"
                            bLocked = True
                        Case "Std Mat"
                            bLocked = True
                        Case "Std Mat O&P"
                            bLocked = True
                        Case "Std L $"
                            bLocked = True
                        Case "Std Labor O&P"
                            bLocked = True
                        Case "Std Eqpt"
                            bLocked = True
                        Case "Std Eqpt O&P"
                            bLocked = True
                        Case "Std Total"
                            bLocked = True
                        Case "Std Total O&P"
                            bLocked = True
                        Case "Opn L Hr"
                            bLocked = True
                        Case "Opn Mat"
                            bLocked = True
                        Case "Opn Mat O&P"
                            bLocked = True
                        Case "Opn L $"
                            bLocked = True
                        Case "Opn Labor O&P"
                            bLocked = True
                        Case "Opn Eqpt"
                            bLocked = True
                        Case "Opn Eqpt O&P"
                            bLocked = True
                        Case "Opn Total"
                            bLocked = True
                        Case "Opn Total O&P"
                            bLocked = True
                        Case "RR L Hr"
                            bLocked = True
                        Case "RR Mat"
                            bLocked = True
                        Case "RR Mat O&P"
                            bLocked = True
                        Case "RR L $"
                            bLocked = True
                        Case "RR Labor O&P"
                            bLocked = True
                        Case "RR Eqpt"
                            bLocked = True
                        Case "RR Eqpt O&P"
                            bLocked = True
                        Case "RR Total"
                            bLocked = True
                        Case "RR Total O&P"
                            bLocked = True
                        Case "Met L Hr"
                            bLocked = True
                        Case "Met Mat"
                            bLocked = True
                        Case "Met Mat O&P"
                            bLocked = True
                        Case "Met L $"
                            bLocked = True
                        Case "Met Labor O&P"
                            bLocked = True
                        Case "Met Eqpt"
                            bLocked = True
                        Case "Met Eqpt O&P"
                            bLocked = True
                        Case "Met Total"
                            bLocked = True
                        Case "Met Total O&P"
                            bLocked = True
                    End Select
                    'Lock additional header fields
                    If strType = "H" Then
                        Select Case .Caption
'                            Case "Book Desc"
'                                bLocked = False
                            Case "Crew Qty"
                                bLocked = False
                            Case "Crew Id"
                                bLocked = True
                            Case "Daily Out"
                                bLocked = True
                            Case "Metric Daily Out"
                                bLocked = True
                            Case "Unit"
                                bLocked = True
                        End Select
                    End If
                End If
                
                If bLocked = True Then
                    bLocked = True
                    CellStyle.Locked = True
                    CellStyle.ForeColor = vbGrayText
                Else
                    CellStyle.ForeColor = vbBlack
                    CellStyle.Locked = False
                End If
            End With
        End If
    End If
    
    ' If the row is highlighted, then let it be
    If (Condition And dbgSelectedRow) = 8 Then
            CellStyle.ForeColor = vbWhite
        Exit Sub
    End If
    ' This col is always white unless an error occurred
    ' NOTE: In this release of the grid, definied styles (like "Normal") cannot be applied to cells
    If Not objRowInfo Is Nothing Then
        If objRowInfo.Error = True Then
            CellStyle.BackColor = vbRed
        Else
        End If
    Else
    End If

End Sub

Private Sub TDBGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If TDBGrid.Columns(TDBGrid.Col).DataField = "book_desc" Or _
            TDBGrid.Columns(TDBGrid.Col).DataField = "metric_book_desc" Or _
            TDBGrid.Columns(TDBGrid.Col).DataField = "tech_desc" Or _
            TDBGrid.Columns(TDBGrid.Col).DataField = "metric_tech_desc" Or _
            TDBGrid.Columns(TDBGrid.Col).DataField = "index_desc" Or _
            TDBGrid.Columns(TDBGrid.Col).DataField = "assembly_book_desc" Or _
            TDBGrid.Columns(TDBGrid.Col).DataField = "metric_assembly_book_desc" Then
            If Len(TDBGrid.Text) + 1 > 75 Then
                KeyAscii = 0
            End If
        End If
    End If

End Sub

Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)
'Validate the Crew Qty based on Crew.type_code
Dim strSelect As String
Dim rec As New ADODB.RecordSet ' Recordset to hold query results
Dim blnReturn As Boolean
If MsgId = 1 Or MsgId = 2 Then
    If TDBGrid.DataChanged = False Then         'Don't need to position if still editing current row.
        m_varCurrentBookmark = TDBGrid.Bookmark
        TDBGrid.Bookmark = m_varSaveBookmark
    End If
End If
Select Case MsgId
Case 1        'Type code was change - reset appropriate values
    Screen.MousePointer = vbHourglass
    If TDBGrid.Columns(0).Value = "E" Or TDBGrid.Columns(0).Value = "H" Then
        TDBGrid.Columns("Std L Hr") = 0
        TDBGrid.Columns("Opn L Hr") = 0
        TDBGrid.Columns("RR L Hr") = 0
        TDBGrid.Columns("Met L Hr") = 0
        TDBGrid.Columns("Metric Unit") = ""
        TDBGrid.Columns("Unit") = ""
    End If
    If TDBGrid.Columns(0).Value = "H" Then
        TDBGrid.Columns("Metric Unit") = ""
        TDBGrid.Columns("Unit") = ""
        TDBGrid.Columns("Std Mat") = 0
        TDBGrid.Columns("Std Mat O&P") = 0
        TDBGrid.Columns("Std L $") = 0
        TDBGrid.Columns("Std Labor O&P") = 0
        TDBGrid.Columns("Std Eqpt") = 0
        TDBGrid.Columns("Std Eqpt O&P") = 0
        TDBGrid.Columns("Std Total") = 0
        TDBGrid.Columns("Std Eqpt O&P") = 0
        TDBGrid.Columns("Std Total") = 0
        TDBGrid.Columns("Std Total O&P") = 0
        TDBGrid.Columns("Opn Mat") = 0
        TDBGrid.Columns("Opn Mat O&P") = 0
        TDBGrid.Columns("Opn L $") = 0
        TDBGrid.Columns("Opn Labor O&P") = 0
        TDBGrid.Columns("Opn Eqpt") = 0
        TDBGrid.Columns("Opn Eqpt O&P") = 0
        TDBGrid.Columns("Opn Total") = 0
        TDBGrid.Columns("Opn Total O&P") = 0
        TDBGrid.Columns("RR Mat") = 0
        TDBGrid.Columns("RR Mat O&P") = 0
        TDBGrid.Columns("RR L $") = 0
        TDBGrid.Columns("RR Labor O&P") = 0
        TDBGrid.Columns("RR Eqpt") = 0
        TDBGrid.Columns("RR Eqpt O&P") = 0
        TDBGrid.Columns("RR Total") = 0
        TDBGrid.Columns("RR Total O&P") = 0
        TDBGrid.Columns("Met Mat") = 0
        TDBGrid.Columns("Met Mat O&P") = 0
        TDBGrid.Columns("Met L $") = 0
        TDBGrid.Columns("Met Labor O&P") = 0
        TDBGrid.Columns("Met Eqpt") = 0
        TDBGrid.Columns("Met Eqpt O&P") = 0
        TDBGrid.Columns("Met Total") = 0
        TDBGrid.Columns("Met Total O&P") = 0
    End If
    If TDBGrid.Columns(0).Value = "M" Then
        MsgBox "Make sure you attach at least one Material to the Unit Cost using the MATERIAL_USAGE button."
    End If
    TDBGrid.Bookmark = m_varCurrentBookmark
    Screen.MousePointer = vbNormal
Case 2      'Crew type code was changed; if type_code = C, set crew qty = 1
    Screen.MousePointer = vbHourglass
    If Trim(TDBGrid.Columns("Crew Id").Value) = "" Then   'Crew ID was cleared, clear Crew Qty, Daily Output, Metric Daily Output
        TDBGrid.Columns("Daily Out") = ""
        TDBGrid.Columns("Metric Daily Out") = ""
        TDBGrid.Columns("Crew Qty") = ""
    Else
        strSelect = "Select type_code from crew where crew.crew_id='" + TDBGrid.Columns("Crew Id").Value + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        ' If it does, copy that data into grid
        If rec.RecordCount > 0 Then
            TDBGrid.Columns("Crew Type") = rec.Fields("type_code").Value
            If rec.Fields("type_code") = "C" Then
                TDBGrid.Columns("Crew Qty") = 1
            End If
        End If
        rec.Close
    End If
    If TDBGrid.DataChanged = False Then
        TDBGrid.Bookmark = m_varCurrentBookmark
    End If
Case 3      'qty Error     - reposition back to original line, saving values
If TDBGrid.DataChanged = False Then
    TDBGrid.Bookmark = m_varSaveBookmark
    MsgBox "Please enter a valid crew quantity."
    TDBGrid.Split = 1
    TDBGrid.Col = TDBGrid.Columns("Crew Qty").ColIndex
    TDBGrid.RefetchRow
End If
    
End Select
Screen.MousePointer = vbNormal

End Sub

' Fired when a row is deleted
Private Sub TDBGrid_UnboundDeleteRow(Bookmark As Variant)
    On Error Resume Next
    Dim strUpdate As String
    Dim blnReturn As Boolean
    Dim strError As String
    Dim clsRowInfo As CRowInfo
    
    ' Position to the desired record and execute delete
    m_rec.Bookmark = Bookmark
    
    Set clsRowInfo = m_objGridMap.rowinfo(m_rec.Bookmark)
    If Not clsRowInfo Is Nothing Then
        If clsRowInfo.State = STATE_NEW Then
            TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
            m_rec.Delete ' Remove it from recordset so it doesn't show.
            MsgBox "Delete successful."
            Exit Sub
        End If
    End If
    
    ' Build SQL statement
    strUpdate = "exec sp_delete_unit_cost "
    strUpdate = strUpdate + "@unit_cost_skey=" + str(m_rec.Fields("unit_cost_skey")) + ","
    strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    
    blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnReturn Then
        m_objGridMap.SetError m_rec.Bookmark, strError
        m_rec.Fields("Error").Value = "Error"
        Bookmark = Null
    Else
        TDBGrid.ApproxCount = TDBGrid.ApproxCount - 1
        m_objGridMap.SetRowState m_rec.Bookmark, STATE_DELETED
        m_rec.Delete ' Remove it from recordset so it doesn't show.
    End If
'    TDBGrid.ReBind
    Exit Sub

NoDelete:
    MsgBox Err.Description
    Bookmark = Null
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
        Dim strSelect As String
        Dim rec As New ADODB.RecordSet ' Recordset to hold query results
        Dim blnReturn As Boolean
        Dim vntMyBookmark As Variant
        Dim i As Integer

    With TDBGrid
        If .Columns(.Col).Caption = "Crew Id" Then
            .Columns(.Col).Value = UCase(.Columns(.Col).Value)
        ElseIf .Columns(.Col).Caption = "Book Desc" _
            Or .Columns(.Col).Caption = "Metric Book Desc" _
            Or .Columns(.Col).Caption = "Tech Desc" _
            Or .Columns(.Col).Caption = "Metric Tech Desc" _
            Or .Columns(.Col).Caption = "Metric Assm Bk Dsc" _
            Or .Columns(.Col).Caption = "Assm Book Desc" _
        Then
            If Len(.Text) > 75 Then .Text = Left(.Text, 75)
        End If
    End With
    
    ' For Unit Cost ID column, fetch Unit Cost data
'    m_blnCrewQtyErr = False
    If TDBGrid.Columns(ColIndex).Caption = "Unit Cost ID" Then

        ' Check to see if the unit_cost_id entered exists already
        strSelect = "Select *, last_update_id as 'last_update_id' from Unit_cost where unit_cost_id='" + Compress_String(TDBGrid.Text) + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        ' If it does, copy that data into grid
        If rec.RecordCount > 0 Then
        
            ' If they entered into the add row
            If OldValue = Empty Then
                m_rec.AddNew ' Add new record to recordset that populates grid
            ' If they changed an existing row
            Else
                m_rec.Bookmark = TDBGrid.Bookmark ' Move current record
            End If
            ' Make a new RowInfo
            m_objGridMap.SetRowState m_rec.Bookmark, STATE_NEW

            ' Copy all values
            On Error Resume Next
            For i = 0 To rec.Fields.Count - 1
                m_rec.Fields(rec.Fields(i).Name) = rec.Fields(i).Value
            Next i

            ' Set some defaults
            m_rec.Fields("pct_ind") = False
            m_rec.Fields("last_update_id") = 0
            m_rec.Fields("last_update_id_x") = 0
            m_rec.Fields("country_code") = "USA"
            m_rec.Fields("region_code") = "NAT"
            m_rec.Fields("last_update_id") = 0
            ' Set these because they are not shown by default and need values


            vntMyBookmark = m_rec.Bookmark ' Save current bookmark
'            TDBGrid.Bookmark = vntMyBookmark
            TDBGrid.ReBind ' Reset grid data
            TDBGrid.ReBind ' Do it twice, don't know why
            TDBGrid.Bookmark = vntMyBookmark ' Set bookmark back
            Cancel = True ' Cancel this event so grid doesn't do any more processing
        ' If equip_id doesn't exist
        Else
            ' If this is an existing row or clone
            If Not OldValue = Empty Then
                m_rec.Bookmark = TDBGrid.Bookmark ' Move current record
                ' Blank out the key
                m_rec.Fields("unit_cost_skey") = 0
            End If
        End If
    End If
    
    Select Case TDBGrid.Columns(ColIndex).Caption
    Case "Crew Id"

        'Validate the Crew Qty based on Crew.type_code
        If Trim(TDBGrid.Columns("Crew Id").Value) = "" Then   'Crew ID was cleared, clear Crew Qty, Daily Output, Metric Daily Output
        Else
            strSelect = "Select type_code from crew  inner join published_crew_rate " + _
                " on crew.crew_skey = published_crew_rate.crew_skey " + _
                "where crew.crew_id='" + TDBGrid.Columns("Crew Id").Value + "'"
            ' Use DAL to perform select
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
            ' If it does, copy that data into grid
            If rec.RecordCount = 0 Then
                MsgBox "Please enter a valid crew.  The crew must also be in the Published Crew Rate table."
                Cancel = True
            End If
            rec.Close
        End If
    Case "Crew Qty"
        If TDBGrid.Columns("Crew Type") = "L" Then
            If Not IsNumeric(TDBGrid.Columns("Crew Qty")) Then
                MsgBox "Please enter a whole number for this type of crew."
                TDBGrid.Columns("Crew Qty") = OldValue
                Cancel = True
            ElseIf Val(CInt(TDBGrid.Columns("Crew Qty"))) <> Val(TDBGrid.Columns("Crew Qty")) Then
                MsgBox "Please enter a whole number for this type of crew."
                TDBGrid.Columns("Crew Qty") = OldValue
                Cancel = True
                Screen.MousePointer = vbNormal
            ElseIf Val(TDBGrid.Columns("Crew Qty")) < 1 Then    'Quantity is required
                MsgBox "Please enter a whole number for this type of crew."
                TDBGrid.Columns("Crew Qty") = OldValue
                Cancel = True
                Screen.MousePointer = vbNormal
            End If
        End If
        

    Case "Type"
        Cancel = Not validate_uc_type_code(TDBGrid.Text, CLng(TDBGrid.Columns("unit_cost_skey")))
    End Select
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    MsgBox "TEST FROM CLASS"
    TDBGrid.Columns("Percent").Value = False
    TDBGrid.Columns("Country").Value = "USA"
    TDBGrid.Columns("Region").Value = "NAT"
    TDBGrid.Columns("last_update_id") = 0
    TDBGrid.Columns("last_update_id_x") = 0
    ' Set these because they are not shown by default and need values.
    
End Sub

Public Sub CloneRow()
    Dim vntBookmark
    vntBookmark = TDBGrid.Bookmark
    
    ' Pass to base class
    m_objGridMap.CloneRow
    
    ' Copy other values not listed in grid
    Dim rec As ADODB.RecordSet ' Copy of existing recordset

    Set rec = m_rec.Clone ' Copy existing recordset so we can have multiple current rows
    rec.Bookmark = vntBookmark ' move current row in copy
    m_rec.Fields("book_desc").Value = rec.Fields("book_desc").Value
    m_rec.Fields("metric_book_desc").Value = rec.Fields("metric_book_desc").Value
    m_rec.Fields("crew_equip_desc").Value = rec.Fields("crew_equip_desc").Value
    m_rec.Fields("crew_equip_desc_plural").Value = rec.Fields("crew_equip_desc_plural").Value
    m_rec.Fields("index_code").Value = rec.Fields("index_code").Value
    m_rec.Fields("index_desc").Value = rec.Fields("index_desc").Value
    m_rec.Fields("metric_crew_equip_desc").Value = rec.Fields("metric_crew_equip_desc").Value
    m_rec.Fields("metric_crew_equip_desc_plural").Value = rec.Fields("metric_crew_equip_desc_plural").Value
    m_rec.Fields("traces_ind").Value = rec.Fields("traces_ind").Value
    m_rec.Fields("indent_code").Value = rec.Fields("indent_code").Value
    m_rec.Fields("format_characters").Value = rec.Fields("format_characters").Value
    m_rec.Fields("format_code").Value = rec.Fields("format_code").Value
    m_rec.Fields("graphic_ref_id").Value = rec.Fields("graphic_ref_id").Value
    m_rec.Fields("table_ref_id").Value = rec.Fields("table_ref_id").Value
    m_rec.Fields("table_ref_col").Value = rec.Fields("table_ref_col").Value
    m_rec.Fields("change_notice_code").Value = rec.Fields("change_notice_code").Value

    ' These fields are associated with exceptions only
    If m_rec.Fields("type_code") = "E" Then
        m_rec.Fields("crew_equip_cost_x").Value = rec.Fields("crew_equip_cost_x").Value
        m_rec.Fields("metric_rent_per_day_x").Value = rec.Fields("metric_rent_per_day_x").Value
        m_rec.Fields("metric_rent_per_week_x").Value = rec.Fields("metric_rent_per_week_x").Value
        m_rec.Fields("metric_rent_per_month_x").Value = rec.Fields("metric_rent_per_month_x").Value
        m_rec.Fields("metric_operating_cost_hrly_x").Value = rec.Fields("metric_operating_cost_hrly_x").Value
        m_rec.Fields("metric_crew_equip_cost_x").Value = rec.Fields("metric_crew_equip_cost_x").Value
    End If
End Sub

Public Function CloneRowRecordset() As ADODB.RecordSet
    Set CloneRowRecordset = m_objGridMap.CloneRowRecordset
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

Public Function Delete() As Boolean
Dim varSelectedBookmarks() As Variant
Dim varCurrentSelectedBookmark As Variant
Dim varButton
Dim blnResult As Boolean
Dim i As Integer
Dim i1 As Integer

'On Error Resume Next

    If TDBGrid.SelBookmarks.Count > 0 Then
        ReDim varSelectedBookmarks(0 To TDBGrid.SelBookmarks.Count)
        i = 0
        For Each varCurrentSelectedBookmark In TDBGrid.SelBookmarks
            varSelectedBookmarks(i) = varCurrentSelectedBookmark
            i = i + 1
        Next varCurrentSelectedBookmark
        For i = 0 To UBound(varSelectedBookmarks) - 1
            TDBGrid.Bookmark = varSelectedBookmarks(i)
            
            'ccd 8.4 (rlh) Created/addeed "check_delete_error" to the IF block
            If check_error(TDBGrid.Bookmark, TDBGrid.Columns("unit_cost_skey")) = True _
            Or check_delete_error(TDBGrid.Bookmark, TDBGrid.Columns("unit_cost_skey")) = True _
            Then
                For i1 = 0 To TDBGrid.SelBookmarks.Count
                    If TDBGrid.SelBookmarks(i1) = varSelectedBookmarks(i) Then
                        'TDBGrid.SelBookmarks.Remove i1
                        'CCD 8.4 (rlh)
                        MsgBox ("Current Unit Cost Line can't be deleted until book coding is removed.  Thank you")
                        'Exit For
                        Exit Function
                    End If
                Next i1
            End If
        Next i
        If TDBGrid.SelBookmarks.Count > 0 Then
            Delete = m_objGridMap.Delete
        Else
            Delete = False
        End If
    ElseIf TDBGrid.Columns("unit_cost_skey") <> "" Then
        If check_error(TDBGrid.Bookmark, TDBGrid.Columns("unit_cost_skey")) = True Then
            Delete = False
        Else
            Delete = m_objGridMap.Delete
        End If
    Else
        Delete = False
    End If
End Function

Private Function check_error(varBookmark As Variant, lngSkey As Long) As Boolean
Dim clsRowInfo As CRowInfo
Dim intErrors As Integer ' Tracks if any errors have occurred
Dim rsTemp As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean
Dim intTraceCount As Integer
Const strError = "This unit cost uses Traces output and may not be deleted. Please remove it before deleting."
                
    strSelect = "select count(*) as traces_count from output_usage where skey_type = 'U' and output_id = 69 and skey = " + CStr(lngSkey)
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsTemp)
    If blnReturn = True Then
        intTraceCount = rsTemp.Fields("traces_count")
    Else
        intTraceCount = 0
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    If intTraceCount > 0 Then
        Set clsRowInfo = m_objGridMap.rowinfo(CInt(varBookmark))
        If clsRowInfo Is Nothing Then
            m_objGridMap.SetRowState CInt(varBookmark), STATE_NONE
            Set clsRowInfo = m_objGridMap.rowinfo(CInt(varBookmark))
        End If
        intErrors = intErrors + 1
        clsRowInfo.Error = True
        clsRowInfo.ErrorText = strError
        check_error = True
    Else
        check_error = False
    End If
End Function

Private Function check_delete_error(varBookmark As Variant, lngSkey As Long) As Boolean
Dim clsRowInfo As CRowInfo
Dim intErrors As Integer ' Tracks if any errors have occurred
Dim rsTemp As ADODB.RecordSet
Dim strSelect As String
Dim blnReturn As Boolean
Dim intDeleteCount As Integer
Const strError = "This unit cost is book coded and may not be deleted. Please remove book coding before deleting."
                
    strSelect = "select count(*) as delete_count from output_usage where skey_type = 'U' and output_id <> 0 and skey = " + CStr(lngSkey)
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsTemp)
    If blnReturn = True Then
        intDeleteCount = rsTemp.Fields("delete_count")
    Else
        intDeleteCount = 0
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    If intDeleteCount > 0 Then
        Set clsRowInfo = m_objGridMap.rowinfo(CInt(varBookmark))
        If clsRowInfo Is Nothing Then
            m_objGridMap.SetRowState CInt(varBookmark), STATE_NONE
            Set clsRowInfo = m_objGridMap.rowinfo(CInt(varBookmark))
        End If
        intErrors = intErrors + 1
        clsRowInfo.Error = True
        clsRowInfo.ErrorText = strError
        check_delete_error = True
    Else
        check_delete_error = False
    End If
End Function

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub

