VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEquipRateMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Equipment Rate grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "EquipmentRate" ' Identity of grid
' Names of some table this grid uses
Const EQUIP_TABLE = "Equipment"
Const EQUIPRATE_TABLE = "Equipment_rate"
Const EQUIP_UPDATE = "Equipment update"
Const EQUIPRATE_UPDATE = "Equipment_rate update"
Const PUBEQUIPRATEEX_TABLE = "Published_equipment_rate_exception"
Const PUBEQUIPRATEEX_UPDATE = "Published_equipment_rate_exception update"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean

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
    m_objGridMap.GRIDTYPE = m_GRIDTYPE
'    m_objGridMap.SetFrozenCols (3)
    
    ' Add the columns to the grid
    m_objGridMap.AddCol "Type", "type_code", EQUIP_TABLE, 500
    m_objGridMap.AddCol "Equip ID", "equip_id", EQUIP_TABLE, 1300
    m_objGridMap.AddCol "Alt Equip ID", "alt_equip_id", EQUIP_TABLE, 1300, False, False
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Tech Desc", "tech_desc", EQUIP_TABLE, 2000
    m_objGridMap.AddCol "Metric Tech Desc", "metric_tech_desc", EQUIP_TABLE, 2000, False, False
    m_objGridMap.AddCol "Contact", "contact_id", EQUIPRATE_TABLE, 1000
    m_objGridMap.AddCol "Opr Hrly", "operating_cost_hrly", EQUIPRATE_TABLE, 1000, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "Week Rent", "rent_per_week", EQUIPRATE_TABLE, 1000, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "Unit", "unit", EQUIP_TABLE, 1000
    m_objGridMap.AddCol "Metric Unit", "metric_unit", EQUIP_TABLE, 1000, False, False
    m_objGridMap.AddCol "Model", "model_name", EQUIP_TABLE, 1200
    m_objGridMap.AddCol "Info Source", "info_source_ref", EQUIPRATE_TABLE, 1000
    m_objGridMap.AddCol "Comment", "comment", EQUIPRATE_TABLE, 2000
    m_objGridMap.AddCol "Estimated", "estimated_ind", EQUIPRATE_TABLE, 1000
    m_objGridMap.AddCol "Factor", "factor_ind", EQUIPRATE_UPDATE, 1000, True, True, False
    m_objGridMap.AddCol "TRACES", "traces_ind", EQUIP_TABLE, 1000
    m_objGridMap.AddCol "Start Date", "start_date", EQUIPRATE_UPDATE, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Term Date", "term_date", vbNullString, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Update Date", "equiprate_last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "equiprate_last_update_person", vbNullString, 1500, True, True, False
    m_objGridMap.AddCol "equip_skey", "equip_skey", ALL_TABLES, 800, True, False, True
    m_objGridMap.AddCol "equiprate_last_update_id", "equiprate_last_update_id", EQUIPRATE_UPDATE, 800, True, False, False
    m_objGridMap.AddCol "equip_last_update_id", "equip_last_update_id", EQUIP_UPDATE, 800, True, False, False
    ' Split
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Op Cost", "operating_cost_hrly_x", PUBEQUIPRATEEX_TABLE, 1000, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "Day Rent", "rent_per_day_x", PUBEQUIPRATEEX_TABLE, 1000, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "Week Rent", "rent_per_week_x", PUBEQUIPRATEEX_TABLE, 1000, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "Month Rent", "rent_per_month_x", PUBEQUIPRATEEX_TABLE, 1000, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "Equip Cost", "crew_equip_cost_x", PUBEQUIPRATEEX_TABLE, 1000, False, True, True, FORMAT_PRICE
    m_objGridMap.AddCol "Metric Op Cost", "metric_operating_cost_hrly_x", PUBEQUIPRATEEX_TABLE, 1000, False, False, True, FORMAT_PRICE
    m_objGridMap.AddCol "Metric Day Rent", "metric_rent_per_day_x", PUBEQUIPRATEEX_TABLE, 1000, False, False, True, FORMAT_PRICE
    m_objGridMap.AddCol "Metric Week Rent", "metric_rent_per_week_x", PUBEQUIPRATEEX_TABLE, 1000, False, False, True, FORMAT_PRICE
    m_objGridMap.AddCol "Metric Month Rent", "metric_rent_per_month_x", PUBEQUIPRATEEX_TABLE, 1000, False, False, True, FORMAT_PRICE
    m_objGridMap.AddCol "Metric Equip Cost", "metric_crew_equip_cost_x", PUBEQUIPRATEEX_TABLE, 1000, False, False, True, FORMAT_PRICE
    m_objGridMap.AddCol "Percent", "pct_ind", PUBEQUIPRATEEX_TABLE, 1000
    m_objGridMap.AddCol "Start Date", "start_date_x", PUBEQUIPRATEEX_UPDATE, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Term Date", "term_date_x", vbNullString, 1200, True, True, False, FORMAT_DATE
    m_objGridMap.AddCol "Country", "country_code", PUBEQUIPRATEEX_TABLE, 1100, False, False, True
    m_objGridMap.AddCol "Region", "region_code", PUBEQUIPRATEEX_TABLE, 1100, False, False, True
    m_objGridMap.AddCol "Update Date", "equiprate_last_update_date_x", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "equiprate_last_update_person_x", vbNullString, 1500, True, True, False
    m_objGridMap.AddCol "equiprate_last_update_id_x", "equiprate_last_update_id_x", PUBEQUIPRATEEX_UPDATE, 800, True, False, False
    
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
        
        ' Unit drop downs
        bRet = g_objDAL.GetRecordset(vbNullString, "select unit from unit_of_measure order by unit", rec)
        If bRet Then
            While Not rec.EOF
                Item.Value = rec.Fields(0)
                TDBGrid.Columns("Unit").ValueItems.Add Item
                TDBGrid.Columns("Metric Unit").ValueItems.Add Item
                rec.MoveNext
            Wend
            TDBGrid.Columns("Unit").ValueItems.Presentation = dbgComboBox
            TDBGrid.Columns("Unit").ValueItems.Validate = True
            TDBGrid.Columns("Metric Unit").ValueItems.Presentation = dbgComboBox
            TDBGrid.Columns("Metric Unit").ValueItems.Validate = True
        End If
        rec.Close
        
        ' Country code drop down
        bRet = g_objDAL.GetRecordset(vbNullString, "select country_code from country order by country_code", rec)
        If bRet Then
            With TDBGrid.Columns("Country")
                While Not rec.EOF
                    Item.Value = rec.Fields(0)
                    .ValueItems.Add Item
                    rec.MoveNext
                Wend
                .ValueItems.Presentation = dbgComboBox
                .ValueItems.Validate = True
            End With
        End If
        rec.Close
        
        ' Region code drop down
        bRet = g_objDAL.GetRecordset(vbNullString, "select region_code from region order by region_code", rec)
        If bRet Then
            With TDBGrid.Columns("Region")
                While Not rec.EOF
                    Item.Value = rec.Fields(0)
                    .ValueItems.Add Item
                    rec.MoveNext
                Wend
                .ValueItems.Presentation = dbgComboBox
                .ValueItems.Validate = True
            End With
        End If
        
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
        
        TDBGrid.Columns("Tech Desc").FetchStyle = True
        
'*** APEX Migration Utility Code Change ***
'        Dim ShadedRow As TrueOleDBGrid60.Style
'*** APEX Migration Utility Code Change ***
'        Dim ShadedRow As TrueOleDBGrid70.Style
        Dim ShadedRow As TrueOleDBGrid80.Style
        ' Define a New Style that will be used within this Application
        Set ShadedRow = TDBGrid.Styles.Add("ShadedRow")
        ShadedRow.BackColor = vb3DDKShadow
    End If
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateEquip As Boolean
    Dim blnUpdateEquipRate As Boolean
    Dim blnUpdateEquipRateEx As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    
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
            blnUpdateEquip = False
            blnUpdateEquipRate = False
            blnUpdateEquipRateEx = False
            ' Loop through the fields to see what has changed
            For Each fld In m_rec.Fields
                ' If the value changed
                If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
                    ' See what table the field is from
                    Set objColDef = m_objGridMap.GetCol(fld.Name)
                    ' Mark the table we should update
                    If objColDef.Table = EQUIP_TABLE Then
                        blnUpdateEquip = True
                    ElseIf objColDef.Table = EQUIPRATE_TABLE Then
                        blnUpdateEquipRate = True
                    ElseIf objColDef.Table = PUBEQUIPRATEEX_TABLE Then
                        blnUpdateEquipRateEx = True
                    End If
                End If
            Next
            
            ' If the row is new
            If clsRowInfo.State = STATE_NEW Then
                ' Depends on type_code
                If m_rec.Fields("type_code").Value = "H" Then
                    strUpdate = "exec sp_insert_equipment "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    ' Reset on success
                    If blnReturn = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                        m_rec.Fields("equip_last_update_date").Value = Now
                        m_rec.Fields("equip_last_update_person").Value = strUserName
                        Reset_Orig_Values m_rec
                    Else
                        intErrors = intErrors + 1
                        clsRowInfo.Error = True
                        clsRowInfo.ErrorText = strError
                    End If
                ElseIf m_rec.Fields("type_code").Value = "M" Then
                    ' If skey doesn't exist, insert equip and rate
                    If m_rec.Fields("equip_skey") = 0 Or m_rec.Fields("equip_skey") = "" Or IsNull(m_rec.Fields("equip_skey")) Then
                        strUpdate = "exec sp_insert_equipment_and_rate "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIPRATE_TABLE
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_date").Value = Now
                            m_rec.Fields("equip_last_update_person").Value = strUserName
                            m_rec.Fields("equiprate_last_update_date").Value = Now
                            m_rec.Fields("equiprate_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' If skey does exist, update equip if needed and insert rate
                    Else
                        ' Insert equipment rate
                        strUpdate = "exec sp_insert_equipment_rate "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIPRATE_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_date").Value = Now
                            m_rec.Fields("equip_last_update_person").Value = strUserName
                            m_rec.Fields("equiprate_last_update_date").Value = Now
                            m_rec.Fields("equiprate_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    End If
                ElseIf m_rec.Fields("type_code").Value = "E" Then
                    ' If skey doesn't exist, insert equip and rate
                    If m_rec.Fields("equip_skey") = 0 Or m_rec.Fields("equip_skey") = "" Or IsNull(m_rec.Fields("equip_skey")) Then
                        strUpdate = "exec sp_insert_equipment_and_rate_x "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUBEQUIPRATEEX_TABLE
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                            m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_date").Value = Now
                            m_rec.Fields("equip_last_update_person").Value = strUserName
                            m_rec.Fields("equiprate_last_update_date_x").Value = Now
                            m_rec.Fields("equiprate_last_update_person_x").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' If skey does exist, update equip if needed and insert rate
                    Else
                        ' Insert equipment rate exception
                        blnReturn = False
                        strUpdate = "exec sp_insert_equipment_rate_x "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUBEQUIPRATEEX_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                            m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_date").Value = Now
                            m_rec.Fields("equip_last_update_person").Value = strUserName
                            m_rec.Fields("equiprate_last_update_date_x").Value = Now
                            m_rec.Fields("equiprate_last_update_person_x").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    End If
                End If
            ' If this row was modified
            ElseIf clsRowInfo.State = STATE_MODIFIED Then
                ' Depends on type_code
                If m_rec.Fields("type_code").Value = "M" Then
                    ' If type_code changed, insert the new rate
                    If Not m_rec.Fields("type_code").Value = m_rec.Fields("type_code").OriginalValue Then
                        ' Insert equip rate
                        strUpdate = "exec sp_insert_equipment_rate "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIPRATE_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_date").Value = Now
                            m_rec.Fields("equip_last_update_person").Value = strUserName
                            m_rec.Fields("equiprate_last_update_date").Value = Now
                            m_rec.Fields("equiprate_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' If type_code didn't change, update both if both changed
                    ElseIf blnUpdateEquipRate And blnUpdateEquip Then
                        ' Update equip rate
                        strUpdate = "exec sp_update_equipment_and_rate "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIPRATE_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIPRATE_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@old_contact_id='" + m_rec.Fields("contact_id").OriginalValue + "', "
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                            m_rec.Fields("equiprate_last_update_date").Value = Now
                            m_rec.Fields("equiprate_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' Update rate if rate changed
                    ElseIf blnUpdateEquipRate Then
                        ' Update equip rate
                        strUpdate = "exec sp_update_equipment_rate "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIPRATE_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIPRATE_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@old_contact_id='" + m_rec.Fields("contact_id").OriginalValue + "', "
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                            m_rec.Fields("equiprate_last_update_date").Value = Now
                            m_rec.Fields("equiprate_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' Update equipment if equipment changed
                    ElseIf blnUpdateEquip Then
                        strUpdate = "exec sp_update_equipment "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@full_update=0, "
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_date").Value = Now
                            m_rec.Fields("equip_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    End If
                ElseIf m_rec.Fields("type_code").Value = "E" Then
                    ' If type_code changed, insert new rate
                    If Not m_rec.Fields("type_code").Value = m_rec.Fields("type_code").OriginalValue Then
                        ' Insert equip rate ex
                        strUpdate = "exec sp_insert_equipment_rate_x "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUBEQUIPRATEEX_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                            m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_date").Value = Now
                            m_rec.Fields("equip_last_update_person").Value = strUserName
                            m_rec.Fields("equiprate_last_update_date_x").Value = Now
                            m_rec.Fields("equiprate_last_update_person_x").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' If type_code didn't change, update both if both changed
                    ElseIf blnUpdateEquipRateEx And blnUpdateEquip Then
                        ' Update equip rate ex
                        strUpdate = "exec sp_update_equipment_and_rate_x "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUBEQUIPRATEEX_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUBEQUIPRATEEX_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@old_region_code='" + m_rec.Fields("region_code").OriginalValue + "', "
                        strUpdate = strUpdate + "@old_country_code='" + m_rec.Fields("country_code").OriginalValue + "', "
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                            m_rec.Fields("equiprate_last_update_date_x").Value = Now
                            m_rec.Fields("equiprate_last_update_person_x").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' Update rate if rate changed
                    ElseIf blnUpdateEquipRateEx Then
                        ' Update equip rate ex
                        strUpdate = "exec sp_update_equipment_rate_x "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUBEQUIPRATEEX_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, PUBEQUIPRATEEX_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@old_region_code='" + m_rec.Fields("region_code").OriginalValue + "', "
                        strUpdate = strUpdate + "@old_country_code='" + m_rec.Fields("country_code").OriginalValue + "', "
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                            m_rec.Fields("equiprate_last_update_date_x").Value = Now
                            m_rec.Fields("equiprate_last_update_person_x").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
                    ' If Equipment changed
                    ElseIf blnUpdateEquip Then
                        strUpdate = "exec sp_update_equipment "
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_TABLE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, EQUIP_UPDATE
                        m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                        strUpdate = strUpdate + "@full_update=0, "
                        strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
                        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                        ' Reset on success
                        If blnReturn = True Then
                            intSuccess = intSuccess + 1
                            clsRowInfo.State = STATE_NONE
                            m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                            m_rec.Fields("equip_last_update_date").Value = Now
                            m_rec.Fields("equip_last_update_person").Value = strUserName
                            Reset_Orig_Values m_rec
                        Else
                            intErrors = intErrors + 1
                            clsRowInfo.Error = True
                            clsRowInfo.ErrorText = strError
                        End If
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
End Function

Private Sub Class_Initialize()
    m_objGridMap.m_colCloneCols.Add 1, "equip_skey"
    m_objGridMap.m_colCloneCols.Add 1, "equip_id"
    m_objGridMap.m_colCloneCols.Add 1, "alt_equip_id"
    m_objGridMap.m_colCloneCols.Add 1, "type_code"
    m_objGridMap.m_colCloneCols.Add 1, "book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "unit"
    m_objGridMap.m_colCloneCols.Add 1, "metric_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_tech_desc"
    m_objGridMap.m_colCloneCols.Add 1, "crew_equip_desc"
    m_objGridMap.m_colCloneCols.Add 1, "crew_equip_desc_plural"
    m_objGridMap.m_colCloneCols.Add 1, "index_code"
    m_objGridMap.m_colCloneCols.Add 1, "index_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_crew_equip_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_crew_equip_desc_plural"
    m_objGridMap.m_colCloneCols.Add 1, "metric_unit"
    m_objGridMap.m_colCloneCols.Add 1, "model_name"
    m_objGridMap.m_colCloneCols.Add 1, "traces_ind"
    m_objGridMap.m_colCloneCols.Add 1, "indent_code"
    m_objGridMap.m_colCloneCols.Add 1, "format_characters"
    m_objGridMap.m_colCloneCols.Add 1, "format_code"
    m_objGridMap.m_colCloneCols.Add 1, "graphic_ref_id"
    m_objGridMap.m_colCloneCols.Add 1, "table_ref_id"
    m_objGridMap.m_colCloneCols.Add 1, "table_ref_col"
    m_objGridMap.m_colCloneCols.Add 1, "change_notice_code"
    m_objGridMap.m_colCloneCols.Add 1, "equip_last_update_id"
    
    m_objGridMap.m_colCloneCols.Add 1, "contact_id"
    m_objGridMap.m_colCloneCols.Add 1, "rent_per_week"
    m_objGridMap.m_colCloneCols.Add 1, "operating_cost_hrly"
    m_objGridMap.m_colCloneCols.Add 1, "estimated_ind"
    m_objGridMap.m_colCloneCols.Add 1, "comment"
    m_objGridMap.m_colCloneCols.Add 1, "info_source_ref"
    
    m_objGridMap.m_colCloneCols.Add 1, "country_code"
    m_objGridMap.m_colCloneCols.Add 1, "region_code"
    m_objGridMap.m_colCloneCols.Add 1, "rent_per_day_x"
    m_objGridMap.m_colCloneCols.Add 1, "rent_per_week_x"
    m_objGridMap.m_colCloneCols.Add 1, "rent_per_month_x"
    m_objGridMap.m_colCloneCols.Add 1, "operating_cost_hrly_x"
    m_objGridMap.m_colCloneCols.Add 1, "metric_rent_per_day_x"
    m_objGridMap.m_colCloneCols.Add 1, "metric_rent_per_week_x"
    m_objGridMap.m_colCloneCols.Add 1, "metric_rent_per_month_x"
    m_objGridMap.m_colCloneCols.Add 1, "metric_operating_cost_hrly_x"
    m_objGridMap.m_colCloneCols.Add 1, "crew_equip_cost_x"
    m_objGridMap.m_colCloneCols.Add 1, "metric_crew_equip_cost_x"
    m_objGridMap.m_colCloneCols.Add 1, "pct_ind"
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    ' This causes FetchRowStyle to fire and color the columns appropriately
'    If TDBGrid.Columns(ColIndex).Caption = "Type" Then
'        Dim vntCurrBookmark As Variant
'        vntCurrBookmark = TDBGrid.Bookmark
'        TDBGrid.Bookmark = vntCurrBookmark
'    End If
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    On Error Resume Next
    
    Dim strType As String
    Dim objRowInfo As CRowInfo
    objRowInfo = Nothing
    Set objRowInfo = m_objGridMap.rowinfo(str(Bookmark))
    
    ' If the row is highlighted, then let it be
    If (Condition And dbgSelectedRow) = 8 Then
        Exit Sub
    End If
    ' This col is always white unless an error occurred
    ' NOTE: In this release of the grid, definied styles (like "Normal") cannot be applied to cells
    If TDBGrid.Columns(col).Caption = "Tech Desc" Or TDBGrid.Columns(col).Caption = "Metric Tech Desc" Then
        If Not objRowInfo Is Nothing Then
            If objRowInfo.Error = True Then
                CellStyle.BackColor = vbRed
            Else
                If Bookmark Mod 2 = 1 Then
                    CellStyle.BackColor = vbWindowBackground
                Else
                    CellStyle.BackColor = g_intAlternateRowColor
                End If
            End If
        Else
            If Bookmark Mod 2 = 1 Then
                CellStyle.BackColor = vbWindowBackground
            Else
                CellStyle.BackColor = g_intAlternateRowColor
            End If
        End If
    End If
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    Dim strType As String
    
    If IsNumeric(Bookmark) Then
        ' Based on Split and type_code value, make row black or leave regular
        m_rec.Bookmark = Bookmark
        strType = m_rec.Fields("type_code").Value
        If strType = "M" Then
            If Split = 1 Then
                RowStyle = "Normal"
            ElseIf Split = 2 Then
                RowStyle = "ShadedRow"
            End If
        ElseIf strType = "E" Then
            If Split = 1 Then
                RowStyle = "ShadedRow"
            ElseIf Split = 2 Then
                RowStyle = "Normal"
            End If
        ElseIf strType = "B" Or strType = "H" Then
            If Split = 1 Or Split = 2 Then
                RowStyle = "ShadedRow"
            End If
        End If
    End If
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
    If m_rec.Fields("type_code").Value = "M" Then
        strUpdate = "exec sp_delete_equipment_rate "
        strUpdate = strUpdate + "@equip_skey=" + str(m_rec.Fields("equip_skey")) + ","
        strUpdate = strUpdate + " @contact_id='" + m_rec.Fields("contact_id") + "',"
        strUpdate = strUpdate + " @start_date='" + Format(m_rec.Fields("start_date"), "mm/dd/yyyy") + "',"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    ElseIf m_rec.Fields("type_code").Value = "E" Then
        strUpdate = "exec sp_delete_equipment_rate_x "
        strUpdate = strUpdate + "@equip_skey=" + str(m_rec.Fields("equip_skey")) + ","
        strUpdate = strUpdate + " @country_code='" + m_rec.Fields("country_code") + "',"
        strUpdate = strUpdate + " @region_code='" + m_rec.Fields("region_code") + "',"
        strUpdate = strUpdate + " @start_date='" + Format(m_rec.Fields("start_date"), "mm/dd/yyyy") + "',"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    End If
    
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
    Exit Sub

NoDelete:
    MsgBox Err.Description
    Bookmark = Null
End Sub

' Called when a cell's data has changed as user moves to new row or presses Return
Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    ' For Equip ID column, fetch Equipment data
    If TDBGrid.Columns(ColIndex).Caption = "Equip ID" Then
        Dim strSelect As String
        Dim rec As New ADODB.RecordSet ' Recordset to hold query results
        Dim blnReturn As Boolean
        Dim vntMyBookmark As Variant
        Dim i As Integer

        ' Check to see if the equip_id entered exists already
        strSelect = "Select *, last_update_id as 'equip_last_update_id' from Equipment where equip_id='" + TDBGrid.Text + "'"
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
            m_rec.Fields("traces_ind") = False
            m_rec.Fields("estimated_ind") = False
            m_rec.Fields("factor_ind") = False
            m_rec.Fields("pct_ind") = False
            m_rec.Fields("equiprate_last_update_id") = 0
            m_rec.Fields("equiprate_last_update_id_x") = 0
            m_rec.Fields("country_code") = "USA"
            m_rec.Fields("region_code") = "NAT"
            m_rec.Fields("equiprate_last_update_id") = 0
            ' Set these because they are not shown by default and need values
            m_rec.Fields("metric_rent_per_day_x") = 0
            m_rec.Fields("metric_rent_per_week_x") = 0
            m_rec.Fields("metric_rent_per_month_x") = 0
            m_rec.Fields("metric_operating_cost_hrly_x") = 0
            m_rec.Fields("metric_crew_equip_cost_x") = 0
            
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
                m_rec.Fields("equip_skey") = 0
            End If
        End If
    ElseIf TDBGrid.Columns(ColIndex).Caption = "Week Rent" Then
        TDBGrid.Columns("factor_ind").Value = False ' Price was changed, so no longer factored
    ElseIf TDBGrid.Columns(ColIndex).Caption = "Opr Hrly" Then
        TDBGrid.Columns("factor_ind").Value = False ' Price was changed, so no longer factored
    End If
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("Estimated").Value = False
    TDBGrid.Columns("Percent").Value = False
    TDBGrid.Columns("Factor").Value = False
    TDBGrid.Columns("Country").Value = "USA"
    TDBGrid.Columns("Region").Value = "NAT"
    TDBGrid.Columns("equip_last_update_id") = 0
    TDBGrid.Columns("equiprate_last_update_id") = 0
    TDBGrid.Columns("equiprate_last_update_id_x") = 0
    ' Set these because they are not shown by default and need values.
    TDBGrid.Columns("metric_rent_per_day_x") = 0
    TDBGrid.Columns("metric_rent_per_week_x") = 0
    TDBGrid.Columns("metric_rent_per_month_x") = 0
    TDBGrid.Columns("metric_operating_cost_hrly_x") = 0
    TDBGrid.Columns("metric_crew_equip_cost_x") = 0
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

Public Sub Factor(dblFactor As Double, intApply As Integer)
    Dim vntBookmark As Variant
    
    For Each vntBookmark In TDBGrid.SelBookmarks
        m_rec.Bookmark = vntBookmark
        If intApply = EQUIP_FACTOR_RENT Or intApply = EQUIP_FACTOR_BOTH Then
            m_rec.Fields("Rent_per_week") = m_rec.Fields("Rent_per_week") + m_rec.Fields("Rent_per_week") * dblFactor / 100
        End If
        If intApply = EQUIP_FACTOR_OPERATING Or intApply = EQUIP_FACTOR_BOTH Then
            m_rec.Fields("Operating_cost_hrly") = m_rec.Fields("Operating_cost_hrly") + m_rec.Fields("Operating_cost_hrly") * dblFactor / 100
        End If
        m_rec.Fields("Factor_ind") = -1
        m_objGridMap.SetRowState Int(vntBookmark), STATE_MODIFIED
    Next
    vntBookmark = TDBGrid.SelBookmarks(0)
    TDBGrid.ReBind ' Reset grid contents
    TDBGrid.Bookmark = vntBookmark ' Set bookmark back again
End Sub

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub


