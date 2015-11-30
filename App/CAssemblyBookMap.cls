VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAssemblyBkMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class is "derived" from CGridMap
' It serves the Unit Cost grid
Option Explicit

'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to grid events
'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to grid events
Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid ' Access to grid events
Attribute TDBGrid.VB_VarHelpID = -1

Const m_GRIDTYPE = "AssemblyBook" ' Identity of grid
' Names of some table this grid uses
Const ASSEMBLY_BOOK_DETAIL_TABLE = "assembly_book_detail"
Const ASSEMBLY_DETAIL_TABLE = "assembly_detail"
Const ALL_TABLES = "All"

Private m_objGridMap As New CGridMap ' "Parent" class
Private m_rec As ADODB.RecordSet ' Recordset that populates grid
Private m_blnPreferences As Boolean
Private m_varSaveBookmark As Variant
Private m_varCurrentBookmark As Variant
Private bln_UseInd As Boolean
'*** APEX Migration Utility Code Change ***
'Public myTDBGrid As TrueOleDBGrid60.TDBGrid
'*** APEX Migration Utility Code Change ***
'Public myTDBGrid As TrueOleDBGrid70.TDBGrid
Public myTDBGrid As TrueOleDBGrid80.TDBGrid

Public Function Delete() As Boolean
Dim varSelectedBookmarks() As Variant
Dim varCurrentSelectedBookmark As Variant
Dim varButton
Dim blnResult As Boolean
Dim i As Integer
Dim i1 As Integer

'On Error Resume Next

If TDBGrid.SelBookmarks.Count > 0 Then
        If TDBGrid.SelBookmarks.Count > 0 Then
            Delete = m_objGridMap.Delete
        Else
            Delete = False
        End If
    Else
        Delete = m_objGridMap.Delete
    End If
End Function


Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub
Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
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
    m_objGridMap.AddCol "Type", "type_code", ASSEMBLY_BOOK_DETAIL_TABLE, 500, True
    m_objGridMap.AddCol "Book ID", "assembly_book_id", ASSEMBLY_BOOK_DETAIL_TABLE, 1200, True, True
    m_objGridMap.AddCol "Alt Book ID", "alt_assembly_book_id", ASSEMBLY_BOOK_DETAIL_TABLE, 1200, False, True
    ' Main Split
    m_objGridMap.AddSplit
    m_objGridMap.AddCol "Assembly ID", "assembly_id", ASSEMBLY_DETAIL_TABLE, 1200, False, True
    m_objGridMap.AddCol "Book Desc", "book_desc", ASSEMBLY_BOOK_DETAIL_TABLE, 2000
    m_objGridMap.AddCol "Comm'l Ind", "coml_ind", ASSEMBLY_BOOK_DETAIL_TABLE, 1000
    m_objGridMap.AddCol "Resi Ind", "resi_ind", ASSEMBLY_BOOK_DETAIL_TABLE, 1000
    
    m_objGridMap.AddCol "Metric Book Desc", "metric_book_desc", ASSEMBLY_BOOK_DETAIL_TABLE, 2000, , False
    m_objGridMap.AddCol "Unit", "unit", ASSEMBLY_BOOK_DETAIL_TABLE, 2000, False, False
    m_objGridMap.AddCol "Metric Unit", "metric_unit", ASSEMBLY_BOOK_DETAIL_TABLE, 2000, True, False
    m_objGridMap.AddCol "Section Head", "section_head_desc", ASSEMBLY_BOOK_DETAIL_TABLE, 3000, False, False, , , True
    m_objGridMap.AddCol "Metric Section Head", "metric_section_head_desc", ASSEMBLY_BOOK_DETAIL_TABLE, 3000, False, False, , , True
    m_objGridMap.AddCol "Index Desc", "index_desc", ASSEMBLY_BOOK_DETAIL_TABLE, 2000, False, False
    m_objGridMap.AddCol "Index Code", "index_code", ASSEMBLY_BOOK_DETAIL_TABLE, 1000, False, False, True, , True
    m_objGridMap.AddCol "Qty", "book_qty", ASSEMBLY_BOOK_DETAIL_TABLE, 1000, False, False
    m_objGridMap.AddCol "Labor Hours", "labor_hour", ASSEMBLY_BOOK_DETAIL_TABLE, 1000, False, False
    m_objGridMap.AddCol "Calc Factor", "calculation_factor", ASSEMBLY_BOOK_DETAIL_TABLE, 1000
    m_objGridMap.AddCol "Metric Calc Factor", "metric_calculation_factor", ASSEMBLY_BOOK_DETAIL_TABLE, 1000, False, True
    m_objGridMap.AddCol "Update Date", "last_update_date", vbNullString, 1650, True, True, False, FORMAT_DATETIME
    m_objGridMap.AddCol "Update Person", "last_update_person", vbNullString, 1500, True, True, False
    m_objGridMap.AddCol "assembly_skey", "assembly_skey", ALL_TABLES, 800, True, False, True
    m_objGridMap.AddCol "assembly_book_skey", "assembly_book_skey", ASSEMBLY_BOOK_DETAIL_TABLE, 800, True, False, True
    m_objGridMap.AddCol "last_update_id", "last_update_id", ASSEMBLY_BOOK_DETAIL_TABLE, 800, True, False, False
    m_objGridMap.AddCol "asbly last_update_id", "asbly_last_update_id", ASSEMBLY_DETAIL_TABLE, 800, True, False, False

    ' Call base class
    If m_blnPreferences Then
        m_objGridMap.BuildPreferenceGrid
    Else
        m_objGridMap.BuildGrid
    End If

    ' Give headings for the splits
    TDBGrid.Splits(0).Caption = "Heading"
    TDBGrid.Splits(0).Size = 4
    
    TDBGrid.Splits(1).Caption = "Book Detail"
    TDBGrid.Splits(1).Size = 5

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

        ' Type Code drop down
        Item.Value = "H"
        TDBGrid.Columns("Type").ValueItems.Add Item
        Item.Value = "S"
        TDBGrid.Columns("Type").ValueItems.Add Item
        Item.Value = "D"
        TDBGrid.Columns("Type").ValueItems.Add Item
        Item.Value = "F"
        TDBGrid.Columns("Type").ValueItems.Add Item
        Item.Value = "N"
        TDBGrid.Columns("Type").ValueItems.Add Item

        TDBGrid.Columns("Type").ValueItems.Presentation = dbgComboBox
        TDBGrid.Columns("Type").ValueItems.Validate = True
        TDBGrid.Columns("Type").AutoDropDown = True

        TDBGrid.Columns("Book Desc").FetchStyle = True
    End If
End Sub

Public Function Update() As Boolean
    Dim strUpdate As String ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateBookDetail As Boolean
    Dim objColDef As CColumnDef
    Dim fld As ADODB.Field
    Dim strError As String ' Error string returned from DAL
    Dim intErrors As Integer ' Tracks if any errors have occurred
    Dim intSuccess As Integer ' Tracks successful updates
    Dim intAssembly_Skey As Integer
    Dim validate_assembly_id As Boolean
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
            blnUpdateBookDetail = False
            ' Loop through the fields to see what has changed
            For Each fld In m_rec.Fields
                ' If the value changed
                If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
                    ' See what table the field is from
                    Set objColDef = m_objGridMap.GetCol(fld.Name)
                    ' Mark the table we should update
                    blnUpdateBookDetail = True
                End If
            Next

            ' If the row is new
            If clsRowInfo.State = STATE_NEW Then
            ' NOTE: Cannot do NEWs in the grid, launches single record view!
'            ' If this row was modified
            ElseIf clsRowInfo.State = STATE_MODIFIED Then

                If blnUpdateBookDetail Then
               Dim strLast_assembly_book_id As String

'    On Error Resume Next
'  If m_rec.Fields("assembly_id").OriginalValue <> m_rec.Fields("assembly_id").Value Then
'        Dim strSelect As String
'
'        Dim rec As New ADODB.RecordSet
'       strLast_assembly_book_id = m_rec.Fields("assembly_id").Value
'
'        ' Validate the entered assembly ID and retrieve the skey.
'        strSelect = "Select assembly_skey from assembly_detail where assembly_id='" + m_rec.Fields("assembly_id").Value + "'"
'        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
'        If rec.RecordCount > 0 Then
'            intAssembly_Skey = rec.GetString
'
'
'            validate_assembly_id = True
'        Else
'            MsgBox "Invalid Assembly Id."
'            validate_assembly_id = False
'        End If
'        rec.Close
'    Else
'        validate_assembly_id = True
'    End If

'               UpdateRecordsetFromForm Me, m_rec
'               m_rec.Fields("Assembly_skey").Value = intAssembly_Skey
            
               UpdateFormFromRecordset Me, m_rec
            'Update all types of rows.  The sp_update_book_detail_driver will update the appropriate records.
                    strUpdate = "exec sp_update_assembly_book_detail "
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ASSEMBLY_BOOK_DETAIL_TABLE
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ALL_TABLES
                    m_objGridMap.BuildUpdate strUpdate, m_rec.Bookmark, ASSEMBLY_DETAIL_TABLE
                    
                    strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                    
          
                  
                    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                 
                    ' Reset on success
                    If blnReturn = True Then
                        intSuccess = intSuccess + 1
                        clsRowInfo.State = STATE_NONE
                        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
                        m_rec.Fields("last_update_date").Value = Now
                        m_rec.Fields("last_update_person").Value = strUserName
                       
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

    RebindTDBGridNow
End Function
Private Sub Class_Initialize()
    m_objGridMap.m_colCloneCols.Add 1, "assembly_book_skey"
    m_objGridMap.m_colCloneCols.Add 1, "assembly_book_id"
    m_objGridMap.m_colCloneCols.Add 1, "alt_assembly_book_id"
     m_objGridMap.m_colCloneCols.Add 1, "assembly_id"
    m_objGridMap.m_colCloneCols.Add 1, "assembly_skey"
   
    m_objGridMap.m_colCloneCols.Add 1, "type_code"
    m_objGridMap.m_colCloneCols.Add 1, "unit"
    m_objGridMap.m_colCloneCols.Add 1, "metric_unit"
    m_objGridMap.m_colCloneCols.Add 1, "book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "metric_book_desc"
    m_objGridMap.m_colCloneCols.Add 1, "section_head_desc"
    m_objGridMap.m_colCloneCols.Add 1, "col_head_1"
    m_objGridMap.m_colCloneCols.Add 1, "col_head_2"
    m_objGridMap.m_colCloneCols.Add 1, "col_head_3"
    m_objGridMap.m_colCloneCols.Add 1, "col_head_4"
    m_objGridMap.m_colCloneCols.Add 1, "col_head_5"
    m_objGridMap.m_colCloneCols.Add 1, "col_head_6"
    m_objGridMap.m_colCloneCols.Add 1, "index_code"
    m_objGridMap.m_colCloneCols.Add 1, "index_desc"
    m_objGridMap.m_colCloneCols.Add 1, "book_qty"
    m_objGridMap.m_colCloneCols.Add 1, "labor_hour"
    m_objGridMap.m_colCloneCols.Add 1, "calculation_factor"
    m_objGridMap.m_colCloneCols.Add 1, "metric_calculation_factor"
    m_objGridMap.m_colCloneCols.Add 1, "asbly_last_update_id"
    
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    Select Case TDBGrid.Columns(ColIndex).Caption
        Case "Type"
            If TDBGrid.Text <> "H" Then
                TDBGrid.Columns("Section Head").Text = ""
                TDBGrid.Columns("Metric Section Head").Text = ""
            End If
        Case "Assembly ID"
            Dim strSelect As String
            Dim rec As New ADODB.RecordSet
            Dim blnReturn As Boolean
            ' Validate the entered assembly ID and retrieve the skey.
            strSelect = "Select assembly_skey, type_code, unit, metric_unit, book_desc, metric_book_desc, last_update_id from assembly_detail where assembly_id='" + TDBGrid.Columns("Assembly ID").Text + "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
            m_rec.Bookmark = TDBGrid.Bookmark
            If rec.RecordCount > 0 Then
                m_rec.Fields("assembly_skey") = rec.Fields("assembly_skey")
                m_rec.Fields("unit") = rec.Fields("unit")
                m_rec.Fields("metric_unit") = rec.Fields("metric_unit")
                If TDBGrid.Columns("Type").Value = "D" Or Len(TDBGrid.Columns("book_desc")) = 0 Then
                    m_rec.Fields("book_desc") = rec.Fields("book_desc")
                End If
                If TDBGrid.Columns("Type").Value = "D" Or Len(TDBGrid.Columns("metric_book_desc")) = 0 Then
                    m_rec.Fields("metric_book_desc") = rec.Fields("metric_book_desc")
                End If
                m_rec.Fields("asbly_last_update_id") = rec.Fields("last_update_id")
                If rec.Fields("type_code") = "M" Then
                    Dim rec2 As New ADODB.RecordSet
                    strSelect = "select labor_hour from published_assembly_cost where assembly_skey = " + CStr(m_rec.Fields("assembly_skey"))
                    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec2)
                    If Not rec2.EOF Then
                           m_rec.Fields("labor_hour") = rec2.Fields("labor_hour")
                    End If
                    rec2.Close
                    Set rec2 = Nothing
                End If
            Else
                m_rec.Fields("asbly_last_update_id") = 0
                m_rec.Fields("unit") = ""
                m_rec.Fields("metric_unit") = ""
                If m_rec.Fields("type_code") = "D" Then
                    m_rec.Fields("book_desc") = ""
                    m_rec.Fields("metric_book_desc") = ""
                End If
                m_rec.Fields("assembly_skey") = 0
                m_rec.Fields("labor_hour") = 0
            End If
            m_rec.Update
            TDBGrid.RefetchCol (TDBGrid.Columns("assembly_skey").ColIndex)
            TDBGrid.RefetchCol (TDBGrid.Columns("unit").ColIndex)
            TDBGrid.RefetchCol (TDBGrid.Columns("metric_unit").ColIndex)
            TDBGrid.RefetchCol (TDBGrid.Columns("book_desc").ColIndex)
            TDBGrid.RefetchCol (TDBGrid.Columns("metric_book_desc").ColIndex)
            TDBGrid.RefetchCol (TDBGrid.Columns("labor_hour").ColIndex)
            TDBGrid.RefetchCol (TDBGrid.Columns("asbly_last_update_id").ColIndex)
            rec.Close
            Set rec = Nothing
    End Select

End Sub

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim blnReturn As Boolean
If m_blnPreferences Then Exit Sub   'Don't validate if in preferences
    Select Case TDBGrid.Columns(ColIndex).Caption
        Case "Index Code"
            If Len(Trim(TDBGrid.Text)) > 0 Then
                If TDBGrid.Text <> "IX" And TDBGrid.Text <> "JX" Then
                    MsgBox "Please enter a valid Index Code (IX or JX)."
                    TDBGrid.Text = OldValue
                    Cancel = True
                End If
            End If
        Case "Assembly ID"
            If Len(Trim(TDBGrid.Text)) > 0 Then
                Dim strSelect As String
                Dim rec As New ADODB.RecordSet
                ' Validate the entered assembly ID and retrieve the skey.
                strSelect = "Select assembly_skey from assembly_detail where assembly_id='" + TDBGrid.Text + "'"
                blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
                If rec.RecordCount = 0 Then
                    MsgBox "Please enter a valid Assembly ID."
                    Cancel = True
                End If
                rec.Close
                Set rec = Nothing
            Else
                If TDBGrid.Columns("Type").Text = "D" Then
                    MsgBox "The assembly ID is required for D rows.  Please enter a valid Assembly ID."
                    Cancel = True
                End If
            End If
        End Select
    If Cancel = False Then
        TDBGrid.PostMsg 1   'Refetch index codes
    End If
    

End Sub

Private Sub TDBGrid_BeforeUpdate(Cancel As Integer)
If m_blnPreferences = False Then
    If TDBGrid.Columns("Comm'l Ind").Value = 0 And TDBGrid.Columns("Resi Ind").Value = 0 Then
                bln_UseInd = True
                Cancel = True
            Else
                bln_UseInd = False
    End If
End If
End Sub

Private Sub TDBGrid_Change()
Dim iSelStart As Integer
  
If TDBGrid.Columns(TDBGrid.col).Caption = "Index Code" Then
    If Len(TDBGrid.Text) > 0 Then
        iSelStart = TDBGrid.SelStart
        TDBGrid.Text = UCase(TDBGrid.Text)
        TDBGrid.SelStart = iSelStart
    End If
End If

End Sub


Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
If bln_UseInd = True Then
    MsgBox "Commercial or Residential Use indicator must be selected."
    TDBGrid.SetFocus
    bln_UseInd = False
End If
Response = 0

End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
    On Error Resume Next
    Dim objRowInfo As CRowInfo
    
    objRowInfo = Nothing
    Set objRowInfo = m_objGridMap.rowinfo(str(Bookmark))

    ' If the row is highlighted, then let it be
    If (Condition And dbgSelectedRow) = 8 Then
        Exit Sub
    End If
    
    If TDBGrid.Columns(col).Caption = "Section Head" Or "Metric Section Head" Then
'-------------CR#924-------------------------------------------
'        If TDBGrid.Columns("Type").CellText(Bookmark) = "D" Then
'            CellStyle.Locked = True
'            CellStyle.ForeColor = vbGrayText
'        End If
'---------------------------------------------------------------
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
        If TDBGrid.Columns(TDBGrid.col).DataField = "book_desc" Or _
            TDBGrid.Columns(TDBGrid.col).DataField = "metric_book_desc" Then
            If Len(TDBGrid.Text) + 1 > 75 Then
                KeyAscii = 0
            End If
        End If
    End If

End Sub

Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)
Dim blnReturn As Boolean
Dim strBookmark As String
Select Case MsgId
    Case 1
        TDBGrid.RefetchCol (TDBGrid.Columns("Section Head").ColIndex)
        TDBGrid.RefetchCol (TDBGrid.Columns("Metric Section Head").ColIndex)
End Select

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
            Exit Sub
        End If
    End If
    
    ' Build SQL statement
    strUpdate = "exec sp_delete_assembly_book_detail "
    strUpdate = strUpdate + "@assembly_book_skey='" + Trim(str(m_rec.Fields("assembly_book_skey"))) + "'"
    
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
    MsgBox Err.description
    Bookmark = Null
End Sub

Private Sub TDBGrid_OnAddNew()
    ' Defaults for new added row
    TDBGrid.Columns("last_update_id") = 0
    ' Set these because they are not shown by default and need values.
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

Public Sub RowWrap(intRowWrap As Integer)
    m_objGridMap.RowWrap (intRowWrap)
End Sub




Private Function validate_assembly_id() As Boolean
Dim strLast_assembly_book_id As String

    On Error Resume Next
  If m_rec.Fields("assembly_id").OriginalValue <> m_rec.Fields("assembly_id").Value Then
        Dim strSelect As String
        Dim blnReturn As Boolean
        Dim intAssembly_Skey As Integer
        Dim rec As New ADODB.RecordSet
       strLast_assembly_book_id = m_rec.Fields("assembly_id").Value

        ' Validate the entered assembly ID and retrieve the skey.
        strSelect = "Select assembly_skey from assembly_detail where assembly_id='" + m_rec.Fields("assembly_id").Value + "'"
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        If rec.RecordCount > 0 Then
            intAssembly_Skey = rec.GetString
        MsgBox "Assembly Skey is " + rec.GetString, vbOKOnly
            
            validate_assembly_id = True
        Else
            MsgBox "Invalid Assembly Id."
            validate_assembly_id = False
        End If
        rec.Close
    Else
        validate_assembly_id = True
    End If


End Function


Private Sub RebindTDBGridNow()
Dim oldRow As Variant

    On Error Resume Next
    oldRow = myTDBGrid.Bookmark
    myTDBGrid.Refresh
    myTDBGrid.Bookmark = oldRow
End Sub
