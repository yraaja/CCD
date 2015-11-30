VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGridMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' <modulename> CGridMap.cls</modulename>
' <functionname>General (Main) </functionname>
'
' <summary>
' Helper class module for any form utilizing the ComponentOne DataGrid control.
'It is the datagrid "factory".
'
'Key sub/functions:
'"   BuildGrid()
'       Dynamically builds datagrid using data from the REGISTRY and some table/column specific hard coded specifications
'"   AddCol()
'       Dynamically adds a column to the datagrid being built
'"   SetGrid()
'"   TDBGrid_BeforeColUpdate
'       BEFORE UPDATING CELL CONTENTS, STRIP OUT ANY NON-PRINT CHARACTERS
'       Set contents of datagrid columns based upon the known contents of specific datagrid columns  (e.g. given a "unit cost id" you can retrieve, "tech_desc", "metric_tech_desc", "unit", "unit_cost_skey" etal)
'"   RecordSet() - property  (set/let)
'       Serves or Sets the underlying recordset that drives the datagrid
'"   BuildUpdate()
'       This sub helps build the SQL string being constructed for an Update query looking at each column name and determining syntax associated with the data type  (i.e. single quotes around string data columns) and removing decimal pts. and blanks from specific field/columns for "unit cost ids"

' </summary>
' <seealso> Any form w/datagrid control </seealso>
'
' <datastruct> TDBGRID</datastruct>
' <datastruct> m_objGridMap</datastruct>
'<datastruct> m_colColumnDefs</datastruct>
'
'
' <storedprocedurename> n/a</storedprocedurename>
'
'<param name="data"> ???a dataset containing all the data for updating ?
'</param>
' <param name="someParameter">
'???? Description of someParameter goes here  updating
' </param>
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
'<code>
'</code>
'<code>
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'"   CGridMap.cls
'"   CCDdal.CRSMDataAccess
'Access to the DAL (data access layer dll) opened in MainModule_Main()
'"   TDBGrid (ComponentOne DataGrid)
'"   TrueOleDbGrid80  (C:\WINDOWS\system32\dodg8.oca)
'Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid  ' Access to Grid events
'"   REGISTRY
'</dependson>


'*** APEX Migration Utility Code Change ***
'Private WithEvents TDBGrid As TrueOleDBGrid60.TDBGrid ' Access to Grid events
'Private WithEvents TDBGrid As TrueOleDBGrid70.TDBGrid ' Access to Grid events

Private WithEvents TDBGrid As TrueOleDBGrid80.TDBGrid  ' Access to Grid events
Attribute TDBGrid.VB_VarHelpID = -1

Private m_rec As ADODB.RecordSet ' The recordset that populates this grid
Private m_intMaxCol As Integer ' Number of columns in the recordset
Private m_strGridType As String ' Names grid
Private m_intFrozenCols As Integer ' Number of columns in grid to be frozen during scrolling
Private m_colColumnDefs As New Collection ' Collection of CColumnDef objects describing grid columns
Private m_colSplits As New Collection ' Collection of locations for splits
Private m_colRowInfos As New Collection ' Collection of CRowInfo objects describing row state
Private m_intGridCols As Integer ' Number of columns in the grid
Public m_colCloneCols As New Collection ' List of fields to be cloned

Dim sngMouseUpX As Single
Dim sngMouseUpY As Single

Public Sub BuildUpdate_OrigValues(ByRef strUpdateClause As String, varCurrentRec As Variant, strTable As String)
    Dim objCol As CColumnDef
    Dim strObj_id As String
    
    m_rec.Bookmark = varCurrentRec
    'Build the update string fo the row, using original values
    'Key fields will not have original values.  Key fields for
    'Unit Cost Long Description are hard-coded, and should be a
    'parameter array if this is to be used for any other updates.
    
    For Each objCol In m_colColumnDefs
        ' Check if it should go into the SQL by looking at table name
        If objCol.Table = strTable Then
            With m_rec.Fields(objCol.TableCol)
                If IsNull(.Value) Then
                    strUpdateClause = strUpdateClause + "@" + .Name + "="
                    Select Case .Type
                    Case adChar
                        strUpdateClause = strUpdateClause + "''" + ", "
                    Case adVarChar
                        strUpdateClause = strUpdateClause + "''" + ", "
                    Case adInteger, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adTinyInt, adSmallInt
                        strUpdateClause = strUpdateClause + "0" + ", "
                    Case adDouble, adNumeric, adDecimal
                        strUpdateClause = strUpdateClause + "0" + ", "
                    Case adBoolean
                        strUpdateClause = strUpdateClause + str(0) + ", "
                    Case adDBTimeStamp
                        strUpdateClause = strUpdateClause + "''" + ", "
                    Case adCurrency
                        strUpdateClause = strUpdateClause + "0" + ", "
                    Case Else
                        MsgBox "Unknown field type encountered in CGridMap.BuildUpdate for " + .Name
                    End Select
                End If
                If Not IsNull(.Value) Then
                If .Name = "obj_id" Then strObj_id = .Value     '@#
                ' Different SQL based on type of field data
                Select Case .Type
                Case adChar
                    If .Name = "unit_cost_id" Or .Name = "ext_unit_cost_id" Or .Name = "mat_id" Or .Name = "assembly_id" Then
                        .Value = Compress_String(.Value)
                    End If
                    strUpdateClause = strUpdateClause + "@" + .Name + "='"
                    strUpdateClause = strUpdateClause + SQLFixString(.Value) + "', "
                Case adVarChar
                    If .Name = "unit_cost_id" Or .Name = "ext_unit_cost_id" Or .Name = "mat_id" Or .Name = "assembly_id" Then
                        .Value = Compress_String(.Value)
                    End If
                    strUpdateClause = strUpdateClause + "@" + .Name + "='"
                    strUpdateClause = strUpdateClause + SQLFixString(.Value) + "', "
                Case adInteger, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adTinyInt, adSmallInt
                    strUpdateClause = strUpdateClause + "@" + .Name + "="
                    strUpdateClause = strUpdateClause + str(.Value) + ", "
                Case adDouble, adNumeric, adDecimal
                    strUpdateClause = strUpdateClause + "@" + .Name + "="
                    strUpdateClause = strUpdateClause + str(.Value) + ", "
                Case adBoolean
                    strUpdateClause = strUpdateClause + "@" + .Name + "="
                    If .Value = True Then
                        strUpdateClause = strUpdateClause + str(1) + ", "
                    Else
                        strUpdateClause = strUpdateClause + str(0) + ", "
                    End If
                Case adDBTimeStamp
                    strUpdateClause = strUpdateClause + "@" + .Name + "='"
                    strUpdateClause = strUpdateClause + Format(.Value, "mm/dd/yyyy") + "', "
                Case adCurrency
                    strUpdateClause = strUpdateClause + "@" + .Name + "="
                    strUpdateClause = strUpdateClause + str(.Value) + ", "
                Case Else
                    MsgBox "Unknown field type encountered in CGridMap.BuildUpdate for " + .Name
                End Select
                End If
                'Do not pass original values for key fields
                If .Name <> "row_meas_sys_cd" And .Name <> "obj_skey" And .Name <> "obj_id" Then
                    If Not IsNull(.OriginalValue) Then
                    ' Different SQL based on type of field data
                        Select Case .Type
                        Case adChar
                            If .Name = "unit_cost_id" Or .Name = "ext_unit_cost_id" Or .Name = "mat_id" Or .Name = "assembly_id" Then
                                .Value = Compress_String(.Value)
                            End If
                            strUpdateClause = strUpdateClause + "@old_" + .Name + "='"
                            strUpdateClause = strUpdateClause + SQLFixString(.OriginalValue) + "', "
                        Case adVarChar
                            If .Name = "unit_cost_id" Or .Name = "ext_unit_cost_id" Or .Name = "mat_id" Or .Name = "assembly_id" Then
                                .OriginalValue = Compress_String(.OriginalValue)
                            End If
                            strUpdateClause = strUpdateClause + "@old_" + .Name + "='"
                            strUpdateClause = strUpdateClause + SQLFixString(.OriginalValue) + "', "
                        Case adInteger, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adTinyInt, adSmallInt
                            strUpdateClause = strUpdateClause + "@old_" + .Name + "="
                            strUpdateClause = strUpdateClause + str(.OriginalValue) + ", "
                        Case adDouble, adNumeric, adDecimal
                            strUpdateClause = strUpdateClause + "@" + .Name + "="
                            strUpdateClause = strUpdateClause + str(.OriginalValue) + ", "
                        Case adBoolean
                            strUpdateClause = strUpdateClause + "@old_" + .Name + "="
                            If .OriginalValue = True Then
                                strUpdateClause = strUpdateClause + str(1) + ", "
                            Else
                                strUpdateClause = strUpdateClause + str(0) + ", "
                            End If
                        Case adDBTimeStamp
                            strUpdateClause = strUpdateClause + "@" + .Name + "='"
                            strUpdateClause = strUpdateClause + Format(.Value, "mm/dd/yyyy") + "', "
                        Case adCurrency
                            strUpdateClause = strUpdateClause + "@old_" + .Name + "="
                            strUpdateClause = strUpdateClause + str(.OriginalValue) + ", "
                        Case Else
                            MsgBox "Unknown field type encountered in CLongDescBuildUpdate for " + .Name
                        End Select
                    End If
                    If IsNull(.OriginalValue) Then
                        strUpdateClause = strUpdateClause + "@old_" + .Name + "="
                        Select Case .Type
                        Case adChar
                            strUpdateClause = strUpdateClause + "''" + ", "
                        Case adVarChar
                            strUpdateClause = strUpdateClause + "''" + ", "
                        Case adInteger, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adTinyInt, adSmallInt
                            strUpdateClause = strUpdateClause + "0" + ", "
                        Case adDouble, adNumeric, adDecimal
                            strUpdateClause = strUpdateClause + "0" + ", "
                        Case adBoolean
                            strUpdateClause = strUpdateClause + str(0) + ", "
                        Case adDBTimeStamp
                            strUpdateClause = strUpdateClause + "''" + ", "
                        Case adCurrency
                            strUpdateClause = strUpdateClause + "0" + ", "
                        Case Else
                            MsgBox "Unknown field type encountered in CGridMap.BuildUpdate for " + .Name
                        End Select
                    End If
                End If
            End With
        End If
       
NextCol:
    Next
End Sub

Public Function Delete() As Boolean
Dim varSelectedBookmarks() As Variant
Dim varCurrentSelectedBookmark As Variant
Dim varButton
Dim blnResult As Boolean
Dim i As Integer
        If TDBGrid.SelBookmarks.Count > 0 Then
            varButton = MsgBox(CStr(TDBGrid.SelBookmarks.Count) + " records will be deleted.  Are you sure you want to delete?", vbYesNo + vbCritical)
            If varButton = vbYes Then
                ReDim varSelectedBookmarks(0 To TDBGrid.SelBookmarks.Count)
                i = 0
                For Each varCurrentSelectedBookmark In TDBGrid.SelBookmarks
                    varSelectedBookmarks(i) = varCurrentSelectedBookmark
                    i = i + 1
                Next varCurrentSelectedBookmark
                For i = 0 To UBound(varSelectedBookmarks) - 1
                    TDBGrid.Bookmark = varSelectedBookmarks(i)
                    If DeleteTDBGridRow() = True Then
                        Delete = True    'Set to true if one or more deletions succeed
                    End If
                Next i
            End If
        Else
            varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
            If varButton = vbYes Then
                Delete = DeleteTDBGridRow()
            End If
        End If
    If Delete = True Then
        DoEvents
        MsgBox "Delete successful."
    End If
End Function

Public Sub SetError(intBookmark As Integer, strError As String)
    On Error Resume Next
    Dim clsRowInfo As CRowInfo
    clsRowInfo = Nothing
    Set clsRowInfo = m_colRowInfos(str(intBookmark))
    If clsRowInfo Is Nothing Then
        Set clsRowInfo = New CRowInfo
        clsRowInfo.Bookmark = intBookmark
        clsRowInfo.Error = True
        clsRowInfo.ErrorText = strError
        m_colRowInfos.Add clsRowInfo, str(intBookmark)
    Else
        clsRowInfo.Error = True
        clsRowInfo.ErrorText = strError
    End If
End Sub

Private Function DeleteTDBGridRow() As Boolean
Dim clsRowInfo As CRowInfo
Dim varBookmark As Variant
On Error Resume Next
varBookmark = TDBGrid.Bookmark  'Save current

TDBGrid.Delete
Set clsRowInfo = Nothing
Set clsRowInfo = m_colRowInfos(str(varBookmark))
'clsRowInfo.Error = True
'clsRowInfo.ErrorText = "test"
    
If Not clsRowInfo Is Nothing Then
    If clsRowInfo.State = STATE_DELETED Then
        DeleteTDBGridRow = True
    End If
End If
    
Exit_Function:
Exit Function

Error_Processing:
MsgBox Error$
Resume 0

End Function

Public Property Get RowInfos() As Collection
    Set RowInfos = m_colRowInfos
End Property

Public Function IsPendingChange() As Boolean
    On Error GoTo Error_Processing
    Dim clsRowInfo As CRowInfo
    Dim i As Integer
    
    If TDBGrid.EditActive = True Or TDBGrid.DataChanged = True Then
        IsPendingChange = True
        Exit Function
    End If

'    TDBGrid.Update
    TDBGrid.MoveFirst
    
    ' If there are no RowInfos, then no pending changes
    If m_colRowInfos.Count = 0 Then
        IsPendingChange = False
        Exit Function
    End If
    
    For i = 1 To m_colRowInfos.Count
        Set clsRowInfo = m_colRowInfos(i)
        If clsRowInfo.State = STATE_NEW Or clsRowInfo.State = STATE_MODIFIED Then
            IsPendingChange = True
            Exit Function
        End If
    Next
    IsPendingChange = False

Exit_Sub:
Exit Function

Error_Processing:
If Err = 91 Then    'ignore if not set
    IsPendingChange = False
    Resume Exit_Sub
Else
    Resume Next
End If

End Function

Public Property Let RecordSet(RecSet As ADODB.RecordSet)
    On Error Resume Next
    Set m_rec = RecSet
    
'    If Not (m_rec.EOF And m_rec.BOF) Then m_rec.MoveLast
    m_rec.MoveLast
    
    ' Remove all existing items
    Dim i As Integer
    For i = 1 To m_colRowInfos.Count
        m_colRowInfos.Remove 1
    Next
    
    m_intMaxCol = m_rec.Fields.Count
    
End Property


Public Property Get RecordSet() As ADODB.RecordSet
    Set RecordSet = m_rec
End Property

' Return column with the given table column name
Public Function GetCol(strColName As String) As CColumnDef
    On Error Resume Next
    Set GetCol = m_colColumnDefs.Item(strColName)
End Function

Public Function RefreshRecord(intKey As Integer) As Boolean
'Refresh fields in the record from the stored sql
'If an error is encountered, refreshrecord will be false
Dim rec As ADODB.RecordSet
Dim fldsNew As ADODB.Fields
Dim fldNew As ADODB.Field
Dim objRowInfo As CRowInfo
Dim strOrigBookmark As String
Dim strSelect As String
Dim blnReturn As Boolean

    On Error Resume Next
    strOrigBookmark = m_rec.Bookmark
    Set objRowInfo = Nothing
    Set objRowInfo = m_colRowInfos(intKey)
    If Not objRowInfo Is Nothing Then
        strSelect = objRowInfo.RowSQL
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        If blnReturn = True Then
    ' If it does, copy that data into grid
            If rec.RecordCount = 1 Then
                Set fldsNew = rec.Fields
                m_rec.Bookmark = objRowInfo.Bookmark
                For Each fldNew In fldsNew
                    m_rec.Fields(fldNew.Name).Value = fldNew.Value
                Next
                rec.Close
                Set rec = Nothing
                RefreshRecord = True
            Else
                RefreshRecord = False
            End If
        Else
            RefreshRecord = False
        End If
    Else
        RefreshRecord = False
    End If
End Function

Public Sub SelectAllRows()
    'On Error Resume Next
    If Not (m_rec.BOF And m_rec.EOF) Then
        m_rec.MoveFirst
        Do While Not m_rec.EOF
            TDBGrid.SelBookmarks.Add m_rec.Bookmark
            m_rec.MoveNext
        Loop
    End If
End Sub

'*** APEX Migration Utility Code Change ***
'Public Sub SetGrid(grd As TrueOleDBGrid60.TDBGrid)
'*** APEX Migration Utility Code Change ***
'Public Sub SetGrid(grd As TrueOleDBGrid70.TDBGrid)
Public Sub SetGrid(grd As TrueOleDBGrid80.TDBGrid)
    Set TDBGrid = grd
End Sub

Public Sub RowWrap(intRowWrap As Integer)
    Dim intCount As Integer
'*** APEX Migration Utility Code Change ***
'    Dim col As TrueOleDBGrid60.Column
'*** APEX Migration Utility Code Change ***
'    Dim Col As TrueOleDBGrid70.Column
    Dim Col As TrueOleDBGrid80.Column
    ' Row Wrap on
    If intRowWrap = 1 Then
        TDBGrid.Splits(0).SizeMode = dbgScalable
        TDBGrid.Splits(0).ScrollGroup = 2 ' Need this to cause vertical scroll bar to appear
        TDBGrid.MultipleLines = 1
        If TDBGrid.Splits.Count > 1 Then
            For intCount = 1 To TDBGrid.Splits.Count - 1
                TDBGrid.Splits(intCount).Size = 0
            Next
            For Each Col In TDBGrid.Splits(0).Columns
                If m_colColumnDefs(Col.DataField).Visible = True Then
                    Col.Visible = True
                End If
            Next
        End If
    Else
    ' Row Wrap off
        TDBGrid.Splits(0).ScrollGroup = 1
        TDBGrid.MultipleLines = 0
        If TDBGrid.Splits.Count > 1 Then
            TDBGrid.Splits(0).SizeMode = dbgNumberOfColumns
            'TDBGrid.Splits(0).Size = m_intFrozenCols
            TDBGrid.Splits(0).Size = m_colSplits(1)
            For intCount = 1 To TDBGrid.Splits.Count - 1
                TDBGrid.Splits(intCount).Size = 1
            Next
            'For intCount = m_intFrozenCols To TDBGrid.Splits(0).Columns.Count - 1
            For intCount = m_colSplits(1) To TDBGrid.Splits(0).Columns.Count - 1
                TDBGrid.Splits(0).Columns(intCount).Visible = False
            Next
        End If
    End If
End Sub

Public Sub BuildGrid_old()
'*** APEX Migration Utility Code Change ***
'    Dim Cols As TrueOleDBGrid60.Columns
'*** APEX Migration Utility Code Change ***
'    Dim Cols As TrueOleDBGrid70.Columns
    Dim Cols As TrueOleDBGrid80.Columns
'*** APEX Migration Utility Code Change ***
'    Dim col As TrueOleDBGrid60.Column
'*** APEX Migration Utility Code Change ***
'    Dim Col As TrueOleDBGrid70.Column
    Dim Col As TrueOleDBGrid80.Column
    Dim Flds As ADODB.Fields
    Dim i As Long
'*** APEX Migration Utility Code Change ***
'    Dim ErrorRow As TrueOleDBGrid60.Style
'*** APEX Migration Utility Code Change ***
'    Dim ErrorRow As TrueOleDBGrid70.Style
    Dim ErrorRow As TrueOleDBGrid80.Style
            
    ' Setup some grid features
    TDBGrid.AlternatingRowStyle = True
    TDBGrid.EvenRowStyle.BackColor = RGB(230, 227, 223) ' Make even rows grey
    TDBGrid.ScrollBars = dbgAutomatic
    TDBGrid.TabAcrossSplits = True
    TDBGrid.TabAction = 1 ' Tab moves from column to column
    TDBGrid.AnchorRightColumn = True
    TDBGrid.FetchRowStyle = True
    TDBGrid.AllowColMove = True
          
    ' Define a New Style that will be used within this Application
    Set ErrorRow = TDBGrid.Styles.Add("ErrorRow")
    ErrorRow.BackColor = vbRed
    
    ' Get column order and visibility from Registry
    ' The index corresponds to m_colNames and the data is the column order
    For i = 1 To m_intGridCols
        With m_colColumnDefs(i)
            Dim strKey As String
            Dim lRet As Long
            Dim hKey As Long
            ' Build key to open
            strKey = CCD_KEY + "\" + m_strGridType + "\" + .TableCol ' .Name
            lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
            If lRet <> 0 Then
                lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
                lRet = RegSetValueExLong(hKey, "Order", 0&, REG_DWORD, i - 1, 4)
                .Order = i - 1
                lRet = RegSetValueExLong(hKey, "Visible", 0&, REG_DWORD, .Visible, 4)
            Else
                Dim lValue As Long
                Dim lSize As Long
                lSize = 4
                lRet = RegQueryValueExLong(hKey, "Order", 0&, REG_DWORD, lValue, lSize)
                .Order = lValue
                lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
                .Visible = lValue
            End If
        End With
    Next i

    ' Remove what is there
    While TDBGrid.Columns.Count
        TDBGrid.Columns.Remove 0
    Wend

    Set Cols = TDBGrid.Columns
    ' Add columns to grid
    For i = 1 To m_intGridCols
        Set Col = Cols.Add(i - 1)
        With m_colColumnDefs(i)
            Col.Caption = .Name
            'Col.Order = .Order
            Col.Visible = .Visible
            Col.DataField = .TableCol
            Col.Locked = .Locked
            If Not .DropDown = vbNullString Then
                Col.DropDown = .DropDown
            End If
            If (.ColWidth > 0) Then
                Col.Width = .ColWidth
            End If
            ' Use check-boxes where necessary
            If Right(.TableCol, 4) = "_ind" Then
                Col.ValueItems.Presentation = 4
            End If
        End With
    Next i
    ' Set order of columns, can't do above
    For i = 1 To m_intGridCols
        Cols(i - 1).Order = m_colColumnDefs(i).Order
    Next i
    
    ' Set-up splits
    If m_intFrozenCols > 0 Then
'*** APEX Migration Utility Code Change ***
'        Dim S As TrueOleDBGrid60.Split
'*** APEX Migration Utility Code Change ***
'        Dim S As TrueOleDBGrid70.Split
        Dim S As TrueOleDBGrid80.Split
        Set S = TDBGrid.Splits.Add(0)
'*** APEX Migration Utility Code Change ***
'        Dim C As TrueOleDBGrid60.Column
'*** APEX Migration Utility Code Change ***
'        Dim C As TrueOleDBGrid70.Column
        Dim c As TrueOleDBGrid80.Column
        Set Cols = TDBGrid.Splits(0).Columns
        ' Make left side split columns un-sizable
        For Each c In Cols
            c.AllowSizing = False
        Next c
        ' Hide all columns that shouldn't appear on left side
        For i = m_intFrozenCols To Cols.Count - 1
            Cols(i).Visible = False
        Next
        With TDBGrid.Splits(0)
            .SizeMode = dbgNumberOfColumns
            .Size = m_intFrozenCols
            .AllowSizing = False ' Split isn't sizeable
        End With
        ' Right Split
        TDBGrid.Splits(1).RecordSelectors = False
        TDBGrid.Splits(1).AllowSizing = False
        Set Cols = TDBGrid.Splits(1).Columns
        ' Second splits hides first column
        For i = 0 To m_intFrozenCols - 1
            Cols(i).Visible = False
            Cols(i).AllowSizing = False
        Next
    End If
    
    ' Reinitialise the grid
    TDBGrid.Bookmark = Null
    TDBGrid.HoldFields
    TDBGrid.ReBind
    
    ' Callibrate VScroll bar
'    TDBGrid.ApproxCount = m_rec.RecordCount

End Sub

Public Sub BuildGrid(Optional bAllowSplitSizing As Boolean, Optional bDoNotUseRegistry As Boolean)
'*** APEX Migration Utility Code Change ***
'    Dim Cols As TrueOleDBGrid70.Columns
    Dim Cols As TrueOleDBGrid80.Columns
'*** APEX Migration Utility Code Change ***
'    Dim Col As TrueOleDBGrid70.Column
    Dim Col As TrueOleDBGrid80.Column
    Dim Flds As ADODB.Fields
    Dim i As Long
'*** APEX Migration Utility Code Change ***
'    Dim ErrorRow As TrueOleDBGrid70.Style
    Dim ErrorRow As TrueOleDBGrid80.Style
'*** APEX Migration Utility Code Change ***
'    Dim ActiveRow As TrueOleDBGrid70.Style
    Dim ActiveRow As TrueOleDBGrid80.Style
'*** APEX Migration Utility Code Change ***
'    Dim ActiveRow2 As TrueOleDBGrid70.Style
    Dim ActiveRow2 As TrueOleDBGrid80.Style
'*** APEX Migration Utility Code Change ***
'    Dim LockedStyle As TrueOleDBGrid70.Style
    Dim LockedStyle As TrueOleDBGrid80.Style
    Dim strKey As String
    Dim lRet As Long
    Dim hKey As Long
    Dim lValue As Long
    
    Dim VisibleVal As Long

    ' Setup some grid features
    TDBGrid.Appearance = dbgXPTheme ' Confirm to themes on XP/standard 3D on 2000
    TDBGrid.DeadAreaBackColor = vbApplicationWorkspace
    TDBGrid.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid.AlternatingRowStyle = False
    'if the preference is for alternating background, set it now
    strKey = CCD_KEY + "\" + m_strGridType
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    If lRet <> 0 Then
        '8/8/2005 RTD
        'No preferences in registry, default to alternating colors
        TDBGrid.AlternatingRowStyle = True
        TDBGrid.OddRowStyle.BackColor = vbWindowBackground
        TDBGrid.EvenRowStyle.BackColor = g_intAlternateRowColor ' Make even rows 'gray'
    Else
        lRet = RegQueryValueExLong(hKey, "Background", 0&, REG_DWORD, lValue, 4)
        If lValue = 0 Then
            TDBGrid.AlternatingRowStyle = True
            'TDBGrid.EvenRowStyle.BackColor = RGB(230, 227, 223) ' Make even rows gray
            TDBGrid.EvenRowStyle.BackColor = g_intAlternateRowColor  ' Make even rows gray
        Else
            TDBGrid.EvenRowStyle.BackColor = vbWindowBackground
        End If
    End If

    TDBGrid.ScrollBars = dbgAutomatic
    TDBGrid.TabAcrossSplits = True
    TDBGrid.TabAction = dbgGridNavigation ' Tab moves from column to column
    TDBGrid.WrapCellPointer = True ' Wrap from end of row to beginning of next row
    TDBGrid.AnchorRightColumn = True
    TDBGrid.FetchRowStyle = True
    TDBGrid.AllowColMove = True
'    TDBGrid.Styles("Normal").WrapText = True
        
    ' Define a New Style that will be used within this Application
    Set ErrorRow = TDBGrid.Styles.Add("ErrorRow")
    ErrorRow.BackColor = vbRed
    
    Set LockedStyle = TDBGrid.Styles.Add("LockedStyle")
    LockedStyle.ForeColor = vbGrayText
    
    Set ActiveRow = TDBGrid.Styles.Add("ActiveRow")
    ActiveRow.BackColor = &HFFFFC0
'    Set ActiveRow2 = TDBGrid.Styles.Add("ActiveRow2")
'    ActiveRow2.BackColor = &HC0C000

    ' Get column order and visibility from Registry
    ' The index corresponds to m_colNames and the data is the column order
    
    'If m_strGridType = "Output" Then Stop       'rlh temporary for debug only!
    
    For i = 1 To m_intGridCols
        With m_colColumnDefs(i)
            If bDoNotUseRegistry Then
                .Order = i - 1
            Else
                '
                '   Build key to open
                strKey = CCD_KEY + "\" + m_strGridType + "\" + .TableCol ' .Name
                
                If m_strGridType = "Output" Then
'                    Debug.Print vbCrLf & "-----------------------"
'                    Debug.Print m_colColumnDefs(I).Name & ": " & strKey 'rlh 06/10/08
                End If
                
                lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
                If lRet <> 0 Then
                    'INITIALIZE (SETUP IN REGISTRY) NEXT COLUMN
                    lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
                    lRet = RegSetValueExLong(hKey, "Order", 0&, REG_DWORD, i - 1, 4)
                    .Order = i - 1
'                    If m_strGridType = "Output" Then Stop       'rlh temporary for debug only!
'
                    If m_strGridType = "Assembly_Usage" Then
                            '########################### rlh 06/11/2010 CCD 8.6 #2 (Assembly Usage Button) #############
                            'PROBLEM: Visible property not being set and RETAINED THROUGH THIS CODE
                            '         You have to MANUALLY go in and set to &HFFFFFFFF   !!!
                            '         This has to be resolved here!!!
                            '########################### rlh 06/17/2010 CCD 8.5 #2 #####################################
                            VisibleVal = &HFFFFFFFF
                            If lValue = 0 Then
                                .Visible = VisibleVal 'rlh 6/11/2010
                                lRet = RegSetValueExLong(hKey, "Visible", 0&, REG_DWORD, VisibleVal, 4)
                            Else
                                .Visible = lValue
                                lRet = RegSetValueExLong(hKey, "Visible", 0&, REG_DWORD, .Visible, 4)
                            End If
                            '############################ rlh - end of changes ###############################################################
                    End If
                    If m_strGridType <> "Assembly_Usage_1" Then
                        lRet = RegSetValueExLong(hKey, "Visible", 0&, REG_DWORD, .Visible, 4)
                    End If
                    lRet = RegSetValueExLong(hKey, "Width", 0&, REG_DWORD, .ColWidth, 4)
                Else
                    Dim lSize As Long
                    lSize = 4
                    '(rlh) Added on 6/10/2008 - Re "Green" output indicator request (K. Rodriguez)
                    Select Case m_strGridType
                    Case "Output"
                        Select Case .TableCol
                         Case "ext_indicator"       'Mst Fmt
                            lRet = RegQueryValueExLong(hKey, "Order", 0&, REG_DWORD, lValue, lSize)
                            .Order = lValue
                            .Order = 5
                            'If m_strGridType = "Output" Then Stop       'rlh temporary for debug only!
                            'Debug.Print ".Order: " & .Order
                            lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
                            .Visible = lValue
                             'Debug.Print "Visible: " & .Visible
                            lRet = RegQueryValueExLong(hKey, "Width", 0&, REG_DWORD, lValue, lSize)
                            .ColWidth = 1000   'rlh
                        Case "tag_code"             'Green
                            lRet = RegQueryValueExLong(hKey, "Order", 0&, REG_DWORD, lValue, lSize)
                            .Order = lValue
                            .Order = 10
                            'If m_strGridType = "Output" Then Stop       'rlh temporary for debug only!
                            'Debug.Print ".Order: " & .Order
                            lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
                            .Visible = lValue
                             'Debug.Print "Visible: " & .Visible
                            lRet = RegQueryValueExLong(hKey, "Width", 0&, REG_DWORD, lValue, lSize)
                            .ColWidth = 650     'RLH
                         Case "graphic_ref_id"
                            lRet = RegQueryValueExLong(hKey, "Order", 0&, REG_DWORD, lValue, lSize)
                            .Order = lValue
                            .Order = 13
                            .Order = 15
                            'If m_strGridType = "Output" Then Stop       'rlh temporary for debug only!
                            'Debug.Print ".Order: " & .Order
                            lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
                            .Visible = lValue
                             'Debug.Print "Visible: " & .Visible
                            lRet = RegQueryValueExLong(hKey, "Width", 0&, REG_DWORD, lValue, lSize)
                            .ColWidth = lValue
                         Case "table_ref_id"
                            lRet = RegQueryValueExLong(hKey, "Order", 0&, REG_DWORD, lValue, lSize)
                            .Order = lValue
                            .Order = 14
                            .Order = 16
                            'If m_strGridType = "Output" Then Stop       'rlh temporary for debug only!
                            'Debug.Print ".Order: " & .Order
                            lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
                            .Visible = lValue
                             Debug.Print "Visible: " & .Visible
                            lRet = RegQueryValueExLong(hKey, "Width", 0&, REG_DWORD, lValue, lSize)
                            .ColWidth = lValue
                        Case Else
                            lRet = RegQueryValueExLong(hKey, "Order", 0&, REG_DWORD, lValue, lSize)
                            .Order = lValue
                            'Debug.Print ".Order: " & .Order
                            lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
                            .Visible = lValue
                             'Debug.Print "Visible: " & .Visible
                            lRet = RegQueryValueExLong(hKey, "Width", 0&, REG_DWORD, lValue, lSize)
                            
                            .ColWidth = lValue
                           
                        End Select
                    
                    Case Else
                        lRet = RegQueryValueExLong(hKey, "Order", 0&, REG_DWORD, lValue, lSize)
                        .Order = lValue
                        Debug.Print ".Order: " & .Order
                        lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
                        
                        If m_strGridType = "Assembly_Usage" Then
                            '########################### rlh 06/11/2010 CCD 8.6 #2 (Assembly Usage Button) #############
                            'PROBLEM: Visible property not being set and RETAINED THROUGH THIS CODE
                            '         You have to MANUALLY go in and set to &HFFFFFFFF   !!!
                            '         This has to be resolved here!!!
                            '########################### rlh 06/17/2010 CCD 8.5 #2 #####################################
                            VisibleVal = &HFFFFFFFF
                            If lValue = 0 Then
                                .Visible = VisibleVal 'rlh 6/11/2010
                                'lRet = RegSetValueExLong(hKey, "Visible", 0&, REG_DWORD, VisibleVal, 4)
                            Else
                                .Visible = lValue
                                'lRet = RegSetValueExLong(hKey, "Visible", 0&, REG_DWORD, VisibleVal, 4)
                            End If
                            '############################ rlh - end of changes ###############################################################
                        Else
                            .Visible = lValue
                        End If
                        
                        
                        lRet = RegQueryValueExLong(hKey, "Width", 0&, REG_DWORD, lValue, lSize)
                        .ColWidth = lValue
                    
                    End Select      'rlh 06/10/08
                End If
            End If
        End With
    Next i
    '
    '   Remove what is there
    While TDBGrid.Columns.Count
        TDBGrid.Columns.Remove 0
    Wend

    Set Cols = TDBGrid.Columns
    ' Add columns to grid
    Dim j As Integer
    For i = 1 To m_intGridCols
        'Add the columns in the correct sequence - can't refresh order
'        J = 1
'        Do While (m_colColumnDefs(J).Order <> i - 1) And J < m_colColumnDefs.Count
'            J = J + 1
'        Loop
        Set Col = Cols.Add(i - 1)
        With m_colColumnDefs(i)
            Col.Caption = .Name
            'Col.Order = .Order
            Col.Visible = .Visible
            Col.DataField = .TableCol
            Col.Locked = .Locked
            Col.FetchStyle = .FetchStyle
            If (Col.Locked = True) And (.Format <> FORMAT_STRING_URL) Then
                Col.AddCellStyle 0, LockedStyle
                Col.AddCellStyle 0, LockedStyle
                Col.AddCellStyle 1, LockedStyle
                Col.AddCellStyle 4, LockedStyle
            End If
            If Not .DropDown = vbNullString Then
                Col.DropDown = .DropDown
            End If
            If (.ColWidth > 0) Then
                Col.Width = .ColWidth
            End If
            Select Case .Format
            Case FORMAT_PRICE, FORMAT_DECIMAL
                Col.Alignment = dbgRight
            Case FORMAT_STRING_URL
                ' ADDED 6/17/2005 RTD FOR VERSION 7.4.0 CR #97
                ' CELL APPEARANCE IS URL STANDARD
                Col.ForeColor = vbBlue
                Col.Font.Underline = True
                'Col.AllowFocus = False
            Case FORMAT_CHECK_BOX
                ' ADDED 8/10/2005 RTD for VERSION 7.5.0
                ' FORCE CELL APPEARANCE TO BE A CHECK BOX
                Col.ValueItems.Presentation = dbgCheckBox
                Col.Alignment = dbgCenter
            End Select
            ' Use check-boxes where necessary
            If Right(.TableCol, 4) = "_ind" Then
                Col.ValueItems.Presentation = 4
            End If
        End With
    Next i
    ' Set order of columns, can't do above
    ' MODIFIED 6/16/2005 RTD FOR VERSION 7.4.0 (CR#1529)
    ' CORRECT RUN-TIME ERROR 6145
    For i = 1 To m_intGridCols
        If m_colColumnDefs(i).Order <= Cols.Count Then
            Cols(i - 1).Order = m_colColumnDefs(i).Order
        Else
            Cols(i - 1).Order = Cols.Count
        End If
    Next i
    
    ' Set-up splits
    Dim intCounter As Integer
    Dim intCols As Integer
    Dim intThisSplit As Integer
'*** APEX Migration Utility Code Change ***
'    Dim S As TrueOleDBGrid60.Split
'*** APEX Migration Utility Code Change ***
'    Dim S As TrueOleDBGrid70.Split
    Dim S As TrueOleDBGrid80.Split
    For intCounter = 1 To m_colSplits.Count
        intThisSplit = m_colSplits(intCounter)
        ' Add the new split
        Set S = TDBGrid.Splits.Add(intCounter - 1)
        ' Turn of record selectors for subsequent splits
        TDBGrid.Splits(intCounter).RecordSelectors = False
        TDBGrid.Splits(intCounter).AllowSizing = True
        ' Hide all previous columns for new split
        Set Cols = TDBGrid.Splits(intCounter).Columns
        For intCols = 0 To intThisSplit - 1
            Cols(intCols).Visible = False
            Cols(intCols).AllowSizing = False
        Next
        ' Hide all next columns for old split
        Set Cols = TDBGrid.Splits(intCounter - 1).Columns
        'Set locked style for visible columns
        For intCols = 0 To intThisSplit - 1
            If Cols(intCols).Locked = True Then
                Cols(intCols).AddCellStyle 0, LockedStyle
                Cols(intCols).AddCellStyle 1, LockedStyle
                Cols(intCols).AddCellStyle 4, LockedStyle
            End If
        Next
        For intCols = intThisSplit To TDBGrid.Columns.Count - 1
            Cols(intCols).Visible = False
            Cols(intCols).AllowSizing = False
        Next
    Next
    
    If Not bAllowSplitSizing Then
        ' If there are splits, make the left-most split un-sizeable
        If m_colSplits.Count > 0 Then
            With TDBGrid.Splits(0)
                .SizeMode = dbgNumberOfColumns
                .Size = m_colSplits(1)
                .AllowSizing = False ' Split isn't sizeable
                Set Cols = .Columns
                For intCols = 0 To m_colSplits(1) - 1
                    Cols(intCols).AllowSizing = False
                Next
                .AllowRowSelect = True
                .AllowColSelect = False
                .SelectedStyle = .EvenRowStyle
            End With
        End If
    Else
        If m_colSplits.Count > 0 Then
            With TDBGrid.Splits(0)
                .SizeMode = dbgNumberOfColumns
                .Size = m_colSplits(1)
                .AllowSizing = True ' Split is sizeable
                Set Cols = .Columns
                For intCols = 0 To m_colSplits(1) - 1
                    Cols(intCols).AllowSizing = True
                Next
                .AllowRowSelect = True
                .AllowColSelect = True
                .SelectedStyle = .EvenRowStyle
            End With
        End If
    End If
    '
    ' Reinitialise the grid
    TDBGrid.Bookmark = Null
    TDBGrid.HoldFields
    TDBGrid.ReBind
    
    ' Callibrate VScroll bar
'    TDBGrid.ApproxCount = m_rec.RecordCount

End Sub

Public Sub BuildPreferenceGrid()
'*** APEX Migration Utility Code Change ***
'    Dim Cols As TrueOleDBGrid60.Columns
'*** APEX Migration Utility Code Change ***
'    Dim Cols As TrueOleDBGrid70.Columns
    Dim Cols As TrueOleDBGrid80.Columns
'*** APEX Migration Utility Code Change ***
'    Dim col As TrueOleDBGrid60.Column
'*** APEX Migration Utility Code Change ***
'    Dim Col As TrueOleDBGrid70.Column
    Dim Col As TrueOleDBGrid80.Column
    Dim Flds As ADODB.Fields
    Dim i As Long
'*** APEX Migration Utility Code Change ***
'    Dim ErrorRow As TrueOleDBGrid60.Style
'*** APEX Migration Utility Code Change ***
'    Dim ErrorRow As TrueOleDBGrid70.Style
    Dim ErrorRow As TrueOleDBGrid80.Style
    ' Setup some grid features
    TDBGrid.Appearance = dbgXPTheme ' Confirm to themes on XP/standard 3D on 2000
    TDBGrid.DeadAreaBackColor = vbApplicationWorkspace
    TDBGrid.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid.EvenRowStyle.BackColor = g_intAlternateRowColor ' Make even rows 'gray'
    TDBGrid.TabAcrossSplits = True
    TDBGrid.TabAction = 1  ' Tab moves from column to column
    TDBGrid.AnchorRightColumn = True
    TDBGrid.AllowColMove = True
    ' Get column order and visibility from Registry
    ' The index corresponds to m_colNames and the data is the column order
    For i = 1 To m_intGridCols
        With m_colColumnDefs(i)
            ' For now, default column order and all visible
            'If .Visible = True then
                Dim strKey As String
                Dim lRet As Long
                Dim hKey As Long
                ' Build key to open
                'code change by Mohan on Jan 09,2012: changing string to constant "RSMeans\CCD" to CCD_KEY
                strKey = CCD_KEY + "\" + m_strGridType + "\" + .TableCol ' .Name
                lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
                If lRet <> 0 Then
                    lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
                    lRet = RegSetValueExLong(hKey, "Order", 0&, REG_DWORD, i - 1, 4)
                    lRet = RegSetValueExLong(hKey, "Visible", 0&, REG_DWORD, .Visible, 4)
                    lRet = RegSetValueExLong(hKey, "Width", 0&, REG_DWORD, .ColWidth, 4)
                Else
                    Dim lValue As Long
                    Dim lSize As Long
                    lSize = 4
                    lRet = RegQueryValueExLong(hKey, "Order", 0&, REG_DWORD, lValue, lSize)
                    .Order = lValue
                    lRet = RegQueryValueExLong(hKey, "Visible", 0&, REG_DWORD, lValue, lSize)
                    .Visible = lValue
                    lRet = RegQueryValueExLong(hKey, "Width", 0&, REG_DWORD, lValue, lSize)
                    If lValue > 0 Then
                        .ColWidth = lValue
                    End If
                End If
            '.Order = i
                'End if
            'End if
        End With
    Next i
    ' Remove what is there
    While TDBGrid.Columns.Count
        TDBGrid.Columns.Remove 0
    Wend
    Set Cols = TDBGrid.Columns
    ' Add columns to grid
    Dim j As Integer
    For i = 1 To m_intGridCols
        'Add the columns in the correct sequence - can't refresh order
        'Revert to original logic to set order - may have been registry data error
'        J = 1
'        Do While m_colColumnDefs(J).Order <> i - 1
'            J = J + 1
'        Loop
        Set Col = Cols.Add(i - 1)
        With m_colColumnDefs(i)
            Col.Caption = .Name
            ' Don't show columns that are used behind the scenes
            If .ColWidth > 0 Then
                Col.Visible = True
            End If
            Col.DataField = .TableCol
            'Col.Locked = .Locked
            If (.ColWidth >= 0) Then
                Col.Width = .ColWidth
            End If
            ' Use check-boxes for all
            Col.ValueItems.Presentation = 4
        End With
    Next i
    Debug.Print "In Set Order Loop"
    ' Set order of columns, can't do above
    For i = 1 To m_intGridCols
    
        If m_colColumnDefs(i).Order > Cols.Count Then
            Cols(i - 1).Order = Cols.Count
        Else
            Cols(i - 1).Order = m_colColumnDefs(i).Order
        End If
        
        Debug.Print "DataField[" & i & "] = " & Cols(i - 1).DataField & ", Order = " & Cols(i - 1).Order & _
            ", Visible = " & Cols(i - 1).Visible & ", Width = " & Cols(i - 1).Width
    Next i
    ' Set-up splits
    Dim intCounter As Integer
    Dim intCols As Integer
    Dim intThisSplit As Integer
'*** APEX Migration Utility Code Change ***
'    Dim S As TrueOleDBGrid60.Split
'*** APEX Migration Utility Code Change ***
'    Dim S As TrueOleDBGrid70.Split
    Dim S As TrueOleDBGrid80.Split
    Debug.Print "Before splits, m_colSplits.Count = " & m_colSplits.Count & _
        ", TDBGrid.Splits = " & TDBGrid.Splits.Count
        
    For intCounter = 1 To m_colSplits.Count
        intThisSplit = m_colSplits(intCounter)
        ' Add the new split
        Set S = TDBGrid.Splits.Add(intCounter - 1)
        
        Debug.Print "In splits loop after add, m_colSplits.Count = " & m_colSplits.Count & _
        ", TDBGrid.Splits = " & TDBGrid.Splits.Count
        ' Turn of record selectors for subsequent splits
        TDBGrid.Splits(intCounter).RecordSelectors = False
        TDBGrid.Splits(intCounter).AllowSizing = True
        ' Hide all previous columns for new split
        Set Cols = TDBGrid.Splits(intCounter).Columns
        For intCols = 0 To intThisSplit - 1
            Cols(intCols).Visible = False
            Cols(intCols).AllowSizing = False
        Next
        ' Hide all next columns for old split
        Set Cols = TDBGrid.Splits(intCounter - 1).Columns
        For intCols = intThisSplit To TDBGrid.Columns.Count - 1
            Cols(intCols).Visible = False
            Cols(intCols).AllowSizing = False
        Next
    Next
    ' If there are splits, make the left-most split un-sizeable
    If m_colSplits.Count > 0 Then
        With TDBGrid.Splits(0)
            .SizeMode = dbgNumberOfColumns
            .Size = m_colSplits(1)
            .AllowSizing = False ' Split isn't sizeable
            Set Cols = .Columns
            For intCols = 0 To m_colSplits(1) - 1
                Cols(intCols).AllowSizing = False
            Next
        End With
    End If
    ' Reinitialise the grid
    'TDBGrid.Bookmark = Null
    TDBGrid.HoldFields
    TDBGrid.ReBind
    
End Sub

Private Sub Class_Initialize()
    m_intGridCols = 0
    m_intFrozenCols = 0
    m_intMaxCol = 0
    Set m_rec = Nothing
End Sub

Public Sub BuildUpdate(ByRef strUpdateClause As String, varCurrentRec As Variant, strTable As String, _
                       Optional excludeList As Collection)
    Dim objCol As CColumnDef
    m_rec.Bookmark = varCurrentRec
    
    For Each objCol In m_colColumnDefs
        ' Check if it should go into the SQL by looking at table name
        If objCol.Table = strTable Then
            ' Process it unless it is supposed to be excluded.
            If (Not ExistsInCollection(excludeList, m_rec.Fields(objCol.TableCol).Name)) Then
            
                With m_rec.Fields(objCol.TableCol)
                    If IsNull(.Value) Then
                        strUpdateClause = strUpdateClause + "@" + .Name + "="
                        Select Case .Type
                        Case adChar
                            strUpdateClause = strUpdateClause + "''" + ", "
                        Case adVarChar
                            strUpdateClause = strUpdateClause + "''" + ", "
                        Case adInteger, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adTinyInt, adSmallInt
                            strUpdateClause = strUpdateClause + "0" + ", "
                        Case adDouble, adNumeric, adDecimal
                            strUpdateClause = strUpdateClause + "0" + ", "
                        Case adBoolean
                            strUpdateClause = strUpdateClause + str(0) + ", "
                        Case adDBTimeStamp
                            strUpdateClause = strUpdateClause + "''" + ", "
                        Case adCurrency
                            strUpdateClause = strUpdateClause + "0" + ", "
                        Case Else
                            MsgBox "Unknown field type encountered in CGridMap.BuildUpdate for " + .Name
                        End Select
    
                        GoTo NextCol ' Skip the rest and move on
                    End If
                    ' Different SQL based on type of field data
                    Select Case .Type
                    Case adChar
                        If .Name = "unit_cost_id" Or .Name = "ext_unit_cost_id" Or .Name = "mat_id" Or .Name = "assembly_id" Then
                            .Value = Compress_String(.Value)
                        End If
                        strUpdateClause = strUpdateClause + "@" + .Name + "='"
                        strUpdateClause = strUpdateClause + SQLFixString(.Value) + "', "
                    Case adVarChar
                        If .Name = "unit_cost_id" Or .Name = "ext_unit_cost_id" Or .Name = "mat_id" Or .Name = "assembly_id" Then
                            .Value = Compress_String(.Value)
                        End If
                        strUpdateClause = strUpdateClause + "@" + .Name + "='"
                        strUpdateClause = strUpdateClause + SQLFixString(.Value) + "', "
                    Case adInteger, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adTinyInt, adSmallInt
                        strUpdateClause = strUpdateClause + "@" + .Name + "="
                        strUpdateClause = strUpdateClause + str(.Value) + ", "
                    Case adDouble, adNumeric, adDecimal
                        Select Case .Name   'rlh 03/22/2010
                        Case "PREVIOUS_2_rent_per_week", "PREVIOUS_3_rent_per_week", "PREVIOUS_4_rent_per_week", "PREVIOUS_5_rent_per_week"
                        
                        Case Else
                            strUpdateClause = strUpdateClause + "@" + .Name + "="
                            strUpdateClause = strUpdateClause + str(.Value) + ", "
                        End Select
                    Case adBoolean
                        strUpdateClause = strUpdateClause + "@" + .Name + "="
                        If .Value = True Then
                            strUpdateClause = strUpdateClause + str(1) + ", "
                        Else
                            strUpdateClause = strUpdateClause + str(0) + ", "
                        End If
                    Case adDBTimeStamp
                        strUpdateClause = strUpdateClause + "@" + .Name + "='"
                        strUpdateClause = strUpdateClause + Format(.Value, "mm/dd/yyyy") + "', "
                    Case adCurrency
                        strUpdateClause = strUpdateClause + "@" + .Name + "="
                        strUpdateClause = strUpdateClause + str(.Value) + ", "
                    Case Else
                        MsgBox "Unknown field type encountered in CGridMap.BuildUpdate for " + .Name
                    End Select
                End With
                
            End If  ' if not exists in collection
        End If  ' if col.table = table
NextCol:
    Next    ' For each column
End Sub

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    ' BEFORE UPDATING CELL CONTENTS, STRIP OUT ANY NON-PRINT CHARACTERS
    ' ADDED 5/26/2005 RTD, FOR CR#1554 ET AL.

    If (InStr(TDBGrid.Columns(ColIndex).Value, vbTab) > 0) Then
        ' text includes a TAB character, replace with a space
        TDBGrid.Columns(ColIndex).Value = Replace(TDBGrid.Columns(ColIndex).Value, vbTab, " ")
    End If
    If InStr(TDBGrid.Columns(ColIndex).Value, vbCr) > 0 Then
        ' text includes a CR character, replace with nothing
        TDBGrid.Columns(ColIndex).Value = Replace(TDBGrid.Columns(ColIndex).Value, vbCr, "")
    End If
    If InStr(TDBGrid.Columns(ColIndex).Value, vbLf) > 0 Then
        ' text includes a LF character, replace with nothing
        TDBGrid.Columns(ColIndex).Value = Replace(TDBGrid.Columns(ColIndex).Value, vbLf, "")
    End If
    
    '######################## rlh ########################
    '####               Get MF2004 ID  rlh added 05/07/08
    '######################## rlh ########################
'''    'Stop 'rlh
'''    If MASTER_FORMAT_ASSEMBLIES = 2004 Then
'''        'added this block on 05/06/08
'''        Dim strSELECT As String
'''        Dim rec As ADODB.RecordSet
'''        Dim blnReturn As Boolean
'''
'''        'rlh - commented out on 12/11/08 after John C. "override descr" failure
'''        'strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey from unit_cost_detail where unit_cost_skey = (select unit_cost_skey from unit_cost_detail_ext where unit_cost_id='" + Compress_String(TDBGrid.Text) + " ')"
'''
'''        'RLH - 12/11/08  - to address John C. "override descr" failure
'''         strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey from unit_cost_detail where unit_cost_skey = (select unit_cost_skey from unit_cost_detail_ext where unit_cost_id='" + Compress_String(TDBGrid.Columns(0)) + " ')"
'''
'''        ' Use DAL to perform select
'''        blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
'''        ' If found, copy that data into grid
'''        If rec.RecordCount > 0 Then
'''            ' Reset 1995/2004 unit_cost_id fields in the grid
'''            TDBGrid.Columns("unit_cost_id") = rec.Fields("unit_cost_id").Value
'''        End If
'''    End If
'''    '#################  down to here ######################   'rlh

'(rlh) ----  block below added on 02/17/2009

Dim strSelect As String
Dim rec As ADODB.RecordSet
Dim blnReturn  As Boolean
Dim i As Integer
'Dim strOrigBookmark As Variant

On Error GoTo ERRLBL        'rlh

Select Case TDBGrid.Columns(ColIndex).Caption
        Case "Unit Cost ID"      'Unit Cost ID
            strSelect = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey " + _
                    "from unit_cost_detail where unit_cost_id='" + Compress_String(TDBGrid.Text) + "'"
                     
            
                ' Use DAL to perform select
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
            ' If found, copy that data into grid
            If rec.RecordCount > 0 Then
                ' Copy all values
                TDBGrid.Columns("UC Tech Desc") = rec.Fields("tech_desc").Value
                TDBGrid.Columns("UC Met Desc") = rec.Fields("metric_tech_desc").Value
                TDBGrid.Columns("Unit") = rec.Fields("unit").Value
                TDBGrid.Columns("unit_cost_skey") = rec.Fields("unit_cost_skey").Value
            End If
            
            If MASTER_FORMAT_ASSEMBLIES = 2004 Then
                '######################## rlh ########################
                '####               Get MF2004 ID
                '######################## rlh ########################
                
                '' added this block on 05/06/08
                strSelect = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey from unit_cost_detail where unit_cost_skey = (select unit_cost_skey from unit_cost_detail_ext where unit_cost_id='" + Compress_String(TDBGrid.Text) + " ')"
                'strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit_cost_skey from unit_cost_detail_ext where unit_cost_skey =" & TDBGrid.Columns("unit_cost_skey")
                ' Use DAL to perform select
                blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
                ' If found, copy that data into grid
                If rec.RecordCount > 0 Then
                    ' Copy all values
                    TDBGrid.Columns("UC Tech Desc") = rec.Fields("tech_desc").Value
                    TDBGrid.Columns("UC Met Desc") = rec.Fields("metric_tech_desc").Value
                    'TDBGrid.Columns("Unit") = rec.Fields("unit").Value
                    TDBGrid.Columns("unit_cost_skey") = rec.Fields("unit_cost_skey").Value
                End If
            End If
            '' down to here on 05/06/08
            
            If rec.State = adStateOpen Then
                rec.Close
            End If
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)

        Case "Unit Cost ID 04"      'Ext Unit Cost ID  (mf04)
            strSelect = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey " + _
                    "from unit_cost_detail where unit_cost_id='" + TDBGrid.Columns("Unit Cost ID 04") + "'"
            
            'rlh 02/05/2009
            strSelect = "Select unit_cost_id, tech_desc, metric_tech_desc, unit_cost_skey " + _
                    "from unit_cost_detail_ext where unit_cost_id='" + Compress_String(TDBGrid.Columns("Unit Cost ID 04")) + "'"
            
            ' Use DAL to perform select
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
            ' If found, copy that data into grid
            If rec.RecordCount > 0 Then
                ' Copy all values
                TDBGrid.Columns("UC Tech Desc") = rec.Fields("tech_desc").Value
                TDBGrid.Columns("UC Met Desc") = rec.Fields("metric_tech_desc").Value
                'TDBGrid.Columns("Unit") = rec.Fields("unit").Value
                TDBGrid.Columns("unit_cost_skey") = rec.Fields("unit_cost_skey").Value
            End If
            
            If MASTER_FORMAT_ASSEMBLIES = 2004 Then
                '######################## rlh ########################
                '####               Get MF2004 ID
                '######################## rlh ########################
                
                '' added this block on 05/06/08
                'strSELECT = "Select unit_cost_id, tech_desc, metric_tech_desc, unit, unit_cost_skey from unit_cost_detail where unit_cost_skey = (select unit_cost_skey from unit_cost_detail_ext where unit_cost_id='" + Compress_String(TDBGrid.Text) + " ')"
               
                strSelect = "Select unit_cost_id, tech_desc, metric_tech_desc,  unit_cost_skey from unit_cost_detail where unit_cost_skey =" & TDBGrid.Columns("unit_cost_skey")
                
                'rlh 02/05/2009 - Get "Unit" only!!!  (we already have the others...)
                strSelect = "Select unit from unit_cost_detail where unit_cost_skey =" & TDBGrid.Columns("unit_cost_skey")
                
                ' Use DAL to perform select
                blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
                ' If found, copy that data into grid
                If rec.RecordCount > 0 Then
                    ' Copy "Unit" value
                    TDBGrid.Columns("Unit") = rec.Fields("unit").Value
                End If
            End If
            '' down to here on 05/06/08
            
            If rec.State = adStateOpen Then
                rec.Close
            End If
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)

    End Select
    
'    If Not (m_rec.BOF = True Or m_rec.EOF = True) Then
'        m_rec.Bookmark = strOrigBookmark
'    End If

    Exit Sub
ERRLBL: 'rlh 05/06/2008
    MsgBox ("(CGridMap):TDBGRID_BeforeColUpdate: " & Err.Description)
'    Resume
'    Stop 'rlh
    
    
End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
    
    If DataError = 6159 And _
        TDBGrid.SelBookmarks.Count > 0 And TDBGrid.Name <> "TDBGridAdds" _
        And TDBGrid.Name <> "TDBGridAssembly" Then
        Delete
        Response = 0
    End If

End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid80.StyleDisp)
    On Error GoTo Skip
    If m_colRowInfos(str(Bookmark)).Error = True Then
        RowStyle = "ErrorRow"
    End If
    Exit Sub
Skip:
End Sub

Private Sub TDBGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCtrl                     As Boolean
    Dim varSelectedBookmarks()      As Variant
    Dim varCurrentSelectedBookmark  As Variant
    Dim varButton
    Dim blnResult                   As Boolean
    Dim i                           As Integer
    
    blnCtrl = Shift And vbCtrlMask
    '
    '   They can't delete Summary Estimate records on the model form
    '   or Buildings.  And we want to use custom deleting for the other grids.
        If TDBGrid.Name = "TDBGridBuilding" Or TDBGrid.Name = "TDBGridModels" _
        Or TDBGrid.Name = "TDBGridSummaryEstimate" Or TDBGrid.Name = "TDBGridAdds" _
        Or TDBGrid.Name = "TDBGridAssembly" Then
    Else
        '
        '   Intercept the delete key - prompt user.
        If KeyCode = 46 Then
            If TDBGrid.SelBookmarks.Count >= 1 Then ' if multiple records are selected, the error procedure traps and processes them.
                varButton = MsgBox(CStr(TDBGrid.SelBookmarks.Count) + " records will be deleted.  Are you sure you want to delete this row(s) permanently?", vbYesNo + vbCritical)
                If varButton <> vbYes Then
                    KeyCode = 0
                End If
'
'5/17/02 ad this doesn't work for any of the grids, you have to select the row to get it to delete
'so this prompts is useless.
'            ElseIf TDBGrid.SelBookmarks.Count = 0 Then
'                varButton = MsgBox("Are you sure you want to delete this row(s) permanently?", vbYesNo + vbCritical)
'                If varButton <> vbYes Then
'                    KeyCode = 0
'                End If
           End If
        End If
    End If
    If blnCtrl And KeyCode = vbKeyX Then
        KeyCode = 0
    End If
End Sub

Private Sub TDBGrid_LostFocus()
On Error Resume Next
'TDBGrid.Update
End Sub

Private Sub TDBGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ADDED 6/17/2005 RTD FOR VERSION 7.4.0 CR #97
' IF MOUSE IS OVER A FORMAT_STRING_URL COLUMN, DISPLAY THE 'HAND' CURSOR
    Dim bOverCell As Boolean
    Dim r As Long, c As Long
    Dim sCellValue As String
    Dim vBookmark As Variant
    Dim bIsUrl As Boolean
    
    bOverCell = TDBGrid.CellContaining(X, Y, r, c)
    If bOverCell Then
        If (m_colColumnDefs(c + 1).Format = FORMAT_STRING_URL) And (r >= 0) Then
            ' hovering over a URL column
            ' if cell at r,c is not empty then change cursor to Hand
            vBookmark = TDBGrid.RowBookmark(r)
            sCellValue = TDBGrid.Columns(c).CellText(vBookmark)
            If sCellValue <> "" Then
                TDBGrid.MousePointer = dbgMPHand
                Status "Click to visit '" & sCellValue & "'..."
                bIsUrl = True
            End If
        End If
    End If
    If Not bIsUrl And (TDBGrid.MousePointer <> dbgMPDefault) Then
        ' We are not over a URL column
        TDBGrid.MousePointer = dbgMPDefault
        Status ""
    End If
    
End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    sngMouseUpX = X
    sngMouseUpY = Y
    TDBGrid.PostMsg tdbgENABLE_SORT
    
    'CHECK IF URL WAS CLICKED
    Dim bOverCell As Boolean
    Dim r As Long, c As Long
    Dim sCellValue As String
    Dim vBookmark As Variant
    Dim bIsUrl As Boolean
    Dim hWnd As Long
    
    bOverCell = TDBGrid.CellContaining(X, Y, r, c)
    If bOverCell Then
        If (m_colColumnDefs(c + 1).Format = FORMAT_STRING_URL) And (r >= 0) Then
            ' hovering over a URL column
            ' if cell at r,c is not empty then change cursor to Hand
            vBookmark = TDBGrid.RowBookmark(r)
            sCellValue = TDBGrid.Columns(c).CellText(vBookmark)
            If sCellValue <> "" Then
                If Left(LCase(sCellValue), 7) <> "http://" Then
                    sCellValue = "http://" & sCellValue
                End If
                hWnd = ShellExecute(0&, vbNullString, sCellValue, vbNullString, vbNullString, SW_SHOWNORMAL)
                Call BringWindowToTop(hWnd)
            End If
        End If
    End If

End Sub

Private Sub TDBGrid_PostEvent(ByVal MsgId As Integer)

    If MsgId = tdbgENABLE_SORT Then
        SetMenuBar
    End If

End Sub

Public Sub SetMenuBar()
Dim iCurSplit As Integer
Dim i As Integer

On Error Resume Next
If Not TDBGrid Is Nothing Then
    If Not m_rec Is Nothing Then
        iCurSplit = TDBGrid.SplitContaining(sngMouseUpX, sngMouseUpY)
        If iCurSplit <> -1 Then
            fMainForm.tbToolBar.Buttons.Item(tbrSORT_DESCENDING).Enabled = False
            fMainForm.tbToolBar.Buttons.Item(tbrSORT_ASCENDING).Enabled = False
            If TDBGrid.Splits.Count > 0 Then
                If m_rec.RecordCount > 0 Then
                    If TDBGrid.Splits(iCurSplit).SelStartCol <> -1 Then
                        fMainForm.tbToolBar.Buttons.Item(tbrSORT_DESCENDING).Enabled = True
                        fMainForm.tbToolBar.Buttons.Item(tbrSORT_ASCENDING).Enabled = True
                        For i = 0 To TDBGrid.Splits.Count - 1
                            If i <> iCurSplit Then
                                TDBGrid.Splits(i).ClearSelCols
                            End If
                        Next i
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Public Sub Sort(intDir As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim i As Integer
    Dim iCurSplit As Integer
    Dim blnfound As Boolean
    Dim intSelColCount As Integer
    Dim intCurrentCol As Integer
    Dim strSort As String

If IsNumeric(TDBGrid.FirstRow) Then
    intRow = TDBGrid.FirstRow
    intCol = TDBGrid.LeftCol
   
    For i = 0 To TDBGrid.Splits.Count - 1
        If TDBGrid.Splits(i).SelStartCol <> -1 Then
            blnfound = True
            iCurSplit = i
            Exit For
        End If
    Next i

    If blnfound = False Then
        MsgBox "You must select a column before sorting."
        Exit Sub
    Else
        For intCurrentCol = TDBGrid.Splits(iCurSplit).SelStartCol To TDBGrid.Splits(iCurSplit).SelEndCol
            For i = 0 To TDBGrid.Columns.Count - 1
                If intCurrentCol = TDBGrid.Columns(i).Order Then
                    Exit For
                End If
            Next i
            If intCurrentCol = TDBGrid.Columns(i).Order Then
                If Len(strSort) > 1 Then
                    strSort = strSort + ", "
                End If
                strSort = strSort + TDBGrid.Columns(i).DataField
                If intDir = SORT_ASCENDING Then
                    strSort = strSort + " ASC"
                Else
                    strSort = strSort + " DESC"
                End If
            End If
        Next intCurrentCol
    End If
    m_rec.Sort = strSort
    TDBGrid.ReBind
    TDBGrid.FirstRow = intRow
    TDBGrid.LeftCol = intCol

End If
End Sub

' Called to serve data to the grid from the recordset
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_UnboundReadDataEx(ByVal RowBuf As TrueOleDBGrid60.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_UnboundReadDataEx(ByVal RowBuf As TrueOleDBGrid70.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
Private Sub TDBGrid_UnboundReadDataEx(ByVal RowBuf As TrueOleDBGrid80.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    Dim intCol As Integer ' Grids internal column number
    Dim intRowsFetched As Integer, lngRow As Long ' Number of rows fetched, row number
    Dim vntStartRow As Variant ' Bookmark of starting row
    
    Dim lngCols As Long ' Number of columns requested
    Dim lngRows As Long ' Number of rows requested
    '
    '   If the columncount = 0 and the rowcount = 1 this event was fired because
    '   the grid wants the bookmark!
    lngCols = RowBuf.COLUMNCOUNT - 1
    lngRows = RowBuf.RowCount - 1
    
    intRowsFetched = 0
    
    On Error GoTo NoRead

    If m_rec Is Nothing Then GoTo NoRead
    
    ' Figure out where to start
    If IsNull(StartLocation) Then
        If offset < 0 Then
            ' StartLocation refers to EOF
            m_rec.MoveLast
            m_rec.MoveNext
        Else
            ' StartLocation refers to BOF
            m_rec.MoveFirst
            m_rec.MovePrevious
        End If
        m_rec.Move offset
    Else
        m_rec.Move offset, StartLocation
    End If
        
    ' vntStartRow is the bookmark of the first requested row
    vntStartRow = m_rec.Bookmark
    Dim pos As Long
    pos = m_rec.AbsolutePosition
    
    For lngRow = 0 To lngRows
        ' If out of bounds quit this loop
        If m_rec.BOF Or m_rec.EOF Then Exit For
        
        On Error Resume Next
        
        ' Only show rows that aren't deleted
'        Dim clsRowInfo As CRowInfo
'        Set clsRowInfo = Nothing
'        Set clsRowInfo = m_colRowInfos(Str(m_rec.Bookmark))
'        If clsRowInfo Is Nothing Or Not clsRowInfo.state = STATE_DELETED Then
            ' Retrieve values
            For intCol = 0 To lngCols
                Select Case m_colColumnDefs(RowBuf.columnName(intCol)).Format
                Case FORMAT_PRICE
                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.columnName(intCol)), "##,##0.00")
                Case FORMAT_DECIMAL3
                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.columnName(intCol)), "##,##0.000")
                Case FORMAT_DECIMAL
                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.columnName(intCol)), "##,##0.00###")
                Case FORMAT_DATE
                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.columnName(intCol)), "mm/dd/yyyy")
                Case FORMAT_DATETIME
                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.columnName(intCol)), "mm/dd/yyyy HH:MM:ss")
                Case FORMAT_MATERIAL
                    'Line of code was changed by Mohan on Jan 05,2012: changed FORMAT_MATERIAL_SRV to FORMAT_MATERIAL_04_SRV
                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.columnName(intCol)), FORMAT_MATERIAL_04_SRV)
                Case FORMAT_UNIT_COST
                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.columnName(intCol)), FORMAT_UNIT_COST_SRV)
                Case FORMAT_UNIT_COST_04
                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.columnName(intCol)), FORMAT_UNIT_COST_04_SRV)
                Case FORMAT_STRING_TRIM
                    RowBuf.Value(lngRow, intCol) = Trim(m_rec.Fields(RowBuf.columnName(intCol)))
                Case FORMAT_STRING_URL
                    RowBuf.Value(lngRow, intCol) = Trim(m_rec.Fields(RowBuf.columnName(intCol)))
                Case Else
                    RowBuf.Value(lngRow, intCol) = m_rec.Fields(RowBuf.columnName(intCol))
                End Select
'                If m_rec.Fields(RowBuf.ColumnName(intCol)).Type = adDBTimeStamp Then
'                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.ColumnName(intCol)), "mm/dd/yyyy HH:MM:ss")
'                ElseIf m_rec.Fields(RowBuf.ColumnName(intCol)).Type = adDBDate Then
'                    RowBuf.Value(lngRow, intCol) = Format(m_rec.Fields(RowBuf.ColumnName(intCol)), "mm/dd/yyyy")
'                Else
'                    RowBuf.Value(lngRow, intCol) = m_rec.Fields(RowBuf.ColumnName(intCol))
'                End If
            Next intCol
            
            ' Assign a bookmark for currently fetched row
            RowBuf.Bookmark(lngRow) = m_rec.Bookmark
            
            intRowsFetched = intRowsFetched + 1
'        Else
'            GoTo NoRead
'        End If
        m_rec.MoveNext
    Next lngRow
    
    RowBuf.RowCount = intRowsFetched
    
    ' Callibrate VScroll bar
    If pos >= 0 Then ApproximatePosition = pos
    
    Exit Sub

NoRead:
    RowBuf.RowCount = 0
    If GRIDTYPE = "Model Assemblies" Then
        Exit Sub
    Else
        If Err.Number = 3021 Then Resume Next
    End If

End Sub

' Fired when a row has been modified
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_UnboundWriteData(ByVal RowBuf As TrueOleDBGrid60.RowBuffer, WriteLocation As Variant)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_UnboundWriteData(ByVal RowBuf As TrueOleDBGrid70.RowBuffer, WriteLocation As Variant)
Private Sub TDBGrid_UnboundWriteData(ByVal RowBuf As TrueOleDBGrid80.RowBuffer, WriteLocation As Variant)
    Dim intCol As Integer
    Dim blnChanges As Boolean
    On Error GoTo NoUpdate
    
    ' Position to the desired row
    m_rec.Bookmark = WriteLocation
    
    ' Write only modified columns
    For intCol = 0 To RowBuf.COLUMNCOUNT - 1
        If Not IsNull(RowBuf.Value(0, intCol)) Then
            If m_rec.Fields(RowBuf.columnName(RowBuf.ColumnIndex(0, intCol))).Type = adVarChar Or m_rec.Fields(RowBuf.columnName(RowBuf.ColumnIndex(0, intCol))).Type = adChar Then
                If StringDataModified(IIf(IsNull(m_rec.Fields(RowBuf.columnName(intCol)).Value), "", m_rec.Fields(RowBuf.columnName(intCol)).Value), CStr(RowBuf.Value(0, intCol)), CStr(RowBuf.columnName(intCol))) = True Then
                    blnChanges = True
                    
                    Dim definedSize As Integer 'rlh
                    definedSize = m_rec.Fields(RowBuf.columnName(RowBuf.ColumnIndex(0, intCol))).definedSize  'rlh
                    'm_rec.Fields(RowBuf.ColumnName(intCol)).Value = Left(StringVerifiedData(CStr(RowBuf.Value(0, intCol)), CStr(RowBuf.ColumnName(intCol))), 255)          'rlh - too large for 75 char field (if it is one)
                    m_rec.Fields(RowBuf.columnName(intCol)).Value = Left(StringVerifiedData(CStr(RowBuf.Value(0, intCol)), CStr(RowBuf.columnName(intCol))), definedSize)   'rlh
                
                End If
            Else
                If CStr(IIf(IsNull(m_rec.Fields(RowBuf.columnName(intCol)).Value), "", m_rec.Fields(RowBuf.columnName(intCol)).Value)) <> CStr(RowBuf.Value(0, intCol)) Then
                    blnChanges = True
                    m_rec.Fields(RowBuf.columnName(intCol)).Value = StringVerifiedData(CStr(RowBuf.Value(0, intCol)), CStr(RowBuf.columnName(intCol)))
                End If
            End If
        End If
    Next intCol
    
    If blnChanges = True Then
        On Error Resume Next
        Dim clsRowInfo As CRowInfo
        clsRowInfo = Nothing
        Set clsRowInfo = m_colRowInfos(str(WriteLocation))
        If clsRowInfo Is Nothing Then
            Set clsRowInfo = New CRowInfo
            clsRowInfo.State = STATE_MODIFIED
            clsRowInfo.Bookmark = WriteLocation
            clsRowInfo.Error = False
            m_colRowInfos.Add clsRowInfo, str(WriteLocation)
        Else
            ' Only make it modified if the row isn't a new one.
            If Not clsRowInfo.State = STATE_NEW Then
                clsRowInfo.State = STATE_MODIFIED
            End If
        End If
    End If
    Exit Sub

NoUpdate:
    If Err.Number <> 0 Then MsgBox "CGridMap.TDBGrid.UnboundWriteData error " & Err.Number & ":" & Err.Description
    'MsgBox "Column " + m_colColumnDefs(RowBuf.ColumnName(intCol)).Name + " contains the invalid value '" + IIf(IsNull(RowBuf.Value(0, intCol)), "Null", RowBuf.Value(0, intCol)) + "'."
    Screen.MousePointer = vbNormal
    'Resume
    RowBuf.RowCount = 0
    TDBGrid.DataChanged = False
End Sub

' Fired upon adding a new row to the grid
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_UnboundAddData(ByVal RowBuf As TrueOleDBGrid60.RowBuffer, NewRowBookmark As Variant)
'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_UnboundAddData(ByVal RowBuf As TrueOleDBGrid70.RowBuffer, NewRowBookmark As Variant)
Private Sub TDBGrid_UnboundAddData(ByVal RowBuf As TrueOleDBGrid80.RowBuffer, NewRowBookmark As Variant)
    Dim intCol As Integer
    Dim adoErrs As ADODB.Errors     'rlh
    Dim i As Integer                'rlh
    
    On Error GoTo NoAdd

    m_rec.AddNew
    
    
    ' Save values in modified columns
    For intCol = 0 To RowBuf.COLUMNCOUNT - 1
       
        If Not IsNull(RowBuf.Value(0, intCol)) Then
            If m_rec.Fields(RowBuf.columnName(RowBuf.ColumnIndex(0, intCol))).Type = adVarChar Or m_rec.Fields(RowBuf.columnName(RowBuf.ColumnIndex(0, intCol))).Type = adChar Then
                'StringVerifiedData reformats IDs and converts null to ""
                
                If Len(CStr(RowBuf.Value(0, intCol))) > 75 Then       'rlh
                    Dim tmpstr As String
                    tmpstr = Mid(RowBuf.Value(0, intCol), 1, 75)
                Else
                    tmpstr = RowBuf.Value(0, intCol)
                End If                                                'rlh
                
                'm_rec.Fields(RowBuf.ColumnName(RowBuf.ColumnIndex(0, intCol))).Value = StringVerifiedData(RowBuf.Value(0, intCol), RowBuf.ColumnName(RowBuf.ColumnIndex(0, intCol)))  'rlh
                 m_rec.Fields(RowBuf.columnName(RowBuf.ColumnIndex(0, intCol))).Value = StringVerifiedData(tmpstr, RowBuf.columnName(RowBuf.ColumnIndex(0, intCol)))                    'rlh
            Else
                m_rec.Fields(RowBuf.columnName(RowBuf.ColumnIndex(0, intCol))).Value = RowBuf.Value(0, intCol)
                
            End If
        Else
            ' Default vallues cna be handled here
        End If
        'MsgBox "Err.Description: " & Err.Description & "  " & intCol
    Next intCol

    On Error Resume Next
    Dim clsRowInfo As CRowInfo
    clsRowInfo = Nothing
    clsRowInfo = m_colRowInfos(str(m_rec.Bookmark))
    If clsRowInfo Is Nothing Then
        Set clsRowInfo = New CRowInfo
        clsRowInfo.State = STATE_NEW
        clsRowInfo.Bookmark = m_rec.Bookmark
        clsRowInfo.Error = False
        m_colRowInfos.Add clsRowInfo, str(m_rec.Bookmark)
    End If

    ' Return the bookmark for this row
    NewRowBookmark = m_rec.Bookmark
    
    TDBGrid.ApproxCount = TDBGrid.ApproxCount + 1
    Exit Sub
    
NoAdd:
    RowBuf.RowCount = 0
    
'''    Debug.Print " "                                                     'rlh
'''    Debug.Print "* * * * * * * CGRIDMAP.UnboundAddData  * * * * * * *"  'rlh
'''
'''    'DISPLAY ADO ERRORS
'''    Set adoErrs = m_rec.ActiveConnection.Errors 'rlh
'''    For I = 0 To (adoErrs.Count - 1)            'rlh
'''        Debug.Print adoErrs(I).Description      'rlh
'''        Debug.Print adoErrs(I).NativeError
'''        Debug.Print adoErrs(I).Number
'''        Debug.Print adoErrs(I).HelpFile
'''    Next                                        'rlh
'''
'''    'ROWBUF
'''    Debug.Print "    "
'''    Debug.Print "* * * * * ROWBUF: * * * * * "
'''    For I = 0 To RowBuf.COLUMNCOUNT - 1         'rlh
'''        Debug.Print "Name: " & RowBuf.ColumnName(I) & "  Value: " & RowBuf.Value(0, I)
'''    Next
'''
'''    'DISPLAY RECORDSET ATTRIBUTES AND VALUES
'''    Debug.Print "   "
'''    Debug.Print "* * * * RECORDSET FIELDS/VALUES * * *"
'''    For I = 0 To m_rec.Fields.Count - 1         'rlh
'''        Debug.Print I & " Name: " & m_rec.Fields(I).Name & " Value: " & m_rec.Fields(I) & "  Type: " & m_rec.Fields(I).Type & " Defined Size: " & m_rec.Fields(I).definedSize
'''    Next
'''
'''    'Resume 'rlh
    'Debug.Print "CGridMap.TDBGrid.UnboundAddData error " & Err.Number & ":" & Err.Description
End Sub

' Add a column to the grid
Public Sub AddCol(strColName As String, strColMap As String, Optional strTable As String = vbNullString, Optional intColWidth As Integer = 0, Optional bLocked As Boolean = False, Optional bVisible As Boolean = True, Optional bClone As Boolean = True, Optional intFormat As Integer = 0, Optional bFetchStyle As Boolean = False) ' Optional strDropDown As String = vbNullString)
    Dim objColumnDef As New CColumnDef
    With objColumnDef
        .Name = strColName
        .TableCol = strColMap
        .Table = strTable
        .ColWidth = intColWidth
        .Locked = bLocked
        .Visible = bVisible
        .Clone = bClone
        .FetchStyle = bFetchStyle
        .Format = intFormat
        '.DropDown = strDropDown
        If intFormat = FORMAT_STRING_URL And Not bLocked Then
            'IF FORMAT_STRING_URL is specified, field should be locked
            .Locked = True
        End If
    End With
    
    m_colColumnDefs.Add objColumnDef, strColMap
    
    m_intGridCols = m_intGridCols + 1
End Sub

' Add a split at current m_intGridCols value
Public Sub AddSplit()
    m_colSplits.Add m_intGridCols
End Sub

' Clone a row from the grid
Public Sub CloneRow(Optional blnMakeNew As Boolean = False)
    On Error Resume Next
'*** APEX Migration Utility Code Change ***
'    Dim objGridCol As TrueOleDBGrid60.Column ' A column
'*** APEX Migration Utility Code Change ***
'    Dim objGridCol As TrueOleDBGrid70.Column ' A column
    Dim objGridCol As TrueOleDBGrid80.Column ' A column
    Dim objColumnDef As CColumnDef
    Dim vntMyBookmark As Variant ' Temp bookmark
    Dim i As Integer
    Dim rec As ADODB.RecordSet ' Copy of existing recordset

    Set rec = m_rec.Clone ' Copy existing recordset so we can have multiple current rows
    rec.Bookmark = TDBGrid.Bookmark ' move current row in copy
'    m_rec.AddNew ' add new record to original

'    For i = 0 To TDBGrid.Columns.Count - 1
'        Set objGridCol = TDBGrid.Columns(i)
'        Set objColumnDef = m_colColumnDefs.Item(objGridCol.DataField)
'        ' Only copy appropriate columns
'        If objColumnDef.Clone Then
'            m_rec.Fields(objGridCol.DataField).Value = rec.Fields(objGridCol.DataField).Value
'        End If
'    Next
    CopyRSFieldsAndData m_rec, rec, m_colCloneCols
    rec.Close ' Close copy

    ' If this row should be marked as new.
    ' Used when clone does not open a single-record view.
    If blnMakeNew = True Then
        Dim clsRowInfo As CRowInfo
        clsRowInfo = Nothing
        clsRowInfo = m_colRowInfos(str(m_rec.Bookmark))
        If clsRowInfo Is Nothing Then
            Set clsRowInfo = New CRowInfo
            clsRowInfo.State = STATE_NEW
            clsRowInfo.Bookmark = m_rec.Bookmark
            clsRowInfo.Error = False
            m_colRowInfos.Add clsRowInfo, str(m_rec.Bookmark)
        End If
    End If

    TDBGrid.ApproxCount = TDBGrid.ApproxCount + 1
'    vntMyBookmark = m_rec.Bookmark ' Save current bookmark
'    TDBGrid.ReBind ' Reset grid contents
'    TDBGrid.Bookmark = vntMyBookmark ' Set bookmark back again
'    m_rec.Bookmark = vntMyBookmark
    TDBGrid.SetFocus
    m_rec.MoveLast
    TDBGrid.ReOpen m_rec.Bookmark
    DoEvents
    TDBGrid.Col = 0
    TDBGrid.Columns(0).Value = TDBGrid.Columns(0).Value     'Set datachanged flag for pencil
    TDBGrid.SetFocus
End Sub

Public Function CloneRowRecordset() As ADODB.RecordSet
    On Error Resume Next
    Dim rec As ADODB.RecordSet ' Copy of existing recordset
    Dim recNew As New ADODB.RecordSet
    Dim intItem As Integer
    Dim fld As ADODB.Field
    Set CloneRowRecordset = Nothing

'    Set rec = m_rec.Clone ' Copy existing recordset so we can have multiple current rows
'    rec.Bookmark = TDBGrid.Bookmark ' move current row in copy
'    m_rec.AddNew ' add new record to original

'    For Each fld In rec.Fields
'        intItem = 0
'        intItem = m_colCloneCols(fld.Name)
'        If Not intItem = 0 Then
'            m_rec.Fields(fld.Name).Value = fld.Value
'        End If
'        recTest.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
'    Next
    
'    recTest.Open
'    recTest.AddNew
'    For Each fld In rec.Fields
'        recTest.Fields(fld.Name).Value = fld.Value
'    Next

    m_rec.Bookmark = TDBGrid.Bookmark
    CopyRSFieldsAndData recNew, m_rec, m_colCloneCols

'    Set CloneRowRecordset = m_rec
    Set CloneRowRecordset = recNew
End Function

Public Function rowinfo(intKey As Integer) As CRowInfo
    On Error Resume Next
    rowinfo = Nothing
    Set rowinfo = m_colRowInfos(str(intKey))
End Function

Public Sub SetRowState(intBookmark As Integer, intState As Integer)
    On Error Resume Next
    Dim clsRowInfo As CRowInfo
    Set clsRowInfo = Nothing
    Set clsRowInfo = m_colRowInfos(str(intBookmark))
    If clsRowInfo Is Nothing Then
        Set clsRowInfo = New CRowInfo
        clsRowInfo.State = intState
        clsRowInfo.Bookmark = intBookmark
        clsRowInfo.Error = False
        m_colRowInfos.Add clsRowInfo, str(intBookmark)
    Else
        clsRowInfo.State = intState
    End If
End Sub

Public Sub SetRowSkey(intBookmark As Integer, lngSkey As Long)
    On Error Resume Next
    Dim clsRowInfo As CRowInfo
    Set clsRowInfo = Nothing
    Set clsRowInfo = m_colRowInfos(str(intBookmark))
    If Not clsRowInfo Is Nothing Then
        clsRowInfo.sKey = lngSkey
    End If
End Sub
Public Sub SetRowSql(intBookmark As Integer, strSQL As String)
    On Error Resume Next
    Dim clsRowInfo As CRowInfo
    Set clsRowInfo = Nothing
    Set clsRowInfo = m_colRowInfos(str(intBookmark))
    If Not clsRowInfo Is Nothing Then
        clsRowInfo.RowSQL = strSQL
    End If
End Sub
Public Sub SetRowSubKey(intBookmark As Integer, strSub_Key As String)
    On Error Resume Next
    Dim clsRowInfo As CRowInfo
    Set clsRowInfo = Nothing
    Set clsRowInfo = m_colRowInfos(str(intBookmark))
    If Not clsRowInfo Is Nothing Then
        clsRowInfo.SubKey = strSub_Key
    End If
End Sub

Public Sub SetFrozenCols(intCols As Integer)
    m_intFrozenCols = intCols
End Sub

Public Property Let GRIDTYPE(ByVal strNewValue As String)
    m_strGridType = strNewValue
End Property

Public Property Get GRIDTYPE() As String
    GRIDTYPE = m_strGridType
End Property

Public Property Get COLUMNCOUNT() As Integer
    COLUMNCOUNT = m_intGridCols
End Property

' When a user clicks on a selection row, give the grid focus
Private Sub TDBGrid_SelChange(Cancel As Integer)
    TDBGrid.SetFocus
End Sub
