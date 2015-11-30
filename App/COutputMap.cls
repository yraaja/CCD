VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COutputMap"
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
Const OUTPUT_TABLE = "output_usage"
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

Dim m_oOutputUsageFormat As OUTPUT_USAGE_FORMAT

'Property to Get/Set Allowable Output Usage Formats
'[1=2004 ONLY; 2=1995 ONLY; 3=BOTH]
Public Property Get OutputUsageFormat() As OUTPUT_USAGE_FORMAT
    OutputUsageFormat = m_oOutputUsageFormat
End Property
Public Property Let OutputUsageFormat(NewValue As OUTPUT_USAGE_FORMAT)
    m_oOutputUsageFormat = NewValue
    RefreshMasterFormatValueItems
End Property

Private Sub Class_Initialize()

    With m_objGridMap.m_colCloneCols
        .Add 1, "output_group_id"
        .Add 1, "output_id"
        .Add 1, "output_desc"
        .Add 1, "ID_Num"
        .Add 1, "selected"
        .Add 1, "graphic_ref_id"
        .Add 1, "table_ref_id"
        .Add 1, "ext_indicator"         ' MasterFormat 2004 indicator
        .Add 1, "ext_graphic_ref_id"    ' MasterFormat 2004 Graphic Ref
        .Add 1, "ext_table_ref_id"      ' MasterFormat 2004 Table Ref
        .Add 1, "table_ref_col"
        .Add 1, "format_code"
        .Add 1, "indent_code"
        .Add 1, "format_characters"
        .Add 1, "last_update_id"
    End With
    
    '8/15/2005 RTD
    'DEFAULT TO ALLOW BOTH FORMATS
    'USED BY MASTERFORMAT VERSION DROP-DOWN
    '(EXT_INDICATOR COLUMN)
    OutputUsageFormat = OUTPUT_BOTH

End Sub

Public Sub InitGrid(ByRef picControl0 As PictureBox, ByRef picControl1 As PictureBox, _
                ByRef picControl2 As PictureBox, iSkeyCount As Integer)
    
    Dim strSELECT   As String
    Dim recTemp     As ADODB.RecordSet
    '
    '   Identifies the Grid.
    With m_objGridMap
        .GRIDTYPE = "Output"
        '
        '   Add the columns to the grid, pointing to their
        '   corresponding db field names so that the grid can
        '   populate directly from the m_rec recordset.
        
        'rlh ######################  WARNING WARNING WARNING #######################
        '
        'ADD COLUMNS (AddCol) in the order that you want to see them in the grid!
        '
        '#######################  END OF WARNING ################################
        
        .AddCol "Output Grp ID", "output_group_id", OUTPUT_TABLE, 600, True, False, True, , True
        .AddCol "Output ID", "output_id", OUTPUT_TABLE, 600, True, False, True, , True
        .AddCol "ID_Num", "ID_Num", OUTPUT_TABLE, 300, False, False, False, , True
        .AddCol "Selected", "selected", OUTPUT_TABLE, 320, False, True, False, , True
        .AddCol "Book", "output_desc", OUTPUT_TABLE, 2000, False, True, False, , True
        '
        '   Main Split, this grid only has 1.
        '
        .AddSplit
        
        .AddCol "Mst Fmt", "ext_indicator", OUTPUT_TABLE, 700, False, True, True, , True
'        .AddCol "Graphic", "graphic_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
'        .AddCol "Table", "table_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
        ' 8/19/2005 RTD - ADD COLUMNS FOR MASTERFORMAT 2004 INDICATOR AND REFS.
        ' 8/25/2005 RTD - CHANGE EXT_INDICATOR COLUMN TO CLONE=TRUE
'        .AddCol "Mst Fmt", "ext_indicator", OUTPUT_TABLE, 700, False, True, True, , True
        .AddCol "Graphic 04", "ext_graphic_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
        
        'rlh (below) 06/10/08  NEW GRID COLUMN
        .AddCol "TAG", "tag_code", OUTPUT_TABLE, 1250, False, True, True, , True
        'rlh end of testing 06/10/08
        
        .AddCol "Table 04", "ext_table_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
        .AddCol "Table Ref Col", "table_ref_col", OUTPUT_TABLE, 1000, True, False, True, , True
        .AddCol "FMT Code", "format_code", OUTPUT_TABLE, 850, False, True, True, , True
        .AddCol "Indent Code", "indent_code", OUTPUT_TABLE, 550, False, True, True, , True
        .AddCol "FMT Char", "format_characters", OUTPUT_TABLE, 800, False, True, True, , True
        .AddCol "Upd ID", "last_update_id", OUTPUT_TABLE, 800, False, False, False, , True
        'TRIAL LOCATION rlh 06/10/08  (moved from beginning of split till end)
        .AddCol "Graphic 95", "graphic_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
        .AddCol "Table 95", "table_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
'
        '
        '   Call base class.
        If m_blnPreferences Then
            .BuildPreferenceGrid
        Else
            .BuildGrid True, False
        End If
        
        RefreshValueItems picControl0, picControl1, picControl2, iSkeyCount
        RefreshMasterFormatValueItems
        
        '################### ADD COMBOBOX TO "TAG" CELLS #############
        '
        RefreshTagValueItems            'rlh 06/12/2008
        '
        '#############################################################
        With TDBGrid
            .ScrollBars = dbgAutomatic
            .AlternatingRowStyle = False
        End With
    End With
    
End Sub

Public Sub InitPreferenceGrid()
    
    Dim strSELECT   As String
    Dim recTemp     As ADODB.RecordSet
    '
    '   Identifies the Grid.
    With m_objGridMap
        .GRIDTYPE = "Output"
        '
        '   Add the columns to the grid, pointing to their
        '   corresponding db field names so that the grid can
        '   populate directly from the m_rec recordset.
        .AddCol "Output Grp ID", "output_group_id", OUTPUT_TABLE, 600, True, False, True, , True
        .AddCol "Output ID", "output_id", OUTPUT_TABLE, 600, True, False, True, , True
        .AddCol "ID_Num", "ID_Num", OUTPUT_TABLE, 300, False, False, False, , True
        .AddCol "Selected", "selected", OUTPUT_TABLE, 320, False, True, False, , True
        .AddCol "Book", "output_desc", OUTPUT_TABLE, 2000, False, True, False, , True
        '
        '   Main Split, this grid only has 1.
        '
        .AddSplit

        .AddCol "Graphic", "graphic_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
        .AddCol "Table", "table_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
        .AddCol "Mst Fmt", "ext_indicator", OUTPUT_TABLE, 700, False, True, False, , True
        .AddCol "Graphic 04", "ext_graphic_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
        .AddCol "Table 04", "ext_table_ref_id", OUTPUT_TABLE, 1250, False, True, False, , True
        .AddCol "Table Ref Col", "table_ref_col", OUTPUT_TABLE, 1000, True, False, True, , True
        .AddCol "FMT Code", "format_code", OUTPUT_TABLE, 850, False, True, True, , True
        .AddCol "Indent Code", "indent_code", OUTPUT_TABLE, 550, False, True, True, , True
        .AddCol "FMT Char", "format_characters", OUTPUT_TABLE, 800, False, True, True, , True
        .AddCol "Upd ID", "last_update_id", OUTPUT_TABLE, 800, False, False, False, , True
        '
        .BuildPreferenceGrid
        '
        'ADDED 8/8/2005 RTD
        'These Columns were hidden and can't be made visible or
        'it breaks the logic
        TDBGrid.Columns("Output Grp ID").Locked = True
        TDBGrid.Columns("Output ID").Locked = True
        TDBGrid.Columns("ID_Num").Locked = True
        TDBGrid.Columns("Output Grp ID").Locked = True
    End With
End Sub

Public Sub RefreshValueItems(ByRef picControl0 As PictureBox, ByRef picControl1 As PictureBox, _
                ByRef picControl2 As PictureBox, iSkeyCount As Integer)
                
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid70.ValueItem
    Dim Item As New TrueOleDBGrid80.ValueItem
    
    With TDBGrid.Columns("Selected").ValueItems
        Item.Value = "0"
        Item.DisplayValue = picControl0.Picture
        .Add Item
        
        Item.Value = "1"
        Item.DisplayValue = picControl1.Picture
        .Add Item
        '
        '   If the user selected more than 1 row for output then
        '   show the grayed image in the list.
        If iSkeyCount > 1 Then
            Item.Value = "2"
            Item.DisplayValue = picControl2.Picture
            .Add Item
        End If
        
        .Translate = True
        .CycleOnClick = True
    End With
    
End Sub

Private Sub RefreshMasterFormatValueItems()
'8/15/2005 RTD
'FORMAT MSTFMT INDICATOR COLUMN AS DROP-DOWN;
'LOAD APPROPRIATE VALUES INTO DROP-DOWN BASED ON VALUE OF
'OutputUsageFormat PROPERTY
    Dim Item As New TrueOleDBGrid80.ValueItem
    
    On Error GoTo Err_Handler
    With TDBGrid.Columns("Mst Fmt").ValueItems
        .Clear
        If OutputUsageFormat = OUTPUT_MF2004_ONLY Or OutputUsageFormat = OUTPUT_BOTH Then
            Item.Value = "1"
            Item.DisplayValue = "2004"
            .Add Item
        End If
        If OutputUsageFormat = OUTPUT_MF1995_ONLY Or OutputUsageFormat = OUTPUT_BOTH Then
            Item.Value = "2"
            Item.DisplayValue = "1995"
            .Add Item
        End If
        If OutputUsageFormat = OUTPUT_BOTH Then
            Item.Value = "3"
            Item.DisplayValue = "BOTH"
            .Add Item
        End If
        .Translate = True
        .Validate = True
        .CycleOnClick = False
        .Presentation = dbgComboBox
    End With
    With TDBGrid.Columns("Mst Fmt")
        .AutoCompletion = True
        .AutoDropDown = True
    End With
    Exit Sub

Err_Handler:
    Exit Sub

End Sub

Private Sub RefreshTagValueItems()
'8/15/2005 RTD
'FORMAT MSTFMT INDICATOR COLUMN AS DROP-DOWN;
'LOAD APPROPRIATE VALUES INTO DROP-DOWN BASED ON VALUE OF
'OutputUsageFormat PROPERTY
    Dim Item As New TrueOleDBGrid80.ValueItem
    
    On Error GoTo Err_Handler
    With TDBGrid.Columns("TAG").ValueItems
        .Clear
        
        'Item.Value = "1"
        Item.Value = " "
        Item.DisplayValue = " "
        .Add Item
        
        'Item.Value = "2"
        Item.Value = "G"
        Item.DisplayValue = "G"
        .Add Item
       
        
        .Translate = True
        .Validate = True
        .CycleOnClick = False
        .Presentation = dbgComboBox
    End With
    
    With TDBGrid.Columns("TAG")
        .AutoCompletion = True
        .AutoDropDown = True
    End With
    Exit Sub

Err_Handler:
    Exit Sub

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
    GRIDTYPE = "Output"
End Property

Private Sub TDBGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim I As Integer
    Dim bLock As Boolean
    
    'On Error Resume Next
    Screen.MousePointer = vbHourglass
    Select Case TDBGrid.Columns(ColIndex).Caption
        Case "Selected"
            If Right(TDBGrid.Columns("selected"), 1) = 0 _
            Or Right(TDBGrid.Columns("selected"), 1) = 1 _
            Or Right(TDBGrid.Columns("selected"), 1) = 2 _
            Then
                m_rec.Bookmark = TDBGrid.Bookmark
                m_rec![Selected] = Right(TDBGrid.Columns("selected"), 1)
                m_rec.Update
                '
                ' Based on value of checkbox lock or unlock
                For I = 0 To TDBGrid.Columns.Count - 1
                    If I = TDBGrid.Columns("Selected").ColIndex Then
                    Else
                        TDBGrid.Columns(I).RefetchCell
                    End If
                Next I
            Else
                Cancel = True
            End If
        Case "Indent Code", "Fmt Char"  'Must be numeric
            If Not IsNumeric(TDBGrid.Text) Then
                MsgBox "The indent code and format characters must be numeric.", vbInformation
                Cancel = True
            End If
        Case "Mst Fmt"
            '8/10/2005 RTD
            'VERIFY MASTERFORMAT OUTPUT USAGE VALUE
            If Val(TDBGrid.Columns("Mst Fmt").Value) < 1 Or Val(TDBGrid.Columns("Mst Fmt").Value) > 3 Then
                MsgBox "The MasterFormat output usage must be set to 1995 Only, 2004 Only, or Both.", vbInformation
                Cancel = True
            End If
    End Select
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub TDBGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'ADDED 9/6/2005 RTD
'INTERCEPT AND PROCESS PASTE KEYBOARD EVENTS (CTRL-V OR SHIFT-INSERT)
'   TO WORK-AROUND TRUE DBGRID 8 PASTE BUG.
'REPORTED BY J. CHIANG ON 9/2/2005
'NOTE: THIS CODE DOES NOT INTERCEPT MOUSE OR WINDOWS PASTE EVENTS.
    Dim bIsPasting As Boolean
    Dim sText As String
    Dim sColumn As String
    Dim sField As String
    
    On Error Resume Next
    'CHECK FOR CONTROL+V OR SHIFT+INSERT
    bIsPasting = ((Shift And vbCtrlMask) = vbCtrlMask And KeyCode = 86) Or ((Shift And vbShiftMask) = vbShiftMask And KeyCode = 45)
    If bIsPasting Then
        'GET THE COLUMN AND DATABASE FIELD AND THE CLIPBOARD TEXT
        sColumn = TDBGrid.Columns(TDBGrid.Col).Caption
        If sColumn = "Table" Or sColumn = "Graphic" Or sColumn = "Table 04" Or sColumn = "Graphic 04" Then
            sText = Clipboard.GetText
        End If
        If sText <> "" And sColumn <> "" Then
            'UPDATE THE DATABASE FIELD, THEN REFETCH THE CELL
            sField = TDBGrid.Columns(TDBGrid.Col).DataField
            m_rec.Bookmark = TDBGrid.Bookmark
            m_rec(sField).Value = sText
            m_rec.Update
            TDBGrid.Columns(sColumn).RefetchCell
        End If
    End If
    
End Sub

Private Sub TDBGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If TDBGrid.Columns(TDBGrid.Col).Caption = "Chg Ntc" Or _
            TDBGrid.Columns(TDBGrid.Col).Caption = "Table" Or _
            TDBGrid.Columns(TDBGrid.Col).Caption = "Graphic" Or _
            TDBGrid.Columns(TDBGrid.Col).Caption = "Table 04" Or _
            TDBGrid.Columns(TDBGrid.Col).Caption = "Graphic 04" Then
            If Len(TDBGrid.Text) + 1 > 12 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TDBGrid_LostFocus()
    On Error Resume Next
    TDBGrid.Update
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    Dim I As Integer
    Dim bLock As Boolean
    
    If ColIndex = TDBGrid.Columns("selected").ColIndex Then
        m_rec.Bookmark = TDBGrid.Bookmark
        m_rec![Selected] = Right(TDBGrid.Columns("selected"), 1)
        m_rec.Update
        ' Based on value of checkbox lock or unlock
        For I = 0 To TDBGrid.Columns.Count - 1
            TDBGrid.Columns(I).RefetchCell
        Next I
    ElseIf ColIndex = TDBGrid.Columns("Mst Fmt").ColIndex Then
        '8/10/2005 RTD
        'UPDATE UNDERLYING RECORDSET AND REFRESH THE MF04 COLUMNS
        m_rec.Bookmark = TDBGrid.Bookmark
        m_rec.Fields("ext_indicator") = TDBGrid.Columns("Mst Fmt").Value
        m_rec.Update
        TDBGrid.Columns("Graphic 04").RefetchCell
        TDBGrid.Columns("Table 04").RefetchCell
    End If

End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
Private Sub TDBGrid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid80.StyleDisp)
'8/9/2005 RTD - MODIFIED TO USE WINDOWS STANDARD COLORS
    Dim blnLocked As Boolean
    ' If the row is highlighted, then let it be
    If (Condition And dbgSelectedRow) = 8 Then
        Exit Sub
    End If
    If IsNumeric(Bookmark) Then
        ' Based on value of checkbox lock or unlock
        m_rec.Bookmark = Bookmark
        If m_rec!Selected = 1 Then
            blnLocked = False
        Else
            blnLocked = True
        End If
        If blnLocked = True Then
            If Col = 3 Then
                CellStyle.Locked = False
                CellStyle.BackColor = vbWindowBackground
            Else
                CellStyle.Locked = True
                CellStyle.ForeColor = vbGrayText
                If Not g_blnUseAlternateDisabledColor Then
                    CellStyle.BackColor = vbButtonFace
                Else
                    CellStyle.BackColor = g_intAlternateRowColor
                End If
            End If
        Else
            If Col = TDBGrid.Columns("Book").ColIndex Then
                'Col 1 (Book Desc) always locked
                CellStyle.Locked = True
                CellStyle.ForeColor = vbWindowText  'vbGrayText
                CellStyle.BackColor = vbWindowBackground
            Else
                CellStyle.Locked = False
                CellStyle.BackColor = vbWindowBackground
                CellStyle.ForeColor = vbWindowText
            End If
            '8/10/2005 RTD
            'IF REFRESHING THE MASTERFORMAT 2004 CELLS,
            ' CHECK THE MF04_INDICATOR FIELD FOR LOCKED/UNLOCKED APPEARANCE
            If Col = TDBGrid.Columns("Table 04").ColIndex Or Col = TDBGrid.Columns("Graphic 04").ColIndex Then
                If Not IsNull(m_rec!ext_indicator) Then
                    'CELLS ARE LOCKED IF EXT_IND = 2 [1995 ONLY]
                    blnLocked = (m_rec!ext_indicator = 2)
                Else
                    blnLocked = True
                End If
                If blnLocked Then
                    CellStyle.Locked = blnLocked
                    CellStyle.ForeColor = vbGrayText
                    CellStyle.BackColor = g_intAlternateRowColor
                Else
                    CellStyle.Locked = blnLocked
                    CellStyle.ForeColor = vbWindowText
                    CellStyle.BackColor = vbWindowBackground
                End If
            End If
        End If
    End If
End Sub

