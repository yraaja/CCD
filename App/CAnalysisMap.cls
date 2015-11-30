VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnalysisMap"
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
Const ANALYSIS_TABLE = "PROJ_BLDG_COMPONENT_COST"
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
Private m_colUpdatedRows        As New Collection
'
'   Used in the SuccessfulUpdates Property Get.
Private m_NumberOfSuccessfulUpdates As Integer
'
'   Used in the UpdateErrors Property Get.
Private m_NumberOfUpdateErrors As Integer
'
'   Indicates a field that will affect overall cost rollups
'   has been changed so we must RefreshCosts if they update.
Dim bRefreshCosts As Boolean

Private Sub Class_Initialize()
    With m_objGridMap.m_colCloneCols
        .Add 1, "ID"
        .Add 1, "Desc"
        .Add 1, "Ct"
        .Add 1, "Bk ID"
    End With
End Sub

Public Sub InitGrid(Optional numcol As Integer)
    Dim I As Integer
    '
    '   Identifies the Grid.
    With m_objGridMap
        .GRIDTYPE = "ANALYSIS"
        '
        '   Add the columns to the grid, pointing to their
        '   corresponding db field names so that the grid can
        '   populate directly from the m_rec recordset.
        .AddCol "Project Skey", "bk_skey", ANALYSIS_TABLE, 800, True, False
        .AddCol "ID", "class_id", ANALYSIS_TABLE, 800, True, True
        .AddCol "Desc", "class_desc", ANALYSIS_TABLE, 1500, True, True
        '
        '   Main Split, this grid only has 1.
        '
        .AddSplit
        .AddCol "Ct", "proj_count", ANALYSIS_TABLE, 1200, True, True
        .AddCol "Bk ID", "bk_id", ANALYSIS_TABLE, 1200, False, True, 1
        .AddCol "Override Book Desc", "override_book_desc", ANALYSIS_TABLE, 1200, False, True, True
        .AddCol "Override Metric Book Desc", "override_metric_book_desc", ANALYSIS_TABLE, 1200, False, True, True
        ' ADDED COLUMN FORMAT FORMAT_PRICE; INCREASED WIDTH TO 700 TO ACCOMODATE
        ' 6/21/2005 RTD FOR VERSION 7.4.0 (CR#1511)
        For I = 1 To numcol
            .AddCol "", "col_" & Trim(str(I)), ANALYSIS_TABLE, 700, True, True, , FORMAT_PRICE
        Next
        '
        '   Call base class.
        .AddCol "Last Update", "last_update_date", ANALYSIS_TABLE, 1200, True, True
        .AddCol "Last Update By", "last_update_person", ANALYSIS_TABLE, 1200, True, True
        If m_blnPreferences Then
            .BuildPreferenceGrid
        Else
            .BuildGrid True, False
        End If
    End With
    '
    '   Give headings for the splits.
    With TDBGrid
        .Splits(0).Caption = "ID"
        .Splits(1).Caption = "Value"
        .Splits(1).Size = 5
    End With
End Sub

Public Property Get GRIDTYPE() As String
    GRIDTYPE = "ANALYSIS"
End Property

Public Sub SetMenuBar()
    m_objGridMap.SetMenuBar
End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Public Sub SelectAllRows()
    m_objGridMap.SelectAllRows
End Sub
Public Function CloneRowRecordset() As ADODB.RecordSet
    Set CloneRowRecordset = m_objGridMap.CloneRowRecordset
End Function

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


Public Sub Update(ByVal strFacilityClassID As String)
    Dim strSQL As String
    Dim strError As String
    Dim I As Integer
    'On Error Resume Next

    With m_rec
        If .State <> adStateClosed Then
            .MoveFirst
            While Not .EOF
                '
                '   Skip the record if it didn't change
                Dim clsRowInfo As CRowInfo
                Set clsRowInfo = m_objGridMap.rowinfo(.Bookmark)
                
                If Not clsRowInfo Is Nothing Then
                    If clsRowInfo.State <> STATE_NONE Then
                        strSQL = "EXEC sp_update_project_book @bookskey = '" & .Fields("bk_skey") & "', @projbookid = '" & .Fields("bk_id") & "', @facility_class_id = '" & strFacilityClassID & "', @class_id = '" & .Fields("class_id") & "', @override_book_desc = '" & Trim(.Fields("override_book_desc")) & "', @override_metric_book_desc = '" & Trim(.Fields("override_metric_book_desc")) & "'"
                        ' exec store proc
                        If g_objDAL.ExecQuery(vbNullString, strSQL, strError) Then
                            If strError <> "" Then
                                MsgBox "Error while updating book id"
                            End If
                        End If
                        clsRowInfo.State = STATE_NONE
                    End If
                End If
                .MoveNext
            Wend
            
            .UpdateBatch
        End If
    End With
End Sub

Public Function IsPendingChange() As Boolean
    IsPendingChange = m_objGridMap.IsPendingChange
End Function
