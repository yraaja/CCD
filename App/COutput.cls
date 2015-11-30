VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_objGridMap As New CGridMap ' "Parent" class
'*** APEX Migration Utility Code Change ***
'Dim m_grd As TrueOleDBGrid60.TDBGrid
'*** APEX Migration Utility Code Change ***
'Dim m_grd As TrueOleDBGrid70.TDBGrid
Dim m_grd As TrueOleDBGrid80.TDBGrid
Dim m_frmCallingForm As Form
Dim m_intGridIndex As Integer

Private m_rec As ADODB.RecordSet ' The recordset that populates this grid
Private m_intMaxCol As Integer ' Number of columns in the recordset
Private m_strGridType As String ' Names grid
Private m_intFrozenCols As Integer ' Number of columns in grid to be frozen during scrolling

Private m_colColumnDefs As New Collection ' Collection of CColumnDef objects describing grid columns
Private m_colSplits As New Collection ' Collection of locations for splits
Private m_colRowInfos As New Collection ' Collection of CRowInfo objects describing row state
Private m_intGridCols As Integer ' Number of columns in the grid
Public m_colCloneCols As New Collection ' List of fields to be cloned

'Support Each Grid on the Output dialogue.
Public Sub SetGrid(frm As Form, Index As Integer)
    ' Pass-through to base class
    Set m_grd = frm.TDBGrid(Index)
    Set TDBGrid = frm.TDBGrid(Index)
    Set m_frmCallingForm = frm
    m_intGridIndex = Index
'    AddCol "graphic_ref_id", "graphic_ref_id"
'    frm.TDBGrid(index).AddCol "table_ref_id", "table_ref_id"
End Sub
Public Sub AddCol(strColName As String, strColMap As String, Optional strTable As String = vbNullString, Optional intColWidth As Integer = 0, Optional bLocked As Boolean = False, Optional bVisible As Boolean = True, Optional bClone As Boolean = True, Optional intFormat As Integer = 0, Optional bFetchStyle As Boolean = False) ' Optional strDropDown As String = vbNullString)
'    Dim objColumnDef As CColumnDef
'    Set objColumnDef = m_frmCallingForm.TDBGrid(0).Columns(0)
    Dim m_rec As ADODB.RecordSet
    Dim sql As String
    sql = "select output_id, output_desc, output_group_id from output where output_group_id=1"
    blnRet = g_objDAL.GetRecordset(vbNullString, sql, m_rec)

    
    Set m_frmCallingForm.TDBGrid(0).RecordSet = m_rec
    
        '.Name = strColName
'        .TableCol = strColMap
'        .Table = strTable
'        .ColWidth = intColWidth
'        .Locked = bLocked
'        .Visible = bVisible
'        .Clone = bClone
'        .FetchStyle = bFetchStyle
'        .Format = intFormat
'        '.DropDown = strDropDown
    'End With
    
    'm_frmCallingForm.TDBGrid.Columns.Add objColumnDef, strColMap
    'm_frmCallingForm.TDBGrid(0).Columns(0).caption =  objColumnDef, strColMap
    'm_intGridCols = m_intGridCols + 1
End Sub

