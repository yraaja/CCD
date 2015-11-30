VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DynaTree 
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   ScaleHeight     =   3615
   ScaleWidth      =   5145
   ToolboxBitmap   =   "MeansCtrl_DynaTree.ctx":0000
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4500
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MeansCtrl_DynaTree.ctx":0312
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   6165
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "DynaTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Modified by Patrick Lane 7/2/99
'Added Labor hierarchy, moved add of top level nodes to FetchChildren
'Added constants for table names

'Modified by Patrick Lane 11/3/99
'Added Uniformat ID hierarchy
'Added constants for table names

'Modified by Patrick Lane 5/5/00
'Added Book Detail hierarchy, logic for Commercial/Residential selection

'Modified by Patrick Lane 9/25/00
'Added Building Selection - no hierarchy

'Modified by Patrick Lane 11/16/00
'Converted material hierarchy to MasterFormat95

'Modified by Patrick Lane 3/22/02
'Converted Assembly, Assembly Book to Uniformat 2 with from/to range - Note: no subordinate count in DB, 0 used.

'Modified by Patrick Lane, 9/7/02
'Added CCI material/equip/labor

'Modified by Patrick Lane, 10/27/02
'Changed CCI material/equip/labor to use hierarchy and new table

'Modified by Sean Tzeng, 11/01/02
'Added Division 17

'Modified by Sean Tzeng, 11/06/02
'Added Project Analysis Tree List

'Modified by Patrick Lane, 11/17/02
'Added CCI Mat'l & Equipment list

'Modified by Rob Durfee, 6/27/05
'Added Initial Support for upcoming MasterFormat 2004
'Upgraded Control Version to 1.10

'Modified by Rob Durfee, 7/29/05, 8/8/2005
'Added Support for MasterFormat 2004
'Added UserControl Property functions to support Design-Time changes
'Upgraded Control Version to 1.20

Dim cn As ADODB.Connection
Dim rs As New ADODB.RecordSet
Dim Top1 As Node
Dim strSelect As String
Dim m_strType As String
Dim m_strHierType As String
Dim m_strTable As String
Dim m_strAssemblyTypeFldName As String
Dim m_blnShowMasterFormatRoot As Boolean
Dim m_strLastKeyMade As String
Dim m_blnDisableRedraw As Boolean

Const CCIMaterialTable = "CCI_MATERIAL"
Const CCIEquipmentTable = "CCI_EQUIPMENT"
Const CCILaborTable = "CCI_TRADE"
Const BuildingTable = "bldg_detail"
Const EquipHierTable = "MASTERFORMAT_ID_HIERARCHY"
Const MasterFormat95Table = "MASTERFORMAT95_ID_HIERARCHY"
Const MasterFormat04Table = "MASTERFORMAT04_ID_HIERARCHY"
Const AssemblyHierTable1 = "UNIFORMAT_ID_HIERARCHY"
Const AssemblyHierTable = "UNIFORMAT2_ID_HIERARCHY"
Const LaborHierTable = "LABOR_ID_HIERARCHY"
Const MaterialHierTable = "MATERIAL_ID_HIERARCHY"

Const strSelect_Start = "SELECT hier_id, hier_desc, subord_count From "
Const strSelect_Start2 = "SELECT hier_id, hier_desc, subord_count_ad From "
Const strSelect_Start3 = "SELECT hier_id, hier_desc, subord_count_abd From "
Const strSelect_Start4 = "SELECT hier_id, hier_desc, mat_subord_count From "
Const strSelect_Start5 = "SELECT uni2_category_id as hier_id, uni2_desc as hier_desc, 0 as subord_count From "

Const strSelect_Where = " where hier_id LIKE "
Const strSelect_Where2 = " where level_id = "
Const strSelect_Where3 = " where hier_type_code = '"
Const strSelect_Where4 = "' and hier_id LIKE "
Const strSelect_Where5 = " where uni2_level = "
Const strSelect_Where6 = "' and uni2_category_id LIKE "

Const TYPE_ASSEMBLY_BK_DTL_COM = "ASBLY_BK_DTL_COMMERCIAL"
Const TYPE_ASSEMBLY_BK_DTL_RESI = "ASBLY_BK_DTL_RESI"
Const TYPE_ASSEMBLY_COMMERCIAL = "ASSEMBLY_COMMERCIAL"
Const TYPE_ASSEMBLY_RESI = "ASSEMBLY_RESI"
Const TYPE_BUILDING = "BUILDING"
Const TYPE_CCI_MATERIAL = "CCI_MATERIAL"
Const TYPE_CCI_EQUIPMENT = "CCI_EQUIPMENT"
Const TYPE_CCI_LABOR = "CCI_LABOR"
Const TYPE_EQUIPMENT = "EQUIPMENT"
Const TYPE_LABOR = "LABOR"
Const TYPE_MATERIAL = "MATERIAL"
Const TYPE_MATERIAL04 = "MATERIAL04"
Const TYPE_UNIT_COST = "UNITCOST"
Const TYPE_UNIT_COST04 = "UNITCOST04"
Const TYPE_CCI_INDEX = "CCI_INDEX"
Const TYPE_PROJECT_LIST = "PROJECT_LIST"
Const TYPE_PROJECT_ANALYSIS = "PROJECT_ANALYSIS"
Const TYPE_CCI_MAT_EQUIP = "CCI_MAT_EQUIP"

Const COMMERCIAL = "C"
Const RESIDENTIAL = "R"
Const COM_FIELD_NAME = "coml_ind"
Const RESI_FIELD_NAME = "resi_ind"

Dim sDebug As String

Const WM_SETREDRAW As Long = &HB
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Event NodeSelected(ByVal strID As String)
Attribute NodeSelected.VB_Description = "Occurs when a tree node is selected"

'PROPERTY DisableRedraw
'ENABLES/DISABLES REDRAWING OF TREE CONTROL TO SPEED UP ADDING/REMOVING OF NODES
Public Property Get DisableRedraw() As Boolean
Attribute DisableRedraw.VB_Description = "Returns/sets whether Windows sends Redraw messages to the Tree control"
Attribute DisableRedraw.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DisableRedraw = m_blnDisableRedraw
End Property
Public Property Let DisableRedraw(NewValue As Boolean)
    m_blnDisableRedraw = NewValue
    SendMessageLong TreeView1.hWnd, WM_SETREDRAW, Not m_blnDisableRedraw, 0
    PropertyChanged "DisableRedraw"
End Property

'PROPERTY READ-ONLY Version
'RETURNS THE VERSION OF ACTIVEX CONTROL
Public Property Get Version() As String
Attribute Version.VB_Description = "Returns the control's Version string"
Attribute Version.VB_ProcData.VB_Invoke_Property = ";Misc"
    Version = App.Major & "." & App.Minor & ".0." & App.Revision
End Property

'PROPERTY READ-ONLY TableName
'RETURNS THE TABLE NAME FROM WHICH THE TREE IS BUILT
Public Property Get TableName() As String
Attribute TableName.VB_Description = "Returns the database table name used as the source for the Tree node data"
Attribute TableName.VB_ProcData.VB_Invoke_Property = ";Data"
    TableName = m_strTable
End Property

'PROPERTY READ-ONLY TreeType
'RETURNS THE TREE TYPE VALUE
Public Property Get TreeType() As String
Attribute TreeType.VB_Description = "Returns the Tree Type loaded into the DynaTree"
Attribute TreeType.VB_ProcData.VB_Invoke_Property = ";Data"
    TreeType = m_strType
End Property

'PROPERTY READ-ONLY SelectedNodeText
'RETURNS THE TEXT VALUE OF THE SELECTED TREE NODE
Public Property Get SelectedNodeText() As String
Attribute SelectedNodeText.VB_Description = "Returns the Text value of the selected node"
Attribute SelectedNodeText.VB_ProcData.VB_Invoke_Property = ";Text"
    SelectedNodeText = TreeView1.SelectedItem.Text
End Property

'PROPERTY READ-ONLY SelectedNodeKey
'RETURNS THE KEY VALUE OF THE SELECTED TREE NODE
Public Property Get SelectedNodeKey() As String
Attribute SelectedNodeKey.VB_Description = "Returns the Key value of the selected node"
Attribute SelectedNodeKey.VB_ProcData.VB_Invoke_Property = ";Text"
    SelectedNodeKey = TreeView1.SelectedItem.Key
End Property

'PROPERTY ShowMasterFormatRoot
'IF TRUE, TREE ROOT WILL DISPLAY "MasterFormat xxxx"
'IF FALSE, TREE ROOT WILL DISPLAY "All"
Public Property Get ShowMasterFormatRoot() As Boolean
Attribute ShowMasterFormatRoot.VB_Description = "Returns/sets whether the Root's Text Value is  'MasterFormat' instead of 'All'"
Attribute ShowMasterFormatRoot.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShowMasterFormatRoot = m_blnShowMasterFormatRoot
End Property
Public Property Let ShowMasterFormatRoot(NewValue As Boolean)
    m_blnShowMasterFormatRoot = NewValue
    PropertyChanged "ShowMasterFormatRoot"
End Property

'PROPERTY Appearance
'GETS/SETS THE APPEARANCE OF THE TREEVIEW CONTROL
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not the control is painted at run-time with 3D or Flat effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = TreeView1.Appearance
End Property
Public Property Let Appearance(NewValue As AppearanceConstants)
    Select Case NewValue
    Case cc3D, ccFlat
        TreeView1.Appearance = NewValue
        TreeView1.Refresh
        PropertyChanged "Appearance"
    Case Else
        Err.Raise 380    ' Invalid Property Value
    End Select
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a repaint of the Tree object"
    UserControl.Refresh
End Sub

Private Sub Add_Assembly_Detail_Nodes(ByVal ID As String, strBlanks As String)
Dim strDisplay As String
Dim strSQL As String

On Error Resume Next
'On Error GoTo error_processing
    ' Get the last four from the Assembly Detail table itself
    ' Select only the specified type
    strSQL = "Select assembly_id, tech_desc from assembly_detail " + _
        "where assembly_id LIKE" + " '" + ID + strBlanks + _
        "' and " + m_strAssemblyTypeFldName + " = 1 " + _
        "order by assembly_id"
    rs.Open strSQL
    TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
    While Not rs.EOF
        strDisplay = rs.Fields("assembly_id") + " : "
        If Not IsNull(rs.Fields("tech_desc")) Then
            strDisplay = strDisplay + rs.Fields("tech_desc")
        End If
        Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("assembly_id")), strDisplay, 1)
        rs.MoveNext
    Wend
'Exit_sub:
'Exit Sub

'error_processing:
'MsgBox Error$
'Resume Exit_sub
End Sub

Private Sub Add_Assembly_Book_Detail_Nodes(ByVal ID As String, strBlanks As String)
Dim strDisplay As String
Dim strSQL As String
On Error Resume Next

    ' Get the last four from the Assembly Book Detail table itself
    strSQL = "Select assembly_book_id, book_desc from assembly_book_detail " + _
    "where assembly_book_id LIKE" + " '" + ID + strBlanks + "' and " _
     + m_strAssemblyTypeFldName + " = 1 " + _
    " order by assembly_book_id"
    rs.Open strSQL
    TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
    While Not rs.EOF
        strDisplay = rs.Fields("assembly_book_id") + " : "
        If Not IsNull(rs.Fields("book_desc")) Then
            strDisplay = strDisplay + rs.Fields("book_desc")
        End If
        Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("assembly_book_id")), strDisplay, 1)
        rs.MoveNext
    Wend
End Sub

Public Sub ClearTree()
Attribute ClearTree.VB_Description = "Clear all tree nodes"
    Dim blnSaveDisableRedraw As Boolean
    
    On Error Resume Next
    blnSaveDisableRedraw = DisableRedraw
    DisableRedraw = True
    TreeView1.Nodes.Clear
    DisableRedraw = blnSaveDisableRedraw
    rs.Close
    
End Sub

Private Sub FocusMatlEquipAsbly(ByVal ID As String)
    On Error Resume Next
    Select Case Len(ID)
    Case 2
        ' Top-level are already fetched
    Case 3
        ' Second-level, to be retrieved
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 2))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 2))
        End If
    Case 4
        ' First check top level
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 2))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 2))
        End If
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 4))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 4))
        End If
    Case 6        ' Third-level, to be retrieved
        ' First check top level
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 2))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 2))
        End If
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 3))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 3))
        End If
    Case 10
         ' First check top level
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 2))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 2))
        End If
        If m_strTable = MaterialHierTable _
            Or m_strTable = EquipHierTable _
            Or m_strHierType = COMMERCIAL Then
            ' Check next level
            If Left(TreeView1.Nodes(MakeKey(Left(ID, 3))).Child.Key, 1) = "Z" Then
                FetchChildren (Left(ID, 3))
            End If
        ElseIf m_strHierType = RESIDENTIAL Then
            ' Check next level
            If Left(TreeView1.Nodes(MakeKey(Left(ID, 4))).Child.Key, 1) = "Z" Then
                FetchChildren (Left(ID, 4))
            End If
        End If
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 6))).Child.Key, 1) = "Z" Then
            ' Children not already loaded
            FetchChildren (Left(ID, 6))
        End If
    End Select
    TreeView1.Nodes(MakeKey(ID)).Selected = True

End Sub

Public Sub InitData(cnShared As ADODB.Connection, Optional strType As String = TYPE_MATERIAL)
Attribute InitData.VB_Description = "Initializes the database connection and fills the Tree with data"
    Dim sTopNodeTitle As String
    Dim bSaveRedraw As Boolean
    
    bSaveRedraw = DisableRedraw
    DisableRedraw = True
    
    m_strType = strType
    sTopNodeTitle = "All"
'    strConnect = "DSN=CCD;UID=sa;PWD=;"
'    cn.Open strConnect
    Set cn = cnShared
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockReadOnly
    ' Use client cursor to enable AbsolutePosition property.
    rs.CursorLocation = adUseClient

    ' Set-up Select string
    Select Case strType
        Case TYPE_EQUIPMENT
            m_strTable = EquipHierTable
        Case TYPE_LABOR
            m_strTable = LaborHierTable
        Case TYPE_UNIT_COST, TYPE_MATERIAL
            m_strTable = MasterFormat95Table
            If m_blnShowMasterFormatRoot Then sTopNodeTitle = "MasterFormat 1995"
        Case TYPE_UNIT_COST04, TYPE_MATERIAL04
            m_strTable = MasterFormat04Table
            If m_blnShowMasterFormatRoot Then sTopNodeTitle = "MasterFormat 2004"
        Case TYPE_ASSEMBLY_COMMERCIAL, _
            TYPE_ASSEMBLY_RESI, _
            TYPE_ASSEMBLY_BK_DTL_COM, _
            TYPE_ASSEMBLY_BK_DTL_RESI
                m_strTable = AssemblyHierTable
                Select Case strType
                    Case TYPE_ASSEMBLY_COMMERCIAL, TYPE_ASSEMBLY_BK_DTL_COM
                        m_strHierType = COMMERCIAL
                        m_strAssemblyTypeFldName = COM_FIELD_NAME
                    Case Else
                        m_strHierType = RESIDENTIAL
                        m_strAssemblyTypeFldName = RESI_FIELD_NAME
                End Select
    End Select

'Return recordset from built SQL
    Select Case strType
        Case TYPE_UNIT_COST, TYPE_UNIT_COST04
            strSelect = strSelect_Start + m_strTable + strSelect_Where2    'Used until form closed
        Case TYPE_MATERIAL, TYPE_MATERIAL04
            strSelect = strSelect_Start4 + m_strTable + strSelect_Where2    'Used until form closed
        Case TYPE_ASSEMBLY_COMMERCIAL, _
            TYPE_ASSEMBLY_RESI, _
            TYPE_ASSEMBLY_BK_DTL_COM, _
            TYPE_ASSEMBLY_BK_DTL_RESI
            strSelect = strSelect_Start5 + m_strTable + strSelect_Where5    'Used until form closed
'            strSelect = strSelect_Start2 + m_strTable + strSelect_Where3 + m_strHierType + strSelect_Where4 'Used until form closed
'        Case TYPE_ASSEMBLY_BK_DTL_COM, _
'            TYPE_ASSEMBLY_BK_DTL_RESI
'            strSelect = strSelect_Start3 + m_strTable + strSelect_Where3 + m_strHierType + strSelect_Where4 'Used until form closed
        Case TYPE_BUILDING
            strSelect = "select distinct substring(bldg_id, 1,1) from bldg_detail" ' top level only
        Case TYPE_CCI_MATERIAL
            strSelect = "select distinct cci_mat_id hier_id from CCI_MATERIAL  order by cci_mat_id" ' top level only
        Case TYPE_CCI_EQUIPMENT
            strSelect = "select distinct cci_equip_id hier_id from CCI_EQUIPMENT  order by cci_equip_id" ' top level only
        Case TYPE_CCI_LABOR
            strSelect = "select distinct trade_id hier_id from LABOR_TRADE inner join CCI_LABOR on CCI_LABOR.trade_skey = labor_trade.trade_skey order by trade_id" ' top level only
        Case TYPE_PROJECT_LIST
            strSelect = "SELECT DISTINCT class_id AS hier_id From CLASSIFICATION WHERE class_system_id = 'F' "
        Case TYPE_PROJECT_ANALYSIS
            strSelect = "SELECT DISTINCT class_id AS hier_id From CLASSIFICATION WHERE class_system_id = 'F' "
        Case Else
        strSelect = strSelect_Start + m_strTable + strSelect_Where    'Used until form closed
    End Select

    TreeView1.ImageList = ImageList1
    Set Top1 = TreeView1.Nodes.Add(, , "Top", sTopNodeTitle, 1)
    Top1.Expanded = True
    If strType = TYPE_BUILDING Then     '0 doesn't work for building since the first valid id is 0
        FetchChildren "~"
    ElseIf strType = TYPE_ASSEMBLY_COMMERCIAL Or strType = TYPE_ASSEMBLY_RESI Or strType = TYPE_ASSEMBLY_BK_DTL_COM Or strType = TYPE_ASSEMBLY_BK_DTL_RESI Then
        FetchChildren 1
    Else
        FetchChildren 0
    End If
    
    DisableRedraw = bSaveRedraw
    
End Sub

Private Sub FocusLabor(ByVal ID As String)
On Error Resume Next
    Select Case Len(ID)
    Case 4
        ' Top-level are already fetched - Labor Trade Id
        TreeView1.Nodes(MakeKey(ID)).Selected = True
    Case Is > 4 And Len(ID) <= 6     'Labor Trade Id + State
        ' Second-level, to be retrieved
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 4))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 4))
        End If
        TreeView1.Nodes(MakeKey(ID)).Selected = True
    Case Is > 6    'Trade Id, State, City
         ' First check top level
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 4))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 4))
        End If
        ' Check next level
        If Left(TreeView1.Nodes(MakeKey(Left(ID, 6))).Child.Key, 1) = "Z" Then
            FetchChildren (Left(ID, 6))
        End If
        TreeView1.Nodes(MakeKey(ID)).Selected = True
    End Select

End Sub

Private Sub FocusUnitCost(ByVal ID As String)
    On Error Resume Next
    
    If Left(TreeView1.Nodes(MakeKey(Left(ID, 2))).Child.Key, 1) = "Z" Then
        FetchChildren (ID)
    End If
    TreeView1.Nodes(MakeKey(ID)).Selected = True

End Sub

Private Sub FocusProject(ByVal ID As String)
    On Error Resume Next
    'If Left(TreeView1.Nodes(MakeKey(Left(id, 2))).Child.Key, 1) = "Z" Then
    '    FetchChildren (id)
    'End If
    TreeView1.Nodes(MakeKey(ID)).Selected = True
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
    Screen.MousePointer = vbHourglass
    If Node.Expanded = True Then
        ' Only fetch if we haven't already
        If Left(Node.Child.Key, 1) = "Z" Then
            FetchChildren (Right(Node.Key, Len(Node.Key) - 1))
        End If
    End If
    Node.Selected = True
    Screen.MousePointer = vbNormal
End Sub

Public Sub FocusItem(ByVal ID As String)
Attribute FocusItem.VB_Description = "Selects a node in the Tree"
    On Error Resume Next

    If Right(ID, 1) = "*" Then 'strip out wildcard
        ID = Left(ID, Len(ID) - 1)
    End If

    Select Case m_strType
        Case TYPE_MATERIAL, TYPE_MATERIAL04
            If Left(ID, 1) = "M" Then
                ID = Right(ID, Len(ID) - 1)
            End If
            FocusMatlEquipAsbly ID
        Case TYPE_EQUIPMENT, TYPE_ASSEMBLY_COMMERCIAL, TYPE_ASSEMBLY_BK_DTL_COM, _
                TYPE_ASSEMBLY_RESI, TYPE_ASSEMBLY_BK_DTL_RESI
            FocusMatlEquipAsbly ID
        Case TYPE_LABOR
            FocusLabor ID
        Case TYPE_UNIT_COST, TYPE_UNIT_COST04
            FocusUnitCost ID
        Case TYPE_PROJECT_ANALYSIS, TYPE_PROJECT_LIST
            FocusProject ID
    End Select

End Sub

Private Function MakeKey(ByVal ID As String) As String
    Dim sNewKey As String
    
    sNewKey = "K" + ID
    m_strLastKeyMade = sNewKey
    MakeKey = sNewKey

End Function

Private Function MakeValue(Row As ADODB.Fields) As String
    Select Case m_strType
    '    Case TYPE_ASSEMBLY_BK_DTL_COM, _
    '        TYPE_ASSEMBLY_BK_DTL_RESI
    '            MakeValue = Row("hier_id").Value + " : " + Trim(Row("hier_desc")) + "  (" + Format(Row("subord_count_abd"), "#") + ")"
    '    Case TYPE_ASSEMBLY_COMMERCIAL, _
    '        TYPE_ASSEMBLY_RESI
    '            MakeValue = Row("hier_id").Value + " : " + Trim(Row("uni2_desc")) + "  (0)"
        Case TYPE_MATERIAL, TYPE_MATERIAL04
            MakeValue = Row("hier_id").Value + " : " + Trim(Row("hier_desc")) + "  (" + Format(Row("mat_subord_count"), "#") + ")"
        Case Else
            If Trim(Row("hier_desc")) <> "" Then
                MakeValue = Trim(Row("hier_id").Value) + " : " + Trim(Row("hier_desc")) + "  (" + Format(Row("subord_count"), "#") + ")"
            Else
                ' There is no hier description
                MakeValue = Trim(Row("hier_id").Value) + "  (" + Format(Row("subord_count"), "#") + ")"
            End If
    End Select
End Function

Private Function FetchChildren(ByVal ID As String) As String
    Dim SelectStmt As String
    Dim strSelect2 As String
    Dim lngLevelID As Long
    Dim strStart As String
    Dim strEnd As String
    Dim strDisplay As String
    Dim iErrorCount As Integer
    Dim sClassSystem As String
    Dim sSelectID As String
    Dim iNodeLevel As Integer
    Dim sKey As String
    Dim i As Integer

    On Error GoTo Error_Processing
    Screen.MousePointer = vbHourglass

    If m_strType = TYPE_UNIT_COST Or m_strType = TYPE_MATERIAL Or _
        m_strType = TYPE_UNIT_COST04 Or m_strType = TYPE_MATERIAL04 Then
    Select Case m_strType
    Case TYPE_UNIT_COST, TYPE_UNIT_COST04
        'Use level for processing
        rs.Open "SELECT level_id, unit_cost_id_start, unit_cost_id_end FROM " + m_strTable + " where hier_id='" + ID + "'", cn
        If Not (rs.EOF And rs.BOF) Then
            lngLevelID = rs.Fields("level_id") + 1
            strStart = rs.Fields("unit_cost_id_start")
            strEnd = rs.Fields("unit_cost_id_end")
        Else
            lngLevelID = 1
        End If
        rs.Close
        If lngLevelID = 1 Then  'Fill top level
            strSelect2 = strSelect + CStr(lngLevelID) + " order by hier_id"
            rs.Open strSelect2, cn
            While Not rs.EOF
                ' Add Node to tree
                Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
        ElseIf lngLevelID = 5 Then  'Fill detail level from Unit_Cost_Detail
            If m_strType = TYPE_UNIT_COST04 Then
                'GET MASTERFORMAT 2004 UNIT_COST_DETAIL_EXT
                strSelect2 = "SELECT unit_cost_id, tech_desc FROM Unit_cost_detail_ext WHERE unit_cost_id BETWEEN '" + strStart + "' and '" + strEnd + "' ORDER BY unit_cost_id"
            Else
                'GET MASTERFORMAT 1995 UNIT_COST_DETAIL
                strSelect2 = "SELECT unit_cost_id, tech_desc FROM Unit_cost_detail WHERE unit_cost_id BETWEEN '" + strStart + "' and '" + strEnd + "' ORDER BY unit_cost_id"
            End If
            rs.Open strSelect2, cn
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            While Not rs.EOF
                If m_strType = TYPE_UNIT_COST04 Then
                    strDisplay = Format(rs.Fields("unit_cost_id"), "@@@@@@.@@ @@@@") & " : "
                Else
                    strDisplay = Format(rs.Fields("unit_cost_id"), "@@@@@@@@@@@@") + " : "
                End If
                If Not IsNull(rs.Fields("tech_desc")) Then
                    strDisplay = strDisplay + rs.Fields("tech_desc")
                End If
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("unit_cost_id")), strDisplay, 1)
                rs.MoveNext
            Wend
        ElseIf lngLevelID = 4 And m_strType = TYPE_UNIT_COST04 Then
            strSelect2 = strSelect + CStr(lngLevelID) + "and hier_id BETWEEN '" + strStart + "' and '" + strEnd + "' order by hier_id"
            rs.Open strSelect2, cn
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            While Not rs.EOF
                strDisplay = Format(Trim(rs.Fields("hier_id").Value), "@@@@@@.@@") + " : " + Trim(rs.Fields("hier_desc")) + "  (" + Format(rs.Fields("subord_count"), "#") + ")"
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), strDisplay, 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
        Else    'Fill next level
            strSelect2 = strSelect + CStr(lngLevelID) + "and hier_id BETWEEN '" + strStart + "' and '" + strEnd + "' order by hier_id"
            rs.Open strSelect2, cn
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            While Not rs.EOF
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
        End If
        rs.Close
    Case TYPE_MATERIAL, TYPE_MATERIAL04
        'Use level for processing
        SelectStmt = "Select level_id, mat_id_start, mat_id_end from " + m_strTable + " where hier_id='" + ID + "'"
        rs.Open SelectStmt, cn
        If Not (rs.EOF And rs.BOF) Then
            lngLevelID = rs.Fields("level_id") + 1
            strStart = Right(rs.Fields("mat_id_start"), Len(rs.Fields("mat_id_start")) - 1) 'need to strip out M
            strEnd = Right(rs.Fields("mat_id_end"), Len(rs.Fields("mat_id_end")) - 1) 'need to strip out M
        Else
            lngLevelID = 1
        End If
        rs.Close
        If lngLevelID = 1 Then  'Fill top level
            strSelect2 = strSelect + CStr(lngLevelID) + " order by hier_id"
            rs.Open strSelect2, cn
            While Not rs.EOF
                ' Add Node to tree
                Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
        ElseIf lngLevelID = 5 Then  'Fill detail level from Material
            Dim dtmToday As Date
            dtmToday = Date
            SelectStmt = "Select mat_id, tech_desc, avg(mp.list_price) as average from Material, Material_price as mp where material.mat_skey = mp.mat_skey and mp.start_date <= '" + Format(dtmToday, "mm/dd/yyyy") + "' and mp.term_date >= '" + Format(dtmToday, "mm/dd/yyyy") + "' and mat_id between 'M" + strStart + "' and 'M" + strEnd + "' group by mat_id, tech_desc"
            rs.Open SelectStmt, cn
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            While Not rs.EOF
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("mat_id")), rs.Fields("mat_id") + " : " + rs.Fields("tech_desc") + "  (" + Format(rs.Fields("average"), "#.##") + ")", 1)
                rs.MoveNext
            Wend
        Else    'Fill next level
            strSelect2 = strSelect + CStr(lngLevelID) + "and hier_id BETWEEN '" + strStart + "' and '" + strEnd + "' order by hier_id"
            rs.Open strSelect2, cn
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            While Not rs.EOF
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
        End If
        rs.Close
    End Select
    ElseIf m_strType = TYPE_CCI_MATERIAL Or m_strType = TYPE_CCI_EQUIPMENT Or m_strType = TYPE_CCI_LABOR Then
       rs.Open strSelect, cn
        While Not rs.EOF
            Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey(rs.Fields("hier_id")), rs.Fields("hier_id"), 1)
            rs.MoveNext
        Wend
        rs.Close
    ElseIf m_strType = TYPE_BUILDING Then
        'First pass - 0, next pass 1st letter - len does not work
        Select Case ID
    '    Case "~" 'Fill first position for building
    '        strSelect = "select distinct substring(bd_s.bldg_id, 1,1) as hier_id, '' as hier_desc, count(dtl.bldg_id) as subord_count from bldg_detail bd_s " + _
    '            "inner join bldg_detail dtl on substring(bd_s.bldg_id, 1,1) = substring(dtl.bldg_id, 1,1)" + _
    '            "group by bd_s.bldg_id"
        
        Case "~" 'Fill first position for building
            'strSelect = "select distinct bd_s.type_code as hier_id, " _
                & "bd_s.bldg_category as hier_desc, count(dtl.bldg_category) " _
                & "as subord_count from bldg_detail bd_s inner join bldg_detail dtl " _
                & "on bd_s.bldg_category = dtl.bldg_category group by bd_s.type_code, " _
                & "bd_s.bldg_category, bd_s.bldg_id"
            '
            '   Initial select needs to be by Commercial or Residential
            strSelect = "SELECT DISTINCT bd.type_code AS hier_id, " _
                & "CASE WHEN bd.type_code = 'C' THEN 'Commercial' " _
                & "Else 'Residential' END AS hier_desc, " _
                & "COUNT(dtl.type_code) As subord_count " _
                & "FROM bldg_detail bd INNER JOIN bldg_detail dtl " _
                & "ON bd.type_code = dtl.type_code " _
                & "GROUP BY bd.type_code, bd.bldg_category , bd.bldg_id"
        Case "C" ', "R"
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            '
            '   Now determine if they click on top level Com/Resi or a Category below.
            '   Does by category
            strSelect = "SELECT DISTINCT bd.bldg_category AS hier_id, " _
                & "'' AS hier_desc, COUNT(dtl.bldg_category) " _
                & "AS subord_count FROM bldg_detail bd INNER JOIN bldg_detail dtl " _
                & "ON bd.bldg_category = dtl.bldg_category WHERE bd.type_code = '" & ID _
                & "' GROUP BY bd.type_code, bd.bldg_category, bd.bldg_id"
    
        Case "R"
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            '
            '   Now determine if they click on top level Com/Resi or a Category below.
            '   Does by category
            strSelect = "SELECT DISTINCT bd.bldg_category AS hier_id, " _
                & "'' AS hier_desc, COUNT(dtl.bldg_category) " _
                & "AS subord_count FROM bldg_detail bd INNER JOIN bldg_detail dtl " _
                & "ON bd.bldg_category = dtl.bldg_category WHERE bd.type_code = '" & ID _
                & "' AND bd.bldg_category = 'Economy' GROUP BY bd.type_code, bd.bldg_category, bd.bldg_id"
                
            rs.Open strSelect, cn
            Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(Trim(rs.Fields("hier_id"))), MakeValue(rs.Fields), 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey(Trim(rs.Fields("hier_id"))), tvwChild, "Z" + rs.Fields("hier_id").Value)
            rs.Close
            '
            '   Now determine if they click on top level Com/Resi or a Category below.
            '   Does by category
            strSelect = "SELECT DISTINCT bd.bldg_category AS hier_id, " _
                & "'' AS hier_desc, COUNT(dtl.bldg_category) " _
                & "AS subord_count FROM bldg_detail bd INNER JOIN bldg_detail dtl " _
                & "ON bd.bldg_category = dtl.bldg_category WHERE bd.type_code = '" & ID _
                & "' AND bd.bldg_category = 'Average' GROUP BY bd.type_code, bd.bldg_category, bd.bldg_id"
                
            rs.Open strSelect, cn
            Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(Trim(rs.Fields("hier_id"))), MakeValue(rs.Fields), 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey(Trim(rs.Fields("hier_id"))), tvwChild, "Z" + rs.Fields("hier_id").Value)
            rs.Close
            '
            '   Now determine if they click on top level Com/Resi or a Category below.
            '   Does by category
            strSelect = "SELECT DISTINCT bd.bldg_category AS hier_id, " _
                & "'' AS hier_desc, COUNT(dtl.bldg_category) " _
                & "AS subord_count FROM bldg_detail bd INNER JOIN bldg_detail dtl " _
                & "ON bd.bldg_category = dtl.bldg_category WHERE bd.type_code = '" & ID _
                & "' AND bd.bldg_category = 'Custom' GROUP BY bd.type_code, bd.bldg_category, bd.bldg_id"
                
            rs.Open strSelect, cn
            Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(Trim(rs.Fields("hier_id"))), MakeValue(rs.Fields), 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey(Trim(rs.Fields("hier_id"))), tvwChild, "Z" + rs.Fields("hier_id").Value)
            rs.Close
            '
            '   Now determine if they click on top level Com/Resi or a Category below.
            '   Does by category
            strSelect = "SELECT DISTINCT bd.bldg_category AS hier_id, " _
                & "'' AS hier_desc, COUNT(dtl.bldg_category) " _
                & "AS subord_count FROM bldg_detail bd INNER JOIN bldg_detail dtl " _
                & "ON bd.bldg_category = dtl.bldg_category WHERE bd.type_code = '" & ID _
                & "' AND bd.bldg_category = 'Luxury' GROUP BY bd.type_code, bd.bldg_category, bd.bldg_id"
              
            rs.Open strSelect, cn
            Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(Trim(rs.Fields("hier_id"))), MakeValue(rs.Fields), 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey(Trim(rs.Fields("hier_id"))), tvwChild, "Z" + rs.Fields("hier_id").Value)
            rs.Close
                
        Case Else
                TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
                'strSelect = "select bldg_id as hier_id, bldg_desc as hier_desc, 1 as subord_count from bldg_detail where bldg_id like '" + Left(id, 1) + "%'"
                strSelect = "select bldg_id as hier_id, bldg_desc as hier_desc, 1 as subord_count from bldg_detail where bldg_category like '" + ID + "%'"
        End Select
        If ID <> "R" Then
            rs.Open strSelect, cn
            While Not rs.EOF
                ' Add Node to tree
                Select Case ID
                Case "~" 'Fill first position for building
                    Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey(Trim(rs.Fields("hier_id"))), MakeValue(rs.Fields), 1)
                    ' Add fake node below so parent can be expanded
                    Set Top1 = TreeView1.Nodes.Add(MakeKey(Trim(rs.Fields("hier_id"))), tvwChild, "Z" + rs.Fields("hier_id").Value)
                Case "C", "R"
                    Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(Trim(rs.Fields("hier_id"))), MakeValue(rs.Fields), 1)
                    ' Add fake node below so parent can be expanded
                    Set Top1 = TreeView1.Nodes.Add(MakeKey(Trim(rs.Fields("hier_id"))), tvwChild, "Z" + rs.Fields("hier_id").Value)
                Case Else
                    Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(Trim(rs.Fields("hier_id"))), MakeValue(rs.Fields), 1)
                End Select
                rs.MoveNext
            Wend
            rs.Close
        End If
    ElseIf m_strType = TYPE_PROJECT_LIST Or m_strType = TYPE_PROJECT_ANALYSIS Then
        Select Case ID
        Case 0
            'NOTE FOR FUTURE REVISION
            'THIS QUERY RUNS REALLY SLOW
            strSelect = "SELECT  C.class_id AS hier_id, C.class_desc AS hier_desc, COUNT(P.proj_bldg_skey) as subord_count " & _
                        " FROM  CLASSIFICATION C LEFT OUTER JOIN PROJECT_BUILDING_DETAIL P " & _
                        "        ON C.class_system_id = P.facility1_class_system_id AND C.class_id = P.facility1_class_id " & _
                        " WHERE  C.class_system_id = 'F' AND C.class_desc is not null GROUP BY C.class_id, C.class_desc, c.sort_order ORDER BY C.sort_order"
        Case Else
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            If m_strType = TYPE_PROJECT_LIST Then
                strSelect = "SELECT  P.proj_bldg_skey as hier_id, M.exterior_material_desc, year(P.bid_date) as year, P.gross_floor_area, P.gross_floor_area_uom" & _
                            "  FROM  PROJECT_BUILDING_DETAIL P INNER JOIN EXTERIOR_MATERIAL M ON P.exterior_mat_code = M.exterior_mat_code " & _
                            " WHERE  P.facility1_class_id = '" & ID & "' and P.facility1_class_system_id = 'F'"
            ElseIf m_strType = TYPE_PROJECT_ANALYSIS Then
            '    strSelect = "SELECT DISTINCT class_desc, class_id as hier_id," & _
            '                "  (SELECT COUNT(PC.proj_bldg_skey)" & _
            '                "     FROM PROJ_BLDG_COMPONENT_COST PC INNER JOIN" & _
            '                "          PROJECT_BUILDING_DETAIL D ON PC.proj_bldg_skey = D.proj_bldg_skey" & _
            '                "    WHERE PC.class_id = C.class_id AND D.facility1_class_id = '" & id & "') AS numproj " & _
            '                "  FROM CLASSIFICATION C " & _
            '                " WHERE class_system_id = 'P1' AND class_desc is not null ORDER BY C.class_id"
                strSelect = "EXEC sp_select_project_analysis_treeview @id = '" & ID & "'"
            End If
        End Select
        rs.Open strSelect, cn
        DisableRedraw = True
        If ID = 0 Then
            Do While Not rs.EOF
                Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey(rs.Fields("hier_id")), Trim(rs.Fields("hier_id")) & " : " & Trim(rs.Fields("hier_desc")) & " (" & rs.Fields("subord_count") & ")", 1)
                 ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Loop
        Else
            If m_strType = TYPE_PROJECT_LIST Then
                Do While Not rs.EOF
                    Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), rs.Fields("hier_id") & " : " & rs.Fields("exterior_material_desc") & " " & rs.Fields("year") & " " & FormatNumber(rs.Fields("gross_floor_area"), 0) & " " & rs.Fields("gross_floor_area_uom"), 1)
                    rs.MoveNext
                Loop
            ElseIf m_strType = TYPE_PROJECT_ANALYSIS Then
                Do While Not rs.EOF
                    Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, ID & "-" & MakeKey(rs.Fields("class_id")), rs.Fields("class_desc") & " (" & rs.Fields("num_proj") & ")", 1)
                    rs.MoveNext
                Loop
            End If
        End If
        DisableRedraw = False
        rs.Close
    ElseIf m_strType = TYPE_ASSEMBLY_COMMERCIAL Or m_strType = TYPE_ASSEMBLY_RESI Or m_strType = TYPE_ASSEMBLY_BK_DTL_COM Or m_strType = TYPE_ASSEMBLY_BK_DTL_RESI Then
        'Use level for processing
        strSelect2 = "Select uni2_level, assembly_id_start, assembly_id_end from " + AssemblyHierTable + " where uni2_category_id='" + ID + "'"
        rs.Open "Select uni2_level, assembly_id_start, assembly_id_end from " + AssemblyHierTable + " where uni2_category_id='" + ID + "'", cn
        If Not (rs.EOF And rs.BOF) Then
            lngLevelID = rs.Fields("uni2_level") + 1
            strStart = rs.Fields("assembly_id_start")
            strEnd = rs.Fields("assembly_id_end")
        Else
            lngLevelID = 1
        End If
        rs.Close
        If lngLevelID = 1 Then  'Fill top level
            strSelect2 = strSelect + CStr(lngLevelID) + " order by uni2_category_id"
            rs.Open strSelect2, cn
            While Not rs.EOF
                ' Add Node to tree
                Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
        Else    'Fill next level
            strSelect2 = strSelect + CStr(lngLevelID) + "and uni2_category_id BETWEEN '" + strStart + "' and '" + strEnd + "' order by uni2_category_id"
           rs.Open strSelect2, cn
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            While Not rs.EOF
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
        End If
        rs.Close
    ElseIf m_strType = TYPE_CCI_INDEX Then
        If ID = "0" Then  'Fill top level
            ' Add MF & UF Nodes to tree
            Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey("MF"), "MasterFormat", 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey("MF"), tvwChild, "Z" + "MasterFormat")
            ' 9/23/2005 RTD - ADDED SUPPORT FOR RES CLASS SYSTEM
            Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey("R1"), "Residential", 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey("R1"), tvwChild, "Z" + "R1")
            Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey("U2"), "Uniformat", 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey("U2"), tvwChild, "Z" + "U2")
        Else
            If ID = "MF" Or ID = "U2" Or ID = "R1" Then
                sClassSystem = ID
                lngLevelID = 1
            Else
                'Retrieve the class system
                Set Top1 = TreeView1.Nodes(MakeKey(ID))
                iNodeLevel = 0
                Do While Top1.Key <> "Top"
                    iNodeLevel = iNodeLevel + 1
                    Set Top1 = Top1.Parent
                Loop
                Set Top1 = TreeView1.Nodes(MakeKey(ID))
                For i = 1 To iNodeLevel - 1
                    Set Top1 = Top1.Parent
                Next i
                ' 9/23/2005 RTD - ADDED SUPPORT FOR RES CLASS SYSTEM
                If Left(Top1.Key, 1) = "K" Then
                    sClassSystem = Mid(Top1.Key, 2)
                End If
                'If Top1.Key = "KMF" Then
                '    sClassSystem = "MF"
                'Else
                '    sClassSystem = "U2"
                'End If
                'Use level for processing
                rs.Open "Select hierarchy_level_code from cci_index_hierarchy where class_id='" + ID + "'", cn
                If Not (rs.EOF And rs.BOF) Then
                    lngLevelID = rs.Fields("hierarchy_level_code") + 1
                End If
                rs.Close
            End If
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key) 'Remove dummy node
            sSelectID = Trim(ID)
            If Right(sSelectID, 1) = "." Then  ' Strip trailing period
                sSelectID = Left(sSelectID, Len(sSelectID) - 1)
            Else
                sSelectID = sSelectID
            End If
            If lngLevelID = 1 Then
                strSelect = "select class_id, class_description from cci_index_hierarchy where class_system_id = '" + sClassSystem + "'" + _
                    " and hierarchy_level_code = '" + CStr(lngLevelID) + "'"
            Else 'for levels after 1, select next hier level with matching id of previous
                strSelect = "select class_id, class_description from cci_index_hierarchy " + _
                "where class_system_id = '" + sClassSystem + "'" + _
                    " and hierarchy_level_code = '" + CStr(lngLevelID) + "'" + _
                    " and substring(class_id, 1, " + CStr(Len(sSelectID)) + ") = '" + sSelectID + "'"
            End If
            rs.Open strSelect, cn
            While Not rs.EOF
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("class_id")), rs.Fields("class_description").Value, 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("class_id")), tvwChild, "Z" + rs.Fields("class_description").Value)
                rs.MoveNext
            Wend
            rs.Close
        End If
    ElseIf m_strType = TYPE_CCI_MAT_EQUIP Then
        If ID = "0" Then  'Fill top level
            ' Add MF & UF Nodes to tree
            Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey("MA"), "Material", 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey("MA"), tvwChild, "Z" + "Material")
            Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey("EQ"), "Equipment", 1)
            ' Add fake node below so parent can be expanded
            Set Top1 = TreeView1.Nodes.Add(MakeKey("EQ"), tvwChild, "Z" + "EQ")
        Else
            If ID = "MA" Then   'Add Material nodes
                strSelect = "select distinct cci_mat_id hier_id from CCI_MATERIAL  order by cci_mat_id" ' top level only
                sKey = MakeKey("MA")
            End If
            If ID = "EQ" Then   'Add Equipment Nodes
                strSelect = "select distinct cci_equip_id hier_id from CCI_EQUIPMENT  order by cci_equip_id" ' top level only
                sKey = MakeKey("EQ")
            End If
            If ID = "MA" Or ID = "EQ" Then    'If Mat or Equip, fill nodes.
                TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key) 'Remove dummy node
                rs.Open strSelect, cn
                While Not rs.EOF
                    Set Top1 = TreeView1.Nodes.Add(sKey, tvwChild, MakeKey(rs.Fields("hier_id")), rs.Fields("hier_id"), 1)
                    rs.MoveNext
                Wend
                rs.Close
            End If
        End If
    Else
        Select Case Len(ID)
        Case 1  'Fill first 2 positions for material, equipment, 4 for labor
            Select Case m_strType
                Case TYPE_LABOR      'Labor
                    ' Get the first 4 characters
                    strSelect2 = strSelect + "'____' order by hier_id"
                Case Else   'Material and Equipment and Assemblies and Assembly Book
                    ' Get the first 2 characters
                    strSelect2 = strSelect + "'__' order by hier_id"
            End Select
            rs.Open strSelect2, cn
            While Not rs.EOF
                ' Add Node to tree
                Set Top1 = TreeView1.Nodes.Add("Top", tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
            rs.Close
        Case 2
            Select Case m_strHierType
            Case RESIDENTIAL    'get the next 2 characters
               rs.Open strSelect + "'" + ID + "__' order by hier_id"
            Case Else
            ' Get the next 1 character
               rs.Open strSelect + "'" + ID + "_' order by hier_id"
            End Select
    
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            While Not rs.EOF
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
            rs.Close
        Case 3
            ' Get the next 3 characters
            rs.Open strSelect + "'" + ID + "___' order by hier_id"
            TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
            While Not rs.EOF
                Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                ' Add fake node below so parent can be expanded
                Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                rs.MoveNext
            Wend
            rs.Close
        Case 4
            Select Case m_strHierType
            Case RESIDENTIAL    'Add the last six from the detail table
                Select Case m_strType
                Case TYPE_ASSEMBLY_RESI
                    Add_Assembly_Detail_Nodes ID, "______"
                Case TYPE_ASSEMBLY_BK_DTL_RESI
                    Add_Assembly_Book_Detail_Nodes ID, "______"
                End Select
            Case Else
                ' Get the next 2 characters - Labor State
                rs.Open strSelect + "'" + ID + "__' order by hier_id"
                TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
                While Not rs.EOF
                    Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                    ' Add fake node below so parent can be expanded
                    Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                    rs.MoveNext
                Wend
            End Select
            rs.Close
        Case 6
            Select Case m_strType
                Case TYPE_EQUIPMENT
                    ' Get the last four from the Equipment table itself
                    strSelect2 = "Select equip_id, tech_desc from Equipment where equip_id LIKE" + " '" + ID + "____' order by equip_id"
                    rs.Open strSelect2
                    TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
                    While Not rs.EOF
                        strDisplay = rs.Fields("equip_id") + " : "
                        If Not IsNull(rs.Fields("tech_desc")) Then
                            strDisplay = strDisplay + rs.Fields("tech_desc")
                        End If
                        Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("equip_id")), strDisplay, 1)
                        rs.MoveNext
                    Wend
                    rs.Close
                Case TYPE_LABOR
                    ' Get the next 7+ characters - Labor City
                    rs.Open strSelect + "'" + ID + "_%' order by hier_id"
                    TreeView1.Nodes.Remove (TreeView1.Nodes(MakeKey(ID)).Child.Key)
                    While Not rs.EOF
                        Set Top1 = TreeView1.Nodes.Add(MakeKey(ID), tvwChild, MakeKey(rs.Fields("hier_id")), MakeValue(rs.Fields), 1)
                        ' Add fake node below so parent can be expanded
    '                    Set Top1 = TreeView1.Nodes.Add(MakeKey(rs.Fields("hier_id")), tvwChild, "Z" + rs.Fields("hier_id").Value)
                        rs.MoveNext
                    Wend
                'Add last four positions from assembly_detail
                Case TYPE_ASSEMBLY_COMMERCIAL
                    Add_Assembly_Detail_Nodes ID, "____"
                Case TYPE_ASSEMBLY_BK_DTL_COM
                    Add_Assembly_Book_Detail_Nodes ID, "____"
            End Select
            rs.Close
        End Select
    End If

Exit_Sub:
    Screen.MousePointer = vbNormal
    Exit Function

Error_Processing:
    'UPDATED 6/27/2005 RTD
    '35601: Element not found (trying to remove node while changing views)
    '3704:  Object not closed - next pass resets
    '3705:  Operation requested not allowed if the object is open.
    '91:    Object not set - next pass resets
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case 3705
        If iErrorCount > 3 Then
            Resume Exit_Sub
        Else
            iErrorCount = iErrorCount + 1
            rs.Close
            Resume Next
        End If
    Case 3704, 91, 35601
        ' object closed or not set, ignore
        Resume Exit_Sub
    Case 35602
        MsgBox "DynaTree:FetchChildren() Nodes.Add Error" & vbCrLf & _
                CStr(Err.Number) & ": " & Err.Description & vbCrLf & _
                "Key: " & m_strLastKeyMade, vbInformation + vbOKOnly
        Resume Exit_Sub
    Case Else
        MsgBox "DynaTree:FetchChildren() Error" & vbCrLf & _
                CStr(Err.Number) & ": " & Err.Description, vbInformation + vbOKOnly
        Resume Exit_Sub
    End Select

End Function

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    RaiseEvent NodeSelected(Right(Node.Key, Len(Node.Key) - 1))
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'A PARENT FORM PROPERTY HAS CHANGED
    Select Case "PropertyName"
    Case "Font"
        Set TreeView1.Font = Ambient.Font
    End Select
End Sub

Private Sub UserControl_InitProperties()
    Set TreeView1.Font = Ambient.Font
    m_blnShowMasterFormatRoot = False
    m_blnDisableRedraw = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_blnShowMasterFormatRoot = PropBag.ReadProperty("ShowMasterFormatRoot", False)
    m_blnDisableRedraw = PropBag.ReadProperty("DisableRedraw", False)
    TreeView1.Appearance = PropBag.ReadProperty("Appearance", cc3D)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ShowMasterFormatRoot", m_blnShowMasterFormatRoot, False
    PropBag.WriteProperty "DisableRedraw", m_blnDisableRedraw, False
    PropBag.WriteProperty "Appearance", TreeView1.Appearance, cc3D
End Sub

Private Sub UserControl_Resize()
    TreeView1.Move TreeView1.Left, TreeView1.Top, ScaleWidth - 2 * TreeView1.Left, ScaleHeight - 2 * TreeView1.Top
End Sub

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Displays Control Version Information"
Attribute ShowAbout.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub

