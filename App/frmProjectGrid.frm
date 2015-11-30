VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmProjectGrid 
   Caption         =   "Project Maintenance"
   ClientHeight    =   6900
   ClientLeft      =   3075
   ClientTop       =   3600
   ClientWidth     =   11805
   Icon            =   "frmProjectGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   11805
   Begin VB.CommandButton buDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9960
      TabIndex        =   38
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton buNew 
      Caption         =   "&New"
      Height          =   615
      Left            =   8520
      TabIndex        =   37
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9480
      TabIndex        =   33
      Top             =   100
      Width           =   2175
      Begin VB.OptionButton optSort 
         Caption         =   "ID"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   35
         Top             =   0
         Width           =   495
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Class"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   34
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Sort By:"
         Height          =   255
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.TextBox txtProjectID 
      Height          =   285
      Left            =   7800
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox comboClassification 
      Height          =   315
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.CheckBox ckUse 
      Caption         =   "Use"
      Height          =   375
      Left            =   10920
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.CheckBox ckWrap 
      Caption         =   "Row Wrap Check"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2950
      Width           =   255
   End
   Begin VB.CommandButton buSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   495
      Left            =   7800
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame fmCostArea 
      Height          =   1335
      Left            =   9240
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
      Begin VB.TextBox txtCostMax 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   700
      End
      Begin VB.TextBox txtAreaMax 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   840
         Width           =   700
      End
      Begin VB.TextBox txtAreaMin 
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Top             =   840
         Width           =   700
      End
      Begin VB.TextBox txtCostMin 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   700
      End
      Begin VB.Label lbMax 
         Alignment       =   2  'Center
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         Top             =   240
         Width           =   700
      End
      Begin VB.Label lbMin 
         Alignment       =   2  'Center
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   700
      End
      Begin VB.Label Label2 
         Caption         =   "Area"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   405
      End
   End
   Begin VB.OptionButton opEQYB 
      Caption         =   "Option1"
      Height          =   255
      Left            =   9240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.OptionButton opGTYB 
      Caption         =   "Option1"
      Height          =   255
      Left            =   9720
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.OptionButton opLTYB 
      Caption         =   "Option1"
      Height          =   255
      Left            =   8760
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.ComboBox comboYearBuilt 
      Height          =   315
      Left            =   7800
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox comboState 
      Height          =   315
      Left            =   7800
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton buAnalysis 
      Caption         =   "Project &Analysis"
      Height          =   615
      Left            =   1440
      TabIndex        =   22
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton buParameters 
      Caption         =   "&Parameters"
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton buUpdate 
      Caption         =   "&Update"
      Height          =   615
      Left            =   7080
      TabIndex        =   23
      Top             =   6120
      Width           =   1215
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridProject 
      Height          =   2715
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3240
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   4789
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(1)._MinWidth=1241540"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      DataMode        =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(23)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(44)  =   "Named:id=29:Normal"
      _StyleDefs(45)  =   ":id=29,.parent=0"
      _StyleDefs(46)  =   "Named:id=30:Heading"
      _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=30,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=31:Footing"
      _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=32:Selected"
      _StyleDefs(52)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=33:Caption"
      _StyleDefs(54)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(55)  =   "Named:id=34:HighlightRow"
      _StyleDefs(56)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(57)  =   "Named:id=35:EvenRow"
      _StyleDefs(58)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=36:OddRow"
      _StyleDefs(60)  =   ":id=36,.parent=29"
      _StyleDefs(61)  =   "Named:id=39:RecordSelector"
      _StyleDefs(62)  =   ":id=39,.parent=30"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=29"
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Project ID:"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lbStatus 
      Height          =   255
      Left            =   6240
      TabIndex        =   31
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label lbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lbEQYB 
      Caption         =   "="
      Height          =   255
      Left            =   9480
      TabIndex        =   27
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lbGTYB 
      Caption         =   ">="
      Height          =   255
      Left            =   9960
      TabIndex        =   26
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lbLTYB 
      Caption         =   "<="
      Height          =   255
      Left            =   9000
      TabIndex        =   25
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lbYearBuilt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Year Built:"
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lbState 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbClassification 
      Caption         =   "Classification:"
      Height          =   255
      Left            =   6840
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11640
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmProjectGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <modulename> frmProjectGrid</modulename>
' <functionname>General (Main) </functionname>
'
' <summary> PROJECT MAINTENANCE
'
' Provides u/i permitting user to do the following:
'
'Add or Change PROJECT (DIV 17) INFORMATION.
'
'Search Criteria:
'
'"   Sort By
'ID              blows up!
'Class
'WARNING: doesn 't work!!!
'"   Year Built:         doesn't work?
'"   State               doesn't work?
'"   Project ID:             works!
'"   Min Max
'o   Cost            doesn't work?
'o   Area            doesn't work?
'
'(Major function buttons)
'
'1.  Search                      (buSearch_Click() )
'Designed not to work w/o the above "search criteria"
'2.  Parameters                  (frmProject)
'3.  Project Analysis                (frmProjectAnalysis)
'4.  Update                  (CProjectMap.Update() /
' m_objGridMap.Update())
'5.  New                     (frmProject.frm)
'6.  Delete                      (CProjectMap.Delete() )
'
'HELPER Class: CProjectMap.Cls
'
'</summary>
'
'<seealso> CProjectMap.cls </seealso>
'<seealso>frmProject.frm</seealso>
'<seealso>frmProjectRpt.frm</seealso>
'
'
' <datastruct>m_objGridMap</datastruct>
'<datastruct>m_rec</datastruct>
'
' <storedprocedurename> sp_select_project_list</storedprocedurename>
'
'
'
'<returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'exec sp_select_building @type_code = '%', @bldg_category = '%', @bldg_id = '001', @bldg_desc = '%' </code>
' <code>
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CProjectMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>



'
'   Class to handle grid
Dim m_objGridMap As New CProjectMap
'
'   Is this the first search we have made
'   on this screen.
Dim m_blnFirstSearch As Boolean
'
'   Recordset to hold query results
Dim m_rec As New ADODB.RecordSet
'
'   True if the Update had errors, used in QueryUnload
Dim m_blnWereErrors As Boolean
'
'   Keeps up with the field that last had focus when form
'   is deactivate, so when activated can set focus.
Dim m_strCurrentFormControl As String
'
'   Notifies that it wants to see changes.
Dim sEventSubscriberID As String
'
'   Keep track of last completed select statement
Dim m_lastSQL As String
Dim m_ClassID As String
'
'   List of Project components
Dim aryClassList() As String

Dim m_blnSortByDesc As Boolean      ' Used for Class Combo sort by Description
Dim m_counter As Integer
Private blnDelete As Boolean

Const USEBOOKMARK = 1
Const USECOORD = 0

' 9/9/2005 RTD
' IF WE RECEIVE AN esnProjectRecordUpdated EVENT, THEN REQUERY THE GRID
Public Sub EventNotify(eNotifyType As EEventSubscriberNotifyType, sAffectedRecordIdentifier As String)
    Dim varBookmark
    
    On Error Resume Next
    ' If the record that was updated is in our grid results
    ' we need to refresh.
    If eNotifyType = esnProjectRecordUpdated Then
        If txtProjectID.Text = "" Or Trim(txtProjectID.Text) = Trim(sAffectedRecordIdentifier) Then
            ' Need to clear fields that could have been updated
            If Trim(txtProjectID.Text) = Trim(sAffectedRecordIdentifier) Then
                comboYearBuilt.Text = ""
                comboState.Text = ""
                txtCostMin.Text = ""
                txtCostMax.Text = ""
                txtAreaMin.Text = ""
                txtAreaMax.Text = ""
            Else
                txtProjectID.Text = Trim(sAffectedRecordIdentifier) 'AK- 6/7/2006
            End If
            varBookmark = TDBGridProject.Bookmark
            buSearch_Click
            TDBGridProject.Bookmark = varBookmark
            TDBGridProject.Refresh
            DoEvents
        End If
    End If
End Sub


Public Sub Set_Value(ID)
    If ID <> "" Then
        FormatTree.SetFocus
        FormatTree.FocusItem ID
        FormatTree_NodeSelected ID
    End If
End Sub

Private Sub buAnalysis_Click()
    Dim frm     As frmProjectAnalysis
    Set frm = New frmProjectAnalysis
    frm.Set_Value (m_ClassID)
    frm.Show
End Sub

Private Sub buDelete_Click()

    m_objGridMap.Delete
    Screen.MousePointer = vbNormal

End Sub

Private Sub buNew_Click()
    Dim frm As frmProject

    Set frm = New frmProject
    With frm
        .NewRow Val(m_ClassID)
        .Show
    End With
    
End Sub

Private Sub buParameters_Click()
    Dim frm     As frmProject
    Dim rec     As ADODB.RecordSet
    
    On Error Resume Next
    If IsNumeric(TDBGridProject.Bookmark) = False Then
        MsgBox "Please select a row.", vbCritical
    Else
        '
        '   Make copy of recordset, using the gridmap NOT 'm_rec.Clone'
        '   so that if they have changed values and not updated the recordset
        '   we pass to the form will contain the original values.
        '
        Set rec = m_objGridMap.CloneRowRecordset
        If Not rec.EOF Then
            Set frm = New frmProject
            '
            '   Pass the current record into the form,
            '   Navigating to single-record view.
            With frm
                .SetRow Trim(m_rec.Fields("proj_bldg_skey"))
                .Show
            End With
        End If
    End If
End Sub

Private Sub ckWrap_Click()
    m_objGridMap.RowWrap (ckWrap)
End Sub

Private Sub comboClassification_Change()
    Dim nRet As Long
    Dim nSelStart As Long

    If blnDelete Then Exit Sub
    nRet = SendMessage(comboClassification.hWnd, CB_FINDSTRING, 0, comboClassification.Text)
    nSelStart = comboClassification.SelStart
    If nRet >= 0 Then
        comboClassification.ListIndex = nRet
        comboClassification.SelStart = nSelStart
        comboClassification.SelLength = Len(comboClassification.Text)
    End If
    
End Sub

Private Sub comboClassification_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete, vbKeyBack
        blnDelete = True
    Case Else
        blnDelete = False
    End Select
End Sub

Private Sub comboClassification_LostFocus()
    Dim iIndex As Long
    
    If comboClassification.ListIndex = -1 And comboClassification.Text <> "" Then
        iIndex = SendMessage(comboClassification.hWnd, CB_FINDSTRING, 0, comboClassification.Text)
        If iIndex >= 0 Then
            comboClassification.ListIndex = iIndex
        End If
    End If

End Sub

Private Sub Form_Load()
       
    On Error Resume Next
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
End Sub

Private Sub Form_Initialize()

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status ("Loading Project Maintenance Grid ...")
    sEventSubscriberID = EventSubscriberAdd(Me)
    m_blnFirstSearch = True
    '
    '   Fill the MasterFormat tree.
    FormatTree.InitData g_cnShared, "PROJECT_LIST"
    '
    '   Initialize grid.
    Dim I As Integer
    I = 0
    If Not g_objDAL.GetRecordset(vbNullString, "SELECT distinct class_id, sort_order FROM CLASSIFICATION WHERE class_system_id = 'P1' and class_id not like 'T%' ORDER BY sort_order", m_rec) Then
        MsgBox "An error occurred while searching for classification code(s)."
    Else
        Do Until m_rec.EOF
            ReDim Preserve aryClassList(I)
            aryClassList(I) = m_rec.Fields("class_id")
            I = I + 1
            m_rec.MoveNext
        Loop
    End If
    m_rec.Close
    ' get exterior material codes
    Dim aryExteriorMaterial() As String
    I = 0
    If Not g_objDAL.GetRecordset(vbNullString, "SELECT distinct exterior_material_desc FROM EXTERIOR_MATERIAL ORDER BY exterior_material_desc", m_rec) Then
        MsgBox "An error occurred while searching for exterior material code(s)."
    Else
        Do Until m_rec.EOF
            ReDim Preserve aryExteriorMaterial(I)
            aryExteriorMaterial(I) = m_rec.Fields("exterior_material_desc")
            I = I + 1
            m_rec.MoveNext
        Loop
    End If
    m_rec.Close
    ' get state code
    Dim aryState() As String
    I = 0
    If Not g_objDAL.GetRecordset(vbNullString, "SELECT distinct state_code FROM state_country ORDER BY state_code", m_rec) Then
        MsgBox "An error occurred while searching for exterior material code(s)."
    Else
        Do Until m_rec.EOF
            ReDim Preserve aryState(I)
            aryState(I) = m_rec.Fields("state_code")
            I = I + 1
            m_rec.MoveNext
        Loop
    End If
    m_rec.Close

    With m_objGridMap
        .SetGrid TDBGridProject
        .InitGrid aryClassList, aryExteriorMaterial, aryState
    End With
    
    '   initialize the default look of the form
    m_blnSortByDesc = True  ' DEFAULT SORT = BY DESCRIPTION
    'optSort(0).Value = Not m_blnSortByDesc
    PopulateComboBox
    opGTYB.Value = True
    buUpdate.Enabled = False
    ckUse.Value = 1
    m_blnFirstSearch = False
    
    Status ("")
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Activate()
    Dim ctl As Control
    
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        If Len(m_strCurrentFormControl) > 0 Then
            For Each ctl In Me.Controls
                If ctl.Name = m_strCurrentFormControl Then
                    ctl.SetFocus
                    Exit For
                End If
            Next ctl
        End If
        OutputView True
        ShowGridSort
       ' m_objGridMap.SetMenuBar
    End If
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button  As String
    
    On Error Resume Next
    '
    '   Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Close Projects Grid Form")
        If Button = vbYes Then
            buUpdate_Click
            '
            '   If there were errors, cancel the close.
            If m_blnWereErrors Then
                Cancel = True
            End If
        ElseIf Button = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    '   Disables & hides the sort buttons on the main form.
    If m_rec.State <> adStateClosed Then m_rec.Close
    Set m_rec = Nothing
    HideGridSort
    EventSubscriberRemove sEventSubscriberID
End Sub

Private Sub Form_LostFocus()
    TDBGridProject.Update
    HideGridSort
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '
    '   Need to place in common routine for all forms.
    '   Possibly place all buttons in a frame like frame1 with
    '   common name and can just place it.
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 11250 Then
            TDBGridProject.Width = Me.Width - (TDBGridProject.Left * 3)
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11250
        End If
        
        If Me.Height >= 7260 Then
            TDBGridProject.Height = Me.Height - 4545
        '    fraGoTo.top = Me.Height - 1260
            buUpdate.Top = Me.Height - 1200
            buParameters.Top = Me.Height - 1200
            buAnalysis.Top = Me.Height - 1200
            buNew.Top = Me.Height - 1200
            buDelete.Top = Me.Height - 1200
        Else
            Me.Height = 7260
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub RefreshGrid(strSQL As String)
    Dim Button  As String
    Dim Refresh As Boolean
    On Error Resume Next
    '
    '   Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNo, "")
        If Button = vbYes Then
            buUpdate_Click
            Refresh = False
        Else
            Refresh = True
        End If
    Else
        Refresh = True
    End If
    
    If Refresh Then
        Dim StartTime
        StartTime = Now
        
        Screen.MousePointer = vbHourglass
        Status ("Populating Project Grid ...")
        buDelete.Enabled = False
        If m_rec.State <> adStateClosed Then m_rec.Close
        If Not g_objDAL.GetRecordset(vbNullString, strSQL, m_rec) Then
            MsgBox "An error occurred while searching for projects(s)."
            lbStatus.Caption = "no record found"
        Else
            If m_rec.RecordCount = 0 Then
                lbStatus.Caption = "no record found"
            Else
                lbStatus.Caption = str(m_rec.RecordCount) & " rows returned in " & _
                                            str(DateDiff("s", StartTime, Now)) + " seconds"
                buDelete.Enabled = True
            End If
            m_objGridMap.RecordSet = m_rec
            '
            '   Reset the grid contents
            With TDBGridProject
                '.Delete    '9/9/2005 RTD - THIS IS NOT NECESSARY WITH REBIND
                .Bookmark = Null
                .ReBind
                .ApproxCount = m_rec.RecordCount
            End With
            m_lastSQL = strSQL
        End If
        Screen.MousePointer = vbNormal
        Status ("")
    End If
End Sub

Private Sub FormatTree_NodeSelected(ByVal strID As String)
    Dim strSELECT   As String
    Dim rec         As New ADODB.RecordSet
    Dim bPass       As Boolean
    
    bPass = True
    strSELECT = "SELECT * FROM CLASSIFICATION WHERE class_id = '" & strID & "' AND class_system_id = 'F'"
    rec.Open strSELECT, g_cnShared
    If rec.EOF Then
        If IsNumeric(strID) Then
            strSELECT = "EXEC sp_select_project_list " & _
                              " @projkey = " & strID & "," & _
                              " @facility_class_id = Null," & _
                              " @bid_year = Null," & _
                              " @datecompare = Null," & _
                              " @state_code = Null," & _
                              " @cost_min = Null," & _
                              " @cost_max = Null," & _
                              " @area_min = Null," & _
                              " @area_max = Null," & _
                              " @use = Null"
            ' 9/9/2005 RTD - PROJECT NODE SELECTED
            '                POPULATE PROJECT ID
            txtProjectID.Text = strID
        Else
            bPass = False
        End If
    Else
        strSELECT = "EXEC sp_select_project_list " & _
                          " @projkey = Null," & _
                          " @facility_class_id = '" & strID & "'," & _
                          " @bid_year = Null," & _
                          " @datecompare = Null," & _
                          " @state_code = Null," & _
                          " @cost_min = Null," & _
                          " @cost_max = Null," & _
                          " @area_min = Null," & _
                          " @area_max = Null," & _
                          " @use = Null"
        m_ClassID = strID
        txtProjectID.Text = ""
    End If
    rec.Close
    Set rec = Nothing
    'MODIFIED 9/8/2005 RTD
    'SELECT THE CLASSIFICATION COMBOBOX ITEM THAT MATCHES THE SELECTED NODE
    If IsNumeric(m_ClassID) Then
        comboClassification.ListIndex = FindComboItemDataIndex(comboClassification, m_ClassID)
    Else
        comboClassification.ListIndex = -1
    End If
    If bPass Then RefreshGrid (strSELECT)
    
End Sub

Private Sub buSearch_Click()
    Dim rec As New ADODB.RecordSet
    Dim bCheck As Boolean
    Dim strSELECT As String, strMsg As String, strYB As String, strSQL As String, strTemp As String
    Dim iIndex As Long

    bCheck = False
    If (txtCostMin.Text = "" Or IsNumeric(txtCostMin.Text)) And (txtCostMax.Text = "" Or IsNumeric(txtCostMax.Text)) And (txtAreaMin.Text = "" Or IsNumeric(txtAreaMin.Text)) And (txtAreaMax.Text = "" Or IsNumeric(txtAreaMax.Text)) Then
        m_counter = 0
        strSELECT = "EXEC sp_select_project_list "
        If txtProjectID.Text = "" Then
            strSELECT = strSELECT & "@projkey = NULL,"
        Else
            strSELECT = strSELECT & "@projkey = '" & txtProjectID.Text & "',"
            bCheck = True
        End If
        'IF USER AUTOCOMPLETED THE COMBO, BUT WE ARRIVED HERE VIA 'DEFAULT' BUTTON BEHAVIOR,
        'LISTINDEX IS NOT SET, SO CHECK IF THE TEXT IS A VALID ITEM
        If comboClassification.ListIndex = -1 And comboClassification.Text <> "" Then
            iIndex = SendMessage(comboClassification.hWnd, CB_FINDSTRING, 0, comboClassification.Text)
            If iIndex >= 0 Then
                comboClassification.ListIndex = iIndex
            End If
        End If
        If comboClassification.ListIndex > 0 Then
            'm_ClassID = Mid(comboClassification.Text, 2, InStr(comboClassification.Text, ")") - 2)
            m_ClassID = comboClassification.ItemData(comboClassification.ListIndex)
            'strSELECT = strSELECT & " @facility_class_id = '" & Mid(comboClassification.Text, 2, InStr(comboClassification.Text, ")") - 2) & "', "
            strSELECT = strSELECT & " @facility_class_id = '" & m_ClassID & "', "
            bCheck = True
            FormatTree.FocusItem m_ClassID
        '9/9/2005 RTD - FIXED CASE ON 'ALL' - COMBOBOX ITEM IS "ALL", NOT 'All"
        ElseIf comboClassification.Text <> "" And comboClassification.Text <> "ALL" Then
            strSQL = "SELECT class_id FROM CLASSIFICATION WHERE class_desc like '" & Replace(comboClassification.Text, "*", "%") & "'"
            If Not g_objDAL.GetRecordset(vbNullString, strSQL, rec) Then
                Screen.MousePointer = vbNormal
                MsgBox "An error occurred while searching for available classification categories."
                Exit Sub
            Else
                If Not rec.EOF Then
                    FormatTree.FocusItem (rec("class_id"))
                    m_ClassID = rec("class_id")
                    While Not rec.EOF
                        strTemp = strTemp & rec("class_id") & ","
                        rec.MoveNext
                    Wend
                    strTemp = Left(strTemp, Len(strTemp) - 1)
                    strSELECT = strSELECT & " @facility_class_id = '" & strTemp & "', "
                    bCheck = True
                Else
                    Screen.MousePointer = vbNormal
                    MsgBox "No classification code found"
                    Exit Sub
                End If
            End If
        Else
            strSELECT = strSELECT & " @facility_class_id = Null, "
            m_ClassID = ""
        End If

        If comboYearBuilt.ListIndex > 0 Then
            strSELECT = strSELECT & " @bid_year = " & comboYearBuilt.Text & ", "
            bCheck = True
        Else
            strSELECT = strSELECT & " @bid_year = Null, "
        End If
        If opLTYB.Value = True Then strYB = "<="
        If opEQYB.Value = True Then strYB = "="
        If opGTYB.Value = True Then strYB = ">="
        If strYB = "" Then
            opEQYB.Value = True
            strYB = "="
        End If
        strSELECT = strSELECT & " @datecompare = '" & strYB & "', "
        
        If comboState.ListIndex > 0 Then
            strSELECT = strSELECT & " @state_code = '" & comboState.Text & "', "
            bCheck = True
        Else
            strSELECT = strSELECT & " @state_code = Null, "
        End If

        If txtCostMin.Text <> "" And IsNumeric(txtCostMin.Text) Then
            strSELECT = strSELECT & " @cost_min = " & txtCostMin.Text & ", "
            bCheck = True
        Else
            strSELECT = strSELECT & " @cost_min = Null, "
        End If
        If txtCostMax.Text <> "" And IsNumeric(txtCostMax.Text) Then
            strSELECT = strSELECT & " @cost_max = " & txtCostMax.Text & ", "
            bCheck = True
        Else
            strSELECT = strSELECT & " @cost_max = Null, "
        End If
        If txtAreaMin.Text <> "" And IsNumeric(txtAreaMin.Text) Then
            strSELECT = strSELECT & " @area_min = " & txtAreaMin.Text & ", "
            bCheck = True
        Else
            strSELECT = strSELECT & " @area_min = Null, "
        End If
        If txtAreaMax.Text <> "" And IsNumeric(txtAreaMax.Text) Then
            strSELECT = strSELECT & " @area_max = " & txtAreaMax.Text & ", "
            bCheck = True
        Else
            strSELECT = strSELECT & " @area_max = Null, "
        End If
        If ckUse = 1 Then
            strSELECT = strSELECT & " @use = 1 "
        Else
            strSELECT = strSELECT & " @use = Null "
        End If
        
        If bCheck Then
            RefreshGrid (strSELECT)
        Else
            MsgBox "Please enter some search criteria.", vbExclamation
        End If
    Else
        strMsg = "Please enter a valid number in "
        If txtCostMin.Text <> "" And Not IsNumeric(txtCostMin.Text) Then strMsg = strMsg & "Cost Min,"
        If txtCostMax.Text <> "" And Not IsNumeric(txtCostMax.Text) Then strMsg = strMsg & "Cost Max,"
        If txtAreaMin.Text <> "" And Not IsNumeric(txtAreaMin.Text) Then strMsg = strMsg & "Area Min,"
        If txtAreaMax.Text <> "" And Not IsNumeric(txtAreaMax.Text) Then strMsg = strMsg & "Area Max,"
        MsgBox Left(strMsg, Len(strMsg) - 1), vbExclamation
    End If
End Sub

Private Sub buUpdate_Click()
    Dim varBookmark As Variant
    Dim bUpdate As Boolean
    Dim strMsg As String
    On Error Resume Next
    bUpdate = True
    Dim Button
    If TDBGridProject.Columns("gross_floor_area") < 1000 Then
        Button = MsgBox("The value you enter for area is less than 1000, do you still want to submit this value?", vbYesNo, "Invalid number for area")
        If Button = vbYes Then
            bUpdate = True
        ElseIf Button = vbNo Then
            bUpdate = False
        End If
    End If
    If TDBGridProject.Columns("proj_bldg_project_tot_cost") < 100000 Then
        Button = MsgBox("The value you enter for total cost is less than 100,000, do you still want to submit this value?", vbYesNo, "Invalid number for area")
        If Button = vbYes Then
            bUpdate = True
        ElseIf Button = vbNo Then
            bUpdate = False
        End If
    End If
    If TDBGridProject.Columns("facility1_class_id") = "" Then
        strMsg = "Please enter a value for class id" & vbCrLf
    End If
    If TDBGridProject.Columns("state_code") = "" Then
        strMsg = strMsg & "Please enter a state for location"
    End If
    If strMsg <> "" Then
        MsgBox strMsg
        bUpdate = False
    End If
    If bUpdate Then
        Screen.MousePointer = vbHourglass
        Status ("Updating Project Details ...")
        m_blnWereErrors = False
        varBookmark = TDBGridProject.Bookmark
        TDBGridProject.Update
        
        With m_objGridMap
            .Update aryClassList, m_ClassID
            Screen.MousePointer = vbNormal
            If .UpdateErrors = 0 Then
                Status ("Project Details Updated Successfully ...")
                MsgBox .SuccessfulUpdates & " rows were updated successfully."
            Else
                Status ("")
                MsgBox .SuccessfulUpdates & " rows were updated successfully." _
                        & vbCrLf & .UpdateErrors & " errors were received."
            End If
        End With
        TDBGridProject.Bookmark = varBookmark
        RefreshGrid (m_lastSQL)
        Status ("Building Tree ...")
        Screen.MousePointer = vbHourglass
        FormatTree.ClearTree
        FormatTree.InitData g_cnShared, "PROJECT_LIST"
        FormatTree.FocusItem m_ClassID
        Screen.MousePointer = vbNormal
        buUpdate.Enabled = False
        Status ("")
    End If
End Sub

Private Sub PopulateClassificationComboBox(Optional bSortByDescription As Boolean = False)
    Dim rec         As New ADODB.RecordSet
    Dim strSELECT   As String
    Dim strText     As String
    Dim I           As Integer
    
    Screen.MousePointer = vbHourglass
    '
    '   Fill the classification combo box
    comboClassification.Clear
    strSELECT = "SELECT DISTINCT facility1_class_id, class_desc " & _
                " FROM vw_project_list "
    If bSortByDescription Then
        strSELECT = strSELECT & " ORDER BY class_desc"
    Else
        strSELECT = strSELECT & " ORDER BY facility1_class_id"
    End If
    If Not g_objDAL.GetRecordset(vbNullString, strSELECT, rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available classification categories."
    Else
        With rec
            If .RecordCount = 0 Then
                comboClassification.AddItem "(unknown)"
            Else
                comboClassification.AddItem "ALL"
                While Not .EOF
                    If bSortByDescription Then
                        strText = Trim(.Fields("class_desc").Value)
                        strText = strText & "  (" & Trim(.Fields("facility1_class_id")) & ")"
                    Else
                        strText = "(" & Trim(.Fields("facility1_class_id")) & ") "
                        For I = Len(strText) To 6
                            strText = strText & " "
                        Next
                        strText = strText & Trim(.Fields("class_desc").Value)
                    End If
                    comboClassification.AddItem strText
                    'USE ITEMDATA PROPERTY TO STORE CLASS ID FOR LATER SEARCHING
                    '6/28/2005 RTD
                    comboClassification.ItemData(comboClassification.NewIndex) = Trim(.Fields("facility1_class_id"))
                    .MoveNext
                Wend
            End If
            .Close
        End With
    End If
    Screen.MousePointer = vbDefault
    Set rec = Nothing
    
End Sub

Private Sub PopulateComboBox()
    Dim rec         As New ADODB.RecordSet
    Dim strSELECT   As String
    Dim strText     As String
    Dim I           As Integer
    
    Screen.MousePointer = vbHourglass
    '
    '   Fill the classification combo box
    optSort(0).Value = Not m_blnSortByDesc
    PopulateClassificationComboBox m_blnSortByDesc
    '
    '   Fill the Year Build combo box
    comboYearBuilt.Clear
    strSELECT = "SELECT DISTINCT year(bid_date) as year FROM vw_project_list WHERE bid_date is not null ORDER BY year(bid_date) desc"
    If Not g_objDAL.GetRecordset(vbNullString, strSELECT, rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available project year."
    Else
        With rec
            If .RecordCount = 0 Then
                comboYearBuilt.AddItem "(unknown)"
            Else
                comboYearBuilt.AddItem "ALL"
                While Not .EOF
                    comboYearBuilt.AddItem Trim(.Fields("year").Value)
                    .MoveNext
                Wend
            End If
            .Close
        End With
    End If
    '
    '   Fill the State combo box
    comboState.Clear
        
    strSELECT = "SELECT DISTINCT state_code FROM vw_project_list ORDER BY state_code"
    If Not g_objDAL.GetRecordset(vbNullString, strSELECT, rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available state code."
    Else
        With rec
            If .RecordCount = 0 Then
                comboState.AddItem "(unknown)"
            Else
                comboState.AddItem "ALL"
                While Not .EOF
                    comboState.AddItem Trim(.Fields("state_code").Value)
                    .MoveNext
                Wend
            End If
            .Close
        End With
    End If
    
    Set rec = Nothing
    Screen.MousePointer = vbNormal
End Sub

Private Sub optSort_Click(Index As Integer)
    m_blnSortByDesc = (optSort(1).Value)
    PopulateClassificationComboBox m_blnSortByDesc
    comboClassification.SetFocus
End Sub

Private Sub TDBGridProject_AfterColUpdate(ByVal ColIndex As Integer)
' MODIFIED 9/8/2005 RTD
' RECALCULATE 'S' COLUMN IF 'Q' OR 'R' COLUMN CHANGED
' COLUMN 'S' IS ALWAYS THE SUM OF 'Q' + 'R' (CR#1512)
    Dim Q As Long
    Dim r As Long
    Dim S As Long
    
    If ColIndex = TDBGridProject.Columns("Q").ColIndex Or ColIndex = TDBGridProject.Columns("R").ColIndex Then
        If TDBGridProject.Columns("Q").Value <> "" Then
            Q = TDBGridProject.Columns("Q").Value
        End If
        If TDBGridProject.Columns("R").Value <> "" Then
            r = TDBGridProject.Columns("R").Value
        End If
        S = Q + r
        TDBGridProject.Columns("S").Value = S
        TDBGridProject.Columns("S").RefetchCell
    End If
    buUpdate.Enabled = True
    
End Sub

Private Sub TDBGridProject_DblClick()
    Dim frm     As frmProject
    Dim rec     As ADODB.RecordSet
    
    On Error Resume Next
    If TDBGridProject.Columns("proj_bldg_id") <> "" Then
        If IsNumeric(TDBGridProject.Bookmark) = False Then
            MsgBox "Please select a row.", vbCritical
        Else
            '
            '   Make copy of recordset, using the gridmap NOT 'm_rec.Clone'
            '   so that if they have changed values and not updated the recordset
            '   we pass to the form will contain the original values.
            '
            Set rec = m_objGridMap.CloneRowRecordset
            If Not rec.EOF Then
                Set frm = New frmProject
                '
                '   Pass the current record into the form,
                '   Navigating to single-record view.
                With frm
                    .SetRow Trim(m_rec.Fields("proj_bldg_skey"))
                    .Show
                End With
            End If
        End If
    End If
End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub

Private Sub TDBGridProject_OnAddNew()
    If m_ClassID <> "" Then
        TDBGridProject.Columns("facility1_class_id") = m_ClassID
        TDBGridProject.Columns("use_ind") = 1
    End If
End Sub

Private Sub TDBGridProject_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If TDBGridProject.Columns("proj_bldg_id") = "" Then
        buParameters.Enabled = False
    Else
        buParameters.Enabled = True
    End If
End Sub

Private Sub TDBGridProject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Dim strErrorMsg As String
        strErrorMsg = m_objGridMap.GetError(TDBGridProject.Bookmark)
        If Len(strErrorMsg) > 0 Then
            MsgBox strErrorMsg
        End If
    End If
End Sub


