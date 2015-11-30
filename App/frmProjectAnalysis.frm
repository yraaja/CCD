VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "Meansctrl.ocx"
Begin VB.Form frmProjectAnalysis 
   Caption         =   "Project Analysis"
   ClientHeight    =   6900
   ClientLeft      =   4365
   ClientTop       =   3915
   ClientWidth     =   11130
   Icon            =   "frmProjectAnalysis.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   11130
   Begin VB.CommandButton buUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   10080
      TabIndex        =   26
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton buClone 
      Caption         =   "Annual &Clone"
      Height          =   495
      Left            =   5400
      TabIndex        =   25
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton buReport 
      Caption         =   "&Reports"
      Height          =   495
      Left            =   2280
      TabIndex        =   24
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton buOutput 
      Caption         =   "&Output"
      Height          =   495
      Left            =   1200
      TabIndex        =   23
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton buCostMaint 
      Caption         =   "Cost &Maint"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   6360
      Width           =   975
   End
   Begin VB.CheckBox ckWrap 
      Caption         =   "Row Wrap"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame frmMeasure 
      Caption         =   "Imperial/Metric"
      Height          =   615
      Left            =   2400
      TabIndex        =   20
      Top             =   2760
      Width           =   2295
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1950
         TabIndex        =   32
         Top             =   240
         Width           =   1955
         Begin VB.OptionButton opImperial 
            Caption         =   "Imperial"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   255
         End
         Begin VB.OptionButton opMetric 
            Caption         =   "Metric"
            Height          =   255
            Left            =   1080
            TabIndex        =   33
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbImperial 
            Caption         =   "Imperial"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lbMetric 
            Caption         =   "Metric"
            Height          =   255
            Left            =   1320
            TabIndex        =   35
            Top             =   0
            Width           =   615
         End
      End
   End
   Begin VB.Frame frmUnit 
      Caption         =   "Cost/Percent"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   2055
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1710
         TabIndex        =   27
         Top             =   240
         Width           =   1715
         Begin VB.OptionButton opDollar 
            Caption         =   "$"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   255
         End
         Begin VB.OptionButton opPercent 
            Caption         =   "%"
            Height          =   255
            Left            =   840
            TabIndex        =   28
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbDollar 
            Caption         =   "$"
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lbPercent 
            Caption         =   "%"
            Height          =   255
            Left            =   1200
            TabIndex        =   30
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Frame fmPublication 
      Caption         =   "For Publication"
      Height          =   1095
      Left            =   8400
      TabIndex        =   6
      Top             =   1440
      Width           =   2655
      Begin VB.TextBox txtPub34 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPub12 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtPub14 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbPub34 
         Alignment       =   2  'Center
         Caption         =   "3/4"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbPub12 
         Alignment       =   2  'Center
         Caption         =   "1/2"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbPub14 
         Alignment       =   2  'Center
         Caption         =   "1/4"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fmSystem 
      Caption         =   "System Generated"
      Height          =   1095
      Left            =   5520
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
      Begin VB.TextBox txtSys34 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtSys12 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtSys14 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbSys34 
         Alignment       =   2  'Center
         Caption         =   "3/4"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbSys12 
         Alignment       =   2  'Center
         Caption         =   "1/2"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbSys14 
         Alignment       =   2  'Center
         Caption         =   "1/4"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton buSearch 
      Caption         =   "&Search"
      Height          =   495
      Left            =   9720
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox comboClassification 
      Height          =   315
      Left            =   6615
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   240
      Width           =   4455
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4895
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridAnalysis 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4895
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
   Begin VB.Label lbClassification 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Classification"
      Height          =   195
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   915
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11100
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "frmProjectAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'   Class to handle grid
Dim m_objGridMap As New CAnalysisMap
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
Dim sEventSubscriberID              As String
'
'   Keep track of last completed select statement
Dim m_lastSQL As String
Dim m_ClassID As String
'
'   List of Project components
Dim m_MaxCol As Integer
Dim m_opDollar As Boolean
Dim m_opPercent As Boolean
Dim m_opImperial As Boolean
Dim m_opMetric As Boolean
Dim m_RowNumber As Integer

Dim m_last_v14 As Integer
Dim m_last_v12 As Integer
Dim m_last_v34 As Integer
'*** APEX Migration Utility Code Change ***
'Dim m_original_style    As TrueOleDBGrid70.Style
Dim m_original_style    As TrueOleDBGrid80.Style
Dim m_last_bookmark As Integer


Dim m_counter As Integer
' declare variables used to track changes in publish values
Dim m_aClassID() As String
Dim m_aQtr1() As String
Dim m_aMed() As String
Dim m_aQtr3() As String
Dim m_oQtr1 As String
Dim m_oMed As String
Dim m_oQtr3 As String
Dim m_ColClassID As String
Dim m_PubBox As TextBox

Const USEBOOKMARK = 1
Const USECOORD = 0

Private Sub buCostMaint_Click()
    Dim frm     As frmProjectGrid
    Set frm = New frmProjectGrid
    frm.Set_Value m_ClassID
    frm.Show
End Sub

Private Sub buOutput_Click()
    Dim varBookmark As Variant
    '
    '   Use current row
    If Not IsNull(TDBGridAnalysis.Bookmark) Then
        'm_rec.Bookmark = TDBGridAnalysis.Bookmark
        
        For Each varBookmark In TDBGridAnalysis.SelBookmarks
            m_rec.Bookmark = varBookmark
            
            If IsNull(m_rec.Fields("bk_skey")) Then
                MsgBox "Please add this row to Book Output by adding a Book ID before using Output Management.", vbCritical
                Exit Sub
            End If
        Next varBookmark

        DoOutput
    End If
End Sub

Public Sub DoOutput()
    Dim frm         As Form
    Dim blnVisible  As Boolean
    Dim strUpdate   As String
    Dim strError    As String
    Dim blnReturn   As Boolean
    Dim varBookmark As Variant

    On Error Resume Next
    If FormOpen("dlgOutput", frm, blnVisible) = False Then
        If Not IsNull(TDBGridAnalysis.Bookmark) Then
            Set frm = New dlgOutput
        Else
            Exit Sub
        End If
    End If
    frm.Visible = True
    
    If Not (TDBGridAnalysis.BOF = True Or TDBGridAnalysis.EOF = True) Then
        strUpdate = "exec sp_temp_output_init"
        '
        '   m_objOutput is a global object for CCDal.CRSMDataAccess
        blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
        '
        '   No rows selected.
        If TDBGridAnalysis.SelBookmarks.Count = 0 Then
            '
            '   Use current row
            If Not IsNull(TDBGridAnalysis.Bookmark) Then
                m_rec.Bookmark = TDBGridAnalysis.Bookmark
                '
                '   Valid values are A = assembly, SF = SquareFoot, E = equipment, U = unit
                strUpdate = "exec sp_temp_add_output_keys @skey_type = 'P', @skey = " _
                            & CStr(m_rec.Fields("bk_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            End If
        Else
            For Each varBookmark In TDBGridAnalysis.SelBookmarks
                m_rec.Bookmark = varBookmark
                '
                '   Valid values are A = assembly, SF = SquareFoot, E = equipment, U = unit
                strUpdate = "exec sp_temp_add_output_keys @skey_type = 'P', @skey = " _
                            & CStr(m_rec.Fields("bk_skey"))
                blnReturn = frm.m_objOutput.ExecQuery(vbNullString, strUpdate, strError)
            Next varBookmark
        End If

        With frm
            .bShowAllFields = True
            '.SetKeys CStr(m_rec.Fields("bk_skey")), "P"
            .FillData
            .Show vbModeless, fMainForm
            .Caption = "Output Usage"
        End With
    End If
End Sub

Private Sub position_output(Optional Y As Single = 0)
    Dim frm             As Form
    Dim blnVisible      As Boolean
    Dim lngCurrentRow   As Long
    '
    '   Only send data to the Output dialog if it is open.
    Screen.MousePointer = vbHourglass
    With TDBGridAnalysis
        If FormOpen("dlgOutput", frm, blnVisible) = True Then
            If Y <> 0 Then
                lngCurrentRow = .RowContaining(Y)
                If lngCurrentRow <> -1 Then
                    .Row = lngCurrentRow
                    m_rec.Bookmark = .Bookmark
                End If
            End If
            DoOutput
            Me.SetFocus
        End If
    End With
    Screen.MousePointer = vbNormal
End Sub

Private Sub buReport_Click()
    Dim dialog As New frmDialog
    Dim frm As New frmProjectRpt
    Dim recProjects As New ADODB.RecordSet
    Dim strSELECT As String
    Dim classid As String
    Dim classdesc As String
    Dim pos As Integer
    Dim sReportType As String
    Dim comparisiondate As String
    Dim bImperial       As Boolean
    Dim variance        As String
               
    With dialog
        .Show vbModal
        
        If .Cancelled = False Then
            sReportType = Trim(.cboReports.Text)
            If sReportType = "PCIS Variance" Then
                comparisiondate = .txtComparisionDate.Text
                variance = .txtVariance.Text
            End If
            pos = InStr(1, .comboClassification.Text, ")")
            classid = Left$(.comboClassification.Text, pos)
            classid = Replace(classid, ")", "")
            classid = Replace(classid, "(", "")
        
            classdesc = Trim(Right$(.comboClassification.Text, Len(.comboClassification.Text) - pos))
            
            bImperial = .opImperial.Value
            
            Unload dialog
        Else
            Unload dialog
            Exit Sub
        End If
    End With
    
    Screen.MousePointer = vbHourglass
    If sReportType = "PCIS Variance" Then
        strSELECT = "exec sp_rpt_pcis_variance "
    Else
        strSELECT = "exec sp_rpt_project_types_components "
    End If
    
    strSELECT = strSELECT & "'" & classdesc & "', '" & classid & "', "

    If bImperial = True Then
        strSELECT = strSELECT & "'imperial'"
    Else
        strSELECT = strSELECT & "'metric'"
    End If
    
    If sReportType = "PCIS Variance" Then
        strSELECT = strSELECT & ", '" & comparisiondate & "'"
        strSELECT = strSELECT & ", '" & variance & "'"
    End If
    '
    '   Use DAL to perform select.
    If Not g_objDAL.GetRecordset(vbNullString, strSELECT, recProjects) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while gathering report data."
    Else
        With frm
            .LoadReport recProjects, sReportType, bImperial
            .RenderReport
            .Show
        End With
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub ckWrap_Click()
    m_objGridMap.RowWrap (ckWrap)
End Sub

Private Sub buProject_Click()
    Dim frm     As frmProjectGrid
    Set frm = New frmProjectGrid
    frm.Show
    frm.Set_Value (m_ClassID)
End Sub

Public Sub Set_Value(ID)
    If ID <> "" Then
        FormatTree.FocusItem ID
        FormatTree.SetFocus
        RefreshGrid ID
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    ColorLockedFields Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button  As String
    
    On Error Resume Next
    '
    '   Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Or UBound(m_aClassID) > 0 Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Close Building Grid Form")
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

Private Sub Form_Resize()
    On Error Resume Next
    '
    '   Need to place in common routine for all forms.
    '   Possibly place all buttons in a frame like frame1 with
    '   common name and can just place it.
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 11250 Then
            TDBGridAnalysis.Width = Me.Width - 255
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11250
        End If
        
        If Me.Height >= 7260 Then
            TDBGridAnalysis.Height = Me.Height - 4545 - 240
            buUpdate.Top = Me.Height - 1050
            buReport.Top = Me.Height - 1050
            buClone.Top = Me.Height - 1050
            buOutput.Top = Me.Height - 1050
            buCostMaint.Top = Me.Height - 1050
        Else
            Me.Height = 7260
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub Form_Initialize()

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status ("Loading Project Analysis Grid ...")
    sEventSubscriberID = EventSubscriberAdd(Me)
    '
    '   Fill the Project tree.
    FormatTree.InitData g_cnShared, "PROJECT_ANALYSIS"
    '
    '   Set up the Grids
    Dim strSQL As String
    Dim objRS As New ADODB.RecordSet
    
    strSQL = "SELECT TOP 1 facility1_class_id, COUNT(proj_bldg_skey) AS proj_count FROM PROJECT_BUILDING_DETAIL P WHERE P.use_ind = 1 AND P.gross_floor_area > 0 GROUP BY facility1_class_id ORDER BY COUNT(proj_bldg_skey) DESC"
    objRS.Open strSQL, g_cnShared
    m_MaxCol = objRS("proj_count")
    objRS.Close
    With m_objGridMap
        .SetGrid TDBGridAnalysis
        .InitGrid m_MaxCol
    End With

    '
    '   Fill the classification combo box
    Dim strText As String
    Dim I As Integer
    comboClassification.Clear
        
    strSQL = "SELECT DISTINCT facility1_class_id as class_id, class_desc FROM vw_project_list WHERE use_ind = 1 ORDER BY facility1_class_id"
    If Not g_objDAL.GetRecordset(vbNullString, strSQL, objRS) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available classification categories."
    Else
        With objRS
            If .RecordCount = 0 Then
                comboClassification.AddItem "(unknown)"
            Else
                While Not .EOF
                    strText = "(" & Trim(.Fields("class_id")) & ") "
                    For I = Len(strText) To 6
                        strText = strText & " "
                    Next
                    strText = strText & Trim(.Fields("class_desc").Value)
                    comboClassification.AddItem strText
                    .MoveNext
                Wend
            End If
            .Close
        End With
    End If
    
    '
    '   Initialize the values on the form
    opDollar.Value = True
    opImperial.Value = True
    buUpdate.Enabled = False
    m_opDollar = opDollar.Value
    m_opPercent = opPercent.Value
    m_opImperial = opImperial.Value
    m_opMetric = opMetric.Value
    Set m_PubBox = txtPub14
    strSQL = "SELECT COUNT(*) AS counter From PUBLISHED_PROJECT_COST WHERE " & _
             "(term_date = '2038-12-31 23:59:59.990') AND  " & _
             "(start_date = LTRIM(STR(YEAR(GETDATE()))) + '-' + LTRIM(STR(MONTH(GETDATE()))) + '-' + LTRIM(STR(DAY(GETDATE()))) + ' 00:00:00.000')"
    If Not g_objDAL.GetRecordset(vbNullString, strSQL, objRS) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while get published project cost."
    Else
        If objRS("counter") > 0 Then
            buClone.Enabled = False
        End If
    End If
    Set objRS = Nothing
    ' save original column style
    Set m_original_style = TDBGridAnalysis.Columns(0).Style
    ReDim m_aClassID(0)
    ReDim m_aQtr1(0)
    ReDim m_aMed(0)
    ReDim m_aQtr3(0)
    
    Status ("")
    Screen.MousePointer = vbNormal
End Sub

Private Sub RefreshGrid(ByVal strID As String)
    Dim Button  As String
    Dim bRefresh As Boolean
    bRefresh = True
    On Error Resume Next
    '
    '   Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Or UBound(m_aClassID) > 0 Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNo, "")
        If Button = vbYes Then
            buUpdate_Click
            '
            '   If there were errors, cancel the close.
            If m_blnWereErrors Then
                MsgBox "An error occur while updating the values"
                bRefresh = False
            End If
        End If
    End If
    If bRefresh = True Then
        Dim objRS As New ADODB.RecordSet
        Dim strSQL As String
        Dim strType As String
        Dim strSystem As String
        ' clear the publication values and arrays
        ReDim m_aClassID(0)
        ReDim m_aQtr1(0)
        ReDim m_aMed(0)
        ReDim m_aQtr3(0)
    
        ' initialized data in radio buttons
        If opDollar.Value = True Then
            strType = "dollar"
        ElseIf opPercent.Value = True Then
            strType = "percent"
        End If
        If opImperial.Value = True Then
            strSystem = "imperial"
        ElseIf opMetric.Value = True Then
            strSystem = "metric"
        End If
        
        Screen.MousePointer = vbHourglass
        Status ("Populating Analysis Grid ...")
        
        strSQL = "SELECT count(*) FROM CLASSIFICATION WHERE class_id = '" & strID & "' AND class_system_id = 'F'"
        objRS.Open strSQL, g_cnShared
        If objRS(0) = 1 Then
            strSQL = "EXEC sp_select_project_analysis @id = '" & strID & "', @type = '" & strType & "', @system = '" & strSystem & "', @class = '0'"
            m_ClassID = strID
        Else
            strSQL = "EXEC sp_select_project_analysis @id = '" & m_ClassID & "', @type = '" & strType & "', @system = '" & strSystem & "', @class = '" & Right(strID, Len(strID) - InStr(strID, "-K") - 1) & "'"
        End If
        Set objRS = Nothing
        If m_rec.State <> adStateClosed Then m_rec.Close
        If Not g_objDAL.GetRecordset(vbNullString, strSQL, m_rec) Then
            MsgBox "An error occurred while searching for projects(s)."
        Else
            m_objGridMap.RecordSet = m_rec
            '
            '   Reset the grid contents
            With TDBGridAnalysis
            '    .ClearFields
                .Bookmark = Null
                .ReBind
                .ApproxCount = m_rec.RecordCount
            End With
            m_lastSQL = strSQL
        End If
        
        TDBGridAnalysis_RowColChange -1, 0
        Screen.MousePointer = vbNormal
        Status ("")
    End If
    

End Sub

Private Sub FormatTree_NodeSelected(ByVal strID As String)
    RefreshGrid strID
End Sub

Private Sub buSearch_Click()
    If comboClassification.Text = "" Then
        MsgBox "Please select a value for classification"
    Else
        RefreshGrid Mid(comboClassification.Text, 2, InStr(comboClassification.Text, ")") - 2)
        FormatTree.FocusItem (Mid(comboClassification.Text, 2, InStr(comboClassification.Text, ")") - 2))
    End If
End Sub

Private Sub opDollar_Click()
    If opDollar.Value <> m_opDollar Then
        If m_ClassID <> "" Then RefreshGrid m_ClassID
        m_opDollar = opDollar.Value
        m_opPercent = opPercent.Value
    End If
End Sub

Private Sub opPercent_Click()
    If opPercent.Value <> m_opPercent Then
        If m_ClassID <> "" Then RefreshGrid m_ClassID
        m_opPercent = opPercent.Value
        m_opDollar = opDollar.Value
    End If
End Sub

Private Sub opImperial_Click()
    If opImperial.Value <> m_opImperial Then
        If m_ClassID <> "" Then RefreshGrid m_ClassID
        m_opImperial = opImperial.Value
        m_opMetric = opMetric.Value
    End If
End Sub

Private Sub opMetric_Click()
    If opMetric.Value <> m_opMetric Then
        If m_ClassID <> "" Then RefreshGrid m_ClassID
        m_opMetric = opMetric.Value
        m_opImperial = opImperial.Value
    End If
End Sub

Private Sub TDBGridAnalysis_DblClick()
    If TDBGridAnalysis.Col >= 7 Then
        m_PubBox.Text = TDBGridAnalysis.Columns(TDBGridAnalysis.Col)
        savePublishValues
    End If
End Sub

Private Sub TDBGridAnalysis_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strSQL As String
    Dim objRS As New ADODB.RecordSet
    Dim v14 As Integer
    Dim v12 As Integer
    Dim v34 As Integer
    
    If TDBGridAnalysis.Bookmark <> m_last_bookmark Or LastRow = -1 Then
        If TDBGridAnalysis.Columns("col_1") = "" Then Exit Sub
        Screen.MousePointer = vbHourglass
        Status ("Calculating System Generated Values ...")
    
        If (TDBGridAnalysis.Columns("ID") = "TA" Or TDBGridAnalysis.Columns("ID") = "TU" Or TDBGridAnalysis.Columns("ID") = "TV" Or TDBGridAnalysis.Columns("ID") = "TM") And opPercent.Value = True Then
            strSQL = "SELECT 0 as proj_num "
        ElseIf TDBGridAnalysis.Columns("ID") = "TA" Then
            strSQL = "SELECT count(D.proj_bldg_skey) as proj_num" & _
                     "  FROM PROJECT_BUILDING_DETAIL D " & _
                     " WHERE D.use_ind = 1 AND D.facility1_class_id = '" & m_ClassID & "' AND D.gross_floor_area > 0"
        ElseIf TDBGridAnalysis.Columns("ID") = "TV" Then
            strSQL = "SELECT count(D.proj_bldg_skey) as proj_num" & _
                     "  FROM PROJECT_BUILDING_DETAIL D " & _
                     " WHERE D.use_ind = 1 AND D.facility1_class_id = '" & m_ClassID & "' AND D.volume > 0"
        ElseIf TDBGridAnalysis.Columns("ID") = "TU" Then
            strSQL = "SELECT count(D.proj_bldg_skey) as proj_num" & _
                     "  FROM PROJECT_BUILDING_DETAIL D " & _
                     " WHERE D.use_ind = 1 AND D.facility1_class_id = '" & m_ClassID & "' AND D.proj_bldg_functional_uom_qty > 0"
        ElseIf TDBGridAnalysis.Columns("ID") = "TM" Then
            strSQL = "SELECT count(D.proj_bldg_skey) as proj_num" & _
                     "  FROM PROJECT_BUILDING_DETAIL D INNER JOIN PROJ_BLDG_COMPONENT_COST C ON D.proj_bldg_skey = C.proj_bldg_skey" & _
                     " WHERE D.use_ind = 1 AND D.facility1_class_id = '" & m_ClassID & "' AND D.proj_bldg_functional_uom_qty > 0 AND C.class_id = 'S'"
        Else
            strSQL = "SELECT count(C.proj_bldg_skey) as proj_num" & _
                     "  FROM PROJ_BLDG_COMPONENT_COST C INNER JOIN PROJECT_BUILDING_DETAIL D ON C.proj_bldg_skey = D.proj_bldg_skey " & _
                     " WHERE D.use_ind = 1 AND D.facility1_class_id = '" & m_ClassID & "' AND C.class_id = '" & TDBGridAnalysis.Columns("ID") & "' "
        End If
        objRS.Open strSQL, g_cnShared

        v12 = Round(objRS("proj_num") / 2 + 0.6) + 6
        v14 = Round(objRS("proj_num") / 4 + 0.6) + 6
        v34 = Round(objRS("proj_num") / 4 * 3 + 0.6) + 6
    
        '  Populate system generated values
        If objRS("proj_num") > 4 Then
            txtSys12.Text = TDBGridAnalysis.Columns(v12)
            txtSys14.Text = TDBGridAnalysis.Columns(v14)
            txtSys34.Text = TDBGridAnalysis.Columns(v34)
        Else
            txtSys12.Text = ""
            txtSys14.Text = ""
            txtSys34.Text = ""
        End If
        '  highlight cells
            TDBGridAnalysis.Columns(m_last_v14).AddCellStyle 8, m_original_style
            TDBGridAnalysis.Columns(m_last_v12).AddCellStyle 8, m_original_style
            TDBGridAnalysis.Columns(m_last_v34).AddCellStyle 8, m_original_style
            TDBGridAnalysis.SelBookmarks.Add (TDBGridAnalysis.Bookmark)
            If objRS("proj_num") > 4 Then
'*** APEX Migration Utility Code Change ***
'                Dim S As New TrueOleDBGrid70.Style
                Dim S As New TrueOleDBGrid80.Style
                S.BackColor = vbRed
                TDBGridAnalysis.Columns(v12).AddCellStyle 8, S
                TDBGridAnalysis.Columns(v14).AddCellStyle 8, S
                TDBGridAnalysis.Columns(v34).AddCellStyle 8, S
            End If
        
        objRS.Close
        strSQL = "SELECT * FROM PUBLISHED_PROJECT_COST WHERE Facility_class_id = '" & m_ClassID & "' AND class_id = '" & TDBGridAnalysis.Columns("ID") & "' AND term_date = '2038-12-31 23:59:59.990'"
        objRS.Open strSQL, g_cnShared
        If Not objRS.EOF Then
            If opDollar.Value = True Then
                If opImperial.Value = True Then
                    txtPub12.Text = objRS("med_unit_cost")
                    txtPub14.Text = objRS("qtr1_unit_cost")
                    txtPub34.Text = objRS("qtr3_unit_cost")
                ElseIf opMetric.Value = True Then
                    txtPub12.Text = objRS("metric_med_unit_cost")
                    txtPub14.Text = objRS("metric_qtr1_unit_cost")
                    txtPub34.Text = objRS("metric_qtr3_unit_cost")
                End If
            ElseIf opPercent.Value = True Then
                txtPub12.Text = objRS("med_total_pct")
                txtPub14.Text = objRS("qtr1_total_pct")
                txtPub34.Text = objRS("qtr3_total_pct")
            End If
        Else
            txtPub12.Text = ""
            txtPub14.Text = ""
            txtPub34.Text = ""
        End If
        objRS.Close
        Set objRS = Nothing
        Dim I
        For I = 0 To UBound(m_aClassID)
            If m_aClassID(I) = TDBGridAnalysis.Columns("class_id") Then
                txtPub12.Text = m_aMed(I)
                txtPub14.Text = m_aQtr1(I)
                txtPub34.Text = m_aQtr3(I)
            End If
        Next
        Screen.MousePointer = vbNormal
        Status ("")
        ' store the col position and row position in global variable for later use
        m_last_v14 = v14
        m_last_v12 = v12
        m_last_v34 = v34
        If Not IsNull(TDBGridAnalysis.Bookmark) Then
            m_last_bookmark = TDBGridAnalysis.Bookmark
        End If
    Else
        TDBGridAnalysis.SelBookmarks.Add (TDBGridAnalysis.Bookmark)
    End If
    '
    '   Populate the dlgOutput form based upon the currently
    '   selected row, only if the user moved to a new row.
    position_output
End Sub

Private Sub TDBGridAnalysis_AfterColUpdate(ByVal ColIndex As Integer)
    buUpdate.Enabled = True
End Sub

Private Sub buUpdate_Click()
    Dim strSQL  As String
    Dim objRS   As New ADODB.RecordSet
    Dim strType As String
    Dim strSystem As String
    Dim qtr1    As String
    Dim med     As String
    Dim qtr3    As String
    Dim varBookmark As Variant
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status ("Updating Analysis Details ...")
    
    If opDollar.Value = True Then
        strType = "dollar"
    ElseIf opPercent.Value = True Then
        strType = "percent"
    Else
        '7/7/2005 SET DEFAULT - CR#1468
        opDollar.Value = True
        strType = "dollar"
    End If
    If opImperial.Value = True Then
        strSystem = "imperial"
    ElseIf opMetric.Value = True Then
        strSystem = "metric"
    Else
        '7/7/2005 SET DEFAULT - CR#1468
        opImperial.Value = True
        strSystem = "imperial"
    End If
    
    Dim I As Integer
    For I = 0 To UBound(m_aClassID)
        If m_aClassID(I) <> "" Then
        ' 8/18/2005 RTD - Test for numeric values; convert to Double to remove formatting
        ' Corrects problem reported by T. Dion/J. Murphy
            If m_aQtr1(I) = "" Or Not IsNumeric(m_aQtr1(I)) Then
                qtr1 = "0"
            Else
                qtr1 = m_aQtr1(I)
            End If
            If m_aMed(I) = "" Or Not IsNumeric(m_aMed(I)) Then
                med = "0"
            Else
                med = m_aMed(I)
            End If
            If m_aQtr3(I) = "" Or Not IsNumeric(m_aQtr3(I)) Then
                qtr3 = "0"
            Else
                qtr3 = m_aQtr3(I)
            End If
            strSQL = "EXEC sp_update_project_analysis @fid = '" & m_ClassID & _
                        "', @cid = '" & m_aClassID(I) & _
                        "', @qtr1 = " & CDbl(qtr1) & _
                        ", @med = " & CDbl(med) & _
                        ", @qtr3 = " & CDbl(qtr3) & _
                        ", @type = '" & strType & _
                        "', @system = '" & strSystem & "'"
            If Not g_objDAL.GetRecordset(vbNullString, strSQL, objRS) Then
                m_blnWereErrors = True
                Screen.MousePointer = vbNormal
                MsgBox "An error occurred while updating the values:" & vbCrLf & g_objDAL.LastErrorDescription, vbExclamation
            Else
                m_blnWereErrors = False
                buClone.Enabled = False
            End If
            objRS.Close
        End If
    Next
    Set objRS = Nothing
    ReDim m_aClassID(0)
    ReDim m_aQtr1(0)
    ReDim m_aMed(0)
    ReDim m_aQtr3(0)
    
    varBookmark = TDBGridAnalysis.Bookmark
    TDBGridAnalysis.Update
    m_objGridMap.Update m_ClassID
    TDBGridAnalysis.Bookmark = varBookmark
    buUpdate.Enabled = False
    Screen.MousePointer = vbNormal
    Status ("")
    ' 8/18/2005 RTD - Don't show Update Complete if an error occurred.
    If Not m_blnWereErrors Then
        MsgBox "Update Completed.", vbInformation
    End If
    
End Sub

Private Sub buClone_Click()
    Dim Button
    Dim strSQL As String
    Dim objRS As New ADODB.RecordSet
    
    Button = MsgBox("WARNING!!!  Running annual clone will close off existing values, please click yes to proceed.", vbYesNo, "Warning!!!  Annual Clone")
    If Button = vbYes Then
        strSQL = "EXEC sp_project_analysis_annual_clone"
        If Not g_objDAL.GetRecordset(vbNullString, strSQL, objRS) Then
            MsgBox "An error occurred while running annual clone"
        Else
            MsgBox "Annual clone completed"
            buClone.Enabled = False
        End If
        objRS.Close
        Set objRS = Nothing
    Else
        MsgBox "Annual clone has been canceled!!"
    End If
End Sub

Private Sub savePublishValues()
    Dim bInList As Boolean
    Dim I
    bInList = False
    For I = 0 To UBound(m_aClassID)
        If m_aClassID(I) = m_ColClassID Then
            bInList = True
            m_aQtr1(I) = txtPub14.Text
            m_aMed(I) = txtPub12.Text
            m_aQtr3(I) = txtPub34.Text
        End If
    Next
    If Not bInList Then
        ReDim Preserve m_aClassID(UBound(m_aClassID) + 1)
        ReDim Preserve m_aQtr1(UBound(m_aQtr1) + 1)
        ReDim Preserve m_aMed(UBound(m_aMed) + 1)
        ReDim Preserve m_aQtr3(UBound(m_aQtr3) + 1)
        m_aClassID(UBound(m_aClassID)) = m_ColClassID
        m_aQtr1(UBound(m_aQtr1)) = txtPub14.Text
        m_aMed(UBound(m_aMed)) = txtPub12.Text
        m_aQtr3(UBound(m_aQtr3)) = txtPub34.Text
    End If
    buUpdate.Enabled = True
End Sub

Private Sub txtPub12_LostFocus()
    If m_oMed <> txtPub12.Text Then
        savePublishValues
    End If
End Sub

Private Sub txtPub12_GotFocus()
    m_ColClassID = TDBGridAnalysis.Columns("class_id")
    m_oMed = txtPub12.Text
    Set m_PubBox = txtPub12
End Sub

Private Sub txtPub14_GotFocus()
    m_ColClassID = TDBGridAnalysis.Columns("class_id")
    m_oQtr1 = txtPub14.Text
    Set m_PubBox = txtPub14
End Sub

Private Sub txtPub14_LostFocus()
    If m_oQtr1 <> txtPub14.Text Then
        savePublishValues
    End If
End Sub

Private Sub txtPub34_GotFocus()
    m_ColClassID = TDBGridAnalysis.Columns("class_id")
    m_oQtr3 = txtPub34.Text
    Set m_PubBox = txtPub34
End Sub

Private Sub txtPub34_LostFocus()
    If m_oQtr3 <> txtPub34.Text Then
        savePublishValues
    End If
End Sub
