VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBuilding 
   Caption         =   "Building Maintenance"
   ClientHeight    =   6600
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11445
   Icon            =   "frmFacility2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   11445
   Begin VB.CommandButton cmdDeleteClone 
      Caption         =   "Delete Clone"
      Height          =   495
      Left            =   9315
      TabIndex        =   74
      Top             =   6000
      Width           =   1000
   End
   Begin VB.Frame fraGoTo 
      Caption         =   "Go To"
      Height          =   775
      Left            =   0
      TabIndex        =   179
      Top             =   5805
      Width           =   5220
      Begin VB.CommandButton cmdCommonAdditiveReport 
         Caption         =   "Common Adds Report"
         Height          =   440
         Left            =   3840
         TabIndex        =   337
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdReports 
         Caption         =   "Summary &Report"
         Height          =   440
         Left            =   2400
         TabIndex        =   78
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdNewModel 
         Caption         =   "&New Model"
         Height          =   440
         Left            =   1320
         TabIndex        =   77
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdModelMaint 
         Caption         =   "&Model"
         Height          =   440
         Left            =   240
         TabIndex        =   76
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   10380
      TabIndex        =   75
      Top             =   6000
      Width           =   1000
   End
   Begin VB.TextBox txtlast_update_person 
      BackColor       =   &H00C0C0C0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   315
      Left            =   6495
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Tag             =   "S"
      Top             =   6240
      Width           =   1290
   End
   Begin VB.TextBox txtlast_update_date 
      BackColor       =   &H00C0C0C0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   315
      Left            =   6495
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2640
   End
   Begin VB.TextBox txtbldg_skey 
      BackColor       =   &H00C0C0C0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   315
      Left            =   8385
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Tag             =   "1N"
      Top             =   6240
      Width           =   750
   End
   Begin TabDlg.SSTab tabBldgAdditions 
      Height          =   2805
      Left            =   0
      TabIndex        =   82
      Top             =   3000
      Width           =   11375
      _ExtentX        =   20055
      _ExtentY        =   4948
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Building Details"
      TabPicture(0)   =   "frmFacility2.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblbldg_id"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblbldg_category"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblbldg_desc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblbldg_part_density"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblbldg_stories"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblbldg_door_density"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblbldg_arch_fees"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblbldg_part_hgt"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblbldg_stories_hgt"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblbldg_wall_factor"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblbldg_fixture_area"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblop_factor"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblgraphic_ref_num"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblgraphic_ref_num2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblbldg_elev_no"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblColumnToBold"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblRowToBold"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblResiBldgType"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblWindowArea"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblMdlRegionCode"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblMdlCountryCode"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtbldg_id"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cbobldg_categoryC"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtbldg_desc"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtbldg_stories"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtbldg_door_density"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtarchitect_fee"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtbldg_stories_hgt"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtbldg_wall_factor"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtbldg_elev_no"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtop_factor"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtbldg_part_density"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtbldg_part_hgt"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtbldg_fixture_area"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtgraphic_ref_id"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtgraphic_ref_id2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cbobldg_categoryR"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cboColumnToBold"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cboRowToBold"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdGraphic1File"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdGraphic2File"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cboResiBldgType"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtwindow_area"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cboMdlRegionCode"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cboMdlCountryCode"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtBldgCostDesc"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "fratype_code"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "fraOPCode"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).ControlCount=   48
      TabCaption(1)   =   "C&ommon Additives"
      TabPicture(1)   =   "frmFacility2.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblComAddsRowCount"
      Tab(1).Control(1)=   "TDBGridAdds"
      Tab(1).Control(2)=   "cmdDeleteAdditive"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).ControlCount=   4
      Begin VB.Frame fraOPCode 
         Caption         =   "OP Code"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   5565
         TabIndex        =   213
         Top             =   360
         Width           =   1755
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1575
            TabIndex        =   341
            Top             =   240
            Width           =   1575
            Begin VB.OptionButton optUnion 
               Caption         =   "Union"
               Height          =   255
               Left            =   0
               TabIndex        =   343
               Top             =   0
               Value           =   -1  'True
               Width           =   795
            End
            Begin VB.OptionButton optOpen 
               Caption         =   "Open"
               Height          =   240
               Left            =   840
               TabIndex        =   342
               Top             =   0
               Width           =   735
            End
         End
      End
      Begin VB.Frame fratype_code 
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2520
         TabIndex        =   180
         Top             =   880
         Width           =   2805
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   2535
            TabIndex        =   338
            Top             =   120
            Width           =   2535
            Begin VB.OptionButton opttype_codeR 
               Caption         =   "Residential"
               Height          =   240
               Left            =   1365
               TabIndex        =   340
               Top             =   0
               Width           =   1155
            End
            Begin VB.OptionButton opttype_codeC 
               Caption         =   "Commercial"
               Height          =   255
               Left            =   0
               TabIndex        =   339
               Top             =   0
               Value           =   -1  'True
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Go To"
         Height          =   675
         Left            =   -74895
         TabIndex        =   334
         Top             =   2070
         Width           =   2730
         Begin VB.CommandButton cmdAssemblyCost 
            Caption         =   "&Assembly Cost"
            Height          =   420
            Left            =   1320
            TabIndex        =   336
            Top             =   205
            Width           =   1245
         End
         Begin VB.CommandButton cmdUnitCost 
            Caption         =   "Unit &Cost"
            Height          =   415
            Left            =   240
            TabIndex        =   335
            Top             =   205
            Width           =   1000
         End
      End
      Begin VB.TextBox txtBldgCostDesc 
         Height          =   525
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   212
         Top             =   360
         Width           =   5225
      End
      Begin VB.ComboBox cboMdlCountryCode 
         Height          =   315
         ItemData        =   "frmFacility2.frx":0182
         Left            =   9950
         List            =   "frmFacility2.frx":0189
         Style           =   2  'Dropdown List
         TabIndex        =   209
         Top             =   490
         Width           =   1035
      End
      Begin VB.ComboBox cboMdlRegionCode 
         Height          =   315
         ItemData        =   "frmFacility2.frx":0192
         Left            =   8085
         List            =   "frmFacility2.frx":0199
         Style           =   2  'Dropdown List
         TabIndex        =   208
         Top             =   490
         Width           =   1035
      End
      Begin VB.TextBox txtwindow_area 
         Height          =   315
         Left            =   6405
         TabIndex        =   65
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox cboResiBldgType 
         Height          =   315
         ItemData        =   "frmFacility2.frx":01A2
         Left            =   8400
         List            =   "frmFacility2.frx":01A4
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1320
         Width           =   2800
      End
      Begin VB.CommandButton cmdGraphic2File 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10815
         TabIndex        =   71
         ToolTipText     =   "Browse for file..."
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdGraphic1File 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10815
         TabIndex        =   69
         ToolTipText     =   "Browse for file..."
         Top             =   1680
         Width           =   375
      End
      Begin VB.ComboBox cboRowToBold 
         Height          =   315
         ItemData        =   "frmFacility2.frx":01A6
         Left            =   1080
         List            =   "frmFacility2.frx":01A8
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   2400
         Width           =   5355
      End
      Begin VB.ComboBox cboColumnToBold 
         Height          =   315
         ItemData        =   "frmFacility2.frx":01AA
         Left            =   7740
         List            =   "frmFacility2.frx":01AC
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   2400
         Width           =   3465
      End
      Begin VB.CommandButton cmdDeleteAdditive 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   -64800
         TabIndex        =   196
         Top             =   2160
         Width           =   1005
      End
      Begin VB.ComboBox cbobldg_categoryR 
         Height          =   315
         ItemData        =   "frmFacility2.frx":01AE
         Left            =   6405
         List            =   "frmFacility2.frx":01B0
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   960
         Width           =   2745
      End
      Begin VB.TextBox txtgraphic_ref_id2 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   9660
         TabIndex        =   70
         Top             =   2040
         Width           =   1140
      End
      Begin VB.TextBox txtgraphic_ref_id 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   9660
         TabIndex        =   68
         Top             =   1680
         Width           =   1140
      End
      Begin VB.TextBox txtbldg_fixture_area 
         Height          =   315
         Left            =   6405
         TabIndex        =   56
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtbldg_part_hgt 
         Height          =   315
         Left            =   2805
         TabIndex        =   61
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtbldg_part_density 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2805
         TabIndex        =   60
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtop_factor 
         Height          =   315
         Left            =   8085
         TabIndex        =   66
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtbldg_elev_no 
         Height          =   315
         Left            =   4590
         TabIndex        =   63
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtbldg_wall_factor 
         Height          =   315
         Left            =   6405
         TabIndex        =   64
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtbldg_stories_hgt 
         Height          =   315
         Left            =   1075
         TabIndex        =   59
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtarchitect_fee 
         Height          =   315
         Left            =   8085
         TabIndex        =   67
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtbldg_door_density 
         Height          =   315
         Left            =   4590
         TabIndex        =   62
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtbldg_stories 
         Height          =   315
         Left            =   1075
         TabIndex        =   58
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtbldg_desc 
         Height          =   315
         Left            =   1075
         MaxLength       =   75
         TabIndex        =   55
         Top             =   1320
         Width           =   4250
      End
      Begin VB.ComboBox cbobldg_categoryC 
         Height          =   315
         ItemData        =   "frmFacility2.frx":01B2
         Left            =   6405
         List            =   "frmFacility2.frx":01B4
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   960
         Width           =   2745
      End
      Begin VB.TextBox txtbldg_id 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1075
         TabIndex        =   52
         Top             =   960
         Width           =   1320
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridAdds 
         Height          =   1695
         Left            =   -74895
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2990
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
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
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
      Begin VB.Label lblMdlCountryCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Country:"
         Height          =   255
         Left            =   9240
         TabIndex        =   211
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lblMdlRegionCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Region:"
         Height          =   255
         Left            =   7350
         TabIndex        =   210
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lblWindowArea 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Window Area:"
         Height          =   255
         Left            =   5355
         TabIndex        =   207
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label lblResiBldgType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Building Type:"
         Height          =   255
         Left            =   7245
         TabIndex        =   206
         Top             =   1365
         Width           =   1065
      End
      Begin VB.Label lblRowToBold 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Row To Bold:"
         Height          =   255
         Left            =   0
         TabIndex        =   199
         Top             =   2475
         Width           =   1065
      End
      Begin VB.Label lblColumnToBold 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Column To Bold:"
         Height          =   255
         Left            =   6510
         TabIndex        =   198
         Top             =   2475
         Width           =   1170
      End
      Begin VB.Label lblComAddsRowCount 
         Alignment       =   2  'Center
         Caption         =   "0 rows returned"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -72270
         TabIndex        =   197
         Top             =   2520
         Width           =   6510
      End
      Begin VB.Label lblbldg_elev_no 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No Elevators:"
         Height          =   255
         Left            =   3540
         TabIndex        =   195
         Top             =   2100
         Width           =   1005
      End
      Begin VB.Label lblgraphic_ref_num2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Graphic 2:"
         Height          =   255
         Left            =   8820
         TabIndex        =   194
         Top             =   2100
         Width           =   795
      End
      Begin VB.Label lblgraphic_ref_num 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Graphic 1:"
         Height          =   255
         Left            =   8820
         TabIndex        =   193
         Top             =   1740
         Width           =   795
      End
      Begin VB.Label lblop_factor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "OP Factor:"
         Height          =   255
         Left            =   7215
         TabIndex        =   192
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label lblbldg_fixture_area 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fixture Area:"
         Height          =   255
         Left            =   5445
         TabIndex        =   191
         Top             =   1365
         Width           =   915
      End
      Begin VB.Label lblbldg_wall_factor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ext Wall Fact:"
         Height          =   255
         Left            =   5355
         TabIndex        =   190
         Top             =   1740
         Width           =   1035
      End
      Begin VB.Label lblbldg_stories_hgt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stories Hgt:"
         Height          =   255
         Left            =   105
         TabIndex        =   189
         Top             =   2100
         Width           =   855
      End
      Begin VB.Label lblbldg_part_hgt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Part Height:"
         Height          =   255
         Left            =   1785
         TabIndex        =   188
         Top             =   2100
         Width           =   945
      End
      Begin VB.Label lblbldg_arch_fees 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Arch Fees:"
         Height          =   255
         Left            =   7125
         TabIndex        =   187
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label lblbldg_door_density 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Door Density:"
         Height          =   255
         Left            =   3540
         TabIndex        =   186
         Top             =   1740
         Width           =   1005
      End
      Begin VB.Label lblbldg_stories 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stories:"
         Height          =   255
         Left            =   105
         TabIndex        =   185
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label lblbldg_part_density 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Part Density:"
         Height          =   255
         Left            =   1785
         TabIndex        =   184
         Top             =   1740
         Width           =   945
      End
      Begin VB.Label lblbldg_desc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Building:"
         Height          =   255
         Left            =   105
         TabIndex        =   183
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label lblbldg_category 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         Height          =   255
         Left            =   5565
         TabIndex        =   182
         Top             =   1020
         Width           =   750
      End
      Begin VB.Label lblbldg_id 
         Alignment       =   1  'Right Justify
         Caption         =   "Building ID:"
         Height          =   255
         Left            =   105
         TabIndex        =   181
         Top             =   1020
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog oDlg 
      Left            =   10710
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Location"
   End
   Begin VB.Frame fraNewBldgModelMatrix 
      Height          =   2655
      Left            =   30
      TabIndex        =   201
      Top             =   120
      Width           =   11340
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   10620
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   10620
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   9975
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   9975
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   120
         Width           =   650
      End
      Begin VB.ComboBox cboFrameType 
         Height          =   315
         Index           =   5
         ItemData        =   "frmFacility2.frx":01B6
         Left            =   4470
         List            =   "frmFacility2.frx":01B8
         TabIndex        =   51
         Text            =   "cboFrameType"
         Top             =   2355
         Width           =   1825
      End
      Begin VB.ComboBox cboWallType 
         Height          =   315
         Index           =   5
         ItemData        =   "frmFacility2.frx":01BA
         Left            =   30
         List            =   "frmFacility2.frx":01BC
         TabIndex        =   50
         Text            =   "cboWallType"
         Top             =   2355
         Width           =   4455
      End
      Begin VB.ComboBox cboFrameType 
         Height          =   315
         Index           =   4
         ItemData        =   "frmFacility2.frx":01BE
         Left            =   4470
         List            =   "frmFacility2.frx":01C0
         TabIndex        =   49
         Text            =   "cboFrameType"
         Top             =   2010
         Width           =   1825
      End
      Begin VB.ComboBox cboWallType 
         Height          =   315
         Index           =   4
         ItemData        =   "frmFacility2.frx":01C2
         Left            =   30
         List            =   "frmFacility2.frx":01C4
         TabIndex        =   48
         Text            =   "cboWallType"
         Top             =   2010
         Width           =   4455
      End
      Begin VB.ComboBox cboFrameType 
         Height          =   315
         Index           =   3
         ItemData        =   "frmFacility2.frx":01C6
         Left            =   4470
         List            =   "frmFacility2.frx":01C8
         TabIndex        =   47
         Text            =   "cboFrameType"
         Top             =   1680
         Width           =   1825
      End
      Begin VB.ComboBox cboWallType 
         Height          =   315
         Index           =   3
         ItemData        =   "frmFacility2.frx":01CA
         Left            =   30
         List            =   "frmFacility2.frx":01CC
         TabIndex        =   46
         Text            =   "cboWallType"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.ComboBox cboFrameType 
         Height          =   315
         Index           =   2
         ItemData        =   "frmFacility2.frx":01CE
         Left            =   4470
         List            =   "frmFacility2.frx":01D0
         TabIndex        =   45
         Text            =   "cboFrameType"
         Top             =   1350
         Width           =   1825
      End
      Begin VB.ComboBox cboWallType 
         Height          =   315
         Index           =   2
         ItemData        =   "frmFacility2.frx":01D2
         Left            =   30
         List            =   "frmFacility2.frx":01D4
         TabIndex        =   44
         Text            =   "cboWallType"
         Top             =   1350
         Width           =   4455
      End
      Begin VB.ComboBox cboFrameType 
         Height          =   315
         Index           =   1
         ItemData        =   "frmFacility2.frx":01D6
         Left            =   4470
         List            =   "frmFacility2.frx":01D8
         TabIndex        =   43
         Text            =   "cboFrameType"
         Top             =   1020
         Width           =   1825
      End
      Begin VB.ComboBox cboWallType 
         Height          =   315
         Index           =   1
         ItemData        =   "frmFacility2.frx":01DA
         Left            =   30
         List            =   "frmFacility2.frx":01DC
         TabIndex        =   42
         Text            =   "cboWallType"
         Top             =   1020
         Width           =   4455
      End
      Begin VB.ComboBox cboFrameType 
         Height          =   315
         Index           =   0
         ItemData        =   "frmFacility2.frx":01DE
         Left            =   4470
         List            =   "frmFacility2.frx":01E0
         TabIndex        =   41
         Text            =   "cboFrameType"
         Top             =   690
         Width           =   1825
      End
      Begin VB.ComboBox cboWallType 
         Height          =   315
         Index           =   0
         ItemData        =   "frmFacility2.frx":01E2
         Left            =   30
         List            =   "frmFacility2.frx":01E4
         TabIndex        =   40
         Text            =   "cboWallType"
         Top             =   690
         Width           =   4455
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4170
         TabIndex        =   18
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4170
         TabIndex        =   29
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4815
         TabIndex        =   19
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4815
         TabIndex        =   30
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5460
         TabIndex        =   20
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5460
         TabIndex        =   31
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6105
         TabIndex        =   21
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6105
         TabIndex        =   32
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6750
         TabIndex        =   22
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6750
         TabIndex        =   33
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   7395
         TabIndex        =   23
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   7395
         TabIndex        =   34
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   8040
         TabIndex        =   24
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   8040
         TabIndex        =   35
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   8685
         TabIndex        =   25
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   8685
         TabIndex        =   36
         Top             =   405
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   9330
         TabIndex        =   26
         Top             =   120
         Width           =   650
      End
      Begin VB.TextBox txtNewBldgPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   9330
         TabIndex        =   37
         Top             =   405
         Width           =   650
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   0
         X2              =   11240
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L.F.Perimeter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2595
         TabIndex        =   205
         Top             =   405
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S.F. Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2595
         TabIndex        =   204
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Exterior Wall Type "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   30
         TabIndex        =   203
         Top             =   120
         Width           =   2565
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&& Structural System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   30
         TabIndex        =   202
         Top             =   405
         Width           =   2565
      End
      Begin VB.Shape shpWhiteBackground 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   1980
         Left            =   6240
         Top             =   720
         Width           =   4995
      End
   End
   Begin VB.Frame fraModelMatrix 
      Height          =   2940
      Left            =   0
      TabIndex        =   86
      Top             =   0
      Width           =   11340
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   10620
         TabIndex        =   17
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   10620
         TabIndex        =   8
         Top             =   120
         Width           =   700
      End
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   9915
         TabIndex        =   16
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   9915
         TabIndex        =   7
         Top             =   120
         Width           =   700
      End
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   9210
         TabIndex        =   15
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   9210
         TabIndex        =   6
         Top             =   120
         Width           =   700
      End
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   8505
         TabIndex        =   14
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   8505
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   7800
         TabIndex        =   13
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   7800
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7095
         TabIndex        =   12
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7095
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6390
         TabIndex        =   11
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6390
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5685
         TabIndex        =   10
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5685
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.TextBox txtPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4980
         TabIndex        =   9
         Top             =   405
         Width           =   700
      End
      Begin VB.TextBox txtArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4980
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.Shape shpSelectedAreaPerimeter 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   285
         Left            =   4980
         Top             =   690
         Width           =   700
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&& Structural System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   178
         Top             =   405
         Width           =   3240
      End
      Begin VB.Label lblCol9_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   10620
         TabIndex        =   177
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblCol8_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   9915
         TabIndex        =   176
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblCol7_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   9210
         TabIndex        =   175
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblCol6_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   8505
         TabIndex        =   174
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblCol5_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7800
         TabIndex        =   173
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblCol4_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7095
         TabIndex        =   172
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblCol3_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   6390
         TabIndex        =   171
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblCol2_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   5685
         TabIndex        =   170
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblCol1_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   4980
         TabIndex        =   169
         Top             =   2610
         Width           =   700
      End
      Begin VB.Label lblWall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   25
         TabIndex        =   168
         Top             =   2610
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   3275
         TabIndex        =   167
         Top             =   2610
         Width           =   1710
      End
      Begin VB.Label lblCol9_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   10620
         TabIndex        =   166
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblCol8_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   9915
         TabIndex        =   165
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblCol7_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   9210
         TabIndex        =   164
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblCol6_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   8505
         TabIndex        =   163
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblCol5_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   7800
         TabIndex        =   162
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblCol4_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   7095
         TabIndex        =   161
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblCol3_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6390
         TabIndex        =   160
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblCol2_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   5685
         TabIndex        =   159
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblCol1_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   4980
         TabIndex        =   158
         Top             =   2325
         Width           =   700
      End
      Begin VB.Label lblWall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   25
         TabIndex        =   157
         Top             =   2325
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   3275
         TabIndex        =   156
         Top             =   2325
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Exterior Wall Type "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   155
         Top             =   120
         Width           =   3240
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   3275
         TabIndex        =   154
         Top             =   2040
         Width           =   1710
      End
      Begin VB.Label lblWall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   25
         TabIndex        =   153
         Top             =   2040
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCol1_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4980
         TabIndex        =   152
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblCol2_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5685
         TabIndex        =   151
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblCol3_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   6390
         TabIndex        =   150
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblCol4_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   7095
         TabIndex        =   149
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblCol5_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   7800
         TabIndex        =   148
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblCol6_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   8505
         TabIndex        =   147
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblCol7_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   9210
         TabIndex        =   146
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblCol8_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   9915
         TabIndex        =   145
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblCol9_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   10620
         TabIndex        =   144
         Top             =   1770
         Width           =   700
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   3275
         TabIndex        =   143
         Top             =   1770
         Width           =   1710
      End
      Begin VB.Label lblWall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   25
         TabIndex        =   142
         Top             =   1770
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3275
         TabIndex        =   141
         Top             =   1500
         Width           =   1710
      End
      Begin VB.Label lblWall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   25
         TabIndex        =   140
         Top             =   1500
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3275
         TabIndex        =   139
         Top             =   1230
         Width           =   1710
      End
      Begin VB.Label lblWall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   25
         TabIndex        =   138
         Top             =   1230
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   3275
         TabIndex        =   137
         Top             =   960
         Width           =   1710
      End
      Begin VB.Label lblWall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   25
         TabIndex        =   136
         Top             =   960
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   25
         TabIndex        =   135
         Top             =   690
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCol1_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   4980
         TabIndex        =   134
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol2_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5685
         TabIndex        =   133
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol3_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   6390
         TabIndex        =   132
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol4_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   7095
         TabIndex        =   131
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol5_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   7800
         TabIndex        =   130
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol6_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   8505
         TabIndex        =   129
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol7_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   9210
         TabIndex        =   128
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol8_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   9915
         TabIndex        =   127
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol9_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   10620
         TabIndex        =   126
         Top             =   2040
         Width           =   700
      End
      Begin VB.Label lblCol9_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   10620
         TabIndex        =   125
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol9_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   10620
         TabIndex        =   124
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol9_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   10620
         TabIndex        =   123
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol9_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   10620
         TabIndex        =   122
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblCol8_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   9915
         TabIndex        =   121
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol8_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   9915
         TabIndex        =   120
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol8_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   9915
         TabIndex        =   119
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol8_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   9915
         TabIndex        =   118
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblCol7_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   9210
         TabIndex        =   117
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol7_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   9210
         TabIndex        =   116
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol7_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   9210
         TabIndex        =   115
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol7_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   9210
         TabIndex        =   114
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblCol6_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   8505
         TabIndex        =   113
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol6_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   8505
         TabIndex        =   112
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol6_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   8505
         TabIndex        =   111
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol6_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   8505
         TabIndex        =   110
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblCol5_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   7800
         TabIndex        =   109
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol5_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   7800
         TabIndex        =   108
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol5_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   7800
         TabIndex        =   107
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol5_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   7800
         TabIndex        =   106
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblCol4_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   7095
         TabIndex        =   105
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol4_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   7095
         TabIndex        =   104
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol4_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   7095
         TabIndex        =   103
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol4_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   7095
         TabIndex        =   102
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblCol3_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   6390
         TabIndex        =   101
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol3_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   6390
         TabIndex        =   100
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol3_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   6390
         TabIndex        =   99
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol3_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   6390
         TabIndex        =   98
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblCol2_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   5685
         TabIndex        =   97
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol2_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   5685
         TabIndex        =   96
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol2_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   5685
         TabIndex        =   95
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol2_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   5685
         TabIndex        =   94
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblCol1_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4980
         TabIndex        =   93
         Top             =   690
         Width           =   700
      End
      Begin VB.Label lblCol1_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4980
         TabIndex        =   92
         Top             =   960
         Width           =   700
      End
      Begin VB.Label lblCol1_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   4980
         TabIndex        =   91
         Top             =   1230
         Width           =   700
      End
      Begin VB.Label lblCol1_TotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4980
         TabIndex        =   90
         Top             =   1500
         Width           =   700
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   3275
         TabIndex        =   89
         Top             =   690
         Width           =   1710
      End
      Begin VB.Label lblSFArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S.F. Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3275
         TabIndex        =   88
         Top             =   120
         Width           =   1710
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L.F.Perimeter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3275
         TabIndex        =   87
         Top             =   405
         Width           =   1710
      End
   End
   Begin VB.Frame fraModelMatrixResi 
      Height          =   2940
      Left            =   0
      TabIndex        =   214
      Top             =   0
      Width           =   11340
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   9900
         TabIndex        =   217
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   9210
         Locked          =   -1  'True
         TabIndex        =   236
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   9210
         TabIndex        =   235
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   234
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   8520
         TabIndex        =   233
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   232
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   7830
         TabIndex        =   231
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   230
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   7140
         TabIndex        =   229
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6450
         Locked          =   -1  'True
         TabIndex        =   228
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6450
         TabIndex        =   227
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   226
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   5760
         TabIndex        =   225
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   224
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5070
         TabIndex        =   223
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   222
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4380
         TabIndex        =   221
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   220
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3690
         TabIndex        =   219
         Top             =   120
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   9900
         Locked          =   -1  'True
         TabIndex        =   218
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   10590
         Locked          =   -1  'True
         TabIndex        =   216
         Top             =   405
         Width           =   690
      End
      Begin VB.TextBox txtAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   10590
         TabIndex        =   215
         Top             =   120
         Width           =   690
      End
      Begin VB.Shape shpSelectedAreaPerimeterResi 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   285
         Left            =   3690
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol9_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   9210
         TabIndex        =   333
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol8_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   8520
         TabIndex        =   332
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol7_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7830
         TabIndex        =   331
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol6_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7140
         TabIndex        =   330
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol5_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   6450
         TabIndex        =   329
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol4_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   5760
         TabIndex        =   328
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol3_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   5070
         TabIndex        =   327
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol2_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   4380
         TabIndex        =   326
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol1_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   3690
         TabIndex        =   325
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   30
         TabIndex        =   324
         Top             =   2610
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCol9_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   9210
         TabIndex        =   323
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol8_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   8520
         TabIndex        =   322
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol7_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   7830
         TabIndex        =   321
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol6_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   7140
         TabIndex        =   320
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol5_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6450
         TabIndex        =   319
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol4_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   5760
         TabIndex        =   318
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol3_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   5070
         TabIndex        =   317
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol2_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   4380
         TabIndex        =   316
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol1_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   3690
         TabIndex        =   315
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   30
         TabIndex        =   314
         Top             =   2325
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblExteriorWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                                               Exterior Wall"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   30
         TabIndex        =   313
         Top             =   120
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   30
         TabIndex        =   312
         Top             =   2040
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCol1_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   3690
         TabIndex        =   311
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol2_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4380
         TabIndex        =   310
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol3_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5070
         TabIndex        =   309
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol4_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5760
         TabIndex        =   308
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol5_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   6450
         TabIndex        =   307
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol6_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   7140
         TabIndex        =   306
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol7_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   7830
         TabIndex        =   305
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol8_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   8520
         TabIndex        =   304
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol9_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   9210
         TabIndex        =   303
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   30
         TabIndex        =   302
         Top             =   1770
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   30
         TabIndex        =   301
         Top             =   1500
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   30
         TabIndex        =   300
         Top             =   1230
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   30
         TabIndex        =   299
         Top             =   960
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWallResi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   298
         Top             =   690
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCol1_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   3690
         TabIndex        =   297
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol2_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   4380
         TabIndex        =   296
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol3_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5070
         TabIndex        =   295
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol4_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5760
         TabIndex        =   294
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol5_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   6450
         TabIndex        =   293
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol6_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   7140
         TabIndex        =   292
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol7_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   7830
         TabIndex        =   291
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol8_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   8520
         TabIndex        =   290
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol9_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   9210
         TabIndex        =   289
         Top             =   2040
         Width           =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   12180
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label lblCol9_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   9210
         TabIndex        =   288
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol9_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   9210
         TabIndex        =   287
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol9_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   9210
         TabIndex        =   286
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol9_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   9210
         TabIndex        =   285
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol8_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   8520
         TabIndex        =   284
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol8_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   8520
         TabIndex        =   283
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol8_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   8520
         TabIndex        =   282
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol8_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   8520
         TabIndex        =   281
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol7_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   7830
         TabIndex        =   280
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol7_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   7830
         TabIndex        =   279
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol7_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   7830
         TabIndex        =   278
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol7_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   7830
         TabIndex        =   277
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol6_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   7140
         TabIndex        =   276
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol6_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   7140
         TabIndex        =   275
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol6_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   7140
         TabIndex        =   274
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol6_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   7140
         TabIndex        =   273
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol5_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   6450
         TabIndex        =   272
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol5_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   6450
         TabIndex        =   271
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol5_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   6450
         TabIndex        =   270
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol5_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   6450
         TabIndex        =   269
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol4_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   5760
         TabIndex        =   268
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol4_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   267
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol4_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   5760
         TabIndex        =   266
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol4_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   5760
         TabIndex        =   265
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol3_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   5070
         TabIndex        =   264
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol3_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   5070
         TabIndex        =   263
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol3_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   5070
         TabIndex        =   262
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol3_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   5070
         TabIndex        =   261
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol2_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4380
         TabIndex        =   260
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol2_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4380
         TabIndex        =   259
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol2_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   4380
         TabIndex        =   258
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol2_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4380
         TabIndex        =   257
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol1_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   3690
         TabIndex        =   256
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol1_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   3690
         TabIndex        =   255
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol1_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3690
         TabIndex        =   254
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol1_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3690
         TabIndex        =   253
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol10_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   9900
         TabIndex        =   252
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol10_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   9900
         TabIndex        =   251
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol10_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   9900
         TabIndex        =   250
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol10_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   9900
         TabIndex        =   249
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol10_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   9900
         TabIndex        =   248
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol10_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   9900
         TabIndex        =   247
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol10_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   9900
         TabIndex        =   246
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol10_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   9900
         TabIndex        =   245
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label lblCol11_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   10590
         TabIndex        =   244
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label lblCol11_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   10590
         TabIndex        =   243
         Top             =   2325
         Width           =   690
      End
      Begin VB.Label lblCol11_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   10590
         TabIndex        =   242
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblCol11_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   10590
         TabIndex        =   241
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblCol11_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   10590
         TabIndex        =   240
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCol11_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   10590
         TabIndex        =   239
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblCol11_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   10590
         TabIndex        =   238
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label lblCol11_TotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   10590
         TabIndex        =   237
         Top             =   1500
         Width           =   690
      End
   End
   Begin VB.Label lbllast_update_person 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Updated By:"
      Height          =   255
      Left            =   5505
      TabIndex        =   85
      Top             =   6315
      Width           =   915
   End
   Begin VB.Label lbllast_update_date 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Updated:"
      Height          =   255
      Left            =   5625
      TabIndex        =   84
      Top             =   5970
      Width           =   795
   End
   Begin VB.Label lblbldg_skey 
      BackStyle       =   0  'Transparent
      Caption         =   "Skey:"
      Height          =   255
      Left            =   7860
      TabIndex        =   83
      Top             =   6315
      Width           =   495
   End
End
Attribute VB_Name = "frmBuilding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'<modulename> frmBuilding</modulename>
'<functionname>General (Main) </functionname>
'
'<summary>
' Provides u/i permitting user to do the following:
'
'This window provides a MATRIX of models as follows:  (m_recModelMatrix)
'
'"Exterior Wall Type" x SF. AREA
'
'"   There are (6) "models" per column starting just under the "L.F. Perimeter" row.
'MODEL 1, 2, 3, 4, 5, 6
'"   There are (9) columns going across (SF. AREA sizes)
'
'Each of these (6) "models" down a given sf (square foot) area column, can be double-clicked upon to reveal more "drill down" information.
'
'Under the MATRIX there is a tab control that has (2) tabs:
'"   Building Details
'"   Common Additives
'
'LOAD THE WINDOW:
'
'"   PopulateModelMatrix    (ado recordset: m_recModelMatrix)
'"   PopulateModelMatrixCommercial
'"   PopulateModelMatrixResi
'
'THE MATRIX:
'This is a group of label control painstakingly built and placed to resemble a cell matrix.
'It has been setup such that upon clicking any one of the labels a sub named:
'
'GotoModel()
'
'
'is executed.
'Immediately a window is instantiated (frmModel) - "Model Maintenance"
'
'NOTE: See "frmModel.frm" commenting for details of setup.
'
'----------------------------------------------
'MATRIX Layout
'It is (9) columns of "label control" arrays as follows:
'
'
'1       2       3           …   9
'1.  LblCol1_TotalOp(1)  LblCol2_TotalOp(1)  LblCol3_TotalOp(1)  . . . LblCol9_TotalOp(1)
'2.  LblCol1_TotalOp(2)  LblCol2_TotalOp(2)  LblCol3_TotalOp(2)  . . . LblCol9_TotalOp(2)
'3.  LblCol1_TotalOp(3)  LblCol2_TotalOp(3)  LblCol3_TotalOp(3)  . . . LblCol9_TotalOp(3)
'4.  LblCol1_TotalOp(4)  LblCol2_TotalOp(4)  LblCol3_TotalOp(4)  . . . LblCol9_TotalOp(4)
'5.  LblCol1_TotalOp(5)  LblCol2_TotalOp(5)  LblCol3_TotalOp(5)  . . . LblCol9_TotalOp(5)
'6.  LblCol1_TotalOp(6)  LblCol2_TotalOp(6)  LblCol3_TotalOp(6)  . . . LblCol9_TotalOp(6)
'
'For each label array (index) there is a CLICK and DOUBLECLICK event defined.
'Sample As follows:
'--------------------------------------------------------------------------------------------
'Private Sub lblCol1_TotalOP_Click(Index As Integer)
'    '
'    '   If they click on a label that is not populated
'    '   don't do anything, unless this is the 1st time we're
'    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
'    If lblCol1_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
'        '
'        '   We are in whatever row the value of index is and the 1st column.
'        sshpSelectedArea = Index & ",0"
'        SetShpTopLocation
'    End If
'End Sub
'
'Private Sub lblCol1_TotalOP_DblClick(Index As Integer)
'    '
'    '   If they click on a label that is not populated
'    '   don't do anything.
'    If lblCol1_TotalOP(Index).Caption <> "" Then
'        GotoModel
'    End If
'End Sub
'Private Sub lblCol2_TotalOP_Click(Index As Integer)
'    '
'    '   If they click on a label that is not populated
'    '   don't do anything, unless this is the 1st time we're
'    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
'    If lblCol2_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
'        '
'        '   We are in whatever row the value of index is and the 2nd column.
'        sshpSelectedArea = Index & ",1"
'        SetShpTopLocation
'    End If
'End Sub
'
'Private Sub lblCol2_TotalOP_DblClick(Index As Integer)
'    '
'    '   If they click on a label that is not populated
'    '   don't do anything.
'    If lblCol2_TotalOP(Index).Caption <> "" Then
'        GotoModel
'    End If
'End Sub
'-------------------------------------------------------------------------------------------------------------
'
'For every building type/prototype there is a cell in the matrix that is "teal" colored.
'This is the model that goes into the books.  Upon double clicking the "teal" colored cell you get another window/form that contains another matrix that has a "teal" colored column of cells.  This column contain the "O&P" costs that you will find in the books.
'
'In addition to the matrix you will find a tab control with (3) tabs:
'"   Assembly Components (PopulateAssemblyComponents() )
'"   Summary Estimates       (PopulateSummaryEstimateComponents() )
'"   Model Details           (PopulateModelDetails() )
'
'
'HELPER Class: CBuildingMap.Cls
'
'NOTE:  file names do not match Project component names (e.g. frmBuildingGrid is really frmFacilityGrid.frm)
'
'</summary>
'
'<seealso> frmBuildingGrid </seealso>
'<seealso>frmModel.frm</seealso>
'
'<datastruct>m_objGridMap</datastruct>
'<datastruct>m_rec</datastruct>
'<datastruct>m_recModels</datastruct>
'<datastruct> m_recModelMatrix</datastruct>
'
'
'<storedprocedurename> sp_select_model</storedprocedurename>
'<storedprocedurename> sp_select_model_basements</storedprocedurename>
'<storedprocedurename> select_building_pub_matrix_cost </storedprocedurename>
'<storedprocedurename> sp_update_commercial_building</storedprocedurename>
'
'
'
'<returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code> * * * RETRIEVES "basic" information about the model * * *
'exec sp_select_model @type_code = '%', @bldg_category = '%', @bldg_id = '001', @bldg_desc = '%', @frame_type = '%', @wall_type = '%', @bldg_model_skey = '2'
'</code>
' <code> * * * POPULATE THE "MODEL MATRIX"
'exec sp_select_building_pub_matrix_cost @bldg_id = '001', @bldg_model_skey = '%', @bldg_desc = 'Apartment, 1-3 Story', @op_code = 'STD', @country_code = 'USA', @region_code = 'NAT', @type_code = 'C'
'</code>
'<code> * * * UPDATES virtually everything seen on the matrix … However,  I'm not sure this ever gets run?!
'
'exec sp_update_commercial_building @bldg_skey = '1',@bldg_id = '001',@bldg_category = 'Commercial',@bldg_desc = 'Apartment, 1-3 Story',@bldg_stories = 4,@bldg_stories_hgt = 10,@bldg_part_density = 9,@bldg_part_hgt = 8,@bldg_door_density = 80,@bldg_type = '0',@bldg_area_std = 22500,
'@bldg_perimeter_std = 400,@bldg_wall_factor = 0.89,@bldg_elev_no = 1,@bldg_fixture_area = 200,
'@window_area = 15,@op_factor = 0.25,@architect_fee = 0.08,@row_to_bold = 2,@col_to_bold = 5,@graphic_ref_id = 'b200-1f',@graphic_ref_id2 = '', @last_update_id_bldg = '8',@bldg_area_1 = '0',@bldg_perimeter_1 = '0',@bldg_orig_area_1 = '0',@area_ind_1 = '0',@last_update_id_area_1 = '0',
'@bldg_area_2 = '0',@bldg_perimeter_2 = '0',@bldg_orig_area_2 = '0',@area_ind_2 = '0',@last_update_id_area_2 = '0',
'@bldg_area_3 = '0',@bldg_perimeter_3 = '0',@bldg_orig_area_3 = '0',@area_ind_3 = '0',@last_update_id_area_3 = '0',@bldg_area_4 = '0',@bldg_perimeter_4 = '0',@bldg_orig_area_4 = '0',@area_ind_4 = '0',@last_update_id_area_4 = '0',@bldg_area_5 = '0',@
'@bldg_perimeter_5 = '0',@bldg_orig_area_5 = '0',@area_ind_5 = '0',@last_update_id_area_5 = '0',@bldg_area_6 = '0',@bldg_perimeter_6 = '0',@bldg_orig_area_6 = '0',@area_ind_6 = '0',@last_update_id_area_6 = '0',@bldg_area_7 = '0',@bldg_perimeter_7 = '0',@bldg_orig_area_7 = '0',@area_ind_7 = '0',@last_update_id_area_7 = '0',@bldg_area_8 = '0',@bldg_perimeter_8 = '0',@bldg_orig_area_8 = '0',@area_ind_8 = '0',@last_update_id_area_8 = '0',@bldg_area_9 = '0',@bldg_perimeter_9 = '0',@bldg_orig_area_9 = '0',@area_ind_9 = '0',@last_update_id_area_9 = '0',@bldg_form = '33', @last_update_person = 'Hancockrl'
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CBuildingMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>


Dim cnTemp As New ADODB.Connection
Dim m_rec As New ADODB.RecordSet
'
'   Common Additives grid recordset.
Dim m_recComAdds As New ADODB.RecordSet
Dim m_recModels As New ADODB.RecordSet
Dim m_recModelMatrix As New ADODB.RecordSet
'
'   Tells if we are doing an insert or update.
Dim m_blnInsert As Boolean
'
'   Indicate if clone is in progress
Dim m_blnClone As Boolean
'
'   Class to handle Common Adds grid.
Dim m_objComAddsGridMap As New CBldgComAdds
'
'   Notifies that it wants to see changes.
Dim sEventSubscriberID              As String
'
'   Tells us we're loading the screen for the 1st time
'   so the cbo clicks won't run.
Dim bIsInitialLoad  As Boolean
'
'   Indicates user has modified the data.
Dim bIsPendingChange As Boolean
'
'   Indicates a field that will affect overall cost rollups
'   has been changed so we must RefreshCosts if they update.
Dim bRefreshCosts As Boolean
'
'   Indicates where the shpSelectedArea is for modelmaint button click.
'   In the format of 1,1 meaning row 1 col area1.
Dim sshpSelectedArea As String

Const USEBOOKMARK = 1
Const USECOORD = 0
'
'   Used to indicate how the bldg costs are calculated as shown in the book.
Const BLDG_COST_DESC = "Model costs calculated for a [BLDG_STORIES] story building with [STORIES_HGT]' " & _
            "story height and" & vbCrLf & " [BLDG_AREA] square feet of floor area"

Private Sub cmdCommonAdditiveReport_Click()
' 6/16/2005 RTD
' CR#1921 VERSION 7.4.0
    Dim fPreviewWindow As New frmReportPreview
    Dim sOpen As String
    
    sOpen = "bldg_desc = """ & Me.txtbldg_desc.Text & """" & vbCrLf & _
                "bldg_stories = """ & Me.txtbldg_stories.Text & """" & vbCrLf & _
                "bldg_stories_hgt = """ & Me.txtbldg_stories_hgt.Text & """" & vbCrLf & _
                "bldg_area = """ & Me.txtArea(Right(sshpSelectedArea, 1)).Text & """"
    fPreviewWindow.ReportName = "Common Additives"
    fPreviewWindow.ReportFile = "rptSummaryEstimate.xml"
    fPreviewWindow.OpenEvent = sOpen
    fPreviewWindow.RecordSet = m_recComAdds
    fPreviewWindow.RenderReport
    fPreviewWindow.Show
    
End Sub

Private Sub Form_Activate()
    ShowPrintToolbar True
End Sub

Private Sub Form_Deactivate()
    ShowPrintToolbar False
End Sub

Private Sub Form_Initialize()

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status ("Loading Building Maintenance ...")
    sEventSubscriberID = EventSubscriberAdd(Me)
    m_blnInsert = False
    sshpSelectedArea = "0,0"
    With cnTemp
        .ConnectionTimeout = 0
        .CommandTimeout = 0
        '.Open "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
        .Open strConnect
    End With
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    '
    '   Place form so that it "hides" the grid
    '   that launched it, so buttons on the grid
    '   don't appear to be on this form.
    Move 150, 200
    Me.Height = 7000
    Me.Width = 11500
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button              As String
    Dim blnPendingChange    As Boolean
    Dim bln_New             As Boolean
    '
    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        '
        '   Can't update Quality Series building variables.
        If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
        Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
        Else
            If bIsPendingChange = True Or m_objComAddsGridMap.IsPendingChange Then
                Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Close Building Form")
                If Button = vbYes Then
                    '
                    '   If there were errors, cancel the close.
                    If Not Update Then
                        Cancel = True
                    End If
                ElseIf Button = vbCancel Then
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ShowPrintToolbar False
    EventSubscriberRemove sEventSubscriberID
    m_rec.Close
    m_recComAdds.Close
    m_recModels.Close
    m_recModelMatrix.Close
    Set m_rec = Nothing
    Set m_recComAdds = Nothing
    Set m_recModels = Nothing
    Set m_recModelMatrix = Nothing
    cnTemp.Close
    Set cnTemp = Nothing
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        Me.Height = 7000
        Me.Width = 11500
    End If
End Sub

Public Sub EventNotify(eNotifyType As EEventSubscriberNotifyType, sAffectedRecordIdentifier As String)
   
    On Error Resume Next
    '
    '   If the record that was updated is for our bldg
    '   we need to refresh.
    If eNotifyType = esnModelRecordUpdated And _
        Trim(txtbldg_id.Text) = Trim(sAffectedRecordIdentifier) Then
        
        SearchForNewBldg txtbldg_id.Text, Replace(Trim(txtbldg_desc.Text), "'", "''")
    End If
End Sub
'
'   This routine is always called to load the form.
Public Function SetRow(sbldg_id As String, Optional blnInsert As Boolean = False, Optional new_rec As ADODB.RecordSet) As Boolean
    Dim strSelect As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    SetRow = True
    bIsInitialLoad = True

    If new_rec Is Nothing Then
        strSelect = "exec sp_select_building @type_code = '%', @bldg_category = '%', "
        strSelect = strSelect & "@bldg_id = '" & sbldg_id & "', @bldg_desc = '%'"
        
        If Not g_objDAL.GetRecordset(vbNullString, strSelect, m_rec) Then
            Screen.MousePointer = vbNormal
            MsgBox "Error attempting to locate building.", vbCritical
            SetRow = False
            Exit Function
        End If
    Else
        Set m_rec = new_rec
    End If
    
    m_blnInsert = blnInsert
    '
    '   If we are inserting/cloning.
    If m_blnInsert Then
        '
        '   Do this so OriginalValue will be set to
        '   the values copied into the row.
        m_rec.UpdateBatch
        
        If Trim(m_rec.Fields("bldg_skey").Value) <> 0 Then
            m_blnClone = True
            '
            '   Initialize grids if this is a clone or existing bldg.
            '   Do this here because if we are on a residential model then we have
            '   to enable the total cost columns in the common adds grid.
            With m_objComAddsGridMap
                .SetGrid TDBGridAdds
                .InitGrid (m_rec.Fields("type_code").Value = "R")
            End With
        Else
            m_blnClone = False
        End If
    Else
        '
        '   Initialize grids if this is a clone or existing bldg.
        '   Do this here because if we are on a residential model then we have
        '   to enable the total cost columns in the common adds grid.
        With m_objComAddsGridMap
            .SetGrid TDBGridAdds
            .InitGrid (m_rec.Fields("type_code").Value = "R")
        End With
    End If
    '
    '   If we are on a Residential building default the
    '   costs to 'OPN'
    If m_rec.Fields("type_code").Value = "R" Then
        optOpen.Value = True
    End If
    
    PopulateScreen
    EnableControls
    tabBldgAdditions.Tab = 0
    bIsInitialLoad = False
    bIsPendingChange = False
    bRefreshCosts = False
    Status ("")
    Screen.MousePointer = vbNormal
End Function

Private Sub PopulateScreen()
    On Error Resume Next
    
    PopulateBldgCategoriesDescriptions
    PopulateResiBldgTypes
    If m_blnInsert And Not m_blnClone Then
        PopulateAvailWallTypesFrameTypes "C"
        TDBGridAdds.ApproxCount = 0
        Me.Caption = "Building Maintenance [New Building]"
    Else
        PopulateBldgDetails
        PopulateModelMatrix
        PopulateComAdds
        Me.Caption = "Building Maintenance [" & Trim(m_rec.Fields("bldg_id").Value) & " | " & Trim(m_rec.Fields("bldg_desc").Value) & "]"
    End If
    
End Sub

Private Sub PopulateBldgCategoriesDescriptions()
    Dim recCategory     As ADODB.RecordSet
    Dim strSelect       As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   Fill the available categories based on the type code.
    cbobldg_categoryC.Clear
    cbobldg_categoryR.Clear
    
    strSelect = "SELECT type_code, bldg_category FROM bldg_category WHERE type_code = 'C' OR type_code = 'R' ORDER BY type_code, bldg_category"
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, recCategory) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching to populate available building categories.", vbCritical
    Else
        With recCategory
            While Not .EOF
                '
                '   Now populate the cboCategories that allow them to update a
                '   bldg's category.
                If Trim(.Fields("type_code").Value) = "C" Then
                    cbobldg_categoryC.AddItem Trim(.Fields("bldg_category").Value)
                Else
                    cbobldg_categoryR.AddItem Trim(.Fields("bldg_category").Value)
                End If
                .MoveNext
            Wend
            .Close
        End With
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateResiBldgTypes()
    
    On Error Resume Next
    With cboResiBldgType
        .Clear
        .AddItem "A  - 1 Story"
        .AddItem "B  - 1 -1/2 Story"
        .AddItem "C  - 2 Story"
        .AddItem "D  - 2 -1/2 Story"
        .AddItem "E  - 3 Story"
        .AddItem "F  - Bi-Level"
        .AddItem "G  - Tri-Level"
        .AddItem "H  - 1 Story Wings & Ells"
        .AddItem "I  - 1 -1/2 Story Wings & Ells"
        .AddItem "J  - 2 Story Wings & Ells"
    End With

End Sub

Private Sub PopulateAvailWallTypesFrameTypes(sTypeCode As String)
    Dim recModels       As ADODB.RecordSet
    Dim strSelect       As String
    Dim i               As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    For i = 0 To 5
        cboWallType(i).Clear
        cboFrameType(i).Clear
    Next i
    
    If sTypeCode = "R" Then
        For i = 0 To 5
            cboFrameType(i).Visible = False
            cboWallType(i).Width = 6280
        Next i
    Else
        For i = 0 To 5
            cboFrameType(i).Visible = True
            cboWallType(i).Width = 4455
        Next i
        strSelect = "SELECT DISTINCT frame_type FROM bldg_model WHERE model_code != '0' AND " _
                        & "model_code != '7' AND model_code != '8' AND frame_type != '' ORDER BY frame_type"
        '
        '   Use DAL to perform select.
        If Not g_objDAL.GetRecordset(vbNullString, strSelect, recModels) Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching to populate available frame types.", vbCritical
        Else
           With recModels
                Do Until .EOF
                    cboFrameType(0).AddItem Trim(.Fields("frame_type").Value)
                    cboFrameType(1).AddItem Trim(.Fields("frame_type").Value)
                    cboFrameType(2).AddItem Trim(.Fields("frame_type").Value)
                    cboFrameType(3).AddItem Trim(.Fields("frame_type").Value)
                    cboFrameType(4).AddItem Trim(.Fields("frame_type").Value)
                    cboFrameType(5).AddItem Trim(.Fields("frame_type").Value)
                    .MoveNext
                Loop
                .Close
            End With
        End If
    End If
    strSelect = "SELECT DISTINCT wall_type FROM bldg_model WHERE model_code != '0' AND " _
                    & "model_code != '7' AND model_code != '8' AND wall_type != '' ORDER BY wall_type"
    '
    '   Use DAL to perform select.
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, recModels) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching to populate available wall types.", vbCritical
    Else
       With recModels
            Do Until .EOF
                cboWallType(0).AddItem Trim(.Fields("wall_type").Value)
                cboWallType(1).AddItem Trim(.Fields("wall_type").Value)
                cboWallType(2).AddItem Trim(.Fields("wall_type").Value)
                cboWallType(3).AddItem Trim(.Fields("wall_type").Value)
                cboWallType(4).AddItem Trim(.Fields("wall_type").Value)
                cboWallType(5).AddItem Trim(.Fields("wall_type").Value)
                .MoveNext
            Loop
            .Close
        End With
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateBldgDetails()
    Dim i               As Integer
    Dim strBldgCostDesc As String
    
    On Error Resume Next
    '
    '   With the bldg record we're using populate the bldg details tab.
    With m_rec
        txtbldg_id.Text = Trim(.Fields("bldg_id").Value)
        '
        '   Keep a copy of the bldg_id in the tag so that if we're
        '   on a clone and the user changes it we can update everything.
        txtbldg_id.Tag = Trim(.Fields("bldg_id").Value)
        
        txtbldg_skey.Text = .Fields("bldg_skey").Value
        txtlast_update_date.Text = .Fields("last_update_date").Value
        txtlast_update_person.Text = .Fields("last_update_person").Value

        If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
        Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
            
            txtBldgCostDesc.Text = "Basement Models are used for all buildings within the Quality Series.  " & _
                "Only their Assemblies may be modified."
        Else
            strBldgCostDesc = Replace(BLDG_COST_DESC, "[BLDG_STORIES]", .Fields("bldg_stories").Value)
            strBldgCostDesc = Replace(strBldgCostDesc, "[STORIES_HGT]", .Fields("bldg_stories_hgt").Value)
            strBldgCostDesc = Replace(strBldgCostDesc, "[BLDG_AREA]", FormatNumber(.Fields("bldg_area_std").Value, 0))
            txtBldgCostDesc.Text = strBldgCostDesc
        End If
       
        If Not IsNull(.Fields("bldg_desc").Value) Then
            txtbldg_desc.Text = Trim(.Fields("bldg_desc").Value)
        Else
            txtbldg_desc.Text = ""
        End If
        
        If Not IsNull(.Fields("bldg_stories").Value) Then
            txtbldg_stories.Text = Trim(.Fields("bldg_stories").Value)
        Else
            txtbldg_stories.Text = ""
        End If
        
        If Not IsNull(.Fields("bldg_stories_hgt").Value) Then
            txtbldg_stories_hgt.Text = Trim(.Fields("bldg_stories_hgt").Value)
        Else
            txtbldg_stories_hgt.Text = ""
        End If

        If Not IsNull(.Fields("bldg_part_density").Value) Then
            txtbldg_part_density.Text = Trim(.Fields("bldg_part_density").Value)
        Else
            txtbldg_part_density.Text = ""
        End If
      
        If Not IsNull(.Fields("bldg_part_hgt").Value) Then
            txtbldg_part_hgt.Text = Trim(.Fields("bldg_part_hgt").Value)
        Else
            txtbldg_part_hgt.Text = ""
        End If
        
        If Not IsNull(.Fields("bldg_wall_factor").Value) Then
            txtbldg_wall_factor.Text = Trim(.Fields("bldg_wall_factor").Value)
        Else
            txtbldg_wall_factor.Text = ""
        End If
    
        If Not IsNull(.Fields("bldg_fixture_area").Value) Then
            txtbldg_fixture_area.Text = Trim(.Fields("bldg_fixture_area").Value)
        Else
            txtbldg_fixture_area.Text = ""
        End If
        
        If Not IsNull(.Fields("window_area").Value) Then
            txtwindow_area.Text = Trim(.Fields("window_area").Value)
        Else
            txtwindow_area.Text = ""
        End If
        
        If Not IsNull(.Fields("bldg_elev_no").Value) Then
            txtbldg_elev_no.Text = Trim(.Fields("bldg_elev_no").Value)
        Else
            txtbldg_elev_no.Text = ""
        End If
       
        If Not IsNull(.Fields("bldg_door_density").Value) Then
            txtbldg_door_density.Text = Trim(.Fields("bldg_door_density").Value)
        Else
            txtbldg_door_density.Text = ""
        End If
   
        If Not IsNull(.Fields("op_factor").Value) Then
            txtop_factor.Text = Trim(.Fields("op_factor").Value)
        Else
            txtop_factor.Text = ""
        End If
  
        If Not IsNull(.Fields("architect_fee").Value) Then
            txtarchitect_fee.Text = Trim(.Fields("architect_fee").Value)
        Else
            txtarchitect_fee.Text = ""
        End If
   
        If Not IsNull(.Fields("graphic_ref_id").Value) Then
            txtgraphic_ref_id.Text = Trim(.Fields("graphic_ref_id").Value)
        Else
            txtgraphic_ref_id.Text = ""
        End If
       
        If Not IsNull(.Fields("graphic_ref_id2").Value) Then
            txtgraphic_ref_id2.Text = Trim(.Fields("graphic_ref_id2").Value)
        Else
            txtgraphic_ref_id2.Text = ""
        End If
        If .Fields("type_code") = "C" Or IsNull(.Fields("type_code")) Then
            opttype_codeC.Value = True
            
            For i = 0 To cbobldg_categoryC.listcount - 1
                If cbobldg_categoryC.List(i) = Trim(.Fields("bldg_category")) Then
                    cbobldg_categoryC.ListIndex = i
                    Exit For
                End If
            Next i
        ElseIf .Fields("type_code") = "R" Then
            opttype_codeR.Value = True
            
            For i = 0 To cbobldg_categoryR.listcount - 1
                If cbobldg_categoryR.List(i) = Trim(.Fields("bldg_category")) Then
                    cbobldg_categoryR.ListIndex = i
                    Exit For
                End If
            Next i
            For i = 0 To cboResiBldgType.listcount - 1
                If Left$(cboResiBldgType.List(i), 1) = Trim(.Fields("bldg_type")) Then
                    cboResiBldgType.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End With
End Sub

Private Sub RePopulateColToBold()
    Dim sCurSelectedCol As String
    Dim i               As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    sCurSelectedCol = cboColumnToBold.ListIndex
    cboColumnToBold.Clear
    '
    '   If we're on a blank bldg use the fraNewBldgModelMatrix.
    If m_blnInsert And m_blnClone = False Then
        '
        '   Populate the cboColToBold with the possible area/perm.
        If opttype_codeC.Value = True Then
            For i = 0 To 8
                If Trim(txtNewBldgArea(i).Text) <> "" And Trim(txtNewBldgPerimeter(i).Text) <> "" Then
                    cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                        txtNewBldgArea(i).Text & " | " & txtNewBldgPerimeter(i).Text
                End If
            Next i
        Else
            If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
                
                For i = 0 To 7
                    If Trim(txtNewBldgArea(i).Text) <> "" And Trim(txtNewBldgPerimeter(i).Text) <> "" Then
                        cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                            txtNewBldgArea(i).Text & " | " & txtNewBldgPerimeter(i).Text
                    End If
                Next i
            Else
                For i = 0 To 10
                    If Trim(txtNewBldgArea(i).Text) <> "" And Trim(txtNewBldgPerimeter(i).Text) <> "" Then
                        cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                            txtNewBldgArea(i).Text & " | " & txtNewBldgPerimeter(i).Text
                    End If
                Next i
            End If
        End If
    ElseIf opttype_codeC.Value = True Then
        '
        '   Populate the cboColToBold with the possible area/perm for Commercial.
        For i = 0 To 8
            If Trim(txtArea(i).Text) <> "" And Trim(txtPerimeter(i).Text) <> "" Then
                cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                    txtArea(i).Text & " | " & txtPerimeter(i).Text
            End If
        Next i
    Else
        If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
            Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
            Left$(Trim(cboResiBldgType.Text), 1) = "J" Then

            For i = 0 To 7
                If Trim(txtAreaResi(i).Text) <> "" And Trim(txtPerimeterResi(i).Text) <> "" Then
                    cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                        txtAreaResi(i).Text & " | " & txtPerimeterResi(i).Text
                End If
            Next i
        Else
            '
            '   Populate the cboColToBold with the possible area/perm for Residential.
            For i = 0 To 10
                If Trim(txtAreaResi(i).Text) <> "" And Trim(txtPerimeterResi(i).Text) <> "" Then
                    cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                        txtAreaResi(i).Text & " | " & txtPerimeterResi(i).Text
                End If
            Next i
        End If
    End If
    cboColumnToBold.ListIndex = sCurSelectedCol
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateNewBldgRowToBold()
    Dim i               As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    cboRowToBold.Clear
    '
    '   Populate the cboRowToBold with the possible models. Exclude format rows.
    For i = 0 To 5
        If opttype_codeC.Value = True Then
            If Trim(cboWallType(i).Text) <> "" Then
                If Trim(cboFrameType(i).Text) <> "" Then
                    cboRowToBold.AddItem "[" & i + 1 & "] " & _
                        Trim(cboWallType(i).Text) & " | " & Trim(cboFrameType(i).Text)
                End If
            End If
        Else
            If Trim(cboWallType(i).Text) <> "" Then
                cboRowToBold.AddItem "[" & i + 1 & "] " & _
                    Trim(cboWallType(i).Text)
            End If
        End If
    Next i
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateModelMatrix()
    Dim strSelect       As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   Make sure it is closed.
    With m_recModelMatrix
        .Close
        '
        '   Set the maximum number to bring back.
        .MaxRecords = MAX_RECORDS
    End With
    '
    '   If we're inserting a blank
    If m_blnInsert And Trim(m_rec.Fields("bldg_id").Value) = "" Then
    Else
        '
        '   IF we are editing a Residential Quality Series building
        '   we have to lock down fields and since there aren't any rows in
        '   published_bldg_matrix_cost for Quality models we populate the matrix differently.
        If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
        Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
            strSelect = "exec sp_select_building_pub_matrix_cost_basements @bldg_id = '"
            strSelect = strSelect & Trim(m_rec.Fields("bldg_id").Value)
            strSelect = strSelect & "', @bldg_model_skey = '%'"
        Else
            strSelect = "exec sp_select_building_pub_matrix_cost @bldg_id = '"
        
            If Len(Trim(m_rec.Fields("bldg_id").Value)) > 0 Then
                strSelect = strSelect & Trim(m_rec.Fields("bldg_id").Value)
            Else
                strSelect = strSelect & "%"
            End If
            
            strSelect = strSelect & "', @bldg_model_skey = '%"
            
            strSelect = strSelect & "', @bldg_desc = '"
            If Len(Trim(txtbldg_desc.Text)) > 0 Then
               '
               '   We never know if we might have apos ' in our
               '   desc so replace for query.
               strSelect = strSelect & Replace(Trim(txtbldg_desc.Text), "'", "''") & "'"
            Else
                strSelect = strSelect & "%'"
            End If
            
            strSelect = strSelect & ", @op_code = '"
            If optUnion.Value = True Then
                strSelect = strSelect & "STD'"
            Else
                strSelect = strSelect & "OPN'"
            End If
                   
            strSelect = strSelect & ", @country_code = '"
            If Len(Trim(cboMdlCountryCode.Text)) > 0 Then
                strSelect = strSelect & cboMdlCountryCode.Text & "'"
            Else
                strSelect = strSelect & "USA'"
            End If
            
            strSelect = strSelect & ", @region_code = '"
            If Len(Trim(cboMdlRegionCode.Text)) > 0 Then
                strSelect = strSelect & cboMdlRegionCode.Text & "'"
            Else
                strSelect = strSelect & "NAT'"
            End If
            
            If opttype_codeC.Value = True Then
                strSelect = strSelect & ", @type_code = 'C'"
            Else
                strSelect = strSelect & ", @type_code = 'R'"
            End If
        End If
        With cnTemp
            .ConnectionTimeout = 0
            .CommandTimeout = 0
            '.Open "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
            .Open strConnect
            
            m_recModelMatrix.CursorLocation = adUseClient
            m_recModelMatrix.Open _
                Source:=strSelect, _
                ActiveConnection:=cnTemp, _
                CursorType:=adOpenStatic, _
                LockType:=adLockBatchOptimistic
    
            If cnTemp.Errors.Count <> 0 Then
                Screen.MousePointer = vbNormal
                MsgBox "Errors in the PopulateModelMatrix routine: " _
                    & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            Else
                cboColumnToBold.Clear
                cboRowToBold.Clear
                '
                '   Populate different frames based upon Commercial or Residential
                If opttype_codeC.Value = True Then
                    PopulateModelMatrixCommercial
                Else
                    PopulateModelMatrixResi
                End If
            End If
        End With
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateModelMatrixCommercial()
    Dim i           As Integer
    
    On Error Resume Next
    With m_recModelMatrix
        If Not .EOF Then
            .MoveFirst
            '
            '   Go through the 1st record to get the
            '   Area & Perimeter combinations. format = 95000|1
            txtArea(0).Text = Trim(.Fields("area1").Value)
            txtArea(0).Tag = txtArea(0).Text & "|" & Trim(.Fields("last_update_id_1").Value)
            txtPerimeter(0).Text = Trim(.Fields("perimeter1").Value)
            txtPerimeter(0).Tag = txtPerimeter(0).Text
            
            txtArea(1).Text = Trim(.Fields("area2").Value)
            txtArea(1).Tag = txtArea(1).Text & "|" & Trim(.Fields("last_update_id_2").Value)
            txtPerimeter(1).Text = Trim(.Fields("perimeter2").Value)
            txtPerimeter(1).Tag = txtPerimeter(1).Text
            
            txtArea(2).Text = Trim(.Fields("area3").Value)
            txtArea(2).Tag = txtArea(2).Text & "|" & Trim(.Fields("last_update_id_3").Value)
            txtPerimeter(2).Text = Trim(.Fields("perimeter3").Value)
            txtPerimeter(2).Tag = txtPerimeter(2).Text
            
            txtArea(3).Text = Trim(.Fields("area4").Value)
            txtArea(3).Tag = txtArea(3).Text & "|" & Trim(.Fields("last_update_id_4").Value)
            txtPerimeter(3).Text = Trim(.Fields("perimeter4").Value)
            txtPerimeter(3).Tag = txtPerimeter(3).Text
            
            txtArea(4).Text = Trim(.Fields("area5").Value)
            txtArea(4).Tag = txtArea(4).Text & "|" & Trim(.Fields("last_update_id_5").Value)
            txtPerimeter(4).Text = Trim(.Fields("perimeter5").Value)
            txtPerimeter(4).Tag = txtPerimeter(4).Text

            txtArea(5).Text = Trim(.Fields("area6").Value)
            txtArea(5).Tag = txtArea(5).Text & "|" & Trim(.Fields("last_update_id_6").Value)
            txtPerimeter(5).Text = Trim(.Fields("perimeter6").Value)
            txtPerimeter(5).Tag = txtPerimeter(5).Text

            txtArea(6).Text = Trim(.Fields("area7").Value)
            txtArea(6).Tag = txtArea(6).Text & "|" & Trim(.Fields("last_update_id_7").Value)
            txtPerimeter(6).Text = Trim(.Fields("perimeter7").Value)
            txtPerimeter(6).Tag = txtPerimeter(6).Text

            txtArea(7).Text = Trim(.Fields("area8").Value)
            txtArea(7).Tag = txtArea(7).Text & "|" & Trim(.Fields("last_update_id_8").Value)
            txtPerimeter(7).Text = Trim(.Fields("perimeter8").Value)
            txtPerimeter(7).Tag = txtPerimeter(7).Text

            txtArea(8).Text = Trim(.Fields("area9").Value)
            txtArea(8).Tag = txtArea(8).Text & "|" & Trim(.Fields("last_update_id_9").Value)
            txtPerimeter(8).Text = Trim(.Fields("perimeter9").Value)
            txtPerimeter(8).Tag = txtPerimeter(8).Text
            '
            '   Populate the cboColToBold with the possible area/perm.
            For i = 0 To 8
                cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                    txtArea(i).Text & " | " & txtPerimeter(i).Text
                If Trim(.Fields("AreaInd").Value) = i + 1 Then
                    cboColumnToBold.ListIndex = i
                End If
            Next i
                        
            PopulateComboCountryRegion Trim(.Fields("country_code").Value), _
                            Trim(.Fields("region_code").Value)
            '
            '   The AreaInd is the Area column to bold ie-(Area5) and the model code
            '   indicates the model and it's row.
            ChangeOpCostBackcolor Trim(.Fields("AreaInd").Value), .Fields("model_code_to_bold").Value
            '
            '   Now for each model returned populate the corresponding costs.
            Do Until .EOF
                lblWall(.Bookmark - 1).Caption = Trim(.Fields("full_desc_wall_type").Value)
                lblFrame(.Bookmark - 1).Caption = Trim(.Fields("frame_type").Value)
                '
                '   Keep the model skey so we can run a summary rpt if needed.
                '   Also keep the format code so we know whether or not to split
                '   desc across 2 lines/rows.
                lblWall(.Bookmark - 1).Tag = Trim(.Fields("format_code").Value) & "|" & _
                                    Trim(.Fields("bldg_model_skey").Value)
                '
                '   In case our text is huge, so all in tooltip.
                lblWall(.Bookmark - 1).ToolTipText = lblWall(.Bookmark - 1).Caption
                lblFrame(.Bookmark - 1).ToolTipText = lblFrame(.Bookmark - 1).Caption
                '
                '   Populate the cboRowToBold with the possible models. Exclude format rows.
                '   Format of [model_code] wall | frame
                If .Fields("model_code").Value <> "7" And .Fields("model_code").Value <> "8" Then
                    cboRowToBold.AddItem "[" & .Fields("model_code").Value & "] " & _
                        lblWall(.Bookmark - 1).Caption & " | " & lblFrame(.Bookmark - 1).Caption
                End If
                
                If .Fields("model_code_to_bold").Value = .Bookmark Then
                    cboRowToBold.ListIndex = .Bookmark - 1
                End If
                
                lblCol1_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col1_total_cost_op").Value), 2)
                lblCol2_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col2_total_cost_op").Value), 2)
                lblCol3_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col3_total_cost_op").Value), 2)
                lblCol4_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col4_total_cost_op").Value), 2)
                lblCol5_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col5_total_cost_op").Value), 2)
                lblCol6_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col6_total_cost_op").Value), 2)
                lblCol7_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col7_total_cost_op").Value), 2)
                lblCol8_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col8_total_cost_op").Value), 2)
                lblCol9_TotalOP(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col9_total_cost_op").Value), 2)
                '
                '   Compute overall costs based on SF Area.
                lblCol1_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol1_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(0).Text)), 2)
                lblCol2_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol2_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(1).Text)), 2)
                lblCol3_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol3_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(2).Text)), 2)
                lblCol4_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol4_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(3).Text)), 2)
                lblCol5_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol5_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(4).Text)), 2)
                lblCol6_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol6_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(5).Text)), 2)
                lblCol7_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol7_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(6).Text)), 2)
                lblCol8_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol8_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(7).Text)), 2)
                lblCol9_TotalOP(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((lblCol9_TotalOP(.Bookmark - 1).Caption * Trim(txtArea(8).Text)), 2)

                .MoveNext
            Loop
            '
            '   Cleanup the text for the wall_type combining lables that share the same type.
            CleanWallLabels .RecordCount
        End If
    End With
End Sub

Private Sub PopulateModelMatrixResi()
    Dim i           As Integer
    
    On Error Resume Next
    With m_recModelMatrix
        If Not .EOF Then
            .MoveFirst
            '
            '   Go through the 1st record to get the
            '   Area & Perimeter combinations. format = 95000|1
            txtAreaResi(0).Text = Trim(.Fields("area1").Value)
            txtAreaResi(0).Tag = txtAreaResi(0).Text & "|" & Trim(.Fields("last_update_id_1").Value)
            txtPerimeterResi(0).Text = Trim(.Fields("perimeter1").Value)
            txtPerimeterResi(0).Tag = txtPerimeterResi(0).Text
            
            txtAreaResi(1).Text = Trim(.Fields("area2").Value)
            txtAreaResi(1).Tag = txtAreaResi(1).Text & "|" & Trim(.Fields("last_update_id_2").Value)
            txtPerimeterResi(1).Text = Trim(.Fields("perimeter2").Value)
            txtPerimeterResi(1).Tag = txtPerimeterResi(1).Text
            
            txtAreaResi(2).Text = Trim(.Fields("area3").Value)
            txtAreaResi(2).Tag = txtAreaResi(2).Text & "|" & Trim(.Fields("last_update_id_3").Value)
            txtPerimeterResi(2).Text = Trim(.Fields("perimeter3").Value)
            txtPerimeterResi(2).Tag = txtPerimeterResi(2).Text
            
            txtAreaResi(3).Text = Trim(.Fields("area4").Value)
            txtAreaResi(3).Tag = txtAreaResi(3).Text & "|" & Trim(.Fields("last_update_id_4").Value)
            txtPerimeterResi(3).Text = Trim(.Fields("perimeter4").Value)
            txtPerimeterResi(3).Tag = txtPerimeterResi(3).Text
            
            txtAreaResi(4).Text = Trim(.Fields("area5").Value)
            txtAreaResi(4).Tag = txtAreaResi(4).Text & "|" & Trim(.Fields("last_update_id_5").Value)
            txtPerimeterResi(4).Text = Trim(.Fields("perimeter5").Value)
            txtPerimeterResi(4).Tag = txtPerimeterResi(4).Text

            txtAreaResi(5).Text = Trim(.Fields("area6").Value)
            txtAreaResi(5).Tag = txtAreaResi(5).Text & "|" & Trim(.Fields("last_update_id_6").Value)
            txtPerimeterResi(5).Text = Trim(.Fields("perimeter6").Value)
            txtPerimeterResi(5).Tag = txtPerimeterResi(5).Text

            txtAreaResi(6).Text = Trim(.Fields("area7").Value)
            txtAreaResi(6).Tag = txtAreaResi(6).Text & "|" & Trim(.Fields("last_update_id_7").Value)
            txtPerimeterResi(6).Text = Trim(.Fields("perimeter7").Value)
            txtPerimeterResi(6).Tag = txtPerimeterResi(6).Text

            txtAreaResi(7).Text = Trim(.Fields("area8").Value)
            txtAreaResi(7).Tag = txtAreaResi(7).Text & "|" & Trim(.Fields("last_update_id_8").Value)
            txtPerimeterResi(7).Text = Trim(.Fields("perimeter8").Value)
            txtPerimeterResi(7).Tag = txtPerimeterResi(7).Text

            txtAreaResi(8).Text = Trim(.Fields("area9").Value)
            txtAreaResi(8).Tag = txtAreaResi(8).Text & "|" & Trim(.Fields("last_update_id_9").Value)
            txtPerimeterResi(8).Text = Trim(.Fields("perimeter9").Value)
            txtPerimeterResi(8).Tag = txtPerimeterResi(8).Text
            
            txtAreaResi(9).Text = Trim(.Fields("area10").Value)
            txtAreaResi(9).Tag = txtAreaResi(9).Text & "|" & Trim(.Fields("last_update_id_10").Value)
            txtPerimeterResi(9).Text = Trim(.Fields("perimeter10").Value)
            txtPerimeterResi(9).Tag = txtPerimeterResi(9).Text
            
            txtAreaResi(10).Text = Trim(.Fields("area11").Value)
            txtAreaResi(10).Tag = txtAreaResi(10).Text & "|" & Trim(.Fields("last_update_id_11").Value)
            txtPerimeterResi(10).Text = Trim(.Fields("perimeter11").Value)
            txtPerimeterResi(10).Tag = txtPerimeterResi(10).Text
            '
            '   Populate the cboColToBold with the possible area/perm.
            If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
            
                For i = 0 To 7
                    cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                        txtAreaResi(i).Text & " | " & txtPerimeterResi(i).Text
                    If Trim(.Fields("AreaInd").Value) = i + 1 Then
                        cboColumnToBold.ListIndex = i
                    End If
                Next i
            Else
                For i = 0 To 10
                    cboColumnToBold.AddItem "[" & i + 1 & "] " & _
                        txtAreaResi(i).Text & " | " & txtPerimeterResi(i).Text
                    If Trim(.Fields("AreaInd").Value) = i + 1 Then
                        cboColumnToBold.ListIndex = i
                    End If
                Next i
            End If
            
            PopulateComboCountryRegion Trim(.Fields("country_code").Value), _
                            Trim(.Fields("region_code").Value)
            '
            '   The AreaInd is the Area column to bold ie-(Area5) and the model code
            '   indicates the model and it's row.
            ChangeOpCostBackcolorResi Trim(.Fields("AreaInd").Value), .Fields("model_code_to_bold").Value
            '
            '   Now for each model returned populate the corresponding costs.
            Do Until .EOF
                lblWallResi(.Bookmark - 1).Caption = Trim(.Fields("wall_type").Value)
                '
                '   Keep the model skey so we can run a summary rpt if needed.
                '   Also keep the format code so we know whether or not to split
                '   desc across 2 lines/rows.
                lblWallResi(.Bookmark - 1).Tag = Trim(.Fields("format_code").Value) & "|" & _
                                    Trim(.Fields("bldg_model_skey").Value)
                '
                '   In case our text is huge, so all in tooltip.
                lblWallResi(.Bookmark - 1).ToolTipText = lblWallResi(.Bookmark - 1).Caption
                '
                '   Populate the cboRowToBold with the possible models. Exclude format rows.
                If .Fields("model_code").Value <> "7" And .Fields("model_code").Value <> "8" Then
                    cboRowToBold.AddItem "[" & .Fields("model_code").Value & "] " & _
                        lblWallResi(.Bookmark - 1).Caption
                End If
                
                If .Fields("model_code_to_bold").Value = .Bookmark Then
                    cboRowToBold.ListIndex = .Bookmark - 1
                End If
                
                lblCol1_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col1_total_cost_op").Value), 2)
                '
                '   Compute overall costs based on SF Area.
                lblCol1_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col1_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
    
                lblCol2_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col2_total_cost_op").Value), 2)
                lblCol2_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col2_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                lblCol3_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col3_total_cost_op").Value), 2)
                lblCol3_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col3_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                lblCol4_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col4_total_cost_op").Value), 2)
                lblCol4_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col4_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                lblCol5_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col5_total_cost_op").Value), 2)
                lblCol5_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col5_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                lblCol6_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col6_total_cost_op").Value), 2)
                lblCol6_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col6_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                lblCol7_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col7_total_cost_op").Value), 2)
                lblCol7_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col7_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                lblCol8_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col8_total_cost_op").Value), 2)
                lblCol8_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col8_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)

                lblCol9_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col9_total_cost_op").Value), 2)
                lblCol9_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col9_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                lblCol10_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col10_total_cost_op").Value), 2)
                lblCol10_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col10_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                lblCol11_TotalOPResi(.Bookmark - 1).Caption = FormatNumber(Trim(.Fields("col11_total_cost_op").Value), 2)
                lblCol11_TotalOPResi(.Bookmark - 1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col11_total_cost_op").Value) * Trim(txtAreaResi(.Bookmark - 1).Text)), 2)
                
                .MoveNext
            Loop
            '
            '   Cleanup the text for the wall_type combining lables that share the same type.
            CleanWallLabelsResi .RecordCount
        End If
    End With
End Sub

Private Sub PopulateComboCountryRegion(sCountryCodeValueToSelect As String, sRegionCodeValueToSelect As String)
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To cboMdlCountryCode.listcount - 1
        If cboMdlCountryCode.List(i) = sCountryCodeValueToSelect Then
            cboMdlCountryCode.ListIndex = i
            Exit For
        End If
    Next i
    
    For i = 0 To cboMdlRegionCode.listcount - 1
        If cboMdlRegionCode.List(i) = sRegionCodeValueToSelect Then
            cboMdlRegionCode.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub ChangeOpCostBackcolor(nAreaInd As Integer, nModelCode As Integer)
    Dim ctrl As Control
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    For Each ctrl In Me.Controls
        '
        '   If backcolor is teal = &HFFFF00
        If ctrl.BackColor = &HFFFF00 Then
            If TypeOf ctrl Is TextBox Or TypeOf ctrl Is Label Then
                ctrl.BackColor = &H80000005 'white
            End If
        End If
    Next ctrl
    '
    '   Set the backcolor for the default model area/perimeter, then execute
    '   the label click to get the fraModelMatrix populated with the costs.
    Select Case nAreaInd
        Case 1
            lblCol1_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol1_TotalOP_Click (nModelCode - 1)
        Case 2
            lblCol2_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol2_TotalOP_Click (nModelCode - 1)
        Case 3
            lblCol3_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol3_TotalOP_Click (nModelCode - 1)
        Case 4
            lblCol4_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol4_TotalOP_Click (nModelCode - 1)
        Case 5
            lblCol5_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol5_TotalOP_Click (nModelCode - 1)
        Case 6
            lblCol6_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol6_TotalOP_Click (nModelCode - 1)
        Case 7
            lblCol7_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol7_TotalOP_Click (nModelCode - 1)
        Case 8
            lblCol8_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol8_TotalOP_Click (nModelCode - 1)
        Case 9
            lblCol9_TotalOP(nModelCode - 1).BackColor = &HFFFF00
            lblCol9_TotalOP_Click (nModelCode - 1)
    End Select
    Screen.MousePointer = vbNormal
End Sub

Private Sub ChangeOpCostBackcolorResi(nAreaInd As Integer, nModelCode As Integer)
    Dim ctrl As Control
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    For Each ctrl In Me.Controls
        '
        '   If backcolor is teal = &HFFFF00
        If ctrl.BackColor = &HFFFF00 Then
            If TypeOf ctrl Is TextBox Or TypeOf ctrl Is Label Then
                ctrl.BackColor = &H80000005 'white
            End If
        End If
    Next ctrl
    '
    '   Set the backcolor for the default model area/perimeter, then execute
    '   the label click to get the fraModelMatrixResi populated with the costs.
    Select Case nAreaInd
        Case 1
            lblCol1_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol1_TotalOPResi_Click (nModelCode - 1)
        Case 2
            lblCol2_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol2_TotalOPResi_Click (nModelCode - 1)
        Case 3
            lblCol3_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol3_TotalOPResi_Click (nModelCode - 1)
        Case 4
            lblCol4_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol4_TotalOPResi_Click (nModelCode - 1)
        Case 5
            lblCol5_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol5_TotalOPResi_Click (nModelCode - 1)
        Case 6
            lblCol6_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol6_TotalOPResi_Click (nModelCode - 1)
        Case 7
            lblCol7_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol7_TotalOPResi_Click (nModelCode - 1)
        Case 8
            lblCol8_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol8_TotalOPResi_Click (nModelCode - 1)
        Case 9
            lblCol9_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol9_TotalOPResi_Click (nModelCode - 1)
        Case 10
            lblCol10_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol10_TotalOPResi_Click (nModelCode - 1)
        Case 11
            lblCol11_TotalOPResi(nModelCode - 1).BackColor = &HFFFF00
            lblCol11_TotalOPResi_Click (nModelCode - 1)
    End Select
    Screen.MousePointer = vbNormal
End Sub

Private Sub CleanWallLabels(nModelCount As Integer)
    Dim spriorfulldesc  As String
    Dim i               As Integer
    
    On Error Resume Next
    '
    '   Turn everything on in case this is not the initial load with
    '   default values but as a result of a bldg change from the cbo's.
    For i = 0 To 7
        lblWall(i).Height = 285
        lblWall(i).Width = 3240
        lblWall(i).Visible = True
        lblFrame(i).Visible = True
        lblCol1_TotalOP(i).Visible = True
        lblCol2_TotalOP(i).Visible = True
        lblCol3_TotalOP(i).Visible = True
        lblCol4_TotalOP(i).Visible = True
        lblCol5_TotalOP(i).Visible = True
        lblCol6_TotalOP(i).Visible = True
        lblCol7_TotalOP(i).Visible = True
        lblCol8_TotalOP(i).Visible = True
        lblCol9_TotalOP(i).Visible = True
    Next i
    
    spriorfulldesc = lblWall(0).Caption
    '
    '   Loop only as many times as we have model records.
    For i = 1 To nModelCount - 1
        '
        '   In the format of format_code|bldg_model_skey
        If spriorfulldesc = lblWall(i).Caption Then
            '
            '   Only split desc across two model lines if the A3 format_code
            '   is followed by an A4 format_code.
            If Left$(lblWall(i - 1).Tag, 2) = "A3" And Left$(lblWall(i).Tag, 2) = "A4" Then
                lblWall(i - 1).Height = 570
                lblWall(i).Visible = False
            End If
        End If
        spriorfulldesc = lblWall(i).Caption
    Next i
    '
    '   In case we only have 3 models, we need to blank out the
    '   rest of the labels.
    For i = nModelCount To 7
        lblWall(i).Visible = False
        lblFrame(i).Visible = False
        lblCol1_TotalOP(i).Visible = False
        lblCol2_TotalOP(i).Visible = False
        lblCol3_TotalOP(i).Visible = False
        lblCol4_TotalOP(i).Visible = False
        lblCol5_TotalOP(i).Visible = False
        lblCol6_TotalOP(i).Visible = False
        lblCol7_TotalOP(i).Visible = False
        lblCol8_TotalOP(i).Visible = False
        lblCol9_TotalOP(i).Visible = False
    Next i
End Sub

Private Sub CleanWallLabelsResi(nModelCount As Integer)
    Dim spriorfulldesc  As String
    Dim i               As Integer
    
    On Error Resume Next
    '
    '   Turn everything on in case this is not the initial load with
    '   default values but as a result of a bldg change from the cbo's.
    For i = 0 To 7
        lblWallResi(i).Height = 285
        lblWallResi(i).Visible = True
        lblCol1_TotalOPResi(i).Visible = True
        lblCol2_TotalOPResi(i).Visible = True
        lblCol3_TotalOPResi(i).Visible = True
        lblCol4_TotalOPResi(i).Visible = True
        lblCol5_TotalOPResi(i).Visible = True
        lblCol6_TotalOPResi(i).Visible = True
        lblCol7_TotalOPResi(i).Visible = True
        lblCol8_TotalOPResi(i).Visible = True
        lblCol9_TotalOPResi(i).Visible = True
        lblCol10_TotalOPResi(i).Visible = True
        lblCol11_TotalOPResi(i).Visible = True
    Next i

    spriorfulldesc = lblWallResi(0).Caption
    '
    '   Loop only as many times as we have model records.
    For i = 1 To nModelCount - 1
        '
        '   In the format of format_code|bldg_model_skey
        If spriorfulldesc = lblWallResi(i).Caption Then
            If Left$(lblWallResi(i - 1).Tag, 2) = "A3" Then
                lblWallResi(i - 1).Height = 570
                lblWallResi(i).Visible = False
            End If
        End If
        spriorfulldesc = lblWallResi(i).Caption
    Next i
    '
    '   In case we only have 3 models, we need to blank out the
    '   rest of the labels.
    For i = nModelCount To 10
        lblWallResi(i).Visible = False
        lblCol1_TotalOPResi(i).Visible = False
        lblCol2_TotalOPResi(i).Visible = False
        lblCol3_TotalOPResi(i).Visible = False
        lblCol4_TotalOPResi(i).Visible = False
        lblCol5_TotalOPResi(i).Visible = False
        lblCol6_TotalOPResi(i).Visible = False
        lblCol7_TotalOPResi(i).Visible = False
        lblCol8_TotalOPResi(i).Visible = False
        lblCol9_TotalOPResi(i).Visible = False
        lblCol10_TotalOPResi(i).Visible = False
        lblCol11_TotalOPResi(i).Visible = False
    Next i
End Sub

Private Sub PopulateComAdds()
    Dim strSelect       As String

    On Error Resume Next
    '
    '   Make sure it is closed.
    With m_recComAdds
        .Close
        '
        '   Set the maximum number to bring back.
        .MaxRecords = MAX_RECORDS
    End With
    '
    '   Was having a problem with the grid refreshing if
    '   the user had changed data and not save and switched to
    '   a different area it would not refresh.  So have to close
    '   and reopen.
    With TDBGridAdds
        .Close
        .ReOpen
    End With

    strSelect = "exec sp_select_bldg_com_add_cst @bldg_skey = '"
    If Len(Trim(txtbldg_skey.Text)) > 0 Then
        strSelect = strSelect & Trim(txtbldg_skey.Text) & "'"
    Else
        'abort can't pull back for all buildings!
        Exit Sub
    End If
    
    With cnTemp
        m_recComAdds.CursorLocation = adUseClient
        m_recComAdds.Open _
            Source:=strSelect, _
            ActiveConnection:=cnTemp, _
            CursorType:=adOpenStatic, _
            LockType:=adLockBatchOptimistic

        If cnTemp.Errors.Count <> 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching to populate Common Adds." _
                    & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            lblComAddsRowCount.Caption = "0 rows returned."
        Else
            '
            '   Pass recordset to handler class.
            m_objComAddsGridMap.RecordSet = m_recComAdds
            '
            '   Need to make sure that the user cannot set
            '   max_records = 0
            With m_recComAdds
                .MoveFirst
            
                If .RecordCount > 0 Then
                    lblComAddsRowCount.Caption = .RecordCount & " rows returned."
                    '
                    ' If the upper bound was hit, inform user.
                    If .RecordCount = MAX_RECORDS And .State = adStateOpen Then
                        MsgBox "The search returned the maximum number of records allowed. More records may be available."
                    End If
                Else
                    lblComAddsRowCount.Caption = "0 rows returned."
                End If
             End With
             DoEvents
             '
             '   Reset the grid contents
             With TDBGridAdds
                 If m_recComAdds.RecordCount = 0 Then
                    .Bookmark = Null
                 Else
                    .Bookmark = 1
                 End If
                 
                 .ReBind
                 .ApproxCount = m_recComAdds.RecordCount
                 m_objComAddsGridMap.SetupInitialSortOrder
            End With
        End If
    End With
End Sub

Private Sub SetShpTopLocation()
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   Get the row we're in.
    Select Case Left$(sshpSelectedArea, 1)
        Case 0
            shpSelectedAreaPerimeter.Top = 690
        Case 1
            shpSelectedAreaPerimeter.Top = 960
        Case 2
            shpSelectedAreaPerimeter.Top = 1235
        Case 3
            shpSelectedAreaPerimeter.Top = 1500
        Case 4
            shpSelectedAreaPerimeter.Top = 1775
        Case 5
            shpSelectedAreaPerimeter.Top = 2040
        Case 6
            shpSelectedAreaPerimeter.Top = 2325
        Case 7
            shpSelectedAreaPerimeter.Top = 2615
    End Select
    '
    '   Get the column we're in.
    Select Case Right$(sshpSelectedArea, 1)
        Case 0
            shpSelectedAreaPerimeter.Left = 4980
        Case 1
            shpSelectedAreaPerimeter.Left = 5685
        Case 2
            shpSelectedAreaPerimeter.Left = 6390
        Case 3
            shpSelectedAreaPerimeter.Left = 7095
        Case 4
            shpSelectedAreaPerimeter.Left = 7800
        Case 5
            shpSelectedAreaPerimeter.Left = 8505
        Case 6
            shpSelectedAreaPerimeter.Left = 9210
        Case 7
            shpSelectedAreaPerimeter.Left = 9915
        Case 8
            shpSelectedAreaPerimeter.Left = 10620
    End Select
    Screen.MousePointer = vbNormal
End Sub


Private Sub SetShpTopLocationResi()
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   Get the row we're in, max 6 models with 2 format rows.
    Select Case Left$(sshpSelectedArea, 1)
        Case 0
            shpSelectedAreaPerimeterResi.Top = 690
        Case 1
            shpSelectedAreaPerimeterResi.Top = 960
        Case 2
            shpSelectedAreaPerimeterResi.Top = 1230
        Case 3
            shpSelectedAreaPerimeterResi.Top = 1500
        Case 4
            shpSelectedAreaPerimeterResi.Top = 1770
        Case 5
            shpSelectedAreaPerimeterResi.Top = 2040
        Case 6
            shpSelectedAreaPerimeterResi.Top = 2325
        Case 7
            shpSelectedAreaPerimeterResi.Top = 2610
    End Select
    '
    '   Get the column we're in.
    Select Case Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))
        Case 0
            shpSelectedAreaPerimeterResi.Left = 3690
        Case 1
            shpSelectedAreaPerimeterResi.Left = 4380
        Case 2
            shpSelectedAreaPerimeterResi.Left = 5070
        Case 3
            shpSelectedAreaPerimeterResi.Left = 5760
        Case 4
            shpSelectedAreaPerimeterResi.Left = 6450
        Case 5
            shpSelectedAreaPerimeterResi.Left = 7140
        Case 6
            shpSelectedAreaPerimeterResi.Left = 7830
        Case 7
            shpSelectedAreaPerimeterResi.Left = 8520
        Case 8
            shpSelectedAreaPerimeterResi.Left = 9210
        Case 9
            shpSelectedAreaPerimeterResi.Left = 9900
        Case 10
            shpSelectedAreaPerimeterResi.Left = 10590
    End Select
    Screen.MousePointer = vbNormal
End Sub

Private Sub EnableControls()
    Dim i   As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If bIsInitialLoad Then
        '
        '   IF we are editing a Residential Quality Series building
        '   we have to lock down fields and since there aren't any rows in
        '   published_bldg_matrix_cost for Quality models we populate the matrix differently.
        If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
        Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
            cmdNewModel.Enabled = False
            cmdReports.Enabled = False
            cmdUpdate.Enabled = False
            bRefreshCosts = False
            txtbldg_id.Locked = True
            txtbldg_id.BackColor = &H8000000F  ' &HC0C0C0
            
            fraNewBldgModelMatrix.Visible = False
            fraModelMatrixResi.Visible = True
            fraModelMatrix.Visible = False
            tabBldgAdditions.TabEnabled(1) = False
            cmdDeleteClone.Enabled = False
            '
            '   Disabled gray color
            cbobldg_categoryR.Enabled = False
            cbobldg_categoryR.BackColor = &H8000000F  ' &HC0C0C0
            fratype_code.Enabled = False
            fratype_code.BackColor = &H8000000F  '&HC0C0C0
            opttype_codeR.Enabled = False
            opttype_codeR.BackColor = &H8000000F  '&HC0C0C0
            opttype_codeC.Enabled = False
            opttype_codeC.BackColor = &H8000000F  '&HC0C0C0
            cboResiBldgType.Enabled = False
            cboResiBldgType.BackColor = &H8000000F  ' &HC0C0C0

            For i = 0 To 10
                txtAreaResi(i).Locked = True
                txtPerimeterResi(i).Locked = True
            Next i
       
            txtbldg_desc.Locked = True
            txtbldg_desc.BackColor = &H8000000F  '&HC0C0C0
            txtbldg_stories.Locked = True
            txtbldg_stories.BackColor = &H8000000F  '&HC0C0C0
            txtbldg_stories_hgt.Locked = True
            txtbldg_stories_hgt.BackColor = &H8000000F  '&HC0C0C0
            txtbldg_part_density.Locked = True
            txtbldg_part_density.BackColor = &H8000000F  '&HC0C0C0
            txtbldg_part_hgt.Locked = True
            txtbldg_part_hgt.BackColor = &H8000000F  '&HC0C0C0
            txtbldg_wall_factor.Locked = True
            txtbldg_wall_factor.BackColor = &H8000000F  '&HC0C0C0
            txtbldg_fixture_area.Locked = True
            txtbldg_fixture_area.BackColor = &H8000000F  '&HC0C0C0
            txtwindow_area.Locked = True
            txtwindow_area.BackColor = &H8000000F  '&HC0C0C0
            txtbldg_elev_no.Locked = True
            txtbldg_elev_no.BackColor = &H8000000F  '&HC0C0C0
            txtbldg_door_density.Locked = True
            txtbldg_door_density.BackColor = &H8000000F  '&HC0C0C0
            txtop_factor.Locked = True
            txtop_factor.BackColor = &H8000000F  '&HC0C0C0
            txtarchitect_fee.Locked = True
            txtarchitect_fee.BackColor = &H8000000F  '&HC0C0C0
            txtgraphic_ref_id.Locked = True
            txtgraphic_ref_id.BackColor = &H8000000F  '&HC0C0C0
            txtgraphic_ref_id2.Locked = True
            txtgraphic_ref_id2.BackColor = &H8000000F  '&HC0C0C0
            cboColumnToBold.Enabled = False
            cboColumnToBold.BackColor = &H8000000F  '&HC0C0C0
            cboRowToBold.Enabled = False
            cboRowToBold.BackColor = &H8000000F  '&HC0C0C0
        Else
            '
            '   Don't allow updates until something is changed.
            cmdUpdate.Enabled = False
            bRefreshCosts = False
            '
            '   If we are NOT inserting.
            If m_blnInsert = False Then
                '
                '   Lock fields that can't be changed.
                txtbldg_id.Locked = True
                txtbldg_id.BackColor = &H8000000F  '&HC0C0C0
                
                fraNewBldgModelMatrix.Visible = False
                
                If opttype_codeC.Value = True Then
                    fraModelMatrix.Visible = True
                    fraModelMatrixResi.Visible = False
                Else
                    If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
                        Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
                        Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
                    
                        fraModelMatrixResi.Width = 9250
                        For i = 8 To 10
                            txtAreaResi(i).Visible = False
                            txtPerimeterResi(i).Visible = False
                        Next i
                        For i = 0 To 7
                            lblCol9_TotalOPResi(i).Visible = False
                            lblCol10_TotalOPResi(i).Visible = False
                            lblCol11_TotalOPResi(i).Visible = False
                        Next i
                    End If
                    fraModelMatrixResi.Visible = True
                    fraModelMatrix.Visible = False
                End If
                tabBldgAdditions.TabEnabled(1) = True
                cmdDeleteClone.Enabled = False
                '
                '   Disabled gray color
                cbobldg_categoryR.Enabled = False
                cbobldg_categoryR.BackColor = &H8000000F  '&HC0C0C0
                cbobldg_categoryC.Enabled = False
                cbobldg_categoryC.BackColor = &H8000000F  '&HC0C0C0
                fratype_code.Enabled = False
                fratype_code.BackColor = &H8000000F  '&HC0C0C0
                opttype_codeR.Enabled = False
                opttype_codeR.BackColor = &H8000000F  '&HC0C0C0
                opttype_codeC.Enabled = False
                opttype_codeC.BackColor = &H8000000F  '&HC0C0C0
                cboResiBldgType.Enabled = False
                cboResiBldgType.BackColor = &H8000000F  '&HC0C0C0
            '
            '   Else if we're INSERTING but not Cloning.
            ElseIf m_blnClone = False Then
                fraNewBldgModelMatrix.Visible = True
                fraModelMatrix.Visible = False
                fraModelMatrixResi.Visible = False
                cmdDeleteClone.Enabled = False
                '
                '   Since we have a blank insert set the
                '   cbo categories to defaults = Commercial.
                opttype_codeC.Value = True
                cbobldg_categoryC.Visible = True
                cbobldg_categoryR.Visible = False
    
                txtNewBldgArea(9).Locked = True
                txtNewBldgArea(10).Locked = True
                txtNewBldgArea(9).Text = ""
                txtNewBldgArea(10).Text = ""
                txtNewBldgArea(9).Visible = False
                txtNewBldgArea(10).Visible = False
                For i = 0 To 8
                    txtNewBldgPerimeter(i).Locked = False
                Next i
                txtNewBldgPerimeter(9).Visible = False
                txtNewBldgPerimeter(10).Visible = False
                
                cboWallType(4).Visible = True
                cboWallType(5).Visible = True
                fraNewBldgModelMatrix.Height = 2715
                fraNewBldgModelMatrix.Width = 10000
                shpWhiteBackground.Height = 1980
                shpWhiteBackground.Width = 3675
    
                tabBldgAdditions.TabEnabled(1) = False
            '
            '   If we're Cloning.
            Else
                fraNewBldgModelMatrix.Visible = False
                If opttype_codeC.Value = True Then
                    fraModelMatrix.Visible = True
                    fraModelMatrixResi.Visible = False
                Else
                    If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
                        Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
                        Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
                    
                        fraModelMatrixResi.Width = 9250
                        For i = 8 To 10
                            txtAreaResi(i).Visible = False
                            txtPerimeterResi(i).Visible = False
                        Next i
                        For i = 0 To 7
                            lblCol9_TotalOPResi(i).Visible = False
                            lblCol10_TotalOPResi(i).Visible = False
                            lblCol11_TotalOPResi(i).Visible = False
                        Next i
                    End If
                    
                    fraModelMatrixResi.Visible = True
                    fraModelMatrix.Visible = False
                End If
                '
                '   If you clone you cannot change the type code for the
                '   building if commercial clone it stays a commercial clone.
                fratype_code.Enabled = False
                fratype_code.BackColor = &HC0C0C0
                opttype_codeR.Enabled = False
                opttype_codeR.BackColor = &H8000000F  '&HC0C0C0
                opttype_codeC.Enabled = False
                opttype_codeC.BackColor = &H8000000F  '&HC0C0C0
                
                tabBldgAdditions.TabEnabled(1) = True
                cmdDeleteClone.Enabled = True
            End If
        End If
    End If
    If opttype_codeR.Value = True Then
        cbobldg_categoryR.Visible = True
        cbobldg_categoryC.Visible = False

        cboResiBldgType.Visible = True
        lblResiBldgType.Visible = True
    Else
        cbobldg_categoryC.Visible = True
        cbobldg_categoryR.Visible = False

        cboResiBldgType.Visible = False
        lblResiBldgType.Visible = False
    End If
    '
    '   No current record - disable buttons
    If TDBGridAdds.Bookmark >= 1 Then
        cmdDeleteAdditive.Enabled = True
    Else
        cmdDeleteAdditive.Enabled = False
    End If
    ColorLockedFields Me
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub GotoModel(Optional bAreaNotClicked As Boolean)
    Dim strSelect               As String
    Dim m_recModels             As ADODB.RecordSet
    Dim nOrginalBookmark        As Variant
    Dim frm                     As frmModel

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Set frm = New frmModel
   
    With m_recModelMatrix
        nOrginalBookmark = .Bookmark
        .MoveFirst
        '
        '   Set bookmark to the row we're in plus 1.
        'need to check on eof
        If Not .EOF Then
            Do Until .Bookmark = Left$(sshpSelectedArea, 1) + 1
                If .EOF Then
                    Exit Do
                Else
                    .MoveNext
                End If
            Loop
        End If
        '
        '   Make sure it is closed.
        With m_recModels
            .Close
            '
            '   Set the maximum number to bring back.
            .MaxRecords = MAX_RECORDS
        End With
        '
        '   IF we are editing a Residential Quality Series building
        '   we have to lock down fields and since there aren't any rows in
        '   published_bldg_matrix_cost for Quality models we populate the matrix differently.
        If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
        Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
            strSelect = "exec sp_select_model_basements @bldg_id = '"
            strSelect = strSelect & Trim(m_rec.Fields("bldg_id").Value)
            strSelect = strSelect & "', @bldg_model_skey = '%" & .Fields("bldg_model_skey").Value & "'"
        Else
            strSelect = "exec sp_select_model @type_code = '%"
            strSelect = strSelect & "', @bldg_category = '%"
            strSelect = strSelect & "', @bldg_id = '"
            If Len(Trim(txtbldg_id.Text)) > 0 Then
                strSelect = strSelect & Trim(txtbldg_id.Text)
            Else
                strSelect = strSelect & "%"
            End If
            
            strSelect = strSelect & "', @bldg_desc = '%"
            strSelect = strSelect & "', @frame_type = '%" '& .Fields("frame_type").Value
            strSelect = strSelect & "', @wall_type = '%" '& .Fields("wall_type").Value
            
            'rlh - 07/14/2009 - BLDG 067 Model 6 is returning (2) rows instead of 1 for %76  (176 and 76)!!!
            'LEGACY CODE
            'strSELECT = strSELECT & "', @bldg_model_skey = '%" & .Fields("bldg_model_skey").Value & "'"
            'NEW CODE (rlh) - remove the % from bldg_model_skey search...
            strSelect = strSelect & "', @bldg_model_skey = '" & .Fields("bldg_model_skey").Value & "'"
        End If
        '
        '   Use DAL to perform select.
        If Not g_objDAL.GetRecordset(vbNullString, strSelect, m_recModels) Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching."
        Else
            '
            '   Pass the current record into the form,
            '   Navigating to single-record view.
            With frm
                '
                '   Don't pass an area if they didn't click it,
                '   that way we won't set focus to the rtpg tab.
                If bAreaNotClicked Then
                    .SetRow m_recModels, , , IIf(optUnion.Value = True, "Union", "Open")
                Else
                    .SetRow m_recModels, , sshpSelectedArea, IIf(optUnion.Value = True, "Union", "Open")
                End If
                .Show
            End With
        End If
        .Bookmark = nOrginalBookmark
    End With
    Screen.MousePointer = vbNormal
End Sub
'
'   Called instead of SetRow since the form is already loaded and we're
'   switching to a different bldg.
Private Sub SearchForNewBldg(sBldgID As String, sBldgDesc As String, _
                    Optional sBldgCategory As String)
    Dim strSelect               As String
    Dim Button                  As String
    
    On Error Resume Next
    If bIsPendingChange = True Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Search For New Building")
        If Button = vbYes Then
            '
            '   If there were errors, cancel the search
            If Not Update Then
                Exit Sub
            End If
        ElseIf Button = vbCancel Then
            '
            ' Cancel the search
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    '
    '   If we just did an insert we have to initialize the common adds grid.
    If m_blnInsert And Not m_blnClone Then
        '
        '   Initialize grids.  Do this here because
        '   if we are on a residential model then we have
        '   to enable the total cost columns in the common adds grid.
        With m_objComAddsGridMap
            .SetGrid TDBGridAdds
            .InitGrid (opttype_codeR.Value = True)
        End With
    End If
    '
    '   If we're searching for a new bldg we aren't
    '   inserting or cloning anymore.
    m_blnInsert = False
    m_blnClone = False
    bRefreshCosts = False
    '
    '   Make sure it is closed.
    With m_rec
        .Close
        '
        '   Set the maximum number to bring back.
        .MaxRecords = MAX_RECORDS
    End With
    '
    '   Don't use type_Code it's too restrictive here.
    strSelect = "exec sp_select_building @type_code = '%',"
    
    strSelect = strSelect & "@bldg_category = '"
    If Trim(sBldgCategory) = "" Then
        If Len(Trim(cbobldg_categoryR.Text)) = 0 Then
            strSelect = strSelect & "%"
        Else
            strSelect = strSelect & Trim(cbobldg_categoryR.Text)
        End If
    Else
        strSelect = strSelect & Trim(sBldgCategory)
    End If
    
    strSelect = strSelect & "', @bldg_id = '"
    strSelect = strSelect & Trim(sBldgID)
    strSelect = strSelect & "', @bldg_desc = '"
    strSelect = strSelect & Trim(sBldgDesc) & "'"
    '
    '   Use DAL to perform select.
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, m_rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for the building."
    Else
        bIsInitialLoad = True
        PopulateScreen
        EnableControls
        bIsInitialLoad = False
        bIsPendingChange = False
    End If
    Screen.MousePointer = vbNormal
End Sub
'
'   Used to enforce business rules and data integrity.
Private Function ValidateCommercialScreen() As Boolean
    Dim recTemp         As New ADODB.RecordSet
    Dim strMessage      As String
    Dim i               As Integer
    Dim strSelect       As String
    
    On Error Resume Next
    ValidateCommercialScreen = True

    If Trim(txtbldg_id.Text) = "" Then
        strMessage = "Please provide a building id."
        txtbldg_id.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_id.Text)) = False Then
        strMessage = "Please provide a numeric building id."
        txtbldg_id.SetFocus
        
    ElseIf Trim(txtbldg_desc.Text) = "" Then
        strMessage = "Please provide a building description that is less than 75 characters."
        txtbldg_desc.SetFocus

    ElseIf Len(Trim(txtbldg_desc.Text)) > 75 Then
        strMessage = "Please provide a building description."
        txtbldg_desc.SetFocus

    ElseIf Trim(txtbldg_stories.Text) = "" Then
        strMessage = "Please provide the number of building stories."
        txtbldg_stories.SetFocus
    
    ElseIf IsNumeric(Trim(txtbldg_stories.Text)) = False Then
        strMessage = "Please provide a numeric number of building stories."
        txtbldg_stories.SetFocus
    
    ElseIf Trim(txtbldg_stories.Text) < 1 Then
       strMessage = "Please provide a number of stories that is greater than or equal to 1."
       txtbldg_stories.SetFocus
     
    ElseIf Trim(txtbldg_stories.Text) > 50 Then
       strMessage = "Maximum number of stories for commercial building is 50."
       txtbldg_stories.SetFocus
    
    ElseIf Trim(txtbldg_stories_hgt.Text) = "" Then
        strMessage = "Please provide the building story height."
        txtbldg_stories_hgt.SetFocus
    
    ElseIf IsNumeric(Trim(txtbldg_stories_hgt.Text)) = False Then
        strMessage = "Please provide a numeric building story height."
        txtbldg_stories_hgt.SetFocus

    ElseIf Trim(txtbldg_stories_hgt.Text) < 1 Then
        strMessage = "Please provide a building story height that is greater than or equal to 1."
        txtbldg_stories_hgt.SetFocus
    
    ElseIf Trim(txtbldg_stories_hgt.Text) > 30 Then
       strMessage = "Maximum story height for commercial building is 30."
       txtbldg_stories_hgt.SetFocus

    ElseIf Trim(txtbldg_part_density.Text) = "" Then
        strMessage = "Please provide a building partition density."
        txtbldg_part_density.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_part_density.Text)) = False Then
        strMessage = "Please provide a numeric building partition density."
        txtbldg_part_density.SetFocus
    
    ElseIf Trim(txtbldg_part_density.Text) < 1 Then
        strMessage = "Please provide a building partition density that is greater than or equal to 1."
        txtbldg_part_density.SetFocus

    ElseIf Trim(txtbldg_part_density.Text) > 999 Then
        strMessage = "Maximum building partition density is 999."
        txtbldg_part_density.SetFocus

    ElseIf Trim(txtbldg_part_hgt.Text) = "" Then
        strMessage = "Please provide a building partition height."
        txtbldg_part_hgt.SetFocus
    
    ElseIf IsNumeric(Trim(txtbldg_part_hgt.Text)) = False Then
        strMessage = "Please provide a numeric building partition height."
        txtbldg_part_hgt.SetFocus
    
    ElseIf Trim(txtbldg_part_hgt.Text) < 1 Then
        strMessage = "Please provide a building partition height that is greater than or equal to 1."
        txtbldg_part_hgt.SetFocus

    ElseIf Trim(txtbldg_part_hgt.Text) > 30 Then
       strMessage = "Maximum building partition height for commercial building is 30."
       txtbldg_part_hgt.SetFocus
       
    ElseIf Trim(txtbldg_door_density.Text) = "" Then
        strMessage = "Please provide a building door density."
        txtbldg_door_density.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_door_density.Text)) = False Then
        strMessage = "Please provide a numeric building door density."
        txtbldg_door_density.SetFocus

    ElseIf Trim(txtbldg_door_density.Text) < 1 Then
        strMessage = "Please provide a building door density that is greater than or equal to 1."
        txtbldg_door_density.SetFocus
    
    ElseIf Trim(txtbldg_door_density.Text) > 9999 Then
       strMessage = "Maximum building door density for commercial building is 9999."
       txtbldg_door_density.SetFocus
    
    'REVISED 6/16/2005 RTD FOR VERSION 7.4.0 CR#1312
    'ElseIf Trim(txtbldg_elev_no.Text) = "" Then
    '    strMessage = "Please provide the number of building elevators."
    '    txtbldg_elev_no.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_elev_no.Text)) = False Then
        strMessage = "Please provide a numeric value for building elevators."
        txtbldg_elev_no.SetFocus

    'ElseIf Trim(txtbldg_elev_no.Text) < 1 Then
    '    strMessage = "Please provide a number of elevators that is greater than or equal to 1."
    '    txtbldg_elev_no.SetFocus
       
    ElseIf Trim(txtbldg_elev_no.Text) > 50 Then
       strMessage = "Maximum number of elevators for a commercial building is 50."
       txtbldg_elev_no.SetFocus

    ElseIf Trim(txtbldg_fixture_area.Text) = "" Then
        strMessage = "Please provide a building fixture area."
        txtbldg_fixture_area.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_fixture_area.Text)) = False Then
        strMessage = "Please provide a numeric building fixture area."
        txtbldg_fixture_area.SetFocus
      
    ElseIf Trim(txtbldg_fixture_area.Text) < 1 Then
       strMessage = "Please provide a building fixture area that is greater than or equal to 1."
       txtbldg_fixture_area.SetFocus
       
    ElseIf Trim(txtbldg_fixture_area.Text) > 99999 Then
       strMessage = "Maximum building fixture area for a commercial building is 99999."
       txtbldg_fixture_area.SetFocus

    ElseIf Trim(txtbldg_wall_factor.Text) = "" Then
        strMessage = "Please provide a building wall factor."
        txtbldg_wall_factor.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_wall_factor.Text)) = False Then
        strMessage = "Please provide a numeric building wall factor."
        txtbldg_wall_factor.SetFocus

    ElseIf Trim(txtbldg_wall_factor.Text) > 1 Then
        strMessage = "Please provide a building wall factor that is less than or equal to 1."
        txtbldg_wall_factor.SetFocus
    
    ElseIf Trim(txtbldg_wall_factor.Text) < 0.01 Then
        strMessage = "Please provide a building wall factor that is greater than or equal to .01."
        txtbldg_wall_factor.SetFocus

    ElseIf Trim(txtwindow_area.Text) = "" Then
        strMessage = "Please provide a building window area."
        txtwindow_area.SetFocus

    ElseIf IsNumeric(Trim(txtwindow_area.Text)) = False Then
        strMessage = "Please provide a numeric building window area."
        txtwindow_area.SetFocus

    ElseIf Trim(txtwindow_area.Text) < 0 Then
       strMessage = "Please provide a building window area that is greater than or equal to zero."
       txtwindow_area.SetFocus

    ElseIf Trim(txtwindow_area.Text) > 99 Then
       strMessage = "Maximum building window area is 99."
       txtwindow_area.SetFocus

    ElseIf Trim(txtop_factor.Text) = "" Then
        strMessage = "Please provide a building overhead & profit factor."
        txtop_factor.SetFocus

    ElseIf IsNumeric(Trim(txtop_factor.Text)) = False Then
        strMessage = "Please provide a numeric building overhead & profit factor."
        txtop_factor.SetFocus

    ElseIf Trim(txtop_factor.Text) > 1 Then
        strMessage = "Maximum building overhead & profit factor is 1."
        txtop_factor.SetFocus

    ElseIf Trim(txtop_factor.Text) < 0.01 Then
        strMessage = "Please provide a building overhead & profit factor that is greater than or equal to .01."
        txtop_factor.SetFocus

    ElseIf Trim(txtarchitect_fee.Text) = "" Then
        strMessage = "Please provide the building architect fees."
        txtarchitect_fee.SetFocus

    ElseIf IsNumeric(Trim(txtarchitect_fee.Text)) = False Then
        strMessage = "Please provide a numeric building architect fees."
        txtarchitect_fee.SetFocus

    ElseIf Trim(txtarchitect_fee.Text) > 1 Then
        strMessage = "Maximum building architect fee is 1."
        txtarchitect_fee.SetFocus
       
    ElseIf Trim(txtarchitect_fee.Text) < 0.01 Then
        strMessage = "Please provide a building architect fee that is greater than or equal to .01"
        txtarchitect_fee.SetFocus
        
    ElseIf Len(Trim(txtgraphic_ref_id.Text)) > 12 Then
        strMessage = "Please alter the file name to be 12 characters or less."
        txtgraphic_ref_id.SetFocus
    
    ElseIf Len(Trim(txtgraphic_ref_id2.Text)) > 12 Then
        strMessage = "Please alter the file name to be 12 characters or less."
        txtgraphic_ref_id2.SetFocus
    
    ElseIf cbobldg_categoryC.Text = "" Then
        strMessage = "Please select a building category."
        cbobldg_categoryC.SetFocus
    
    ElseIf cboRowToBold.listcount > 0 And cboRowToBold.Text = "" Then
        strMessage = "Please provide a row to bold."
        cboRowToBold.SetFocus
    
    ElseIf cboColumnToBold.listcount > 0 And cboColumnToBold.Text = "" Then
        strMessage = "Please provide a column to bold."
        cboColumnToBold.SetFocus
    Else
        '
        '   We're in a blank bldg have to reference fraNewBldgModelMatrix.
        If m_blnInsert And m_blnClone = False Then
            '
            '   Have to choose at least 1 wall type.
            If Trim(cboWallType(0).Text) = "" Then
                strMessage = "Please select at least 1 wall type."
                cboWallType(0).SetFocus
            '
            '   If type code is commercial have to choose at least 1 frame type.
            ElseIf Trim(cboFrameType(0).Text) = "" Then
                strMessage = "Please select at least 1 frame type."
                cboFrameType(0).SetFocus
            Else
                For i = 0 To 8
                    If Trim(txtNewBldgArea(i).Text) = "" Then
                        strMessage = "Please provide 9 total S.F. Areas."
                        txtNewBldgArea(i).SetFocus
                        Exit For
                    ElseIf IsNumeric(Trim(txtNewBldgArea(i).Text)) = False Then
                        strMessage = "Please provide a numeric S.F. Area."
                        txtNewBldgArea(i).SetFocus
                        Exit For
                    ElseIf Trim(txtNewBldgArea(i).Text) = 0 Then
                        strMessage = "Please provide a numeric S.F. Area that is greater than zero."
                        txtNewBldgArea(i).SetFocus
                        Exit For
                    ElseIf Trim(txtNewBldgPerimeter(i).Text) = "" Then
                        strMessage = "Please provide 9 total L.F. Perimeters."
                        txtNewBldgPerimeter(i).SetFocus
                        Exit For
                    ElseIf IsNumeric(Trim(txtNewBldgPerimeter(i).Text)) = False Then
                        strMessage = "Please provide a numeric L.F. Perimeter."
                        txtNewBldgPerimeter(i).SetFocus
                        Exit For
                    ElseIf Trim(txtNewBldgPerimeter(i).Text) = 0 Then
                        strMessage = "Please provide a numeric L.F. Perimeter that is greater than zero."
                        txtNewBldgPerimeter(i).SetFocus
                        Exit For
                    End If
                Next i
                For i = 0 To 5
                    If Trim(cboWallType(i).Text) <> "" And Trim(cboFrameType(i).Text) = "" Then
                        strMessage = "Please provide select frame type for the wall type."
                        cboFrameType(i).SetFocus
                        Exit For
                    ElseIf Trim(cboWallType(i).Text) = "" And Trim(cboFrameType(i).Text) <> "" Then
                        strMessage = "Please provide select wall type for the frame type."
                        cboWallType(i).SetFocus
                        Exit For
                    End If
                Next i
            End If
        '
        '   If Updating Commercial...
        Else
            For i = 0 To 8
                If Trim(txtArea(i).Text) = "" Then
                    strMessage = "Please provide 9 total S.F. Areas."
                    txtArea(i).SetFocus
                    Exit For
                ElseIf IsNumeric(Trim(txtArea(i).Text)) = False Then
                    strMessage = "Please provide a numeric S.F. Area."
                    txtArea(i).SetFocus
                    Exit For
                ElseIf Trim(txtArea(i).Text) = 0 Then
                    strMessage = "Please provide a numeric S.F. Area that is greater than zero."
                    txtArea(i).SetFocus
                    Exit For
                ElseIf Trim(txtPerimeter(i).Text) = "" Then
                    strMessage = "Please provide 9 total L.F. Perimeters."
                    txtPerimeter(i).SetFocus
                    Exit For
                ElseIf IsNumeric(Trim(txtPerimeter(i).Text)) = False Then
                    strMessage = "Please provide a numeric L.F. Perimeter."
                    txtPerimeter(i).SetFocus
                    Exit For
                ElseIf Trim(txtPerimeter(i).Text) = 0 Then
                    strMessage = "Please provide a numeric L.F. Perimeter that is greater than zero."
                    txtPerimeter(i).SetFocus
                    Exit For
                End If
            Next i
        End If
        '
        '   If we are inserting make sure the bldg_id is unique.  But if we are cloning then
        '   only do the check if the bldg_id is not the same as what the clone process generated.
        If strMessage = "" And m_blnInsert And Not m_blnClone Or _
            m_blnClone And Trim(txtbldg_id.Text) <> Trim(m_rec.Fields("bldg_id").Value) Then
            
            strSelect = "SELECT bldg_id FROM bldg_detail WHERE bldg_id = '" & Trim(txtbldg_id.Text) & "'"
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                strMessage = "An error occurred while searching to validate that the bldg_id was unique."
            ElseIf recTemp.RecordCount <> 0 Then
                strMessage = "Bldg_id already exists in the database, please provide a unique bldg_id."
                txtbldg_id.SetFocus
            End If
        ElseIf strMessage = "" And Not m_blnInsert And txtwindow_area.Text = 0 Then
            strSelect = "SELECT formula_code FROM assembly_usage JOIN bldg_model bm ON bm.bldg_skey = '" _
                & Trim(txtbldg_skey.Text) & "' WHERE parent_skey = bldg_model_skey" _
                & " AND formula_code = 'WW'"
                
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                strMessage = "An error occurred while searching to validate that algorithm 'WW' did not exist with window area of zero."
            ElseIf recTemp.RecordCount <> 0 Then
                strMessage = "Please provide a different Algorithm or change the building window area.  " _
                        & "Algorithm 'WW' is not allowed when the building window area is zero. "
                txtwindow_area.SetFocus
            End If
        End If
    End If
    '
    '   Can't insert common adds until building 1st saved.
    If m_blnInsert And m_blnClone = False Then
    ElseIf strMessage = "" Then
        strMessage = ValidateCommonAdds
    End If

    If strMessage <> "" Then
        Screen.MousePointer = vbNormal
        ValidateCommercialScreen = False
        MsgBox strMessage, vbCritical
    End If
End Function

Private Function ValidateResidentialScreen() As Boolean
    Dim recTemp         As New ADODB.RecordSet
    Dim strMessage      As String
    Dim i               As Integer
    Dim strSelect       As String
    
    On Error Resume Next
    ValidateResidentialScreen = True

    If Trim(txtbldg_id.Text) = "" Then
        strMessage = "Please provide a building id."
        txtbldg_id.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_id.Text)) = False Then
        strMessage = "Please provide a numeric building id."
        txtbldg_id.SetFocus
        
    ElseIf Trim(txtbldg_desc.Text) = "" Then
        strMessage = "Please provide a building description that is less than 75 characters."
        txtbldg_desc.SetFocus

    ElseIf Len(Trim(txtbldg_desc.Text)) > 75 Then
        strMessage = "Please provide a building description."
        txtbldg_desc.SetFocus

    ElseIf Trim(txtbldg_stories.Text) = "" Then
        strMessage = "Please provide the number of building stories."
        txtbldg_stories.SetFocus
    
    ElseIf IsNumeric(Trim(txtbldg_stories.Text)) = False Then
        strMessage = "Please provide a numeric number of building stories."
        txtbldg_stories.SetFocus
    
    ElseIf Trim(txtbldg_stories.Text) < 1 Then
       strMessage = "Please provide a number of stories that is greater than or equal to 1."
       txtbldg_stories.SetFocus
  
    ElseIf Trim(txtbldg_stories.Text) > 3 Then
       strMessage = "Maximum number of stories for residential building is 3."
       txtbldg_stories.SetFocus
    
    ElseIf Trim(txtbldg_stories_hgt.Text) = "" Then
        strMessage = "Please provide the building story height."
        txtbldg_stories_hgt.SetFocus
    
    ElseIf IsNumeric(Trim(txtbldg_stories_hgt.Text)) = False Then
        strMessage = "Please provide a numeric building story height."
        txtbldg_stories_hgt.SetFocus

    ElseIf Trim(txtbldg_stories_hgt.Text) < 1 Then
        strMessage = "Please provide a building story height that is greater than or equal to 1."
        txtbldg_stories_hgt.SetFocus

    ElseIf Trim(txtbldg_stories_hgt.Text) > 12 Then
       strMessage = "Maximum story height for residential building is 12."
       txtbldg_stories_hgt.SetFocus

    ElseIf Trim(txtbldg_part_density.Text) = "" Then
        strMessage = "Please provide a building partition density."
        txtbldg_part_density.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_part_density.Text)) = False Then
        strMessage = "Please provide a numeric building partition density."
        txtbldg_part_density.SetFocus
    
    ElseIf Trim(txtbldg_part_density.Text) < 1 Then
        strMessage = "Please provide a building partition density that is greater than or equal to 1."
        txtbldg_part_density.SetFocus

    ElseIf Trim(txtbldg_part_density.Text) > 999 Then
        strMessage = "Maximum building partition density is 999."
        txtbldg_part_density.SetFocus

    ElseIf Trim(txtbldg_part_hgt.Text) = "" Then
        strMessage = "Please provide a building partition height."
        txtbldg_part_hgt.SetFocus
    
    ElseIf IsNumeric(Trim(txtbldg_part_hgt.Text)) = False Then
        strMessage = "Please provide a numeric building partition height."
        txtbldg_part_hgt.SetFocus
    
    ElseIf Trim(txtbldg_part_hgt.Text) < 1 Then
        strMessage = "Please provide a building partition height that is greater than or equal to 1."
        txtbldg_part_hgt.SetFocus

    ElseIf Trim(txtbldg_part_hgt.Text) > 12 Then
       strMessage = "Maximum building partition height for residential building is 12."
       txtbldg_part_hgt.SetFocus
       
    ElseIf Trim(txtbldg_door_density.Text) = "" Then
        strMessage = "Please provide a building door density."
        txtbldg_door_density.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_door_density.Text)) = False Then
        strMessage = "Please provide a numeric building door density."
        txtbldg_door_density.SetFocus

    ElseIf Trim(txtbldg_door_density.Text) < 1 Then
        strMessage = "Please provide a building door density that is greater than or equal to 1."
        txtbldg_door_density.SetFocus
    
    ElseIf Trim(txtbldg_door_density.Text) > 999 Then
       strMessage = "Maximum building door density for residential building is 999."
       txtbldg_door_density.SetFocus
      
    ElseIf Trim(txtbldg_elev_no.Text) = "" Then
        strMessage = "Please provide the number of building elevators."
        txtbldg_elev_no.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_elev_no.Text)) = False Then
        strMessage = "Please provide a numeric value for building elevators."
        txtbldg_elev_no.SetFocus
    
    'REVISED 6/16/2005 RTD FOR VERSION 7.4.0 CR#1312
    'ElseIf Trim(txtbldg_elev_no.Text) < 1 Then
    '    strMessage = "Please provide a number of elevators that is greater than or equal to 1."
    '    txtbldg_elev_no.SetFocus
    
    ElseIf Trim(txtbldg_elev_no.Text) > 5 Then
       strMessage = "Maximum number of elevators for a residential building is 5."
       txtbldg_elev_no.SetFocus
    
    ElseIf Trim(txtbldg_fixture_area.Text) = "" Then
        strMessage = "Please provide a building fixture area."
        txtbldg_fixture_area.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_fixture_area.Text)) = False Then
        strMessage = "Please provide a numeric building fixture area."
        txtbldg_fixture_area.SetFocus

    ElseIf Trim(txtbldg_fixture_area.Text) < 0 Then
       strMessage = "Please provide a building fixture area that is greater than or equal to zero."
       txtbldg_fixture_area.SetFocus

    ElseIf Trim(txtbldg_fixture_area.Text) > 99 Then
       strMessage = "Maximum building fixture area for a residential building is 99."
       txtbldg_fixture_area.SetFocus
       
    ElseIf Trim(txtbldg_wall_factor.Text) = "" Then
        strMessage = "Please provide a building wall factor."
        txtbldg_wall_factor.SetFocus

    ElseIf IsNumeric(Trim(txtbldg_wall_factor.Text)) = False Then
        strMessage = "Please provide a numeric building wall factor."
        txtbldg_wall_factor.SetFocus

    ElseIf Trim(txtbldg_wall_factor.Text) > 1 Then
        strMessage = "Please provide a building wall factor that is less than or equal to 1."
        txtbldg_wall_factor.SetFocus
    
    ElseIf Trim(txtbldg_wall_factor.Text) < 0 Then
        strMessage = "Please provide a building wall factor that is greater than or equal to 0."
        txtbldg_wall_factor.SetFocus

    ElseIf Trim(txtwindow_area.Text) = "" Then
        strMessage = "Please provide a building window area."
        txtwindow_area.SetFocus

    ElseIf IsNumeric(Trim(txtwindow_area.Text)) = False Then
        strMessage = "Please provide a numeric building window area."
        txtwindow_area.SetFocus

    ElseIf Trim(txtwindow_area.Text) < 1 Then
       strMessage = "Please provide a building window area that is greater than or equal to 1."
       txtwindow_area.SetFocus

    ElseIf Trim(txtwindow_area.Text) > 99 Then
       strMessage = "Maximum building window area is 99."
       txtwindow_area.SetFocus

    ElseIf Trim(txtop_factor.Text) = "" Then
        strMessage = "Please provide a building overhead & profit factor."
        txtop_factor.SetFocus

    ElseIf IsNumeric(Trim(txtop_factor.Text)) = False Then
        strMessage = "Please provide a numeric building overhead & profit factor."
        txtop_factor.SetFocus

    ElseIf Trim(txtop_factor.Text) > 1 Then
        strMessage = "Maximum building overhead & profit factor is 1."
        txtop_factor.SetFocus

    ElseIf Trim(txtop_factor.Text) < 0 Then
        strMessage = "Please provide a building overhead & profit factor that is greater than or equal to zero."
        txtop_factor.SetFocus

    ElseIf Trim(txtarchitect_fee.Text) = "" Then
        strMessage = "Please provide the building architect fees."
        txtarchitect_fee.SetFocus

    ElseIf IsNumeric(Trim(txtarchitect_fee.Text)) = False Then
        strMessage = "Please provide a numeric building architect fees."
        txtarchitect_fee.SetFocus

    ElseIf Trim(txtarchitect_fee.Text) > 1 Then
        strMessage = "Maximum building architect fee is 1."
        txtarchitect_fee.SetFocus

    ElseIf Trim(txtarchitect_fee.Text) < 0 Then
        strMessage = "Please provide a building architect fee that is greater than or equal to zero."
        txtarchitect_fee.SetFocus
               
    ElseIf Len(Trim(txtgraphic_ref_id.Text)) > 12 Then
        strMessage = "Please alter the file name to be 12 characters or less."
        txtgraphic_ref_id.SetFocus
    
    ElseIf Len(Trim(txtgraphic_ref_id2.Text)) > 12 Then
        strMessage = "Please alter the file name to be 12 characters or less."
        txtgraphic_ref_id2.SetFocus
    
    ElseIf cbobldg_categoryR.Text = "" Then
        strMessage = "Please select a building category."
        cbobldg_categoryR.SetFocus

    ElseIf cboResiBldgType.Text = "" Then
        strMessage = "Please select a residential building type."
        cboResiBldgType.SetFocus
    
    ElseIf cboRowToBold.listcount > 0 And cboRowToBold.Text = "" Then
        strMessage = "Please provide a row to bold."
        cboRowToBold.SetFocus
    
    ElseIf cboColumnToBold.listcount > 0 And cboColumnToBold.Text = "" Then
        strMessage = "Please provide a column to bold."
        cboColumnToBold.SetFocus
    Else
        '
        '   We're in a blank bldg have to reference fraNewBldgModelMatrix.
        If m_blnInsert And m_blnClone = False Then
            '
            '   Have to choose at least 1 wall type.
            If Trim(cboWallType(0).Text) = "" Then
                strMessage = "Please select at least 1 wall type."
                cboWallType(0).SetFocus
            Else
                If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or Left$(Trim(cboResiBldgType.Text), 1) = "I" _
                    Or Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
                        
                    For i = 0 To 7
                        If Trim(txtNewBldgArea(i).Text) = "" Then
                            strMessage = "Please provide 8 total S.F. Areas."
                            txtNewBldgArea(i).SetFocus
                            Exit For
                        ElseIf IsNumeric(Trim(txtNewBldgArea(i).Text)) = False Then
                            strMessage = "Please provide a numeric S.F. Area."
                            txtNewBldgArea(i).SetFocus
                            Exit For
                        ElseIf Trim(txtNewBldgArea(i).Text) < 50 Then
                             strMessage = "Please provide a S.F. Area that is greater than or equal to 50."
                             txtNewBldgArea(i).SetFocus
                             Exit For
                        ElseIf Trim(txtNewBldgArea(i).Text) > 2000 Then
                             strMessage = "Please provide a numeric S.F. Area that is less than or equal to 2000."
                             txtNewBldgArea(i).SetFocus
                             Exit For
                        End If
                    Next i
                Else

                    For i = 0 To 10
                        If Trim(txtNewBldgArea(i).Text) = "" Then
                            strMessage = "Please provide 11 total S.F. Areas."
                            txtNewBldgArea(i).SetFocus
                            Exit For
                        ElseIf IsNumeric(Trim(txtNewBldgArea(i).Text)) = False Then
                            strMessage = "Please provide a numeric S.F. Area."
                            txtNewBldgArea(i).SetFocus
                            Exit For
                        ElseIf Trim(txtNewBldgArea(i).Text) = 0 Then
                            strMessage = "Please provide a numeric S.F. Area that is greater than zero."
                            txtNewBldgArea(i).SetFocus
                            Exit For
                        Else
                            If Trim(cbobldg_categoryR.Text) <> "Luxury" And Trim(txtNewBldgArea(i).Text) < 600 Then
                                strMessage = "Please provide a S.F. Area that is greater than or equal to 600."
                                txtNewBldgArea(i).SetFocus
                                Exit For
                                
                            ElseIf Trim(cbobldg_categoryR.Text) = "Luxury" And Trim(txtNewBldgArea(i).Text) < 1000 Then
                                strMessage = "Please provide a S.F. Area that is greater than or equal to 1000."
                                txtNewBldgArea(i).SetFocus
                                Exit For
                                               
                            ElseIf Trim(cbobldg_categoryR.Text) = "Economy" And Trim(txtNewBldgArea(i).Text) > 5000 Then
                                strMessage = "Please provide a S.F. Area that is less than or equal to 5000."
                                txtNewBldgArea(i).SetFocus
                                Exit For
                            
                            ElseIf Trim(cbobldg_categoryR.Text) <> "Economy" And Trim(txtNewBldgArea(i).Text) > 6000 Then
                                strMessage = "Please provide a S.F. Area that is less than or equal to 6000."
                                txtNewBldgArea(i).SetFocus
                                Exit For
                            End If
                        End If
                    Next i
                End If
            End If
        '
        '   If Updating Residential...
        Else
            '            Bldg Area Ranges
            '            Main Building               Wing & Ells = H, I, J
            '            Min          Max           Min           Max
            'Economy      600        5000            50            2000
            'Average      600        6000            50            2000
            'Custom       600        6000            50            2000
            'Luxury      1000        6000            50            2000
            If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or Left$(Trim(cboResiBldgType.Text), 1) = "I" _
                Or Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
                
                For i = 0 To 7
                    If Trim(txtAreaResi(i).Text) = "" Then
                        strMessage = "Please provide 8 total S.F. Areas."
                        txtNewBldgArea(i).SetFocus
                        Exit For
                    ElseIf IsNumeric(Trim(txtAreaResi(i).Text)) = False Then
                        strMessage = "Please provide a numeric S.F. Area."
                        txtAreaResi(i).SetFocus
                        Exit For
                    ElseIf Trim(txtAreaResi(i).Text) < 50 Then
                        strMessage = "Please provide a S.F. Area that is greater than or equal to 50."
                        txtAreaResi(i).SetFocus
                        Exit For
                    ElseIf Trim(txtAreaResi(i).Text) > 2000 Then
                        strMessage = "Please provide a numeric S.F. Area that is less than or equal to 2000."
                        txtAreaResi(i).SetFocus
                        Exit For
                    End If
                Next i
            Else
                For i = 0 To 10
                    If Trim(txtAreaResi(i).Text) = "" Then
                        strMessage = "Please provide 11 total S.F. Areas."
                        txtAreaResi(i).SetFocus
                        Exit For
                    ElseIf IsNumeric(Trim(txtAreaResi(i).Text)) = False Then
                        strMessage = "Please provide a numeric S.F. Area."
                        txtAreaResi(i).SetFocus
                        Exit For
                    ElseIf Trim(txtAreaResi(i).Text) = 0 Then
                        strMessage = "Please provide a numeric S.F. Area that is greater than zero."
                        txtAreaResi(i).SetFocus
                        Exit For
                    Else
                        If Trim(cbobldg_categoryR.Text) <> "Luxury" And Trim(txtAreaResi(i).Text) < 600 Then
                            strMessage = "Please provide a S.F. Area that is greater than or equal to 600."
                            txtAreaResi(i).SetFocus
                            Exit For
                            
                        ElseIf Trim(cbobldg_categoryR.Text) = "Luxury" And Trim(txtAreaResi(i).Text) < 1000 Then
                            strMessage = "Please provide a S.F. Area that is greater than or equal to 1000."
                            txtAreaResi(i).SetFocus
                            Exit For
                                           
                        ElseIf Trim(cbobldg_categoryR.Text) = "Economy" And Trim(txtAreaResi(i).Text) > 5000 Then
                            strMessage = "Please provide a S.F. Area that is less than or equal to 5000."
                            txtAreaResi(i).SetFocus
                            Exit For
                        
                        ElseIf Trim(cbobldg_categoryR.Text) <> "Economy" And Trim(txtAreaResi(i).Text) > 6000 Then
                            strMessage = "Please provide a S.F. Area that is less than or equal to 6000."
                            txtAreaResi(i).SetFocus
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If
        '
        '   If we are inserting make sure the bldg_id is unique.  But if we are cloning then
        '   only do the check if the bldg_id is not the same as what the clone process generated.
        If strMessage = "" And m_blnInsert And Not m_blnClone Or _
            m_blnClone And Trim(txtbldg_id.Text) <> Trim(m_rec.Fields("bldg_id").Value) Then
            
            strSelect = "SELECT bldg_id FROM bldg_detail WHERE bldg_id = '" & Trim(txtbldg_id.Text) & "'"
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                strMessage = "An error occurred while searching to validate that the bldg_id was unique."
            ElseIf recTemp.RecordCount <> 0 Then
                strMessage = "Bldg_id already exists in the database, please provide a unique bldg_id."
                txtbldg_id.SetFocus
            End If
        ElseIf strMessage = "" And Not m_blnInsert And txtwindow_area.Text = 0 Then
            strSelect = "SELECT formula_code FROM assembly_usage JOIN bldg_model bm ON bm.bldg_skey = '" _
                & Trim(txtbldg_skey.Text) & "' WHERE parent_skey = bldg_model_skey" _
                & " AND formula_code = 'WW'"
                
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                strMessage = "An error occurred while searching to validate that algorithm 'WW' did not exist with window area of zero."
            ElseIf recTemp.RecordCount <> 0 Then
                strMessage = "Please provide a different Algorithm or change the building window area.  " _
                        & "Algorithm 'WW' is not allowed when the building window area is zero. "
                txtwindow_area.SetFocus
            End If
        End If
    End If
    '
    '   Can't insert common adds until building 1st saved.
    If m_blnInsert And m_blnClone = False Then
    ElseIf strMessage = "" Then
        strMessage = ValidateCommonAdds
    End If

    If strMessage <> "" Then
        Screen.MousePointer = vbNormal
        ValidateResidentialScreen = False
        MsgBox strMessage, vbCritical
    End If
End Function

Private Function ValidateCommonAdds() As String

    On Error Resume Next
    With TDBGridAdds
        If m_objComAddsGridMap.bInsertInProcess Then
            ValidateCommonAdds = "Your last common add insert is still in progress.  " & vbCrLf & "Please move to another row allowing the insert to save within the grid, then click the Update button."
            '
            '   If we failed on this common add validation then
            '   re-enable the cmdupdate since the user doesn't
            '   have to actually change text just move to another
            '   row and then click update.
            If ValidateCommonAdds <> "" Then
                cmdUpdate.Enabled = True
            End If
        Else
            .MoveFirst
        
            Do Until IsNull(.Bookmark)
                If Trim(.Columns("Sort Order").Value) = "" Then
                    ValidateCommonAdds = "Please provide a Sort Order for ID: " & Trim(.Columns("ID").Value)
                    Exit Do
                ElseIf IsNumeric(Trim(.Columns("Sort Order").Value)) = False Then
                    ValidateCommonAdds = "Please provide a numeric Sort Order for ID: " & Trim(.Columns("ID").Value)
                    Exit Do
                ElseIf .Columns("Sort Order").Value = 0 Then
                    ValidateCommonAdds = "Please provide a numeric sort order that is greater than zero."
                    Exit Do
                ElseIf Trim(.Columns("Format Code").Value) = "" Then
                    ValidateCommonAdds = "Please provide a format code for ID: " & Trim(.Columns("ID").Value)
                    Exit Do
                ElseIf Trim(.Columns("Indent Code").Value) = "" Then
                    ValidateCommonAdds = "Please provide an indent code for ID: " & Trim(.Columns("ID").Value)
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
    End With
End Function

Private Sub cmdNewModel_Click()
    Dim recTemp     As New ADODB.RecordSet
    Dim rec         As New ADODB.RecordSet
    Dim strSelect   As String
    Dim frm         As frmModel

    On Error Resume Next
    strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & Trim(txtbldg_skey.Text) & "' AND model_code != '7' AND model_code != '8'"
    
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching."
    ElseIf opttype_codeC.Value = True And recTemp.RecordCount >= 6 Then
        MsgBox "Maximum number of models per commercial building is 6.  Currently this building has " & recTemp.RecordCount & " models.  " & vbCrLf & _
                "Inserting 1 additional model will exceed the maximum. ", vbCritical
    
    ElseIf opttype_codeR.Value = True And recTemp.RecordCount >= 4 Then
        MsgBox "Maximum number of models per residential building is 4.  Currently this building has " & recTemp.RecordCount & " models.  " & vbCrLf & _
                "Inserting 1 additional model will exceed the maximum. ", vbCritical
    Else
        recTemp.Close
        strSelect = "exec sp_select_model @type_code = '%"
        strSelect = strSelect & "', @bldg_category = '%"
        strSelect = strSelect & "', @bldg_id = '"
        If Len(Trim(txtbldg_id.Text)) > 0 Then
            strSelect = strSelect & Trim(txtbldg_id.Text)
        Else
            strSelect = strSelect & "%"
        End If
        
        strSelect = strSelect & "', @bldg_desc = '%"
        strSelect = strSelect & "', @frame_type = '%"
        strSelect = strSelect & "', @wall_type = '%"
        strSelect = strSelect & "', @bldg_model_skey = '%'"
        '
        '   Use DAL to perform select.
        If g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
        
            CopyRSFields rec, recTemp
            '
            '   Open empty single record view
            Set frm = New frmModel

            With frm
                rec.Fields("bldg_id").Value = Trim(txtbldg_id.Text)
                If Trim(rec.Fields("bldg_id").Value) <> "" Then
                    .SetRow rec, True
                    .Show
                End If
            End With
        Else
            MsgBox "An error occurred while searching for new model information.", vbCritical
        End If
    End If

    recTemp.Close
End Sub

Private Sub cmdReports_Click()
    If opttype_codeC.Value = True Then
        '
        '   They cannot run a report for model_code 7 & 8 for Commercial
        '   assemblies are same as model_code 1 anyway.
        If Trim(lblWall(Left$(sshpSelectedArea, 1)).Caption) = "Perimeter Adj., Add or Deduct" _
            Or Trim(lblWall(Left$(sshpSelectedArea, 1)).Caption) = "Story Hgt. Adj., Add or Deduct" Then
        
            MsgBox "Reports are not available on format rows (model codes 7 & 8) for Commercial models.", vbCritical
        Else
            RunReportCommercial
        End If
    Else
        RunReportResidential
    End If
End Sub

Private Sub RunReportCommercial()
    Dim strSelect               As String
    'Dim frm                     As New frmSummaryEstimateRpt

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    strSelect = "exec sp_rpt_summary_estimate_commercial @bldg_model_skey = '"
    '
    '   Get the model skey from the lblwall which is always populated whether
    '   Commercial or Residential.
    strSelect = strSelect & Right$(Trim(lblWall(Left$(sshpSelectedArea, 1)).Tag), Len(Trim(lblWall(Left$(sshpSelectedArea, 1)).Tag)) - InStr(1, Trim(lblWall(Left$(sshpSelectedArea, 1)).Tag), "|"))
    '
    '   Indicates where the shpSelectedArea is for modelmaint button click.
    '   In the format of 1,1 meaning row 1 col area1.
    strSelect = strSelect & "', @bldg_area = '"
    strSelect = strSelect & txtArea(Right$(sshpSelectedArea, 1)).Text
         
    strSelect = strSelect & "', @op_code = '"
    If optUnion.Value = True Then
        strSelect = strSelect & "STD"
    Else
        strSelect = strSelect & "OPN"
    End If
    
    strSelect = strSelect & "', @country_code = '"
    If Len(Trim(cboMdlCountryCode.Text)) = 0 Then
        strSelect = strSelect & "USA"
    Else
        strSelect = strSelect & SQLChangeWildcard(cboMdlCountryCode.Text)
    End If
    
    strSelect = strSelect & "', @region_code = '"
    If Len(Trim(cboMdlRegionCode.Text)) = 0 Then
        strSelect = strSelect & "NAT'"
    Else
        strSelect = strSelect & SQLChangeWildcard(cboMdlRegionCode.Text) & "'"
    End If
    
    Screen.MousePointer = vbNormal
    'frm.RunReportCommercial strSELECT
    CommercialEstimatePreview strSelect
    
End Sub

Private Sub RunReportResidential()
    Dim strSelect               As String
    'Dim frm                     As New frmSummaryEstimateRpt

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    strSelect = "exec sp_rpt_summary_estimate_residential @bldg_model_skey = '"
    '
    '   Get the model skey from the lblwall which is always populated whether
    '   Commercial or Residential.
    strSelect = strSelect & Right$(Trim(lblWallResi(Left$(sshpSelectedArea, 1)).Tag), Len(Trim(lblWallResi(Left$(sshpSelectedArea, 1)).Tag)) - InStr(1, Trim(lblWallResi(Left$(sshpSelectedArea, 1)).Tag), "|"))
    '
    '   Indicates where the shpSelectedArea is for modelmaint button click.
    '   In the format of 1,1 meaning row 1 col area1.
    strSelect = strSelect & "', @bldg_area = '"
    strSelect = strSelect & txtAreaResi(Right$(sshpSelectedArea, 1)).Text
         
    strSelect = strSelect & "', @op_code = '"
    If optUnion.Value = True Then
        strSelect = strSelect & "STD"
    Else
        strSelect = strSelect & "OPN"
    End If
    
    strSelect = strSelect & "', @country_code = '"
    If Len(Trim(cboMdlCountryCode.Text)) = 0 Then
        strSelect = strSelect & "USA"
    Else
        strSelect = strSelect & SQLChangeWildcard(cboMdlCountryCode.Text)
    End If
    
    strSelect = strSelect & "', @region_code = '"
    If Len(Trim(cboMdlRegionCode.Text)) = 0 Then
        strSelect = strSelect & "NAT'"
    Else
        strSelect = strSelect & SQLChangeWildcard(cboMdlRegionCode.Text) & "'"
    End If
    
    Screen.MousePointer = vbNormal
    'frm.RunReportResidential strSELECT
    ResidentialEstimatePreview strSelect
    
End Sub

Private Sub cmdGraphic1File_Click()
    On Error Resume Next
    oDlg.InitDir = "C:\"
    oDlg.ShowOpen
    If oDlg.FileTitle <> "" Then
        txtgraphic_ref_id.Text = oDlg.FileTitle
    End If
End Sub

Private Sub cmdGraphic2File_Click()
    On Error Resume Next
    oDlg.InitDir = "C:\"
    oDlg.ShowOpen
    If oDlg.FileTitle <> "" Then
        txtgraphic_ref_id2.Text = oDlg.FileTitle
    End If
End Sub

Private Sub cmdAssemblyCost_Click()
    Dim ID      As String
    Dim frm     As frmAssemblyGrid
    
    On Error Resume Next
    '
    '   Open single record view with data from row selected.
    With TDBGridAdds
        If Not IsNull(.Bookmark) Then
            Set frm = New frmAssemblyGrid
    
            If Len(Trim(.Columns("ID").Value)) = 14 Then
                ID = Left$(.Columns("ID").Value, 5) & Right$(Left$(.Columns("ID").Value, 9), 3) & Right$(.Columns("ID").Value, 4)
            ElseIf Len(Trim(.Columns("ID").Value)) = 12 And InStr(1, Trim(.Columns("ID").Value), " ") <> 0 Then
                ID = Left$(.Columns("ID").Value, 3) & Right$(Left$(.Columns("ID").Value, 7), 3) & Right$(.Columns("ID").Value, 4)
            Else
                ID = .Columns("ID").Value
            End If
            frm.JumpIn Trim(ID)
        End If
    End With
End Sub

Private Sub cmdUnitCost_Click()
    Dim ID      As String
    Dim frm     As frmUCostUsageGrid
    
    On Error Resume Next
    '
    '   Open single record view with data from row selected.
    With TDBGridAdds
        If Not IsNull(.Bookmark) Then
            Set frm = New frmUCostUsageGrid
    
            If Len(Trim(.Columns("ID").Value)) = 14 Then
                ID = Left$(.Columns("ID").Value, 5) & Right$(Left$(.Columns("ID").Value, 9), 3) & Right$(.Columns("ID").Value, 4)
            ElseIf Len(Trim(.Columns("ID").Value)) = 12 And InStr(1, Trim(.Columns("ID").Value), " ") <> 0 Then
                ID = Left$(.Columns("ID").Value, 3) & Right$(Left$(.Columns("ID").Value, 7), 3) & Right$(.Columns("ID").Value, 4)
            Else
                ID = .Columns("ID").Value
            End If
            'frm.MasterFormat = 1995
            frm.JumpIn Trim(ID)
        End If
    End With
End Sub

Private Sub cmdModelMaint_Click()
    Screen.MousePointer = vbHourglass
    GotoModel
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdDeleteAdditive_Click()
    Dim varButton
    
    On Error Resume Next
    With TDBGridAdds
        If .SelBookmarks.Count = 1 Then
            varButton = MsgBox(CStr(.SelBookmarks.Count) + " records will be deleted.  Are you sure you want to delete this row permanently?", vbYesNo + vbCritical)
            If varButton = vbYes Then
                .Delete
            End If
        '
        '   If multiple records are selected.
        ElseIf .SelBookmarks.Count > 1 Then
            varButton = MsgBox(CStr(.SelBookmarks.Count) + " records will be deleted.  Are you sure you want to delete these rows permanently?", vbYesNo + vbCritical)
            If varButton = vbYes Then
                m_objComAddsGridMap.Delete
            End If
        End If
    End With
End Sub

Private Sub cmdDeleteClone_Click()
    Dim Button
    Dim strUpdate       As String
    Dim strError        As String
    
    On Error Resume Next
    Button = MsgBox("Are you sure you want to delete this cloned building?", vbYesNo + vbCritical)
    If Button = vbYes Then
      Screen.MousePointer = vbHourglass

      strUpdate = "exec sp_delete_building @bldg_skey = '"
      If Trim(txtbldg_skey.Text) = "" Then
          Exit Sub
      Else
          strUpdate = strUpdate & Trim(txtbldg_skey.Text) & "'"
      End If
      
      If Not g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
          MsgBox "Error deleting building clone." & vbCrLf & strError & ".", vbCritical
      Else
          '
          '   Always refresh forms that are listening for changes in case part of the update succeeded.
          '   ie -the bldg was updated and the grid has an old last_update_id.
          EventSubscriberNotify esnBuildingRecordUpdated, m_rec.Fields("bldg_id").Value
          Me.Hide
          Unload Me
      End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdUpdate_Click()
    '
    '   Set cmdUpdate to not enabled so they must change
    '   data to update again or if they fail screen validation
    '   they must change before updating.
    '
    '   NOTE that if they fail the common add validation insert in progress
    '   then validatecommonadds will re-enable the cmdupdate since the user doesn't
    '   have to actually change text just move to another
    '   row and then click update.
    cmdUpdate.Enabled = False
    Update
End Sub

Private Function Update() As Boolean
    Dim sParen              As String
    Dim sPipe               As String
    Dim sbldg_area_std      As String
    Dim sbldg_perimeter_std As String
    Dim nbldg_form          As Integer
    
    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    
    sParen = InStr(1, cboColumnToBold.Text, "]")
    sPipe = InStr(1, cboColumnToBold.Text, "|")
    '
    'in the format of [1] area | perimeter
    sbldg_area_std = Trim(Mid(cboColumnToBold.Text, sParen + 1, sPipe - 1 - sParen - 1))
    sbldg_perimeter_std = Trim(Right$(cboColumnToBold.Text, Len(cboColumnToBold.Text) - sPipe))
    
    If opttype_codeC.Value = True Then
        If ValidateCommercialScreen Then
            '
            '   Now branch between inserting a new building or updating an existing building.
            If m_blnInsert = True And m_blnClone = False Then
                Status ("Updating Building Details ...")
                If InsertCommercial(sbldg_area_std, sbldg_perimeter_std) Then
                    Update = True
                End If
            Else
                Status ("Getting Temporary Building Form ID ...")
                '
                '   Get the form id that uniquely identifies us so that if we
                '   added Common Adds we can get them from the
                '   tmp_common_add_book_detail table.  This is a required parameter
                '   for the update sp's for Commercial & Residential.
                nbldg_form = GetFormID
                If nbldg_form = 0 Then
                    '
                    '   There was an error so we can't update.
                    Exit Function
                End If
                
                Status ("Updating Temporary Common Adds Table ...")
                '
                '   Insert the common adds in the temp table that the
                '   building update sp uses to commit finals from.
                If UpdateCommonAdds(nbldg_form) Then
                    Status ("Updating Building Details ...")
                    
                    If UpdateCommercial(sbldg_area_std, sbldg_perimeter_std, nbldg_form) Then
                        If bRefreshCosts Then
                            If RefreshCostsCommercial Then
                                Update = True
                            End If
                        Else
                            Update = True
                        End If
                    End If
                End If
            End If
            Status ("Cleaning Temporary Tables ...")
            '
            '   Regardless if we updated ok or not, cleanup the tmp_common_add_book_detail
            '   and tmp_form_id tables for our common adds.
            CleanupTmpCommonAdds nbldg_form

            If Update Then
                Status ("Building Details Updated Successfully ...")
                MsgBox "Building Updated Successfully", vbInformation
                Status ("Refreshing Building Maintenance Screen ...")
                bIsPendingChange = False
                '
                '   Pass in the bldg_category to search with since it may not match the
                '   top cbobldgcategory.
                '
                '   Never know if apos ' will be in bldg_desc.
                SearchForNewBldg txtbldg_id.Text, Replace(Trim(txtbldg_desc.Text), "'", "''"), cbobldg_categoryC.Text
            End If
        End If
    Else
        If ValidateResidentialScreen Then
            '
            '   Now branch between inserting a new building or updating an existing building.
            If m_blnInsert = True And m_blnClone = False Then
                Status ("Updating Building Details ...")
                If InsertResidential(sbldg_area_std, sbldg_perimeter_std) Then
                    Update = True
                End If
            Else
                Status ("Getting Temporary Building Form ID ...")
                '
                '   Get the form id that uniquely identifies us so that if we
                '   added Common Adds we can get them from the
                '   tmp_common_add_book_detail table.  This is a required parameter
                '   for the update sp's for Commercial & Residential.
                nbldg_form = GetFormID
                If nbldg_form = 0 Then
                    '
                    '   There was an error so we can't update.
                    Exit Function
                End If

                Status ("Updating Temporary Tables ...")
                '
                '   Insert the common adds in the temp table that the
                '   building update sp uses to commit finals from.
                If UpdateCommonAdds(nbldg_form) Then
                    Status ("Updating Building Details ...")

                    If UpdateResidential(sbldg_area_std, sbldg_perimeter_std, nbldg_form) Then
                        If bRefreshCosts Then
                            If RefreshCostsResidential Then
                                Update = True
                            End If
                        Else
                            Update = True
                        End If
                    End If
                End If
            End If
            Status ("Cleaning Temporary Common Add Table ...")
            '
            '   Regardless if we updated ok or not, cleanup the tmp_common_add_book_detail
            '   and tmp_form_id tables for our common adds.
            CleanupTmpCommonAdds nbldg_form
                
            If Update Then
                Status ("Building Details Updated Successfully ...")
                MsgBox "Building Updated Successfully", vbInformation
                Status ("Refreshing Building Maintenance Screen ...")
                bIsPendingChange = False
                '
                '   Pass in the bldg_category to search with since it may not match the
                '   top cbobldgcategory.
                '
                '   Never know if apos ' will be in bldg_desc.
                SearchForNewBldg txtbldg_id.Text, Replace(Trim(txtbldg_desc.Text), "'", "''"), cbobldg_categoryR.Text
            End If
        End If
    End If
    '
    '   Always refresh forms that are listening for changes in case part of the update succeeded.
    '   ie -the bldg was updated and the grid has an old last_update_id.
    EventSubscriberNotify esnBuildingRecordUpdated, m_rec.Fields("bldg_id").Value
    Status ("")
    Screen.MousePointer = vbNormal
    Exit Function
    
errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the Update routine: " & Err.Description, vbCritical
    Status ("")
End Function

Private Function GetFormID() As Integer
    Dim recTemp         As New ADODB.RecordSet
    Dim sErrorDesc      As String
    Dim bOK             As Boolean
    Dim org_form_id     As Integer
    Dim strUpdate       As String
    
    On Error Resume Next
    Do Until bOK
        '
        '   Get bldg_form number to use in tmp table to indentify common adds as ours.
        If g_objDAL.GetRecordset(vbNullString, "SELECT MAX(form_id) AS form_id FROM form_id", recTemp) Then
            If sErrorDesc = "" Then
                org_form_id = recTemp.Fields("form_id").Value
                recTemp.Close
            
                strUpdate = "INSERT INTO form_id(form_id, form_type) VALUES('" & org_form_id + 1 & "', 'B')"
                If g_objDAL.ExecQuery(vbNullString, strUpdate, sErrorDesc) Then
                    If sErrorDesc = "" Then
                        GetFormID = org_form_id + 1
                        bOK = True
                    '
                    '   If the error is due to primary key constraint then just get the next
                    '   form_id and try again.  Otherwise we have a real error so exit.
                    ElseIf sErrorDesc = "[Microsoft][ODBC SQL Server Driver][SQL Server]Violation of PRIMARY KEY constraint 'PK_tmp_form_id'. Cannot insert duplicate key in object 'tmp_form_id'." Then
                        sErrorDesc = ""
                    Else
                        Screen.MousePointer = vbNormal
                        MsgBox "Error setting Form ID in temporary table 'form_id' " _
                            & vbCrLf & "Error: " & sErrorDesc, vbCritical
                        GetFormID = 0
                        Exit Do
                    End If
                Else
                    Screen.MousePointer = vbNormal
                    MsgBox "Error setting Form ID in temporary table 'form_id' " _
                        & vbCrLf & "Error: " & sErrorDesc, vbCritical
                    GetFormID = 0
                    Exit Do
                End If
            Else
                Screen.MousePointer = vbNormal
                MsgBox "Error selecting Form ID in temporary table 'form_id' " _
                    & vbCrLf & "Error: " & sErrorDesc, vbCritical
                GetFormID = 0
                Exit Do
            End If
        Else
            Screen.MousePointer = vbNormal
            MsgBox "Error selecting Form ID in temporary table 'form_id' " _
                & vbCrLf & "Error: " & sErrorDesc, vbCritical
            GetFormID = 0
            Exit Do
        End If
    Loop
End Function

Private Function UpdateCommonAdds(nbldg_form As Integer) As Boolean
    
    On Error GoTo errorHandler:
    '
    '   They can't add common adds until the bldg is saved.
    With TDBGridAdds
        .MoveFirst
        .Update
        UpdateCommonAdds = m_objComAddsGridMap.Update(Trim(txtbldg_skey.Text), IIf(opttype_codeC = True, "C", "R"), nbldg_form)
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the UpdateCommonAdds routine: " & Err.Description, vbCritical
    Status ("")
End Function
   
Private Function UpdateCommercial(sbldg_area_std As String, sbldg_perimeter_std As String, _
                                nbldg_form As Integer) As Boolean
    
    Dim cmdTemp         As New ADODB.Command
    Dim strError        As String
    Dim strUpdate       As String
    Dim sBldgDesc       As String
    Dim nCounter        As Integer
    Dim i               As Integer
    
    On Error GoTo errorHandler:
    
    strUpdate = "exec sp_update_commercial_building @bldg_skey = '" & Trim(txtbldg_skey.Text) & "',"
    strUpdate = strUpdate & "@bldg_id = '" & Trim(txtbldg_id.Text) & "',"
    strUpdate = strUpdate & "@bldg_category = '" & Trim(cbobldg_categoryC.Text) & "',"
    '
    '   Since we might have ' marks in our book desc, need to replace with '' for SQL.
    sBldgDesc = Trim(txtbldg_desc.Text)
    sBldgDesc = Replace(sBldgDesc, "'", "''", 1)
    strUpdate = strUpdate & "@bldg_desc = '" & sBldgDesc & "',"
    
    strUpdate = strUpdate & "@bldg_stories = " & Trim(txtbldg_stories.Text) & ","
    strUpdate = strUpdate & "@bldg_stories_hgt = " & Trim(txtbldg_stories_hgt.Text) & ","
    strUpdate = strUpdate & "@bldg_part_density = " & Trim(txtbldg_part_density.Text) & ","
    strUpdate = strUpdate & "@bldg_part_hgt = " & Trim(txtbldg_part_hgt.Text) & ","
    strUpdate = strUpdate & "@bldg_door_density = " & Trim(txtbldg_door_density.Text) & ","
    strUpdate = strUpdate & "@bldg_type = '0',"
    strUpdate = strUpdate & "@bldg_area_std = " & sbldg_area_std & ","
    strUpdate = strUpdate & "@bldg_perimeter_std = " & sbldg_perimeter_std & ","
    strUpdate = strUpdate & "@bldg_wall_factor = " & Trim(txtbldg_wall_factor.Text) & ","
    strUpdate = strUpdate & "@bldg_elev_no = " & Trim(txtbldg_elev_no.Text) & ","
    strUpdate = strUpdate & "@bldg_fixture_area = " & Trim(txtbldg_fixture_area.Text) & ","
    strUpdate = strUpdate & "@window_area = " & Trim(txtwindow_area.Text) & ","
    strUpdate = strUpdate & "@op_factor = " & Trim(txtop_factor.Text) & ","
    strUpdate = strUpdate & "@architect_fee = " & Trim(txtarchitect_fee.Text) & ","
    '
    'in the format of [1] wall | frame
    strUpdate = strUpdate & "@row_to_bold = " & IIf(Right$(Left$(cboRowToBold.Text, 2), 1) = "", 1, Right$(Left$(cboRowToBold.Text, 2), 1)) & ","
    strUpdate = strUpdate & "@col_to_bold = " & Left$(Trim(Replace(cboColumnToBold.Text, "[", "")), InStr(1, Trim(Replace(cboColumnToBold.Text, "[", "")), "]") - 1) & ","
    strUpdate = strUpdate & "@graphic_ref_id = '" & Trim(txtgraphic_ref_id.Text) & "',"
    strUpdate = strUpdate & "@graphic_ref_id2 = '" & Trim(txtgraphic_ref_id2.Text) & "',"
    
    strUpdate = strUpdate & " @last_update_id_bldg = '" & Trim(m_rec.Fields("last_update_id").Value) & "',"
    '
    '   Now update any areas in our grid txtArea(0) & txtPerimeter(0)
    '   that have changed.
    For i = 0 To 8
        If Left$(txtArea(i).Tag, 1) = "C" Or Left(txtPerimeter(i).Tag, 1) = "C" Then
            nCounter = nCounter + 1
            
            strUpdate = strUpdate & "@bldg_area_" & nCounter & " = '" & Trim(txtArea(i).Text) & "',"
            strUpdate = strUpdate & "@bldg_perimeter_" & nCounter & " = '" & Trim(txtPerimeter(i).Text) & "',"
            '
            '   If the Area changed get the original from the tag.
            If Left$(txtArea(i).Tag, 1) = "C" Then
                strUpdate = strUpdate & "@bldg_orig_area_" & nCounter & " = '" & Left$(Right$(txtArea(i).Tag, Len(txtArea(i).Tag) - 1), InStr(1, txtArea(i).Tag, "|") - 2) & "',"
            Else
                strUpdate = strUpdate & "@bldg_orig_area_" & nCounter & " = '" & Trim(txtArea(i).Text) & "',"
            End If
            strUpdate = strUpdate & "@area_ind_" & nCounter & " = '" & IIf(txtArea(i).Text = sbldg_area_std, 1, 0) & "',"
            strUpdate = strUpdate & " @last_update_id_area_" & nCounter & " = '" & Right$(txtArea(i).Tag, Len(txtArea(i).Tag) - InStr(1, txtArea(i).Tag, "|")) & "',"
        End If
    Next i
    For i = IIf(nCounter = 0, 1, nCounter + 1) To 9
        strUpdate = strUpdate & "@bldg_area_" & i & " = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_" & i & " = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_" & i & " = '0',"
        strUpdate = strUpdate & "@area_ind_" & i & " = '0',"
        strUpdate = strUpdate & "@last_update_id_area_" & i & " = '0',"
    Next i
    '
    '   common_add_book_detail -ensures that we only add records that our form inserted.
    strUpdate = strUpdate & "@bldg_form = '" & nbldg_form & "',"
    strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"

    With cnTemp
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            .Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the UpdateCommercial routine. " _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            UpdateCommercial = True
        End If
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the UpdateCommercial routine: " & Err.Description, vbCritical
    Status ("")
End Function
    
Private Function InsertCommercial(sbldg_area_std As String, sbldg_perimeter_std As String) As Boolean
    Dim cmdTemp         As New ADODB.Command
    Dim strError        As String
    Dim strUpdate       As String
    Dim sBldgDesc       As String
    Dim s2ndWallType    As String
    Dim sTempWallType   As String
    Dim i               As Integer
    
    On Error GoTo errorHandler:
        
    strUpdate = "exec sp_insert_commercial_building @bldg_id= '" & Trim(txtbldg_id.Text) & "',"
    strUpdate = strUpdate & "@type_code = 'C',"
    strUpdate = strUpdate & "@bldg_category = '" & Trim(cbobldg_categoryC.Text) & "',"
    '
    '   Since we might have ' marks in our book desc, need to replace with '' for SQL.
    sBldgDesc = Trim(txtbldg_desc.Text)
    sBldgDesc = Replace(sBldgDesc, "'", "''", 1)
    strUpdate = strUpdate & "@bldg_desc = '" & sBldgDesc & "',"
    
    strUpdate = strUpdate & "@bldg_stories = " & Trim(txtbldg_stories.Text) & ","
    strUpdate = strUpdate & "@bldg_stories_hgt = " & Trim(txtbldg_stories_hgt.Text) & ","
    strUpdate = strUpdate & "@bldg_part_density = " & Trim(txtbldg_part_density.Text) & ","
    strUpdate = strUpdate & "@bldg_part_hgt = " & Trim(txtbldg_part_hgt.Text) & ","
    strUpdate = strUpdate & "@bldg_door_density = " & Trim(txtbldg_door_density.Text) & ","
    strUpdate = strUpdate & "@bldg_type = '0',"
    strUpdate = strUpdate & "@bldg_area_std = " & sbldg_area_std & ","
    strUpdate = strUpdate & "@bldg_perimeter_std = " & sbldg_perimeter_std & ","
    strUpdate = strUpdate & "@bldg_wall_factor = " & Trim(txtbldg_wall_factor.Text) & ","
    strUpdate = strUpdate & "@bldg_elev_no = " & Trim(txtbldg_elev_no.Text) & ","
    strUpdate = strUpdate & "@bldg_fixture_area = " & Trim(txtbldg_fixture_area.Text) & ","
    strUpdate = strUpdate & "@window_area = " & Trim(txtwindow_area.Text) & ","
    strUpdate = strUpdate & "@op_factor = " & Trim(txtop_factor.Text) & ","
    strUpdate = strUpdate & "@architect_fee = " & Trim(txtarchitect_fee.Text) & ","
    '
    'in the format of [1] wall | frame
    strUpdate = strUpdate & "@row_to_bold = " & IIf(Right$(Left$(cboRowToBold.Text, 2), 1) = "", 1, Right$(Left$(cboRowToBold.Text, 2), 1)) & ","
    strUpdate = strUpdate & "@col_to_bold = " & Left$(Trim(Replace(cboColumnToBold.Text, "[", "")), InStr(1, Trim(Replace(cboColumnToBold.Text, "[", "")), "]") - 1) & ","
    strUpdate = strUpdate & "@graphic_ref_id = '" & Trim(txtgraphic_ref_id.Text) & "',"
    strUpdate = strUpdate & "@graphic_ref_id2 = '" & Trim(txtgraphic_ref_id2.Text) & "',"
    '
    '   Now update any areas in our grid txtArea(0) & txtPerimeter(0)
    '   that have changed.
    For i = 0 To 8
        strUpdate = strUpdate & "@bldg_area_" & i + 1 & " = '" & Trim(txtNewBldgArea(i).Text) & "',"
        strUpdate = strUpdate & "@bldg_perimeter_" & i + 1 & " = '" & Trim(txtNewBldgPerimeter(i).Text) & "',"
        strUpdate = strUpdate & "@area_ind_" & i + 1 & " = '" & IIf(txtNewBldgArea(i).Text = sbldg_area_std, 1, 0) & "',"
    Next i
    '
    '   Update the bldg_model table.
    For i = 0 To 5
        If Trim(cboWallType(i).Text) <> "" Then
            strUpdate = strUpdate & " @frame_type_" & i + 1 & " = '" & Trim(cboFrameType(i).Text) & "',"
            strUpdate = strUpdate & " @wall_type_" & i + 1 & " = '" & Trim(cboWallType(i).Text) & "',"
            strUpdate = strUpdate & " @format_code_" & i + 1 & " = 'A4',"
        Else
            strUpdate = strUpdate & "@frame_type_" & i + 1 & " = '',"
            strUpdate = strUpdate & "@wall_type_" & i + 1 & " = '',"
            strUpdate = strUpdate & "@format_code_" & i + 1 & " = '',"
        End If
    Next i

    strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"

    With cnTemp
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            .Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the InsertCommercial routine." _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            InsertCommercial = True
        End If
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the InsertCommercial routine: " & Err.Description, vbCritical
    Status ("")
End Function

Private Function RefreshCostsCommercial() As Boolean
    Dim strUpdate           As String
    Dim strSelect           As String
    Dim sbldg_model_skey    As String
    Dim i                   As Integer
    Dim recTemp             As New ADODB.RecordSet
    Dim cmdTemp             As New ADODB.Command
    
    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    RefreshCostsCommercial = True
    Status ("Updating Building Cost Information ...")
    With cnTemp
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
        '
        '   If we're inserting and we just added the models, we have to
        '   query to get the model skey's 1st.
        If m_blnInsert And m_blnClone = False Then
            strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & Trim(txtbldg_skey.Text) _
                        & "' AND model_code != '7' AND model_code != '8'"
        
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                Screen.MousePointer = vbNormal
                MsgBox "Errors in the RefreshCostsCommercial routine searching for bldg_model_skey's.", vbCritical
                RefreshCostsCommercial = False
            Else
                With recTemp
                    If .RecordCount > 0 Then
                        Do Until .EOF
                            Status ("Updating Building Cost Information For Model: " & Trim(.Fields("bldg_model_skey").Value) & " ...")
                            DoEvents
                            strUpdate = "exec sp_update_bldg_model @bldg_model_skey = '"
                            strUpdate = strUpdate & Trim(.Fields("bldg_model_skey").Value) & "',"
                            strUpdate = strUpdate & "@op_code = 'STD',"
                            strUpdate = strUpdate & "@country_code = '" & Trim(cboMdlCountryCode.Text) & "',"
                            strUpdate = strUpdate & "@region_code = '" & Trim(cboMdlRegionCode.Text) & "'"
                            With cmdTemp
                                .CommandTimeout = 50000
                                .CommandType = adCmdText
                                .CommandText = strUpdate
                                .Execute adExecuteNoRecords
                            End With
                            DoEvents
                            If cnTemp.Errors.Count = 0 Then
                                strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
                                With cmdTemp
                                    .CommandTimeout = 50000
                                    .CommandType = adCmdText
                                    .CommandText = strUpdate
                                    .Execute adExecuteNoRecords
                                End With
                                DoEvents
                                If cnTemp.Errors.Count <> 0 Then
                                    Screen.MousePointer = vbNormal
                                    MsgBox "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                                        & Trim(.Fields("bldg_model_skey").Value) & " " & vbCrLf & cnTemp.Errors(0).Description _
                                        & vbCrLf & "RefreshCostsCommercial routine will continue for other models.", vbCritical
                                    Screen.MousePointer = vbHourglass
                                End If
                            Else
                                Screen.MousePointer = vbNormal
                                MsgBox "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                                    & Trim(.Fields("bldg_model_skey").Value) & " " & vbCrLf & cnTemp.Errors(0).Description _
                                    & vbCrLf & "RefreshCostsCommercial routine will continue for other models.", vbCritical
                                Screen.MousePointer = vbHourglass
                            End If
                            .MoveNext
                        Loop
                        RefreshCostsCommercial = True
                    Else
                        Screen.MousePointer = vbNormal
                        MsgBox "Errors in the RefreshCostsCommercial routine, unable to locate bldg_model_skey's associated with the building", vbCritical
                        RefreshCostsCommercial = False
                    End If
                End With
            End If
        Else
            For i = 1 To 6
                sbldg_model_skey = Right$(Trim(lblWall(i - 1).Tag), Len(Trim(lblWall(i - 1).Tag)) - InStr(1, Trim(lblWall(i - 1).Tag), "|"))
                If sbldg_model_skey <> "" Then
                    Status ("Updating Building Cost Information For Model: " & sbldg_model_skey & " ...")
                    DoEvents
                    strUpdate = "exec sp_update_bldg_model @bldg_model_skey = '" & sbldg_model_skey & "',"
                    strUpdate = strUpdate & "@op_code = 'STD',"
                    strUpdate = strUpdate & "@country_code = '" & Trim(cboMdlCountryCode.Text) & "',"
                    strUpdate = strUpdate & "@region_code = '" & Trim(cboMdlRegionCode.Text) & "'"
                    
                    With cmdTemp
                        .CommandTimeout = 50000
                        .CommandType = adCmdText
                        .CommandText = strUpdate
                        .Execute adExecuteNoRecords
                    End With
                    DoEvents
                    If cnTemp.Errors.Count = 0 Then
                        strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
                        With cmdTemp
                            .CommandTimeout = 50000
                            .CommandType = adCmdText
                            .CommandText = strUpdate
                            .Execute adExecuteNoRecords
                        End With
                        DoEvents
                        If cnTemp.Errors.Count <> 0 Then
                            Screen.MousePointer = vbNormal
                            MsgBox "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                                & sbldg_model_skey & " " & vbCrLf & cnTemp.Errors(0).Description _
                                & vbCrLf & "RefreshCostsCommercial routine will continue for other models.", vbCritical
                            Screen.MousePointer = vbHourglass
                        End If
                    Else
                        Screen.MousePointer = vbNormal
                        MsgBox "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                            & sbldg_model_skey & " " & vbCrLf & cnTemp.Errors(0).Description _
                            & vbCrLf & "RefreshCostscCommercial routine will continue for other models.", vbCritical
                        Screen.MousePointer = vbHourglass
                    End If
                End If
                sbldg_model_skey = ""
            Next i
            RefreshCostsCommercial = True
        End If
    End With
    Screen.MousePointer = vbNormal
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    RefreshCostsCommercial = False
    MsgBox "Errors in the RefreshCostsCommercial routine: " & Err.Description, vbCritical
    Status ("")
End Function
  
Private Function UpdateResidential(sbldg_area_std As String, sbldg_perimeter_std As String, _
                                nbldg_form As Integer) As Boolean

    Dim cmdTemp         As New ADODB.Command
    Dim strError        As String
    Dim strUpdate       As String
    Dim sBldgDesc       As String
    Dim nCounter        As Integer
    Dim i               As Integer
    
    On Error GoTo errorHandler:
        
    strUpdate = "exec sp_update_residential_building @bldg_skey = '" & Trim(txtbldg_skey.Text) & "',"
    strUpdate = strUpdate & "@bldg_id = '" & Trim(txtbldg_id.Text) & "',"
    strUpdate = strUpdate & "@bldg_category = '" & Trim(cbobldg_categoryR.Text) & "',"
    '
    '   Since we might have ' marks in our book desc, need to replace with '' for SQL.
    sBldgDesc = Trim(txtbldg_desc.Text)
    sBldgDesc = Replace(sBldgDesc, "'", "''", 1)
    strUpdate = strUpdate & "@bldg_desc = '" & sBldgDesc & "',"
    
    strUpdate = strUpdate & "@bldg_stories = " & Trim(txtbldg_stories.Text) & ","
    strUpdate = strUpdate & "@bldg_stories_hgt = " & Trim(txtbldg_stories_hgt.Text) & ","
    strUpdate = strUpdate & "@bldg_part_density = " & Trim(txtbldg_part_density.Text) & ","
    strUpdate = strUpdate & "@bldg_part_hgt = " & Trim(txtbldg_part_hgt.Text) & ","
    strUpdate = strUpdate & "@bldg_door_density = " & Trim(txtbldg_door_density.Text) & ","
    strUpdate = strUpdate & "@bldg_type = '" & Left(Trim(cboResiBldgType.Text), 1) & "',"
    strUpdate = strUpdate & "@bldg_area_std = " & sbldg_area_std & ","
    strUpdate = strUpdate & "@bldg_perimeter_std = " & sbldg_perimeter_std & ","
    strUpdate = strUpdate & "@bldg_wall_factor = " & Trim(txtbldg_wall_factor.Text) & ","
    strUpdate = strUpdate & "@bldg_elev_no = " & Trim(txtbldg_elev_no.Text) & ","
    strUpdate = strUpdate & "@bldg_fixture_area = " & Trim(txtbldg_fixture_area.Text) & ","
    strUpdate = strUpdate & "@window_area = " & Trim(txtwindow_area.Text) & ","
    strUpdate = strUpdate & "@op_factor = " & Trim(txtop_factor.Text) & ","
    strUpdate = strUpdate & "@architect_fee = " & Trim(txtarchitect_fee.Text) & ","
    '
    'in the format of [1] wall | frame
    strUpdate = strUpdate & "@row_to_bold = " & IIf(Right$(Left$(cboRowToBold.Text, 2), 1) = "", 1, Right$(Left$(cboRowToBold.Text, 2), 1)) & ","
    strUpdate = strUpdate & "@col_to_bold = " & Left$(Trim(Replace(cboColumnToBold.Text, "[", "")), InStr(1, Trim(Replace(cboColumnToBold.Text, "[", "")), "]") - 1) & ","
    strUpdate = strUpdate & "@graphic_ref_id = '" & Trim(txtgraphic_ref_id.Text) & "',"
    strUpdate = strUpdate & "@graphic_ref_id2 = '" & Trim(txtgraphic_ref_id2.Text) & "',"
    
    strUpdate = strUpdate & " @last_update_id_bldg = '" & Trim(m_rec.Fields("last_update_id").Value) & "',"
    '
    '   Now update any areas in our grid txtArea(0) & txtPerimeter(0)
    '   that have changed.
    For i = 0 To 10
        If Left$(txtAreaResi(i).Tag, 1) = "C" Or Left(txtPerimeterResi(i).Tag, 1) = "C" Then
            nCounter = nCounter + 1
            
            strUpdate = strUpdate & "@bldg_area_" & nCounter & " = '" & Trim(txtAreaResi(i).Text) & "',"
            strUpdate = strUpdate & "@bldg_perimeter_" & nCounter & " = '" & Trim(txtPerimeterResi(i).Text) & "',"
            '
            '   If the Area changed get the original from the tag.
            If Left$(txtAreaResi(i).Tag, 1) = "C" Then
                strUpdate = strUpdate & "@bldg_orig_area_" & nCounter & " = '" & Left$(Right$(txtAreaResi(i).Tag, Len(txtAreaResi(i).Tag) - 1), InStr(1, txtAreaResi(i).Tag, "|") - 2) & "',"
            Else
                strUpdate = strUpdate & "@bldg_orig_area_" & nCounter & " = '" & Trim(txtAreaResi(i).Text) & "',"
            End If
            strUpdate = strUpdate & "@area_ind_" & nCounter & " = '" & IIf(txtAreaResi(i).Text = sbldg_area_std, 1, 0) & "',"
            strUpdate = strUpdate & " @last_update_id_area_" & nCounter & " = '" & Right$(txtAreaResi(i).Tag, Len(txtAreaResi(i).Tag) - InStr(1, txtAreaResi(i).Tag, "|")) & "',"
        End If
    Next i
    For i = IIf(nCounter = 0, 1, nCounter + 1) To 11
        strUpdate = strUpdate & "@bldg_area_" & i & " = '0',"
        strUpdate = strUpdate & "@bldg_perimeter_" & i & " = '0',"
        strUpdate = strUpdate & "@bldg_orig_area_" & i & " = '0',"
        strUpdate = strUpdate & "@area_ind_" & i & " = '0',"
        strUpdate = strUpdate & "@last_update_id_area_" & i & " = '0',"
    Next i
    '
    '   common_add_book_detail -ensures that we only add records that our form inserted.
    strUpdate = strUpdate & "@bldg_form = '" & nbldg_form & "',"
    strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"

    With cnTemp
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            .Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the UpdateResidential routine. " _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            UpdateResidential = True
        End If
    End With
    Exit Function
    
errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the UpdateResidential routine: " & Err.Description, vbCritical
    Status ("")
End Function
    
Private Function InsertResidential(sbldg_area_std As String, sbldg_perimeter_std As String) As Boolean
    Dim cmdTemp         As New ADODB.Command
    Dim strError        As String
    Dim strUpdate       As String
    Dim sBldgDesc       As String
    Dim s2ndWallType    As String
    Dim sTempWallType   As String
    Dim i               As Integer
    
    On Error GoTo errorHandler:
        
    strUpdate = "exec sp_insert_residential_building @bldg_id= '" & Trim(txtbldg_id.Text) & "',"
    strUpdate = strUpdate & "@type_code = 'R',"
    strUpdate = strUpdate & "@bldg_category = '" & Trim(cbobldg_categoryR.Text) & "',"
    '
    '   Since we might have ' marks in our book desc, need to replace with '' for SQL.
    sBldgDesc = Trim(txtbldg_desc.Text)
    sBldgDesc = Replace(sBldgDesc, "'", "''", 1)
    strUpdate = strUpdate & "@bldg_desc = '" & sBldgDesc & "',"
    
    strUpdate = strUpdate & "@bldg_stories = " & Trim(txtbldg_stories.Text) & ","
    strUpdate = strUpdate & "@bldg_stories_hgt = " & Trim(txtbldg_stories_hgt.Text) & ","
    strUpdate = strUpdate & "@bldg_part_density = " & Trim(txtbldg_part_density.Text) & ","
    strUpdate = strUpdate & "@bldg_part_hgt = " & Trim(txtbldg_part_hgt.Text) & ","
    strUpdate = strUpdate & "@bldg_door_density = " & Trim(txtbldg_door_density.Text) & ","
    strUpdate = strUpdate & "@bldg_type = '" & Left(Trim(cboResiBldgType.Text), 1) & "',"
    strUpdate = strUpdate & "@bldg_area_std = " & sbldg_area_std & ","
    strUpdate = strUpdate & "@bldg_perimeter_std = " & sbldg_perimeter_std & ","
    strUpdate = strUpdate & "@bldg_wall_factor = " & Trim(txtbldg_wall_factor.Text) & ","
    strUpdate = strUpdate & "@bldg_elev_no = " & Trim(txtbldg_elev_no.Text) & ","
    strUpdate = strUpdate & "@bldg_fixture_area = " & Trim(txtbldg_fixture_area.Text) & ","
    strUpdate = strUpdate & "@window_area = " & Trim(txtwindow_area.Text) & ","
    strUpdate = strUpdate & "@op_factor = " & Trim(txtop_factor.Text) & ","
    strUpdate = strUpdate & "@architect_fee = " & Trim(txtarchitect_fee.Text) & ","
    '
    'in the format of [1] wall | frame
    strUpdate = strUpdate & "@row_to_bold = " & IIf(Right$(Left$(cboRowToBold.Text, 2), 1) = "", 1, Right$(Left$(cboRowToBold.Text, 2), 1)) & ","
    strUpdate = strUpdate & "@col_to_bold = " & Left$(Trim(Replace(cboColumnToBold.Text, "[", "")), InStr(1, Trim(Replace(cboColumnToBold.Text, "[", "")), "]") - 1) & ","
    strUpdate = strUpdate & "@graphic_ref_id = '" & Trim(txtgraphic_ref_id.Text) & "',"
    strUpdate = strUpdate & "@graphic_ref_id2 = '" & Trim(txtgraphic_ref_id2.Text) & "',"
    '
    '   Now update any areas in our grid txtArea(0) & txtPerimeter(0)
    '   that have changed.
    '   Note if Wings & Ells we only have 8 Areas.
    If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or Left$(Trim(cboResiBldgType.Text), 1) = "I" Or Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
        For i = 0 To 7
            strUpdate = strUpdate & "@bldg_area_" & i + 1 & " = '" & Trim(txtNewBldgArea(i).Text) & "',"
            strUpdate = strUpdate & "@bldg_perimeter_" & i + 1 & " = '" & Trim(txtNewBldgPerimeter(i).Text) & "',"
            strUpdate = strUpdate & "@area_ind_" & i + 1 & " = '" & IIf(txtNewBldgArea(i).Text = sbldg_area_std, 1, 0) & "',"
        Next i
        For i = 9 To 11
            strUpdate = strUpdate & "@bldg_area_" & i & " = '0',"
            strUpdate = strUpdate & "@bldg_perimeter_" & i & " = '0',"
            strUpdate = strUpdate & "@area_ind_" & i & " = '0',"
        Next i
    Else
        For i = 0 To 10
            strUpdate = strUpdate & "@bldg_area_" & i + 1 & " = '" & Trim(txtNewBldgArea(i).Text) & "',"
            strUpdate = strUpdate & "@bldg_perimeter_" & i + 1 & " = '" & Trim(txtNewBldgPerimeter(i).Text) & "',"
            strUpdate = strUpdate & "@area_ind_" & i + 1 & " = '" & IIf(txtNewBldgArea(i).Text = sbldg_area_std, 1, 0) & "',"
        Next i
    End If
    '
    '   Update the bldg_model table.
    For i = 0 To 3
        If Trim(cboWallType(i).Text) <> "" Then
            '
            '   Frame type is not populated for Resi models, it's included within the wall type.
            strUpdate = strUpdate & " @frame_type_" & i + 1 & " = '',"
            strUpdate = strUpdate & " @wall_type_" & i + 1 & " = '" & Trim(cboWallType(i).Text) & "',"
            strUpdate = strUpdate & " @format_code_" & i + 1 & " = 'A4',"
        Else
            strUpdate = strUpdate & "@frame_type_" & i + 1 & " = '',"
            strUpdate = strUpdate & "@wall_type_" & i + 1 & " = '',"
            strUpdate = strUpdate & "@format_code_" & i + 1 & " = '',"
        End If
    Next i

    strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"

    With cnTemp
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            .Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the InsertResidential routine." _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            InsertResidential = True
        End If
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the InsertResidential routine: " & Err.Description, vbCritical
    Status ("")
End Function

Private Function RefreshCostsResidential() As Boolean
    Dim strError            As String
    Dim strUpdate           As String
    Dim strSelect           As String
    Dim sbldg_model_skey    As String
    Dim i                   As Integer
    Dim recTemp             As New ADODB.RecordSet
    Dim cmdTemp             As New ADODB.Command
    
    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    RefreshCostsResidential = True
    Status ("Updating Building Cost Information ...")
    With cnTemp
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
        '
        '   If we're inserting and we just added the models, we have to
        '   query to get the model skey's 1st.
        If m_blnInsert And m_blnClone = False Then
            strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & Trim(txtbldg_skey.Text) _
                & "' AND model_code != '7' AND model_code != '8'"
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Then
                Screen.MousePointer = vbNormal
                MsgBox "Errors in the RefreshCostsResidential routine searching for bldg_model_skey's.", vbCritical
                RefreshCostsResidential = False
            Else
                With recTemp
                    If .RecordCount > 0 Then
                        Do Until .EOF
                            Status ("Updating Building Cost Information For Model: " & Trim(.Fields("bldg_model_skey").Value) & " ...")
                            DoEvents
                            strUpdate = "exec sp_update_bldg_model_resi @bldg_model_skey = '"
                            strUpdate = strUpdate & Trim(.Fields("bldg_model_skey").Value) & "',"
                            strUpdate = strUpdate & "@op_code = 'STD',"
                            '
                            'allow to update & change order of models?
                            strUpdate = strUpdate & "@country_code = 'USA',"
                            strUpdate = strUpdate & "@region_code = 'NAT'"
                            With cmdTemp
                                .CommandTimeout = 50000
                                .CommandType = adCmdText
                                .CommandText = strUpdate
                                .Execute adExecuteNoRecords
                            End With
                            DoEvents
                            If cnTemp.Errors.Count = 0 Then
                                strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
                                With cmdTemp
                                    .CommandTimeout = 50000
                                    .CommandType = adCmdText
                                    .CommandText = strUpdate
                                    .Execute adExecuteNoRecords
                                End With
                    
                                DoEvents
                                If cnTemp.Errors.Count <> 0 Then
                                    Screen.MousePointer = vbNormal
                                    MsgBox "Errors in the RefreshCostsResidential routine for Building Model skey: " _
                                        & Trim(.Fields("bldg_model_skey").Value) & " " & vbCrLf & cnTemp.Errors(0).Description _
                                        & vbCrLf & "RefreshCostsResidential routine will continue for other models.", vbCritical
                                    Screen.MousePointer = vbHourglass
                                End If
                            Else
                                Screen.MousePointer = vbNormal
                                MsgBox "Errors in the RefreshCostsResidential routine for Building Model skey: " _
                                    & Trim(.Fields("bldg_model_skey").Value) & " " & vbCrLf & cnTemp.Errors(0).Description _
                                    & vbCrLf & "RefreshCostsResidential routine will continue for other models.", vbCritical
                                Screen.MousePointer = vbHourglass
                            End If
                            .MoveNext
                        Loop
                    Else
                        Screen.MousePointer = vbNormal
                        MsgBox "Errors in the RefreshCostsResidential routine, unable to locate bldg_model_skey's associated with the building", vbCritical
                        RefreshCostsResidential = False
                    End If
                End With
            End If
        Else
            '
            '   Don't refresh costs for model_code 7 & 8,
            '   they use assemblies at the Quality Series bldg level, but Tom has not
            '   modified the sp to re-calculate yet.
            For i = 1 To 4
                '
                '   In the format A4|1078 which is format_code pipe and bldg_model_skey.
                sbldg_model_skey = Right$(Trim(lblWallResi(i - 1).Tag), Len(Trim(lblWallResi(i - 1).Tag)) - InStr(1, Trim(lblWallResi(i - 1).Tag), "|"))
                If sbldg_model_skey <> "" Then
                    '
                    '   Now only refresh if not a basement
                    If Trim(lblWallResi(i - 1).Caption) <> "Finished Basement, Add" And _
                        Trim(lblWallResi(i - 1).Caption) <> "Unfinished Basement, Add" Then
                    
                        Status ("Updating Building Cost Information For Model: " & sbldg_model_skey & " ...")
                        strUpdate = "exec sp_update_bldg_model_resi @bldg_model_skey = '" & sbldg_model_skey & "',"
                        strUpdate = strUpdate & "@op_code = 'STD',"
                        '
                        'allow to update & change order of models?
                        strUpdate = strUpdate & "@country_code = 'USA',"
                        strUpdate = strUpdate & "@region_code = 'NAT'"
                        With cmdTemp
                            .CommandTimeout = 50000
                            .CommandType = adCmdText
                            .CommandText = strUpdate
                            .Execute adExecuteNoRecords
                        End With
                        DoEvents
                        If cnTemp.Errors.Count = 0 Then
                            strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
                            With cmdTemp
                                .CommandTimeout = 50000
                                .CommandType = adCmdText
                                .CommandText = strUpdate
                                .Execute adExecuteNoRecords
                            End With
                
                            DoEvents
                            If cnTemp.Errors.Count <> 0 Then
                                Screen.MousePointer = vbNormal
                                MsgBox "Errors in the RefreshCostsResidential routine for Building Model skey: " _
                                    & sbldg_model_skey & " " & vbCrLf & cnTemp.Errors(0).Description _
                                    & vbCrLf & "RefreshCostsResidential routine will continue for other models.", vbCritical
                                Screen.MousePointer = vbHourglass
                            End If
                        Else
                            Screen.MousePointer = vbNormal
                            MsgBox "Errors in the RefreshCostsResidential routine for Building Model skey: " _
                                & sbldg_model_skey & " " & vbCrLf & cnTemp.Errors(0).Description _
                                & vbCrLf & "RefreshCostsResidential routine will continue for other models.", vbCritical
                            Screen.MousePointer = vbHourglass
                        End If
                    End If
                End If
                sbldg_model_skey = ""
                DoEvents
            Next i
        End If
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    RefreshCostsResidential = False
    MsgBox "Errors in the RefreshCostsResidential routine: " & Err.Description, vbCritical
    Status ("")
End Function

Private Sub CleanupTmpCommonAdds(nbldg_form As Integer)
    Dim strUpdate   As String
    Dim sErrorDesc  As String
    
    On Error Resume Next
    strUpdate = "DELETE FROM common_add_book_detail_holding_table WHERE bldg_form = '" & nbldg_form & "'"
    If Not g_objDAL.ExecQuery(vbNullString, strUpdate, sErrorDesc) Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Common Add temporary table 'common_add_book_detail_holding_table' for form_id: " & nbldg_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    ElseIf sErrorDesc <> "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Common Add temporary table 'common_add_book_detail_holding_table' for form_id: " & nbldg_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    End If
    
    strUpdate = "DELETE FROM form_id WHERE form_id = '" & nbldg_form & "'"
    If Not g_objDAL.ExecQuery(vbNullString, strUpdate, sErrorDesc) Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Form ID temporary table 'form_id' for form_id: " & nbldg_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    ElseIf sErrorDesc <> "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Form ID temporary table 'form_id' for form_id: " & nbldg_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    End If
End Sub
'
'   Called from frmMain when the user clicks on the
'   toolbar buttons for sorting.
Public Sub Sort(intDir As Integer)
    m_objComAddsGridMap.Sort intDir
End Sub

Private Sub lblCol1_TotalOP_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol1_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 1st column.
        sshpSelectedArea = Index & ",0"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol1_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol1_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol2_TotalOP_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol2_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 2nd column.
        sshpSelectedArea = Index & ",1"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol2_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol2_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol3_TotalOP_Click(Index As Integer)
   '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol3_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 3rd column.
        sshpSelectedArea = Index & ",2"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol3_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol3_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol4_TotalOP_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol4_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 4th column.
        sshpSelectedArea = Index & ",3"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol4_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol4_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol5_TotalOP_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol5_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 5th column.
        sshpSelectedArea = Index & ",4"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol5_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol5_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol6_TotalOP_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol6_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 6th column.
        sshpSelectedArea = Index & ",5"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol6_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol6_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol7_TotalOP_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol7_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 7th column.
        sshpSelectedArea = Index & ",6"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol7_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol7_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol8_TotalOP_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol8_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 8th column.
        sshpSelectedArea = Index & ",7"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol8_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol8_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol9_TotalOP_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol9_TotalOP(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 9th column.
        sshpSelectedArea = Index & ",8"
        SetShpTopLocation
    End If
End Sub

Private Sub lblCol9_TotalOP_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol9_TotalOP(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol1_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol1_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 1st column.
        sshpSelectedArea = Index & ",0"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol1_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol1_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol2_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol2_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 2nd column.
        sshpSelectedArea = Index & ",1"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol2_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol2_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol3_TotalOPResi_Click(Index As Integer)
   '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol3_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 3rd column.
        sshpSelectedArea = Index & ",2"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol3_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol3_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol4_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol4_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 4th column.
        sshpSelectedArea = Index & ",3"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol4_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol4_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol5_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol5_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 5th column.
        sshpSelectedArea = Index & ",4"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol5_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol5_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol6_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol6_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 6th column.
        sshpSelectedArea = Index & ",5"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol6_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol6_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol7_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol7_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 7th column.
        sshpSelectedArea = Index & ",6"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol7_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol7_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol8_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol8_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 8th column.
        sshpSelectedArea = Index & ",7"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol8_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol8_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol9_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol9_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 9th column.
        sshpSelectedArea = Index & ",8"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol9_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol9_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol10_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol10_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 10th column.
        sshpSelectedArea = Index & ",9"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol10_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol10_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblCol11_TotalOPResi_Click(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblCol11_TotalOPResi(Index).Caption <> "" Or bIsInitialLoad Then
        '
        '   We are in whatever row the value of index is and the 11th column.
        sshpSelectedArea = Index & ",10"
        SetShpTopLocationResi
    End If
End Sub

Private Sub lblCol11_TotalOPResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblCol11_TotalOPResi(Index).Caption <> "" Then
        GotoModel
    End If
End Sub

Private Sub lblFrame_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblFrame(Index).Caption <> "" Then
        '
        '   We are in whatever row the value of index is and the 1st column.
        sshpSelectedArea = Index & ",0"
        SetShpTopLocation
        GotoModel True
    End If
End Sub

Private Sub lblWall_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblWall(Index).Caption <> "" Then
        '
        '   We are in whatever row the value of index is and the 1st column.
        sshpSelectedArea = Index & ",0"
        SetShpTopLocation
        GotoModel True
    End If
End Sub

Private Sub lblWallResi_DblClick(Index As Integer)
    '
    '   If they click on a label that is not populated
    '   don't do anything.
    If lblWallResi(Index).Caption <> "" Then
        '
        '   We are in whatever row the value of index is and the 1st column.
        sshpSelectedArea = Index & ",0"
        SetShpTopLocationResi
        GotoModel True
    End If
End Sub

Private Sub cbobldg_categoryC_Click()
    bIsPendingChange = True
    cmdUpdate.Enabled = True
End Sub

Private Sub cbobldg_categoryR_Click()
    bIsPendingChange = True
    cmdUpdate.Enabled = True
    '
    '   Have to default the op_factor for Resi.
    If Trim(txtop_factor) = "" Then
        Select Case LCase(Trim(cbobldg_categoryR.Text))
            Case "economy"
                txtop_factor.Text = ".05"
            Case "average"
                txtop_factor.Text = ".10"
            Case "custom"
                txtop_factor.Text = ".15"
            Case "luxury"
                txtop_factor.Text = ".30"
        End Select
    End If
End Sub

Private Sub cboResiBldgType_Click()
    If opttype_codeR.Value = True Then
        ComputePerimeterResi True, (m_blnInsert = True And m_blnClone = False)
        SetupNewModelMatrix "R"
    End If
End Sub

Private Sub opttype_codeC_Click()
    
    On Error Resume Next
    EnableControls
    bIsPendingChange = True
    cmdUpdate.Enabled = True
    If m_blnInsert And Not m_blnClone Then
        cboRowToBold.Clear
        cboColumnToBold.Clear
        PopulateAvailWallTypesFrameTypes "C"
        SetupNewModelMatrix "C"
    End If
End Sub

Private Sub SetupNewModelMatrix(sTypeCode As String)
    Dim i As Integer
    
    On Error Resume Next
    
    If sTypeCode = "C" Then
        cboWallType(4).Visible = True
        cboWallType(5).Visible = True
        
        For i = 0 To 10
            txtNewBldgPerimeter(i).Visible = True
            txtNewBldgPerimeter(i).Locked = False
            txtNewBldgPerimeter(i).Text = ""
            txtNewBldgArea(i).Visible = True
        Next i
        For i = 9 To 10
            txtNewBldgArea(i).Visible = False
            txtNewBldgPerimeter(i).Visible = False
        Next i
        
        fraNewBldgModelMatrix.Height = 2715
        fraNewBldgModelMatrix.Width = 10000
        shpWhiteBackground.Height = 1980
        shpWhiteBackground.Width = 3675
    Else
        cboWallType(4).Visible = False
        cboWallType(5).Visible = False
        cboWallType(4).Text = ""
        cboWallType(5).Text = ""

        For i = 0 To 10
            txtNewBldgPerimeter(i).Visible = True
            txtNewBldgPerimeter(i).Locked = True
            txtNewBldgArea(i).Visible = True
            txtNewBldgArea(i).Locked = False
            txtNewBldgPerimeter(i).Text = ""
        Next i
        
        If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
            Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
            Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
        
            For i = 8 To 10
                txtNewBldgArea(i).Visible = False
                txtNewBldgPerimeter(i).Visible = False
            Next i
   
            fraNewBldgModelMatrix.Height = 2055
            fraNewBldgModelMatrix.Width = 9375
            shpWhiteBackground.Height = 1300
            shpWhiteBackground.Width = 3050
        Else
            fraNewBldgModelMatrix.Height = 2055
            fraNewBldgModelMatrix.Width = 11340
            shpWhiteBackground.Height = 1300
            shpWhiteBackground.Width = 5000
        End If
        ComputePerimeterResi True, True
    End If
End Sub

Private Sub opttype_codeR_Click()

    On Error Resume Next
    EnableControls
    bIsPendingChange = True
    cmdUpdate.Enabled = True
    If m_blnInsert And Not m_blnClone Then
        cboRowToBold.Clear
        cboColumnToBold.Clear
        PopulateAvailWallTypesFrameTypes "R"
        SetupNewModelMatrix "R"
    End If
End Sub

Private Sub optOpen_Click()
    PopulateModelMatrix
End Sub

Private Sub optUnion_Click()
    PopulateModelMatrix
End Sub

Private Sub cboColumnToBold_Click()
    Dim nArea As Integer
    Dim sArea As String
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        '
        'in the format of [1] area | perimeter
        sArea = Left$(Trim(Replace(cboColumnToBold.Text, "[", "")), InStr(1, Trim(Replace(cboColumnToBold.Text, "[", "")), "]") - 1)
    
        If sArea <> Trim(m_rec.Fields("col_to_bold").Value) Then
            bRefreshCosts = True
            bIsPendingChange = True
            cmdUpdate.Enabled = True
            '
            '   Force indicator that area_ind changed into
            '   txtarea tag so that updateArea is called.
            If opttype_codeC.Value = True Then
                Call txtArea_Change(CInt(sArea) - 1)
            Else
                Call txtAreaResi_Change(CInt(sArea) - 1)
            End If
        End If
    End If
End Sub

Private Sub cboRowToBold_Click()
    If Not bIsInitialLoad Then
        '
        '   Format of [model_code] wall | frame
        If Right(Left$(Trim(cboRowToBold.Text), 2), 1) <> Trim(m_rec.Fields("row_to_bold").Value) Then
            bIsPendingChange = True
            cmdUpdate.Enabled = True
        End If
    End If
End Sub

Private Sub txtgraphic_ref_id_GotFocus()
    HiliteTextBox txtgraphic_ref_id
End Sub

Private Sub txtgraphic_ref_id_LostFocus()
    txtgraphic_ref_id2.Text = RemoveCharacters(txtgraphic_ref_id2.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtgraphic_ref_id_Change()
    If Trim(txtgraphic_ref_id.Text) <> Trim(m_rec.Fields("graphic_ref_id").Value) Then
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtgraphic_ref_id2_GotFocus()
    HiliteTextBox txtgraphic_ref_id2
End Sub

Private Sub txtgraphic_ref_id2_LostFocus()
    txtgraphic_ref_id2.Text = RemoveCharacters(txtgraphic_ref_id2.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtgraphic_ref_id2_Change()
    If Trim(txtgraphic_ref_id2.Text) <> Trim(m_rec.Fields("graphic_ref_id2").Value) Then
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtop_factor_GotFocus()
    HiliteTextBox txtop_factor
End Sub

Private Sub txtop_factor_LostFocus()
    txtop_factor.Text = RemoveCharacters(txtop_factor.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtop_factor_Change()
    If Trim(txtop_factor.Text) <> Trim(m_rec.Fields("op_factor").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtarchitect_fee_GotFocus()
    HiliteTextBox txtarchitect_fee
End Sub

Private Sub txtarchitect_fee_LostFocus()
    txtarchitect_fee.Text = RemoveCharacters(txtarchitect_fee.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtarchitect_fee_Change()
    If Trim(txtarchitect_fee.Text) <> Trim(m_rec.Fields("architect_fee").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_door_density_GotFocus()
    HiliteTextBox txtbldg_door_density
End Sub

Private Sub txtbldg_door_density_LostFocus()
    txtbldg_door_density.Text = RemoveCharacters(txtbldg_door_density.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtbldg_door_density_Change()
    If Trim(txtbldg_door_density.Text) <> Trim(m_rec.Fields("bldg_door_density").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_elev_no_GotFocus()
    HiliteTextBox txtbldg_elev_no
End Sub

Private Sub txtbldg_elev_no_LostFocus()
    txtbldg_elev_no.Text = RemoveCharacters(txtbldg_elev_no.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtbldg_elev_no_Change()
    If Trim(txtbldg_elev_no.Text) <> Trim(m_rec.Fields("bldg_elev_no").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_fixture_area_GotFocus()
    HiliteTextBox txtbldg_fixture_area
End Sub

Private Sub txtbldg_fixture_area_LostFocus()
    txtbldg_fixture_area.Text = RemoveCharacters(txtbldg_fixture_area.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtbldg_fixture_area_Change()
    If Trim(txtbldg_fixture_area.Text) <> Trim(m_rec.Fields("bldg_fixture_area").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtwindow_area_GotFocus()
    HiliteTextBox txtwindow_area
End Sub

Private Sub txtwindow_area_LostFocus()
    txtwindow_area.Text = RemoveCharacters(txtwindow_area.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtwindow_area_Change()
    If Trim(txtwindow_area.Text) <> Trim(m_rec.Fields("window_area").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_part_density_GotFocus()
    HiliteTextBox txtbldg_part_density
End Sub

Private Sub txtbldg_part_density_LostFocus()
    txtbldg_part_density.Text = RemoveCharacters(txtbldg_part_density.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtbldg_part_density_Change()
    If Trim(txtbldg_part_density.Text) <> Trim(m_rec.Fields("bldg_part_density").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_part_hgt_GotFocus()
    HiliteTextBox txtbldg_part_hgt
End Sub

Private Sub txtbldg_part_hgt_LostFocus()
    txtbldg_part_hgt.Text = RemoveCharacters(txtbldg_part_hgt.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtbldg_part_hgt_Change()
    If Trim(txtbldg_part_hgt.Text) <> Trim(m_rec.Fields("bldg_part_hgt").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_stories_GotFocus()
    HiliteTextBox txtbldg_stories
End Sub

Private Sub txtbldg_stories_LostFocus()
    txtbldg_stories.Text = RemoveCharacters(txtbldg_stories.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtbldg_stories_Change()
    If Trim(txtbldg_stories.Text) <> Trim(m_rec.Fields("bldg_stories").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_stories_hgt_GotFocus()
    HiliteTextBox txtbldg_stories_hgt
End Sub

Private Sub txtbldg_stories_hgt_LostFocus()
    txtbldg_stories_hgt.Text = RemoveCharacters(txtbldg_stories_hgt.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtbldg_stories_hgt_Change()
    If Trim(txtbldg_stories_hgt.Text) <> Trim(m_rec.Fields("bldg_stories_hgt").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_wall_factor_GotFocus()
    HiliteTextBox txtbldg_wall_factor
End Sub

Private Sub txtbldg_wall_factor_LostFocus()
    txtbldg_wall_factor.Text = RemoveCharacters(txtbldg_wall_factor.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^()")
End Sub

Private Sub txtbldg_wall_factor_Change()
    If Trim(txtbldg_wall_factor.Text) <> Trim(m_rec.Fields("bldg_wall_factor").Value) Then
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_desc_GotFocus()
    HiliteTextBox txtbldg_desc
End Sub

Private Sub txtbldg_desc_LostFocus()
    txtbldg_desc.Text = RemoveCharacters(txtbldg_desc.Text, "\?/|[]{}*&><"":?,~`_=+@!#$%^()")
End Sub

Private Sub txtbldg_desc_Change()
    If Trim(txtbldg_desc.Text) <> Trim(m_rec.Fields("bldg_desc").Value) Then
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtbldg_id_GotFocus()
    HiliteTextBox txtbldg_id
End Sub

Private Sub txtbldg_id_LostFocus()
    txtbldg_id.Text = RemoveCharacters(txtbldg_id.Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^().")
End Sub

Private Sub txtbldg_id_Change()
    If Trim(txtbldg_id.Text) <> Trim(m_rec.Fields("bldg_id").Value) Then
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtArea_Change(Index As Integer)
    On Error Resume Next
    '
    '   Need to keep the orig area and the last_update_id ie- 15000|1 in the
    '   tag, but we only want to update those that have a C prefix for Changed.
    If Not bIsInitialLoad Then
        If Left$(txtArea(Index).Tag, 1) = "C" Then
            txtArea(Index).Tag = "C" & Right$(txtArea(Index).Tag, Len(txtArea(Index).Tag) - 1)
        Else
            txtArea(Index).Tag = "C" & txtArea(Index).Tag
        End If
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtArea_GotFocus(Index As Integer)
    HiliteTextBox txtArea(Index)
End Sub

Private Sub txtArea_LostFocus(Index As Integer)
    txtArea(Index).Text = RemoveCharacters(txtArea(Index).Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^().")
    If m_blnInsert = False And Trim(txtArea(Index).Text) <> "" Then
        RePopulateColToBold
    End If
End Sub

Private Sub txtAreaResi_Change(Index As Integer)
    On Error Resume Next
    If Not bIsInitialLoad Then
        If Left$(txtAreaResi(Index).Tag, 1) = "C" Then
            txtAreaResi(Index).Tag = "C" & Right$(txtAreaResi(Index).Tag, Len(txtAreaResi(Index).Tag) - 1)
        Else
            txtAreaResi(Index).Tag = "C" & txtAreaResi(Index).Tag
        End If
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtAreaResi_GotFocus(Index As Integer)
    HiliteTextBox txtAreaResi(Index)
End Sub

Private Sub txtAreaResi_LostFocus(Index As Integer)
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    txtAreaResi(Index).Text = RemoveCharacters(txtAreaResi(Index).Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^().")
    If m_blnInsert = False And Trim(txtAreaResi(Index).Text) <> "" Then
        If opttype_codeR.Value = True Then
            ComputePerimeterResi False, False, Index
        End If
        RePopulateColToBold
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub txtNewBldgArea_Change(Index As Integer)
    bIsPendingChange = True
    cmdUpdate.Enabled = True
End Sub

Private Sub txtNewBldgArea_GotFocus(Index As Integer)
    HiliteTextBox txtNewBldgArea(Index)
End Sub

Private Sub txtNewBldgArea_LostFocus(Index As Integer)
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    txtNewBldgArea(Index).Text = RemoveCharacters(txtNewBldgArea(Index).Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^().")
    If m_blnInsert And Trim(txtNewBldgArea(Index).Text) <> "" Then
        If opttype_codeR.Value = True Then
            ComputePerimeterResi False, True, Index
        End If
        RePopulateColToBold
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub ComputePerimeterResi(bAllPerimeters As Boolean, bInserting As Boolean, Optional Index As Integer)
    Dim i           As Integer
    Dim AreaTempA   As Double
    Dim AreaTempB   As Double

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    If bInserting Then
        If bAllPerimeters Then
            If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
                '
                '   Only 8 Areas for Wings & Ells.
                For i = 0 To 7
                    AreaTempA = Trim(txtNewBldgArea(i).Text) / 1.5
                    AreaTempA = Sqr(AreaTempA)
                    AreaTempB = AreaTempA * 1.5
                    txtNewBldgPerimeter(i).Text = RoundToNearest10(AreaTempA + AreaTempB * 2)
                Next i
            Else
                For i = 0 To 10
                    AreaTempA = Trim(txtNewBldgArea(i).Text) / 1.5
                    AreaTempA = Sqr(AreaTempA)
                    AreaTempB = AreaTempA * 1.5
                    txtNewBldgPerimeter(i).Text = RoundToNearest10(AreaTempA * 2 + AreaTempB * 2)
                Next i
            End If
        Else
            If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
            
                AreaTempA = Trim(txtNewBldgArea(Index).Text) / 1.5
                AreaTempA = Sqr(AreaTempA)
                AreaTempB = AreaTempA * 1.5
                txtNewBldgPerimeter(Index).Text = RoundToNearest10(AreaTempA + AreaTempB * 2)
            Else
                AreaTempA = Trim(txtNewBldgArea(Index).Text) / 1.5
                AreaTempA = Sqr(AreaTempA)
                AreaTempB = AreaTempA * 1.5
                txtNewBldgPerimeter(Index).Text = RoundToNearest10(AreaTempA * 2 + AreaTempB * 2)
            End If
        End If
    Else
        If bAllPerimeters Then
            If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
                '
                '   Only 8 Areas for Wings & Ells.
                For i = 0 To 7
                    AreaTempA = Trim(txtAreaResi(i).Text) / 1.5
                    AreaTempA = Sqr(AreaTempA)
                    AreaTempB = AreaTempA * 1.5
                    txtPerimeterResi(i).Text = RoundToNearest10(AreaTempA + AreaTempB * 2)
                Next i
                
            Else
                For i = 0 To 10
                    AreaTempA = Trim(txtAreaResi(i).Text) / 1.5
                    AreaTempA = Sqr(AreaTempA)
                    AreaTempB = AreaTempA * 1.5
                    txtPerimeterResi(i).Text = RoundToNearest10(AreaTempA * 2 + AreaTempB * 2)
                Next i
            End If
        Else
            If Left$(Trim(cboResiBldgType.Text), 1) = "H" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "I" Or _
                Left$(Trim(cboResiBldgType.Text), 1) = "J" Then
            
                AreaTempA = Trim(txtAreaResi(Index).Text) / 1.5
                AreaTempA = Sqr(AreaTempA)
                AreaTempB = AreaTempA * 1.5
                txtPerimeterResi(Index).Text = RoundToNearest10(AreaTempA + AreaTempB * 2)
            Else
                AreaTempA = Trim(txtAreaResi(Index).Text) / 1.5
                AreaTempA = Sqr(AreaTempA)
                AreaTempB = AreaTempA * 1.5
                txtPerimeterResi(Index).Text = RoundToNearest10(AreaTempA * 2 + AreaTempB * 2)
            End If
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Function RoundToNearest10(sValue As String) As String
    Dim nPos1   As String
    Dim nPos2   As String
    Dim nPos3   As String
    Dim nPos4   As String
    
    On Error Resume Next
    sValue = Round(Trim(sValue), 0)
    
    RoundToNearest10 = sValue
    nPos1 = Left$(sValue, 1)
    Select Case Len(sValue)
        Case 1
            Select Case nPos1
                Case "1" To "9"
                    RoundToNearest10 = "10"
            End Select
            
        Case 2
            nPos2 = Right$(Left$(sValue, 2), 1)
            Select Case nPos2
                Case "1" To "9"
                    If nPos1 = "9" Then
                        RoundToNearest10 = "100"
                    Else
                        RoundToNearest10 = nPos1 + 1 & "0"
                    End If
            End Select
        
        Case 3
            nPos2 = Right$(Left$(sValue, 2), 1)
            nPos3 = Right$(Right$(Left$(sValue, 3), 2), 1)

            Select Case nPos3
                Case "1" To "9"
                    If nPos2 = "9" Then
                        If nPos1 = "9" Then
                            RoundToNearest10 = "1000"
                        Else
                            RoundToNearest10 = nPos1 + 1 & "00"
                        End If
                    Else
                        RoundToNearest10 = nPos1 & nPos2 + 1 & "0"
                    End If
            End Select
        Case 4
            nPos2 = Right$(Left$(sValue, 2), 1)
            nPos3 = Right$(Right$(Left$(sValue, 3), 2), 1)
            nPos4 = Right$(sValue, 1)
            
            Select Case nPos4
                Case "1" To "9"
                    If nPos3 = "9" Then
                        If nPos2 = "9" Then
                            If nPos1 = "9" Then
                                RoundToNearest10 = "10000"
                            Else
                                RoundToNearest10 = nPos1 + 1 & "000"
                            End If
                        Else
                            RoundToNearest10 = nPos1 & nPos2 + 1 & "00"
                        End If
                    Else
                        RoundToNearest10 = nPos1 & nPos2 & nPos3 + 1 & "0"
                    End If
            End Select
    End Select
End Function

Private Sub txtNewBldgPerimeter_Change(Index As Integer)
    bIsPendingChange = True
    cmdUpdate.Enabled = True
End Sub

Private Sub txtNewBldgPerimeter_GotFocus(Index As Integer)
    HiliteTextBox txtNewBldgPerimeter(Index)
End Sub

Private Sub txtNewBldgPerimeter_LostFocus(Index As Integer)
    txtNewBldgPerimeter(Index).Text = RemoveCharacters(txtNewBldgPerimeter(Index).Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^().")
    If m_blnInsert And Trim(txtNewBldgPerimeter(Index).Text) <> "" Then
        RePopulateColToBold
    End If
End Sub

Private Sub txtPerimeter_Change(Index As Integer)
    On Error Resume Next
    If Not bIsInitialLoad Then
        If Left$(txtPerimeter(Index).Tag, 1) = "C" Then
            txtPerimeter(Index).Tag = "C" & Right$(txtPerimeter(Index).Tag, Len(txtPerimeter(Index).Tag) - 1)
        Else
            txtPerimeter(Index).Tag = "C" & txtPerimeter(Index).Tag
        End If
        bRefreshCosts = True
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtPerimeter_GotFocus(Index As Integer)
    HiliteTextBox txtPerimeter(Index)
End Sub

Private Sub txtPerimeter_LostFocus(Index As Integer)
    If m_blnInsert And Trim(txtPerimeter(Index).Text) = "" Then
    Else
        txtPerimeter(Index).Text = RemoveCharacters(txtPerimeter(Index).Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^().")
        RePopulateColToBold
    End If
End Sub

Private Sub cboFrameType_LostFocus(Index As Integer)
    cboFrameType(Index).Text = RemoveCharacters(cboFrameType(Index).Text, "\?/|[]{}*&><"":;?,~`_'-=+@!#$%^().")
    PopulateNewBldgRowToBold
End Sub

Private Sub cboFrameType_Click(Index As Integer)
    bIsPendingChange = True
    cmdUpdate.Enabled = True
End Sub

Private Sub cboFrameType_Change(Index As Integer)
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub cboWallType_LostFocus(Index As Integer)
    
    On Error Resume Next
    cboWallType(Index).Text = RemoveCharacters(cboWallType(Index).Text, "\?/|[]{}*&"":;?,~`_'-=+@!#$%^().")
    PopulateNewBldgRowToBold
End Sub

Private Sub cboWallType_Click(Index As Integer)
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub cboWallType_Change(Index As Integer)
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

'************************
Private Sub TDBGridAdds_GotFocus()
    TDBGridAdds.TabStop = True
End Sub

Private Sub TDBGridAdds_LostFocus()
    TDBGridAdds.TabStop = False
End Sub

Private Sub TDBGridAdds_KeyUp(KeyCode As Integer, Shift As Integer)
    EnableControls
End Sub

Private Sub TDBGridAdds_AfterDelete()
    EnableControls
End Sub

Private Sub TDBGridAdds_AfterInsert()
    bIsPendingChange = True
    cmdUpdate.Enabled = True
    EnableControls
End Sub

Private Sub TDBGridAdds_DblClick()
    If TDBGridAdds.Columns("Skey Type").Value = "A" Then
        cmdAssemblyCost_Click
    ElseIf TDBGridAdds.Columns("Skey Type").Value = "U" Then
        cmdUnitCost_Click
    End If
End Sub

Private Sub TDBGridAdds_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGridAdds.DataChanged = False
End Sub

Private Sub TDBGridAdds_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With TDBGridAdds
        If Button = vbRightButton And IsNumeric(.Bookmark) Then
            If Len(m_objComAddsGridMap.GetError(.Bookmark)) > 0 Then
                MsgBox m_objComAddsGridMap.GetError(.Bookmark)
            End If
        End If
    End With
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGridAdds_UnboundAddData(ByVal RowBuf As TrueOleDBGrid70.RowBuffer, NewRowBookmark As Variant)
Private Sub TDBGridAdds_UnboundAddData(ByVal RowBuf As TrueOleDBGrid80.RowBuffer, NewRowBookmark As Variant)
    lblComAddsRowCount.Caption = TDBGridAdds.ApproxCount + 1 & " rows."
End Sub

Private Sub TDBGridAdds_UnboundDeleteRow(Bookmark As Variant)
    lblComAddsRowCount.Caption = TDBGridAdds.ApproxCount - 1 & " rows."
End Sub

Private Sub TDBGridAdds_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If Not bIsInitialLoad Then
        With TDBGridAdds
            If .ApproxCount <> 0 Then
                If .Columns(ColIndex).Caption = "Book Desc" Then
                    .Columns(ColIndex).Value = Trim(.Columns(ColIndex).Value)
                End If
            End If
            .Columns("Format Code").Value = UCase(Trim(.Columns("Format Code").Value))
        End With
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub TDBGridAdds_AfterColUpdate(ByVal ColIndex As Integer)
    bIsPendingChange = True
    cmdUpdate.Enabled = True
    EnableControls
End Sub

Public Sub PrintReport()
    cmdReports_Click
End Sub

Public Sub PreviewReport()
    cmdReports_Click
End Sub

Public Sub ShowPrintToolbar(ByVal bShowIcons As Boolean)

    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
    fMainForm.mnuFilePageSetup.Enabled = bShowIcons
    fMainForm.mnuFilePrint.Enabled = bShowIcons
    fMainForm.mnuFilePrintPreview.Enabled = bShowIcons

End Sub
