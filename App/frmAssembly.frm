VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAssembly 
   Caption         =   "Assembly Maintenance"
   ClientHeight    =   6360
   ClientLeft      =   1965
   ClientTop       =   480
   ClientWidth     =   10680
   Icon            =   "frmAssembly.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   10680
   Begin VB.PictureBox picUCUsage 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   10680
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3210
      Width           =   10680
      Begin VB.CommandButton cmdMatUsageDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   2020
         Width           =   1150
      End
      Begin VB.ListBox lstValidate 
         Height          =   255
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   109
         Top             =   2040
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Frame fraAssemblyUnitCostUsage 
         Caption         =   "Assembly Unit Cost Usage"
         Height          =   2475
         Left            =   120
         TabIndex        =   31
         Top             =   0
         Width           =   10455
         Begin TrueOleDBGrid80.TDBGrid TDBGrid 
            Height          =   1695
            Left            =   120
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   10200
            _ExtentX        =   17992
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
            DeadAreaBackColor=   -2147483636
            RowDividerColor =   -2147483632
            RowSubDividerColor=   -2147483632
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
      End
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10680
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5745
      Width           =   10680
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   495
         Left            =   7380
         TabIndex        =   81
         Top             =   60
         Width           =   1150
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   8820
         TabIndex        =   83
         Top             =   60
         Visible         =   0   'False
         Width           =   1150
      End
      Begin VB.TextBox last_update_person 
         BackColor       =   &H8000000F&
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "S"
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox last_update_date 
         BackColor       =   &H8000000F&
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox assembly_skey 
         BackColor       =   &H8000000F&
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "1N"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "By:"
         Height          =   255
         Left            =   3180
         TabIndex        =   29
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Updated:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Skey:"
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   180
         Width           =   495
      End
   End
   Begin VB.TextBox opn_change_ind 
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
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   22
      Tag             =   "1N"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox std_change_ind 
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
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   21
      Tag             =   "1N"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox ad_change_ind 
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
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   20
      Tag             =   "1N"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox rr_change_ind 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   19
      Tag             =   "1N"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox std_last_update_id 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   18
      Tag             =   "2N"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox opn_last_update_id 
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
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   17
      Tag             =   "2N"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox rr_last_update_id 
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
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   16
      Tag             =   "4N"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox ad_last_update_id 
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
      Left            =   7860
      Locked          =   -1  'True
      TabIndex        =   0
      Tag             =   "1N"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   10680
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   0
      Width           =   10680
      Begin VB.ComboBox prsv_maint_const_cd 
         Height          =   315
         ItemData        =   "frmAssembly.frx":0442
         Left            =   9120
         List            =   "frmAssembly.frx":0452
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "1S"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox rev_uni2_L6 
         Height          =   315
         Left            =   5880
         TabIndex        =   113
         Tag             =   "1S"
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox rev_uni2_L5 
         Height          =   315
         Left            =   6000
         TabIndex        =   112
         Tag             =   "1S"
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox rev_uni2_L3 
         Height          =   315
         Left            =   6120
         TabIndex        =   111
         Tag             =   "1S"
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox alt_assembly_id 
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
         Left            =   4320
         TabIndex        =   2
         Tag             =   "1S"
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox assembly_id 
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
         Left            =   1200
         TabIndex        =   1
         Tag             =   "1S"
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox type_code 
         Height          =   315
         ItemData        =   "frmAssembly.frx":0475
         Left            =   6720
         List            =   "frmAssembly.frx":047F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "1S"
         Top             =   120
         Width           =   735
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4471
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Assembly"
         TabPicture(0)   =   "frmAssembly.frx":0489
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "labor_equip_ind"
         Tab(0).Control(1)=   "metric_book_desc"
         Tab(0).Control(2)=   "metric_tech_desc"
         Tab(0).Control(3)=   "tech_desc"
         Tab(0).Control(4)=   "book_desc"
         Tab(0).Control(5)=   "unit"
         Tab(0).Control(6)=   "coml_ind"
         Tab(0).Control(7)=   "resi_ind"
         Tab(0).Control(8)=   "metric_unit"
         Tab(0).Control(9)=   "comment"
         Tab(0).Control(10)=   "Label45"
         Tab(0).Control(11)=   "Label46"
         Tab(0).Control(12)=   "Label41"
         Tab(0).Control(13)=   "Label42"
         Tab(0).Control(14)=   "Label30"
         Tab(0).Control(15)=   "Label28"
         Tab(0).Control(16)=   "Label20"
         Tab(0).ControlCount=   17
         TabCaption(1)   =   "Costs"
         TabPicture(1)   =   "frmAssembly.frx":04A5
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label63"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label64"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label65"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label66"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label67"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label68"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label69"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label70"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label71"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label72"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label73"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label74"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Label75"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Label76"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Line1"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "linLaborHours"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "lblLaborHours"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "linHeading"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "std_mat_cost"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "std_labor_cost"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "opn_mat_cost"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "opn_labor_cost"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "rr_mat_cost"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "rr_labor_cost"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "metric_mat_cost"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "metric_labor_cost"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "std_total_cost_op"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "opn_total_cost_op"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "rr_total_cost_op"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "metric_total_cost_op"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "std_mat_cost_op"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "std_labor_cost_op"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "std_equip_cost_op"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "opn_mat_cost_op"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).Control(36)=   "opn_labor_cost_op"
         Tab(1).Control(36).Enabled=   0   'False
         Tab(1).Control(37)=   "opn_equip_cost_op"
         Tab(1).Control(37).Enabled=   0   'False
         Tab(1).Control(38)=   "rr_mat_cost_op"
         Tab(1).Control(38).Enabled=   0   'False
         Tab(1).Control(39)=   "rr_labor_cost_op"
         Tab(1).Control(39).Enabled=   0   'False
         Tab(1).Control(40)=   "rr_equip_cost_op"
         Tab(1).Control(40).Enabled=   0   'False
         Tab(1).Control(41)=   "metric_mat_cost_op"
         Tab(1).Control(41).Enabled=   0   'False
         Tab(1).Control(42)=   "metric_labor_cost_op"
         Tab(1).Control(42).Enabled=   0   'False
         Tab(1).Control(43)=   "metric_equip_cost_op"
         Tab(1).Control(43).Enabled=   0   'False
         Tab(1).Control(44)=   "std_labor_hour"
         Tab(1).Control(44).Enabled=   0   'False
         Tab(1).Control(45)=   "opn_labor_hour"
         Tab(1).Control(45).Enabled=   0   'False
         Tab(1).Control(46)=   "rr_labor_hour"
         Tab(1).Control(46).Enabled=   0   'False
         Tab(1).Control(47)=   "metric_labor_hour"
         Tab(1).Control(47).Enabled=   0   'False
         Tab(1).Control(48)=   "std_equip_cost"
         Tab(1).Control(48).Enabled=   0   'False
         Tab(1).Control(49)=   "std_total_cost"
         Tab(1).Control(49).Enabled=   0   'False
         Tab(1).Control(50)=   "opn_equip_cost"
         Tab(1).Control(50).Enabled=   0   'False
         Tab(1).Control(51)=   "opn_total_cost"
         Tab(1).Control(51).Enabled=   0   'False
         Tab(1).Control(52)=   "rr_equip_cost"
         Tab(1).Control(52).Enabled=   0   'False
         Tab(1).Control(53)=   "rr_total_cost"
         Tab(1).Control(53).Enabled=   0   'False
         Tab(1).Control(54)=   "metric_equip_cost"
         Tab(1).Control(54).Enabled=   0   'False
         Tab(1).Control(55)=   "metric_total_cost"
         Tab(1).Control(55).Enabled=   0   'False
         Tab(1).Control(56)=   "std_inst_cost"
         Tab(1).Control(56).Enabled=   0   'False
         Tab(1).Control(57)=   "opn_inst_cost"
         Tab(1).Control(57).Enabled=   0   'False
         Tab(1).Control(58)=   "rr_inst_cost"
         Tab(1).Control(58).Enabled=   0   'False
         Tab(1).Control(59)=   "metric_inst_cost"
         Tab(1).Control(59).Enabled=   0   'False
         Tab(1).Control(60)=   "std_inst_cost_op"
         Tab(1).Control(60).Enabled=   0   'False
         Tab(1).Control(61)=   "opn_inst_cost_op"
         Tab(1).Control(61).Enabled=   0   'False
         Tab(1).Control(62)=   "rr_inst_cost_op"
         Tab(1).Control(62).Enabled=   0   'False
         Tab(1).Control(63)=   "metric_inst_cost_op"
         Tab(1).Control(63).Enabled=   0   'False
         Tab(1).Control(64)=   "pct_ind"
         Tab(1).Control(64).Enabled=   0   'False
         Tab(1).ControlCount=   65
         Begin VB.CheckBox labor_equip_ind 
            Caption         =   "Labor Equip Ind"
            Height          =   255
            Left            =   -74760
            TabIndex        =   110
            Tag             =   "1"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox metric_book_desc 
            Height          =   285
            Left            =   -72240
            MaxLength       =   75
            MultiLine       =   -1  'True
            TabIndex        =   14
            Tag             =   "1S"
            Top             =   1734
            Width           =   7335
         End
         Begin VB.TextBox metric_tech_desc 
            Height          =   285
            Left            =   -72240
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   13
            Tag             =   "1S"
            Top             =   1428
            Width           =   7335
         End
         Begin VB.TextBox tech_desc 
            Height          =   285
            Left            =   -72240
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   11
            Tag             =   "1S"
            Top             =   816
            Width           =   7335
         End
         Begin VB.TextBox book_desc 
            Height          =   285
            Left            =   -72240
            MaxLength       =   75
            MultiLine       =   -1  'True
            TabIndex        =   12
            Tag             =   "1S"
            Top             =   1122
            Width           =   7335
         End
         Begin VB.CheckBox pct_ind 
            Caption         =   "&Percent"
            Height          =   255
            Left            =   8280
            TabIndex        =   34
            Tag             =   "5S"
            Top             =   320
            Width           =   975
         End
         Begin VB.TextBox metric_inst_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   7605
            Locked          =   -1  'True
            TabIndex        =   76
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox rr_inst_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   7605
            Locked          =   -1  'True
            TabIndex        =   54
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox opn_inst_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   7605
            Locked          =   -1  'True
            TabIndex        =   65
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox std_inst_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   7605
            Locked          =   -1  'True
            TabIndex        =   43
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox metric_inst_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   3225
            Locked          =   -1  'True
            TabIndex        =   71
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox rr_inst_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   3225
            Locked          =   -1  'True
            TabIndex        =   49
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox opn_inst_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   3225
            Locked          =   -1  'True
            TabIndex        =   60
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox std_inst_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   3225
            Locked          =   -1  'True
            TabIndex        =   38
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox metric_total_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   72
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox metric_equip_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   70
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox rr_total_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   50
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox rr_equip_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   48
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox opn_total_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   61
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox opn_equip_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   59
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox std_total_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   39
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox std_equip_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   37
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox metric_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   9460
            Locked          =   -1  'True
            TabIndex        =   78
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox rr_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   9460
            Locked          =   -1  'True
            TabIndex        =   56
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox opn_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   9460
            Locked          =   -1  'True
            TabIndex        =   67
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox std_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   9460
            Locked          =   -1  'True
            TabIndex        =   45
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox metric_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   6735
            Locked          =   -1  'True
            TabIndex        =   75
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox metric_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   74
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox metric_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4980
            Locked          =   -1  'True
            TabIndex        =   73
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox rr_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   6735
            Locked          =   -1  'True
            TabIndex        =   53
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox rr_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   52
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox rr_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4980
            Locked          =   -1  'True
            TabIndex        =   51
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox opn_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   6735
            Locked          =   -1  'True
            TabIndex        =   64
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox opn_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   63
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox opn_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4980
            Locked          =   -1  'True
            TabIndex        =   62
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox std_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   6735
            Locked          =   -1  'True
            TabIndex        =   42
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox std_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   41
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox std_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   4980
            Locked          =   -1  'True
            TabIndex        =   40
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox metric_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   8490
            Locked          =   -1  'True
            TabIndex        =   77
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox rr_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   8490
            Locked          =   -1  'True
            TabIndex        =   55
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox opn_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   8490
            Locked          =   -1  'True
            TabIndex        =   66
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox std_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   8490
            Locked          =   -1  'True
            TabIndex        =   44
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox metric_labor_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   69
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox metric_mat_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   68
            Tag             =   "2G"
            Top             =   2040
            Width           =   795
         End
         Begin VB.TextBox rr_labor_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   47
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox rr_mat_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   46
            Tag             =   "3G"
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox opn_labor_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   58
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox opn_mat_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   57
            Tag             =   "4G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox std_labor_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   36
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox std_mat_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   35
            Tag             =   "2G"
            Top             =   960
            Width           =   795
         End
         Begin VB.ComboBox unit 
            Height          =   315
            Left            =   -72240
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "1S"
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox coml_ind 
            Caption         =   "&Commercial Use"
            Height          =   255
            Left            =   -67920
            TabIndex        =   9
            Tag             =   "1"
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox resi_ind 
            Caption         =   "&Residential Use"
            Height          =   255
            Left            =   -66320
            TabIndex        =   10
            Tag             =   "1"
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox metric_unit 
            Height          =   315
            Left            =   -69600
            TabIndex        =   8
            Tag             =   "1S"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox comment 
            Height          =   285
            Left            =   -72240
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   15
            Tag             =   "1S"
            Top             =   2040
            Width           =   7335
         End
         Begin VB.Line linHeading 
            X1              =   600
            X2              =   9360
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric Long Desc:"
            Height          =   255
            Left            =   -74160
            TabIndex        =   105
            Top             =   1440
            Width           =   1755
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric Book Desc:"
            Height          =   255
            Left            =   -74160
            TabIndex        =   104
            Top             =   1740
            Width           =   1755
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "Long Desc:"
            Height          =   255
            Left            =   -73920
            TabIndex        =   103
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "Book Desc:"
            Height          =   255
            Left            =   -73920
            TabIndex        =   102
            Top             =   1140
            Width           =   1515
         End
         Begin VB.Label lblLaborHours 
            Alignment       =   2  'Center
            Caption         =   "Labor Hours"
            Height          =   375
            Left            =   9465
            TabIndex        =   101
            Top             =   400
            Width           =   795
         End
         Begin VB.Line linLaborHours 
            X1              =   9360
            X2              =   9360
            Y1              =   1125
            Y2              =   2205
         End
         Begin VB.Line Line1 
            X1              =   4935
            X2              =   4935
            Y1              =   1050
            Y2              =   2130
         End
         Begin VB.Label Label76 
            Alignment       =   1  'Right Justify
            Caption         =   "Total"
            Height          =   255
            Left            =   8520
            TabIndex        =   100
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label75 
            Alignment       =   1  'Right Justify
            Caption         =   "Equipment"
            Height          =   255
            Left            =   6720
            TabIndex        =   99
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            Caption         =   "Labor"
            Height          =   255
            Left            =   5880
            TabIndex        =   98
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            Caption         =   "Material"
            Height          =   255
            Left            =   4920
            TabIndex        =   97
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label72 
            Alignment       =   2  'Center
            Caption         =   "Total"
            Height          =   255
            Left            =   4080
            TabIndex        =   96
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label71 
            Alignment       =   2  'Center
            Caption         =   "Equipment"
            Height          =   255
            Left            =   2280
            TabIndex        =   95
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            Caption         =   "Labor"
            Height          =   255
            Left            =   1560
            TabIndex        =   94
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Caption         =   "Material"
            Height          =   255
            Left            =   600
            TabIndex        =   93
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label68 
            Alignment       =   2  'Center
            Caption         =   "Overhead && Profit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   92
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label67 
            Caption         =   "Bare Costs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   91
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label66 
            Caption         =   "Metric"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "Open"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label64 
            Caption         =   "R&&R"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label63 
            Caption         =   "Std"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Install"
            Height          =   255
            Left            =   3120
            TabIndex        =   86
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Install"
            Height          =   255
            Left            =   7560
            TabIndex        =   85
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Unit:"
            Height          =   255
            Left            =   -72840
            TabIndex        =   84
            Top             =   540
            Width           =   435
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric Unit:"
            Height          =   255
            Left            =   -70620
            TabIndex        =   82
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Comment:"
            Height          =   255
            Left            =   -73320
            TabIndex        =   80
            Top             =   2040
            Width           =   915
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Prsv Maint Const:"
         Height          =   255
         Left            =   7680
         TabIndex        =   4
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Alt Assembly  ID:"
         Height          =   255
         Left            =   3000
         TabIndex        =   108
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Assembly ID:"
         Height          =   255
         Left            =   160
         TabIndex        =   107
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   255
         Left            =   6120
         TabIndex        =   106
         Top             =   180
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAssembly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <modulename> frmAssembly</modulename>
' <functionname>General (Main) </functionname>
'
' <summary>
' Upon clicking the "Assembly" button at the bottom of the "Assembly Maintenance Grid" or upon double-clicking an assembly "line" on the datagrid you will be presented
'with a "child" form with a "tab" control representing:
'1.  the fundamentals of the selected assembly
'2.  "costs" across (4) op codes:  STD, OPN, METRIC, RR
'
'(TABS)
'
'ASSEMBLY
'"   You can add/change "description" fields or "unit" fields on this tab
'COSTS
'    You can do any of (3) things on this tab:
'"   Review "bare costs" and "Overhead & profit" costs
'"   Add a new Unit Cost Line
'"   Delete an existing Unit Cost Line
'
'HELPER Class: CAsUCUsageMap.Cls
'</summary>
'
'<seealso>CAsUCUsageMap.cls</seealso>
'<seealso>frmAssemblyGrid.frm</seealso>
'
' <datastruct>m_objGridMap</datastruct>
'<datastruct>m_rec</datastruct>
'
'<storedprocedurename>sp_select_assembly</storedprocedurename>
'<storedprocedurename> sp_delete_assembly</storedprocedurename>
'<storedprocedurename>sp_update_assembly_driver </storedprocedurename>
'<storedprocedurename> </storedprocedurename>
'
'
'<returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
'<code>
'Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
'Dim m_blnClone As Boolean  'Indicate if clone is in progress
'Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
'</code>
' <code>
'exec sp_select_assembly @start_assembly_id='D10100000000', @end_assembly_id='D10109999999', @alt_assembly_id='%', @tech_desc='%', @assembly_type = 0
'</code>
' <code>
'exec sp_update_assembly_driver @type_code='M', @assembly_skey= 20776, @assembly_id='E10101100100', @alt_assembly_id='1112001100', @rev_uni2_L3='     ', @rev_uni2_L5='    ', @rev_uni2_L6='    ', @book_desc='Bank equipment,
'drive up window, drawer & mike, no glazing, economy', @metric_book_desc='', @tech_desc='Architectural equipment, bank equipment drive up window, drawer & mike, no glazing, economy', @metric_tech_desc='Architectural equipment,
'bank equipment drive up window, drawer and mike, no glazing, economy', @coml_ind= 1, @resi_ind= 0, @labor_equip_ind= 0, @comment='', @unit='Day', @metric_unit='Ea.', @ad_change_ind= 1, @std_mat_cost= 4950, @std_inst_cost= 660,
'@std_equip_cost= 0, @std_labor_cost= 660, @std_total_cost= 5610, @std_mat_cost_op= 5450, @std_inst_cost_op= 1200, @std_equip_cost_op= 0, @std_labor_cost_op= 1200, @std_total_cost_op= 6650, @std_labor_hour= 16, @std_change_ind= 0,
'@opn_mat_cost= 4950, @opn_inst_cost= 495, @opn_equip_cost= 0, @opn_labor_cost= 495, @opn_total_cost= 5445, @opn_mat_c
'st_op= 5450, @opn_inst_cost_op= 970, @opn_equip_cost_op= 0, @opn_labor_cost_op= 970, @opn_total_cost_op= 6420, @opn_labor_hour= 16, @opn_change_ind= 0, @rr_mat_cost= 4950, @rr_inst_cost= 660, @rr_equip_cost= 0, @rr_labor_cost= 660,
'@rr_total_cost= 5610, @rr_mat_cost_op= 5450, @rr_inst_cost_op= 1250, @rr_equip_cost_op= 0, @rr_labor_cost_op= 1250, @rr_total_cost_op= 6700, @rr_labor_hour= 16, @rr_change_ind= 0, @metric_mat_cost= 4950, @metric_inst_cost= 660,
'@metric_equip_cost= 0, @metric_labor_cost= 660, @metric_total_cost= 5610, @metric_mat_cost_op= 5450, @metric_inst_cost_op= 1200, @metric_equip_cost_op= 0, @metric_labor_cost_op= 1200, @metric_total_cost_op= 6650, @metric_labor_hour= 16,
'@pct_ind= 0, @ad_last_update_id= 3, @std_last_update_id= 1, @opn_last_update_id= 1, @rr_last_update_id= 1, @last_update_person='Hancockrl',  @update_unitcost_usage_ind=0, @cost_change_ind=0
'</code>
'<code>
'
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CAsUCUsageMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>



Dim m_rec As ADODB.RecordSet
Dim m_rec2 As New ADODB.RecordSet   'Unit Cost Usage grid
Dim m_blnRecFlag As Boolean ' True if a populated RecordSet was passed, then we show data
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
Dim m_blnClone As Boolean  'Indicate if clone is in progress
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
Dim strLast_assembly_id As String ' Holds last unit cost_id so we know if it changed
Dim m_objGridMap As New CAsUCUsageMap ' Class to handle grid
Dim m_recUsage As ADODB.RecordSet
Dim m_lngOriginalSkey  As Long
Dim m_blnSortReqd As Boolean
Public frmCallingForm As Form
'*** APEX Migration Utility Code Change ***
'Public tdbCols As TrueOleDBGrid60.Columns
'*** APEX Migration Utility Code Change ***
'Public tdbCols As TrueOleDBGrid70.Columns
Public tdbCols As TrueOleDBGrid80.Columns
'*** APEX Migration Utility Code Change ***
'Public myTDBGrid As TrueOleDBGrid60.TDBGrid
'*** APEX Migration Utility Code Change ***
'Public myTDBGrid As TrueOleDBGrid70.TDBGrid
Public myTDBGrid As TrueOleDBGrid80.TDBGrid
Dim tdbOldCols As Variant
Dim m_type_code As String
Dim m_iAssemblyType As Integer
Const COMMERCIAL_ASSEMBLIES = 0
Const RESIDENTIAL_ASSEMBLIES = 1

Private Function ValidGridRow() As Boolean

'If Len(Trim(TDBGrid.Columns("unit cost id"))) = 0 Then  'RLH 02/04/2009
If Len(Trim(TDBGrid.Columns("unit cost id 04"))) = 0 Then 'rlh 02/04/2009
    MsgBox "Both the Assembly and Unit Cost IDs must be entered."
    TDBGrid.SetFocus
    ValidGridRow = False
Else
    If AssemblyUCSortRequired(assembly_skey) = True And Len(Trim(TDBGrid.Columns("Sort"))) = 0 Then
        TDBGrid.SetFocus
        ValidGridRow = False
        MsgBox "The Sort Order is required."
    Else
        ValidGridRow = True
    End If
End If

End Function


Private Sub ExecUpdate(strUpdate As String)
Dim blnRet As Boolean
Dim strError As String
Dim rec As ADODB.RecordSet
Dim strSelect1 As String
Dim strSelect As String

'Update the database with the current update sql string.
'If the update fails, display a message, otherwise increment the last update Id

On Error Resume Next
    blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)  'Update the Unit Cost Usage for assembly
    If blnRet = False Then
        MsgBox strError
        m_blnWereErrors = True
    Else
        'Select the last update id from the appropriate table if it was updated - required due to close of prior
        'record and creation of new one
        If ad_change_ind = True Then
            ad_last_update_id.Text = CInt(ad_last_update_id.Text) + 1
            m_rec.Fields("ad_last_update_id").Value = ad_last_update_id.Text
        End If
        If type_code = "E" Then
            strSelect1 = "select last_update_id from published_assembly_exception "
        Else    'must be M
            strSelect1 = "select last_update_id from published_assembly_cost "
        End If
        strSelect1 = strSelect1 + " where assembly_skey = " + _
            assembly_skey.Text + " and start_date = '" + Format(Now(), "yyyy") + _
            "-" + Format(Now(), "mm") + "-" + Format(Now(), "dd") + _
            "' and country_code = 'USA' and region_code = 'NAT' and op_code = '"
        If std_change_ind = True Then
            strSelect = strSelect1 + "STD'"
            blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
            If blnRet = True Then
                std_last_update_id.Text = rec.Fields("last_update_id")
                m_rec.Fields("std_last_update_id").Value = std_last_update_id.Text
            End If
            rec.Close
        End If
        
        If opn_change_ind = True Then
            strSelect = strSelect1 + "OPN'"
            blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
            If blnRet = True Then
                opn_last_update_id.Text = rec.Fields("last_update_id")
                m_rec.Fields("opn_last_update_id").Value = opn_last_update_id.Text
                rec.Close
            End If
        End If
        
        If rr_change_ind = True Then
            strSelect = strSelect1 + "RR'"
            blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
            If blnRet = True Then
                rr_last_update_id.Text = rec.Fields("last_update_id")
                m_rec.Fields("rr_last_update_id").Value = rr_last_update_id.Text
                rec.Close
            End If
        End If
        Set rec = Nothing
    End If

End Sub

Public Sub PrintReport()

End Sub

Public Sub PreviewReport()

End Sub

Private Sub RebindTDBGridNow()
    Dim oldRow As Variant
    oldRow = myTDBGrid.Bookmark
    myTDBGrid.ReBind
    myTDBGrid.Bookmark = oldRow
End Sub

Private Sub format_costs()
std_equip_cost = Format(std_equip_cost, "##,##0.00")
std_equip_cost_op = Format(std_equip_cost_op, "##,##0.00")
std_labor_cost = Format(std_labor_cost, "##,##0.00")
std_labor_cost_op = Format(std_labor_cost_op, "##,##0.00")
std_mat_cost = Format(std_mat_cost, "##,##0.00")
std_mat_cost_op = Format(std_mat_cost_op, "##,##0.00")
std_total_cost = Format(std_total_cost, "##,##0.00")
std_total_cost_op = Format(std_total_cost_op, "##,##0.00")
opn_equip_cost = Format(opn_equip_cost, "##,##0.00")
opn_equip_cost_op = Format(opn_equip_cost_op, "##,##0.00")
opn_labor_cost = Format(opn_labor_cost, "##,##0.00")
opn_labor_cost_op = Format(opn_labor_cost_op, "##,##0.00")
opn_mat_cost = Format(opn_mat_cost, "##,##0.00")
opn_mat_cost_op = Format(opn_mat_cost_op, "##,##0.00")
'rlh  10/22/08 (added formatting)
opn_inst_cost = Format(opn_inst_cost, "##,##0.00")
opn_inst_cost_op = Format(opn_inst_cost_op, "##,##0.00")
'end of rlh
opn_total_cost = Format(opn_total_cost, "##,##0.00")
opn_total_cost_op = Format(opn_total_cost_op, "##,##0.00")
rr_equip_cost = Format(rr_equip_cost, "##,##0.00")
rr_equip_cost_op = Format(rr_equip_cost_op, "##,##0.00")
rr_labor_cost = Format(rr_labor_cost, "##,##0.00")
rr_labor_cost_op = Format(rr_labor_cost_op, "##,##0.00")
rr_mat_cost = Format(rr_mat_cost, "##,##0.00")
'rlh 10/22/08  (added formatting)
rr_inst_cost_op = Format(rr_inst_cost_op, "##,##0.00")
rr_inst_cost = Format(rr_inst_cost, "##,##0.00")
'end of rlh
rr_mat_cost_op = Format(rr_mat_cost_op, "##,##0.00")
rr_total_cost = Format(rr_total_cost, "##,##0.00")
rr_total_cost_op = Format(rr_total_cost_op, "##,##0.00")
metric_equip_cost = Format(metric_equip_cost, "##,##0.00")
metric_equip_cost_op = Format(metric_equip_cost_op, "##,##0.00")
metric_labor_cost = Format(metric_labor_cost, "##,##0.00")
metric_labor_cost_op = Format(metric_labor_cost_op, "##,##0.00")
metric_mat_cost = Format(metric_mat_cost, "##,##0.00")
metric_mat_cost_op = Format(metric_mat_cost_op, "##,##0.00")
metric_total_cost = Format(metric_total_cost, "##,##0.00")
metric_total_cost_op = Format(metric_total_cost_op, "##,##0.00")
End Sub

Private Sub m_rec_unformatfields()
m_rec.Fields("std_equip_cost") = Format(std_equip_cost, "####0.00")
m_rec.Fields("std_equip_cost_op") = Format(std_equip_cost_op, "####0.00")
m_rec.Fields("std_labor_cost") = Format(std_labor_cost, "####0.00")
m_rec.Fields("std_labor_cost_op") = Format(std_labor_cost_op, "####0.00")
m_rec.Fields("std_mat_cost") = Format(std_mat_cost, "####0.00")
m_rec.Fields("std_mat_cost_op") = Format(std_mat_cost_op, "####0.00")
m_rec.Fields("std_total_cost") = Format(std_total_cost, "####0.00")
m_rec.Fields("std_total_cost_op") = Format(std_total_cost_op, "####0.00")
m_rec.Fields("opn_equip_cost") = Format(opn_equip_cost, "####0.00")
m_rec.Fields("opn_equip_cost_op") = Format(opn_equip_cost_op, "####0.00")
m_rec.Fields("opn_labor_cost") = Format(opn_labor_cost, "####0.00")
m_rec.Fields("opn_labor_cost_op") = Format(opn_labor_cost_op, "####0.00")
m_rec.Fields("opn_mat_cost") = Format(opn_mat_cost, "####0.00")
m_rec.Fields("opn_mat_cost_op") = Format(opn_mat_cost_op, "####0.00")
m_rec.Fields("opn_total_cost") = Format(opn_total_cost, "####0.00")
m_rec.Fields("opn_total_cost_op") = Format(opn_total_cost_op, "####0.00")
m_rec.Fields("rr_equip_cost") = Format(rr_equip_cost, "####0.00")
m_rec.Fields("rr_equip_cost_op") = Format(rr_equip_cost_op, "####0.00")
m_rec.Fields("rr_labor_cost") = Format(rr_labor_cost, "####0.00")
m_rec.Fields("rr_labor_cost_op") = Format(rr_labor_cost_op, "####0.00")
m_rec.Fields("rr_mat_cost") = Format(rr_mat_cost, "####0.00")
m_rec.Fields("rr_mat_cost_op") = Format(rr_mat_cost_op, "####0.00")
m_rec.Fields("rr_total_cost") = Format(rr_total_cost, "####0.00")
m_rec.Fields("rr_total_cost_op") = Format(rr_total_cost_op, "####0.00")
m_rec.Fields("metric_equip_cost") = Format(metric_equip_cost, "####0.00")
m_rec.Fields("metric_equip_cost_op") = Format(metric_equip_cost_op, "##,##0.00")
m_rec.Fields("metric_labor_cost") = Format(metric_labor_cost, "####0.00")
m_rec.Fields("metric_labor_cost_op") = Format(metric_labor_cost_op, "####0.00")
m_rec.Fields("metric_mat_cost") = Format(metric_mat_cost, "####0.00")
m_rec.Fields("metric_mat_cost_op") = Format(metric_mat_cost_op, "####0.00")
m_rec.Fields("metric_total_cost") = Format(metric_total_cost, "####0.00")
m_rec.Fields("metric_total_cost_op") = Format(metric_total_cost_op, "####0.00")

End Sub

' Fills all fields with data
Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If Not m_rec.Fields("assembly_skey") = 0 Then
        m_blnRecFlag = True
    End If
End Sub

Private Sub alt_assembly_id_Validate(Cancel As Boolean)
    Dim bln_New As Boolean
    If m_blnInsert Or m_blnClone Then
        bln_New = True
    End If
    If alt_assembly_id <> "" Then
        If Invalid_Assembly_id_Format(alt_assembly_id, "alt_assembly_id", m_rec, bln_New, ConvertAssemblySkey(assembly_skey)) = True Then
            Cancel = True
        End If
    End If
End Sub

Private Sub assembly_id_LostFocus()
    m_objGridMap.AssemblyID = assembly_id
End Sub

Private Sub assembly_id_Validate(Cancel As Boolean)
    On Error Resume Next
    Dim bln_New As Boolean

    If m_blnInsert Or m_blnClone Then
        bln_New = True
    End If
    If assembly_id <> "" Then
        If Invalid_Assembly_id_Format(assembly_id, "assembly_id", m_rec, bln_New, ConvertAssemblySkey(assembly_skey)) = True Then
            Cancel = True
        Else
            If m_blnClone = False And m_blnInsert = True Then
                If InStr(1, UCase(assembly_id), "X") > 0 Then  'Resi
                    resi_ind.Value = 1
                    coml_ind.Value = 0
                Else
                    resi_ind.Value = 0
                    coml_ind.Value = 1
                End If
            End If
        End If
    End If

End Sub

Private Sub book_desc_Change()
    book_desc.Text = StripControlCharacters(book_desc.Text)
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String

    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbNo Then
        Exit Sub
    End If

    strUpdate = "exec sp_delete_assembly "
    strUpdate = strUpdate + "@assembly_skey=" + str(Me.Controls("assembly_skey")) + ","
    strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    
    blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
    Else
        MsgBox "Delete successful."
        m_rec.Delete
        RebindTDBGridNow
        m_blnDeleted = True
        Unload Me
    End If
End Sub

Private Sub Store_Grid_Old_Values()
    Dim i As Integer
    ReDim tdbOldCols(tdbCols.Count - 1)
    For i = 0 To tdbCols.Count - 1
        tdbOldCols(i) = tdbCols.Item(i).Value
    Next
End Sub

Private Sub RestoreGridValues()
    ' this restores the grid back to its positioin if the user did not choose to save
    Dim i As Integer
    On Error Resume Next
    If m_blnInsert = False Then
        For i = 1 To tdbCols.Count - 1
            If tdbCols.Item(i).Value <> tdbOldCols(i) Then
                tdbCols.Item(i).Value = tdbOldCols(i)
            End If
        Next i
        'myTDBGrid.RefetchRow
        myTDBGrid.DataChanged = False
        myTDBGrid.RefreshRow
        DoEvents
    End If
End Sub

Private Sub cmdMatUsageDelete_Click()
    
On Error Resume Next
    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbYes Then
        If TDBGrid.AddNewMode > 0 Then
            TDBGrid.ReBind
        Else
            TDBGrid.Delete
        End If
    End If
End Sub

Private Sub cmdUpdate_Click()
    Const ASSEMBLY_DETAIL_TABLE = 1
    Const ASSEMBLY_EXCEPTION_STD = 2
    Const ASSEMBLY_EXCEPTION_OPN = 3
    Const ASSEMBLY_EXCEPTION_RR = 4
    Const PERCENT = 5
    Dim blnRet As Boolean
    Dim blnUpdateAssembly As Boolean
    Dim ctr As Control
    Dim fld As ADODB.Field
    Dim rec As New ADODB.RecordSet
    Dim strError As String
    Dim strSelect As String
    Dim strUpdate As String
    Dim strSaveUpdate As String
    Dim intStart As Integer
    Dim varSaveBookmark As Variant
    Dim i As Integer

    On Error Resume Next
    m_blnWereErrors = False
    If TDBGrid.AddNewMode = dbgAddNewPending Or TDBGrid.DataChanged Then
        m_blnWereErrors = Not ValidGridRow()
        If m_blnWereErrors = False Then
            TDBGrid.Update
        End If
    End If
    If m_blnWereErrors = False Then
        m_blnWereErrors = CheckEntryErrors()
    End If
    If m_blnWereErrors = False Then
        Screen.MousePointer = vbHourglass
        TDBGrid.Update
        Dim recClone As ADODB.RecordSet
        Set recClone = m_rec.Clone
        recClone.AddNew
        UpdateRecordsetFromForm Me, recClone
        For Each fld In m_rec.Fields
            ' If the value changed
            If Not Trim(fld.Value) = Trim(recClone.Fields(fld.Name).Value) Or ((IsNull(fld.Value) Or Trim(fld.Value) = "") Xor (recClone.Fields(fld.Name).Value = "")) Then
                Set ctr = Nothing
                Set ctr = Me.Controls(fld.Name)
                If Not ctr Is Nothing Then
                    ' See what table the field is from
                    Select Case Left(Me.Controls(fld.Name).Tag, 1)
                        Case ASSEMBLY_DETAIL_TABLE
                            m_rec.Fields("ad_change_ind") = True
                            ad_change_ind = True
                        Case ASSEMBLY_EXCEPTION_STD
                            m_rec.Fields("std_change_ind") = True
                            std_change_ind = True
                        Case ASSEMBLY_EXCEPTION_OPN
                            m_rec.Fields("opn_change_ind") = True
                            opn_change_ind = True
                        Case ASSEMBLY_EXCEPTION_RR
                            m_rec.Fields("rr_change_ind") = True
                            rr_change_ind = True
                        Case PERCENT
                            m_rec.Fields("ad_change_ind") = True
                            m_rec.Fields("std_change_ind") = True
                            m_rec.Fields("opn_change_ind") = True
                            m_rec.Fields("rr_change_ind") = True
                            ad_change_ind = True
                            std_change_ind = True
                            opn_change_ind = True
                            rr_change_ind = True
                    End Select
                End If
            End If
        Next
        ' Undo the changes made by the UpdateRecordsetFromForm call above
        recClone.CancelUpdate
        recClone.Close
        Set recClone = Nothing
        If ad_change_ind = True Or std_change_ind = True Or opn_change_ind = True Or rr_change_ind = True _
            Or m_objGridMap.IsPendingChange Or (m_rec2.RecordCount > 0 And m_blnClone = True) Then
            If ad_last_update_id.Text = "" Then ad_last_update_id.Text = 0
            If std_last_update_id.Text = "" Then std_last_update_id.Text = 0
            If opn_last_update_id.Text = "" Then opn_last_update_id.Text = 0
            If rr_last_update_id.Text = "" Then rr_last_update_id.Text = 0
            strUpdate = "exec sp_update_assembly_driver "
            BuildStoredProcSQL Me, strUpdate, ASSEMBLY_DETAIL_TABLE, m_rec
            BuildStoredProcSQL Me, strUpdate, ASSEMBLY_EXCEPTION_STD, m_rec
            BuildStoredProcSQL Me, strUpdate, ASSEMBLY_EXCEPTION_OPN, m_rec
            BuildStoredProcSQL Me, strUpdate, ASSEMBLY_EXCEPTION_RR, m_rec
            BuildStoredProcSQL Me, strUpdate, PERCENT, m_rec
            strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
            strUpdate = strUpdate + " @last_update_person='" + strUserName + "', "
            strUpdate = strUpdate + " @cost_change_ind=0"
            strSaveUpdate = strUpdate
            strUpdate = strUpdate + ", @update_unitcost_usage_ind=0"
            m_blnWereErrors = False
            If m_blnClone = True Or m_blnInsert = True Or type_code <> "M" Then   'Need to get skey if adding or update <> M
                ExecUpdate strUpdate
                If m_blnWereErrors = False Then
                    strSelect = "select assembly_skey from assembly_detail where assembly_id = '" & assembly_id.Text & "'"
                    g_objDAL.GetRecordset vbNullString, strSelect, rec
                    If (rec.EOF And rec.BOF) Then
                        MsgBox "Record not added."
                        m_blnWereErrors = True
                        Exit Sub
                    Else
                        assembly_skey.Text = rec.Fields("assembly_skey").Value
                        blnUpdateAssembly = True
                    End If
                End If
            End If
            If type_code = "M" Or type_code = "E" Then
                    'Process changes or deletions
                If m_objGridMap.IsPendingChange Or (m_blnClone = False And m_blnInsert = False) Or (m_blnClone = True And m_rec2.RecordCount > 0) Then
                    'If cloning, update the assembly_skey in all records in the grid.
                    If m_blnClone = True Or m_blnInsert = True Then
                        m_objGridMap.AssemblySKey = assembly_skey.Text
                        If m_rec2.RecordCount > 0 Then
                            m_rec2.MoveFirst
                            Do Until m_rec2.EOF
                                m_rec2.Fields("parent_skey") = assembly_skey.Text
                                m_rec2.MoveNext
                            Loop
                        End If
                    End If
                    
                    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                    '(rlh) 02/05/2009
                    'THIS IS WHERE THE UNIT COST LINE/GRID ADDS, CHANGES, OR DELETES HAPPEN!!!
                    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                    blnUpdateAssembly = m_objGridMap.Update() 'RLH 02/05/2009 !!!!!!!!!!!!!!!!!!
                    
                    m_blnWereErrors = Not blnUpdateAssembly
    '                If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
                        'Update the parameter to trigger Unit Cost Usage cost update
                        strUpdate = "exec sp_update_assembly_driver "
                        BuildStoredProcSQL Me, strUpdate, ASSEMBLY_DETAIL_TABLE, m_rec
                        BuildStoredProcSQL Me, strUpdate, ASSEMBLY_EXCEPTION_STD, m_rec
                        BuildStoredProcSQL Me, strUpdate, ASSEMBLY_EXCEPTION_OPN, m_rec
                        BuildStoredProcSQL Me, strUpdate, ASSEMBLY_EXCEPTION_RR, m_rec
                        BuildStoredProcSQL Me, strUpdate, PERCENT, m_rec
                        
                        strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
                        strUpdate = strUpdate + " @last_update_person='" + strUserName + "', "
                        strUpdate = strUpdate + " @cost_change_ind=0"
                        strUpdate = strUpdate + ", @update_unitcost_usage_ind=1"
                        strUpdate = ReplaceSkey(strUpdate, assembly_skey.Text)
                        ExecUpdate strUpdate
                        If m_blnWereErrors = False Then
                            blnUpdateAssembly = True
                        End If
                   'End If
                End If
            End If
            m_blnClone = False  ' no longer cloning if we were
            If m_blnWereErrors = False Then
                ' Put latest data into source recordset
                UpdateRecordsetFromForm Me, m_rec
                m_rec_unformatfields
    
                If type_code = "M" Then 'And g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
                    ''Retrieve latest amounts from Unit Cost Usage changes.
                    strSelect = "exec sp_select_assembly @assembly_id='" + _
                    assembly_id + "', @tech_desc='%'," + _
                    " @assembly_type = " + CStr(m_iAssemblyType)
                    g_objDAL.GetRecordset vbNullString, strSelect, rec
                    If Not (rec.EOF And rec.BOF) Then
                        For i = 0 To rec.Fields.Count
                            m_rec.Fields(rec.Fields(i).Name) = rec.Fields(i).Value
                        Next i
                    End If
                End If
                UpdateFormFromRecordset Me, m_rec
            End If
            If m_blnWereErrors = False And blnUpdateAssembly = True Then
                MsgBox "Assembly Update successful."
            End If
            RebindTDBGridNow
            varSaveBookmark = TDBGrid.Bookmark
            TDBGrid.Refresh
            TDBGrid.Bookmark = varSaveBookmark
        Else
            MsgBox "You must modify a field before updating."
        End If
        Screen.MousePointer = vbNormal
    Else
        Screen.MousePointer = vbNormal
    End If

End Sub

Private Function ReplaceSkey(strString, strSkey As String) As String
Dim iStart As Integer
Dim iEnd As Integer
Dim strTemp As String

iStart = InStr(1, strString, "@assembly_skey=")
If iStart > 0 Then
    iEnd = InStr(iStart, strString, ",")
    strTemp = Left(strString, iStart + 14) + strSkey + Right(strString, Len(strString) - iEnd + 1)
    ReplaceSkey = strTemp
End If

End Function

Private Sub comment_Change()
    comment.Text = StripControlCharacters(comment.Text)
End Sub

Private Sub Form_Activate()
    OutputView False
End Sub

Private Sub Form_Initialize()
    m_blnInsert = False
    m_blnDeleted = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    
    Move START_LEFT, START_TOP ' , 10305, 8115
    Me.Height = 6510
    Me.Width = 10785
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    
    
    If Not m_rec.State = adStateClosed Then
        UpdateFormFromRecordset Me, m_rec
        format_costs
        g_objDAL.GetRecordset vbNullString, "select unit from unit_of_measure order by unit", rec
        While Not rec.EOF
            unit.AddItem (rec.Fields("unit").Value)
            If Trim(m_rec.Fields("unit")) = Trim(rec.Fields("unit").Value) Then
                unit.Text = unit.List(unit.NewIndex)
            End If
            metric_unit.AddItem (rec.Fields("unit").Value)
            If Trim(m_rec.Fields("metric_unit")) = Trim(rec.Fields("unit").Value) Then
                metric_unit.Text = metric_unit.List(metric_unit.NewIndex)
            End If
            rec.MoveNext
        Wend
        rec.Close
    End If
    ' If we are NOT inserting
    If m_blnInsert = False Then
        ' Lock fields that can't be changed
        assembly_id.Locked = True
        'LockField Me, "assembly_id"
        'assembly_id.BackColor = LTGREY
        Me.Caption = Me.Caption + " [" + m_rec.Fields("assembly_id").Value + "]"
    Else
        ' If we are inserting and not showing data
        ' Set some defaults
        assembly_skey.Text = ""
        If Not m_blnRecFlag Then
'            active_status_ind.Value = 1
            Me.Caption = Me.Caption + " [New]"
            ' ADDED 6/30/2005 RTD FOR VERSION 7.4.0 PER B. BALBONI
            Me.type_code.Text = "M"
        Else
            Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("assembly_id").Value + "]"
            m_blnClone = True
            m_lngOriginalSkey = m_rec.Fields("assembly_skey").Value
        End If
    End If
    If Not m_blnClone Then
        strLast_assembly_id = m_rec.Fields("assembly_id").Value
    End If
    m_type_code = type_code
    ' Make the form show the right fields based on current type_code
    type_code_LostFocus
    m_objGridMap.AssemblyType = type_code
    'Stop 'rlh
    FillUsageGrid
    If coml_ind = 1 Then
        m_iAssemblyType = COMMERCIAL_ASSEMBLIES
    Else
        If resi_ind = 1 Then
            m_iAssemblyType = RESIDENTIAL_ASSEMBLIES
        End If
    End If
    m_blnSortReqd = AssemblyUCSortRequired(assembly_skey)
    ColorLockedFields Me
    
End Sub

Private Sub FillUsageGrid()

    On Error GoTo Error_Processing
    Dim strSelect As String
    Dim blnReturn As Boolean
    
    Const strSQL1 = "Select cu.parent_skey, cu.unit_cost_skey, cu.skey_type, override_book_desc, " + _
    "cu.override_metric_book_desc, override_book_qty, cu.override_metric_book_qty, cu.sort_order, " + _
    "cu.usage_unit, cu.usage_unit_qty, cu.adj_factor, cu.usage_metric_unit, cu.usage_metric_unit_qty, " + _
    "cu.metric_adj_factor, cu.last_update_person, cu.last_update_date, cu.last_update_id, u.unit_cost_id, " + _
    "u.ext_unit_cost_id, tech_desc, metric_tech_desc, call_out_id " + _
    "from unit_cost_usage as cu, vw_unit_cost_detail as u " + _
    "where cu.skey_type = 'A' and cu.unit_cost_skey = u.unit_cost_skey and cu.parent_skey = "

        ' Retrieve data for the usage grid using the appropriate skey in the selection criteria:
        '   Clone:  use original Assembly Skey
        '   New:    use 0
        '   Change: use assembly skey

    If m_blnClone = True Then
        'If cloning, set skey to 0 after reading data
            strSelect = strSQL1 + str(m_lngOriginalSkey)
    Else
        If assembly_skey.Text = "" Then
            strSelect = strSQL1 + " 0"
        Else
            strSelect = strSQL1 + assembly_skey.Text
        End If
    End If
    strSelect = strSelect + " order by cu.sort_order, u.unit_cost_id"
    ' Use DAL to perform select
    m_rec2.Close
    'Stop 'rlh
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec2)
    If Not IsNumeric(assembly_skey.Text) Then
        assembly_skey.Text = 0
    End If
    m_objGridMap.AssemblySKey = CLng(assembly_skey.Text)
    m_objGridMap.AssemblyID = assembly_id
    If assembly_skey.Text = "0" And m_rec2.RecordCount > 0 Then
        m_rec2.MoveFirst
        Do Until m_rec2.EOF
            m_rec2.Fields("parent_skey") = 0
            m_rec2.MoveNext
        Loop
    End If

    m_objGridMap.RecordSet = m_rec2.Clone
    If m_blnClone = True Then
        blnReturn = m_objGridMap.SetRowStateNew
    Else
        blnReturn = m_objGridMap.SetRowStateNone
    End If
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec2.RecordCount
    
Exit_Sub:
    Exit Sub

Error_Processing:
    If Err = 3704 Then ' object closed, ignore
        Resume Next
    Else
        MsgBox Error$
        Resume Exit_Sub
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button
    Dim blnPendingChange As Boolean
    Dim fld As Control
    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        Set fld = Me.ActiveControl
        blnPendingChange = IsControlChanged(Me, m_rec)
        If blnPendingChange = True Or m_objGridMap.IsPendingChange Then
            Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
            If Button = vbYes Then
'                If TDBGrid.AddNewMode = dbgAddNewPending Or TDBGrid.DataChanged Then
'                    If ValidGridRow() = True Then
'                        TDBGrid.Update
'                        m_blnWereErrors = CheckEntryErrors()
'                    Else
'                        m_blnWereErrors = True
'                    End If
'                If m_blnWereErrors Then
'                    Cancel = True
'                    TDBGrid.Refresh
'                    fld.SetFocus
'                Else
                    cmdUpdate_Click
                    ' If there were errors, cancel the close
                    If m_blnWereErrors Then
                        Cancel = True
                        fld.SetFocus
                    Else
                        RestoreGridValues
                    End If
'                End If
            ElseIf Button = vbCancel Then
                Cancel = True
                Exit Sub
            ElseIf m_blnInsert = True Then
                m_rec.Delete
            End If
        End If
    End If
End Sub

Private Function CheckEntryErrors() As Boolean
    Dim bln_New As Boolean
    Dim i As Integer
    Dim strError As String
    Dim strItem As String
    Dim varBookmarks() As Variant
    
    CheckEntryErrors = False
    If m_blnInsert Or m_blnClone Then
        bln_New = True
    End If
    'rlh 05/02/2008  - Check on existing alt_assembly_id is no longer necessary (Kathy R. and Gary H.)
'    If alt_assembly_id <> "" Then
'        If Invalid_Assembly_id_Format(alt_assembly_id, "alt_assembly_id", m_rec, bln_New, ConvertAssemblySkey(assembly_skey)) = True Then
'            CheckEntryErrors = True
'            Exit Function
'        End If
'    End If
    If Invalid_Assembly_id_Format(assembly_id, "assembly_id", m_rec, _
        bln_New, ConvertAssemblySkey(assembly_skey)) = True And CheckEntryErrors = False Then
        CheckEntryErrors = True
        Exit Function
    End If
    ' ADDED 6/16/2005 RTD FOR VERSION 7.4.0 CR#1356
    If type_code.Text = "" Then
        MsgBox "A Type Code is required.", vbCritical
        CheckEntryErrors = True
        Exit Function
    End If
    If resi_ind.Value = 0 And coml_ind.Value = 0 And CheckEntryErrors = False Then
        MsgBox "The Commercial Use indicator or Residential Use indicator must be checked.", vbCritical
        CheckEntryErrors = True
        Exit Function
    End If
    If TDBGrid.AddNewMode = dbgAddNewPending Or TDBGrid.DataChanged Then
        If TDBGrid.Columns(TDBGrid.Col).Caption = "Unit Cost ID" Then
            strError = AsblyUCGridError_UnitCostID(Compress_String(TDBGrid.Text), assembly_id, _
            type_code)
        End If
        If strError <> Empty Then
            MsgBox strError
            CheckEntryErrors = True
        Else
            TDBGrid.Update
        End If
    End If
    
    If (m_rec2.RecordCount > 0) And (CheckEntryErrors = False) And (m_rec2.Fields("sort_order") > " ") Then
        If m_blnSortReqd = True Then
            m_rec2.MoveFirst
            Do Until m_rec2.EOF
               If Len(Trim(m_rec2.Fields("sort_order"))) = 0 Then
                    m_objGridMap.SetError m_rec2.Bookmark, "This assembly has an assembly book system line:  The sort order is required."
                    CheckEntryErrors = True
                End If
                m_rec2.MoveNext
            Loop
        End If
    End If

    '*
    ReDim varBookmarks(0 To m_rec2.RecordCount)

    'Validate a unique unit_cost_id/sort_order
    If (m_rec2.RecordCount > 0) And (CheckEntryErrors = False) Then
        lstValidate.Clear
        m_rec2.MoveFirst
        i = 0
        Do Until m_rec2.EOF
            strItem = ""
            If Not IsNull(m_rec2.Fields("unit_cost_id")) Then
                strItem = strItem + m_rec2.Fields("unit_cost_id")
            End If
            If Not IsNull(m_rec2.Fields("sort_order")) Then
                strItem = strItem + Trim(m_rec2.Fields("sort_order"))
            End If
            lstValidate.AddItem (strItem)
            varBookmarks(i) = m_rec2.Bookmark
            i = i + 1
            m_rec2.MoveNext
        Loop
        'Start at the second item, compare each to the prior (list is sorted)
        For i = 1 To lstValidate.listcount - 1
            If lstValidate.List(i) = lstValidate.List(i - 1) Then
                CheckEntryErrors = True
                MsgBox "The Unit Cost ID/Sort ID must be unique."
                m_rec2.Bookmark = varBookmarks(i - 1)
                m_objGridMap.SetError m_rec2.Bookmark, "The Unit Cost ID/Sort ID must be unique."
                TDBGrid.Bookmark = varBookmarks(i - 1)
                TDBGrid.RefetchRow (varBookmarks(i - 1))
                Exit For
            End If
        Next i
    End If
        
End Function

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.Height >= 6510 Then
        picUCUsage.Height = Me.Height - picTop.Height - picFooter.Height - 480
        fraAssemblyUnitCostUsage.Height = picUCUsage.Height - 200
        TDBGrid.Height = fraAssemblyUnitCostUsage.Height - 800
        cmdMatUsageDelete.Top = fraAssemblyUnitCostUsage.Top + fraAssemblyUnitCostUsage.Height - cmdMatUsageDelete.Height - 100
    Else
        picUCUsage.Height = Me.Height - picTop.Height - picFooter.Height - 480
    End If
    fraAssemblyUnitCostUsage.Width = Me.Width - (fraAssemblyUnitCostUsage.Left * 3)
    TDBGrid.Width = fraAssemblyUnitCostUsage.Width - (TDBGrid.Left * 2)
    ResizeForm Me

End Sub

Private Sub metric_book_desc_Change()
    metric_book_desc.Text = StripControlCharacters(metric_book_desc.Text)
End Sub

Private Sub metric_tech_desc_Change()
    metric_tech_desc.Text = StripControlCharacters(metric_tech_desc.Text)
End Sub

Private Sub tech_desc_Change()
' ADDED 6/16/2005 RTD FOR VERSION 7.4.0 CR#1544
    tech_desc.Text = StripControlCharacters(tech_desc.Text)
End Sub

Private Sub type_code_GotFocus()
    If type_code <> "M" Then
        m_type_code = type_code
    End If
End Sub

Private Sub type_code_LostFocus()
    Dim blnResult As Boolean
    On Error Resume Next
    If m_type_code <> type_code Then
        'reset last update indicators
        std_last_update_id = "0"
        opn_last_update_id = "0"
        rr_last_update_id = "0"
        If m_rec2.RecordCount > 0 Then
            m_rec2.MoveFirst
            Do Until m_rec2.EOF
                m_rec2.Delete
                m_rec2.MoveFirst
            Loop
        End If
        MsgBox "Make sure you attach at least one unit cost.", vbInformation
        m_type_code = type_code
        m_objGridMap.AssemblyType = type_code
        TDBGrid.ReBind
    End If

    If type_code.Text = "E" Then
'        fraAssemblyUnitCostUsage.Visible = False
        SSTab1.TabEnabled(1) = True
        pct_ind.Visible = True
    ElseIf type_code.Text = "M" Then
'        If m_type_code <> type_code Then
'        End If
'        FillUsageGrid
        SSTab1.TabEnabled(1) = True
'        fraAssemblyUnitCostUsage.Move 120, 3660
        fraAssemblyUnitCostUsage.Visible = True
        pct_ind.Visible = False
    Else
'        fraAssemblyUnitCostUsage.Visible = False
        If SSTab1.Tab = 1 Then
            SSTab1.Tab = 0
        End If
        SSTab1.TabEnabled(1) = False
    End If
    If type_code.Text = "E" Then
        blnResult = UnLockField(Me, "std_mat_cost")
        blnResult = UnLockField(Me, "std_labor_cost")
        blnResult = UnLockField(Me, "std_equip_cost")
        blnResult = UnLockField(Me, "std_total_cost")
        
        blnResult = UnLockField(Me, "std_mat_cost_op")
        blnResult = UnLockField(Me, "std_labor_cost_op")
        blnResult = UnLockField(Me, "std_equip_cost_op")
        blnResult = UnLockField(Me, "std_total_cost_op")
'        blnResult = UnLockField(Me, "std_labor_hour")
        
        blnResult = UnLockField(Me, "opn_mat_cost")
        blnResult = UnLockField(Me, "opn_labor_cost")
        blnResult = UnLockField(Me, "opn_equip_cost")
        blnResult = UnLockField(Me, "opn_total_cost")
        
        blnResult = UnLockField(Me, "opn_mat_cost_op")
        blnResult = UnLockField(Me, "opn_labor_cost_op")
        blnResult = UnLockField(Me, "opn_equip_cost_op")
        blnResult = UnLockField(Me, "opn_total_cost_op")
'        blnResult = UnLockField(Me, "opn_labor_hour")
        
        blnResult = UnLockField(Me, "rr_mat_cost")
        blnResult = UnLockField(Me, "rr_labor_cost")
        blnResult = UnLockField(Me, "rr_equip_cost")
        blnResult = UnLockField(Me, "rr_total_cost")
        
        blnResult = UnLockField(Me, "rr_mat_cost_op")
        blnResult = UnLockField(Me, "rr_labor_cost_op")
        blnResult = UnLockField(Me, "rr_equip_cost_op")
        blnResult = UnLockField(Me, "rr_total_cost_op")
'        blnResult = UnLockField(Me, "rr_labor_hour")
        
        blnResult = UnLockField(Me, "metric_mat_cost")
        blnResult = UnLockField(Me, "metric_labor_cost")
        blnResult = UnLockField(Me, "metric_equip_cost")
        blnResult = UnLockField(Me, "metric_total_cost")

        blnResult = UnLockField(Me, "metric_mat_cost_op")
        blnResult = UnLockField(Me, "metric_labor_cost_op")
        blnResult = UnLockField(Me, "metric_equip_cost_op")
        blnResult = UnLockField(Me, "metric_total_cost_op")
'        blnResult = UnLockField(Me, "metric_labor_hour")
        
        blnResult = UnLockField(Me, "std_inst_cost")
        blnResult = UnLockField(Me, "std_inst_cost_op")
        blnResult = UnLockField(Me, "opn_inst_cost")
        blnResult = UnLockField(Me, "opn_inst_cost_op")
        blnResult = UnLockField(Me, "rr_inst_cost")
        blnResult = UnLockField(Me, "rr_inst_cost_op")
        blnResult = UnLockField(Me, "metric_inst_cost")
        blnResult = UnLockField(Me, "metric_inst_cost_op")
        blnResult = UnLockField(Me, "metric_unit")

    'Hide hour fields -  N/A for E rows
        lblLaborHours.Visible = False
        std_labor_hour.Visible = False
        rr_labor_hour.Visible = False
        opn_labor_hour.Visible = False
        metric_labor_hour.Visible = False
        linLaborHours.Visible = False
'        linHeading.X2 = 9360
    Else
    'Show hour fields for all except E rows
        lblLaborHours.Visible = True
        std_labor_hour.Visible = True
        rr_labor_hour.Visible = True
        opn_labor_hour.Visible = True
        metric_labor_hour.Visible = True
        linLaborHours.Visible = True
'        linHeading.X2 = 10200
        
        blnResult = LockField(Me, "std_mat_cost")
        blnResult = LockField(Me, "std_labor_cost")
        blnResult = LockField(Me, "std_equip_cost")
        blnResult = LockField(Me, "std_total_cost")
        
        blnResult = LockField(Me, "std_mat_cost_op")
        blnResult = LockField(Me, "std_labor_cost_op")
        blnResult = LockField(Me, "std_equip_cost_op")
        blnResult = LockField(Me, "std_total_cost_op")
        blnResult = LockField(Me, "std_labor_hour")
        
        blnResult = LockField(Me, "opn_mat_cost")
        blnResult = LockField(Me, "opn_labor_cost")
        blnResult = LockField(Me, "opn_equip_cost")
        blnResult = LockField(Me, "opn_total_cost")
        
        blnResult = LockField(Me, "opn_mat_cost_op")
        blnResult = LockField(Me, "opn_labor_cost_op")
        blnResult = LockField(Me, "opn_equip_cost_op")
        blnResult = LockField(Me, "opn_total_cost_op")
        blnResult = LockField(Me, "opn_labor_hour")
        
        blnResult = LockField(Me, "rr_mat_cost")
        blnResult = LockField(Me, "rr_labor_cost")
        blnResult = LockField(Me, "rr_equip_cost")
        blnResult = LockField(Me, "rr_total_cost")
        
        blnResult = LockField(Me, "rr_mat_cost_op")
        blnResult = LockField(Me, "rr_labor_cost_op")
        blnResult = LockField(Me, "rr_equip_cost_op")
        blnResult = LockField(Me, "rr_total_cost_op")
        blnResult = LockField(Me, "rr_labor_hour")
        
        blnResult = LockField(Me, "metric_mat_cost")
        blnResult = LockField(Me, "metric_labor_cost")
        blnResult = LockField(Me, "metric_equip_cost")
        blnResult = LockField(Me, "metric_total_cost")

        blnResult = LockField(Me, "std_inst_cost")
        blnResult = LockField(Me, "std_inst_cost_op")
        blnResult = LockField(Me, "opn_inst_cost")
        blnResult = LockField(Me, "opn_inst_cost_op")
        blnResult = LockField(Me, "rr_inst_cost")
        blnResult = LockField(Me, "rr_inst_cost_op")
        blnResult = LockField(Me, "metric_inst_cost")
        blnResult = LockField(Me, "metric_inst_cost_op")

        blnResult = LockField(Me, "metric_mat_cost_op")
        blnResult = LockField(Me, "metric_labor_cost_op")
        blnResult = LockField(Me, "metric_equip_cost_op")
        blnResult = LockField(Me, "metric_total_cost_op")
        blnResult = LockField(Me, "metric_labor_hour")
        blnResult = LockField(Me, "metric_unit")
    End If

End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Dim strErrorMsg As String
        strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
        If Len(strErrorMsg) > 0 Then
            MsgBox strErrorMsg
        End If
    End If
End Sub

