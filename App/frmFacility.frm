VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBuilding 
   Caption         =   "Building Maintenance"
   ClientHeight    =   6780
   ClientLeft      =   1965
   ClientTop       =   495
   ClientWidth     =   9240
   Icon            =   "frmFacility.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   9240
   Begin VB.PictureBox picGrid 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3480
      Left            =   0
      ScaleHeight     =   3480
      ScaleWidth      =   9240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2760
      Width           =   9240
      Begin VB.Frame fraUnitCost 
         BackColor       =   &H00C0C000&
         Caption         =   "Common Additives"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3315
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   8895
         Begin VB.CommandButton cmdAdditiveDelete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   135
            TabIndex        =   139
            Top             =   2760
            Width           =   1150
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid 
            Height          =   2430
            Left            =   120
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   300
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4286
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
            PrintInfos(0)._StateFlags=   0
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
            DeadAreaBackColor=   8388608
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
            _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
            _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=31"
            _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
            _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
            _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
            _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
            _StyleDefs(16)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
            _StyleDefs(17)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
            _StyleDefs(18)  =   "Splits(0).Style:id=11,.parent=1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=13,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=15,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
            _StyleDefs(38)  =   "Named:id=29:Normal"
            _StyleDefs(39)  =   ":id=29,.parent=0"
            _StyleDefs(40)  =   "Named:id=30:Heading"
            _StyleDefs(41)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(42)  =   ":id=30,.wraptext=-1"
            _StyleDefs(43)  =   "Named:id=31:Footing"
            _StyleDefs(44)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(45)  =   "Named:id=32:Selected"
            _StyleDefs(46)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(47)  =   "Named:id=33:Caption"
            _StyleDefs(48)  =   ":id=33,.parent=30,.alignment=2"
            _StyleDefs(49)  =   "Named:id=34:HighlightRow"
            _StyleDefs(50)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(51)  =   "Named:id=35:EvenRow"
            _StyleDefs(52)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
            _StyleDefs(53)  =   "Named:id=36:OddRow"
            _StyleDefs(54)  =   ":id=36,.parent=29"
            _StyleDefs(55)  =   "Named:id=39:RecordSelector"
            _StyleDefs(56)  =   ":id=39,.parent=30"
            _StyleDefs(57)  =   "Named:id=42:FilterBar"
            _StyleDefs(58)  =   ":id=42,.parent=29"
         End
      End
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   9240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6210
      Width           =   9240
      Begin VB.TextBox col_to_bold 
         Height          =   285
         Left            =   3240
         TabIndex        =   142
         Tag             =   "1N"
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox type_code 
         Height          =   285
         Left            =   5160
         TabIndex        =   91
         Tag             =   "1S"
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   6696
         TabIndex        =   140
         Top             =   0
         Width           =   1150
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   7920
         TabIndex        =   141
         Top             =   0
         Width           =   1150
      End
      Begin VB.TextBox last_update_person 
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
         Height          =   315
         Left            =   3558
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "S"
         Top             =   60
         Width           =   1215
      End
      Begin VB.TextBox last_update_date 
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
         Height          =   315
         Left            =   806
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   60
         Width           =   1695
      End
      Begin VB.TextBox bldg_skey 
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
         Height          =   315
         Left            =   5410
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "1N"
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Updated By:"
         Height          =   255
         Left            =   2572
         TabIndex        =   28
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Updated:"
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Skey:"
         Height          =   255
         Left            =   4844
         TabIndex        =   26
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox crew_type_code 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox last_update_id 
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
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "0N"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   0
      ScaleHeight     =   2760
      ScaleWidth      =   9240
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   0
      Width           =   9240
      Begin VB.Frame Frame1 
         Caption         =   "Building Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   450
         Left            =   2760
         TabIndex        =   2
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton optBldg_Type 
            Caption         =   "Com'l"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optBldg_Type 
            Caption         =   "Resi"
            Height          =   255
            Index           =   1
            Left            =   825
            TabIndex        =   4
            Top             =   165
            Width           =   645
         End
      End
      Begin VB.TextBox bldg_id 
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
         Left            =   1080
         TabIndex        =   1
         Tag             =   "1S"
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox bldg_category 
         Height          =   315
         ItemData        =   "frmFacility.frx":0442
         Left            =   6135
         List            =   "frmFacility.frx":0452
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "1S"
         Top             =   75
         Width           =   1575
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2235
         Left            =   75
         TabIndex        =   6
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3942
         _Version        =   393216
         TabHeight       =   520
         ForeColor       =   8388608
         TabCaption(0)   =   "Building"
         TabPicture(0)   =   "frmFacility.frx":0462
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Shape1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label16"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label18"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label2"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label11"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label17"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label14"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label13"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label12"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label33"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label32"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label31"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label3"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label7"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label30"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "graphic_ref_id2"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "graphic_ref_id"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "bldg_fixture_area"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "bldg_area_std"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "bldg_part_hgt"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "bldg_part_density"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "op_factor"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "bldg_elevator_no"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "bldg_wall_factor"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "bldg_perimeter_std"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "bldg_stories_hgt"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "bldg_arch_fees"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "bldg_door_density"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "bldg_stories"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).ControlCount=   29
         TabCaption(1)   =   "Model Areas/Perimeters"
         TabPicture(1)   =   "frmFacility.frx":047E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "picResi"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "picCommercialArea"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Description"
         TabPicture(2)   =   "frmFacility.frx":049A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label41"
         Tab(2).Control(1)=   "bldg_desc"
         Tab(2).ControlCount=   2
         Begin VB.TextBox bldg_desc 
            Height          =   285
            Left            =   -73200
            MaxLength       =   75
            TabIndex        =   79
            Tag             =   "1S"
            Top             =   750
            Width           =   6915
         End
         Begin VB.TextBox bldg_stories 
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Tag             =   "1S"
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox bldg_door_density 
            Height          =   315
            Left            =   5880
            TabIndex        =   17
            Tag             =   "1S"
            Top             =   1740
            Width           =   735
         End
         Begin VB.TextBox bldg_arch_fees 
            Height          =   315
            Left            =   8040
            TabIndex        =   18
            Tag             =   "1S"
            Top             =   1785
            Width           =   735
         End
         Begin VB.TextBox bldg_stories_hgt 
            Height          =   315
            Left            =   1560
            TabIndex        =   11
            Tag             =   "1S"
            Top             =   1305
            Width           =   735
         End
         Begin VB.TextBox bldg_perimeter_std 
            Height          =   315
            Left            =   6720
            TabIndex        =   35
            Tag             =   "1S"
            Top             =   390
            Width           =   735
         End
         Begin VB.TextBox bldg_wall_factor 
            Height          =   315
            Left            =   8040
            TabIndex        =   10
            Tag             =   "1S"
            Top             =   915
            Width           =   735
         End
         Begin VB.TextBox bldg_elevator_no 
            Height          =   315
            Left            =   8040
            TabIndex        =   14
            Tag             =   "1S"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox op_factor 
            Height          =   315
            Left            =   1560
            TabIndex        =   15
            Tag             =   "1S"
            Top             =   1755
            Width           =   735
         End
         Begin VB.TextBox bldg_part_density 
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
            Left            =   3840
            TabIndex        =   8
            Tag             =   "1S"
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox bldg_part_hgt 
            Height          =   315
            Left            =   3840
            TabIndex        =   12
            Tag             =   "1S"
            Top             =   1305
            Width           =   735
         End
         Begin VB.TextBox bldg_area_std 
            Height          =   315
            Left            =   4680
            TabIndex        =   34
            Tag             =   "1S"
            Top             =   390
            Width           =   735
         End
         Begin VB.TextBox bldg_fixture_area 
            Height          =   315
            Left            =   3840
            TabIndex        =   16
            Tag             =   "1S"
            Top             =   1755
            Width           =   735
         End
         Begin VB.TextBox graphic_ref_id 
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
            Left            =   5880
            TabIndex        =   9
            Tag             =   "1S"
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox graphic_ref_id2 
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
            Left            =   5880
            TabIndex        =   13
            Tag             =   "1S"
            Top             =   1320
            Width           =   735
         End
         Begin VB.PictureBox picCommercialArea 
            BackColor       =   &H00C0C000&
            Height          =   1800
            Left            =   -74880
            ScaleHeight     =   1740
            ScaleWidth      =   8715
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   8775
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0C000&
               Caption         =   "Area to &Bold"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   405
               Left            =   120
               TabIndex        =   81
               Top             =   1335
               Width           =   8535
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   9
                  Left            =   7920
                  TabIndex        =   90
                  Top             =   165
                  Width           =   255
               End
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   8
                  Left            =   6976
                  TabIndex        =   89
                  Top             =   165
                  Width           =   255
               End
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   7
                  Left            =   6033
                  TabIndex        =   88
                  Top             =   165
                  Width           =   255
               End
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   6
                  Left            =   5090
                  TabIndex        =   87
                  Top             =   165
                  Width           =   255
               End
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   5
                  Left            =   4147
                  TabIndex        =   86
                  Top             =   165
                  Width           =   255
               End
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   4
                  Left            =   3204
                  TabIndex        =   85
                  Top             =   165
                  Width           =   255
               End
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   3
                  Left            =   2261
                  TabIndex        =   84
                  Top             =   165
                  Width           =   255
               End
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   2
                  Left            =   1318
                  TabIndex        =   83
                  Top             =   165
                  Width           =   255
               End
               Begin VB.OptionButton optBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   1
                  Left            =   375
                  TabIndex        =   82
                  Top             =   165
                  Width           =   255
               End
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   9
               Left            =   7800
               TabIndex        =   78
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   8
               Left            =   6844
               TabIndex        =   77
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   7
               Left            =   5892
               TabIndex        =   76
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   6
               Left            =   4940
               TabIndex        =   75
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   5
               Left            =   3988
               TabIndex        =   74
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   3036
               TabIndex        =   73
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   3
               Left            =   2084
               TabIndex        =   72
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   2
               Left            =   1132
               TabIndex        =   71
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   1
               Left            =   180
               TabIndex        =   70
               Tag             =   "2S"
               Top             =   1065
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   9
               Left            =   7800
               TabIndex        =   69
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   8
               Left            =   6844
               TabIndex        =   68
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   7
               Left            =   5892
               TabIndex        =   67
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   6
               Left            =   4940
               TabIndex        =   66
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   5
               Left            =   3988
               TabIndex        =   65
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   3036
               TabIndex        =   64
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   3
               Left            =   2084
               TabIndex        =   63
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   2
               Left            =   1132
               TabIndex        =   62
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.TextBox SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   1
               Left            =   180
               TabIndex        =   61
               Tag             =   "2S"
               Top             =   585
               Width           =   735
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00C00000&
               X1              =   7695
               X2              =   7695
               Y1              =   585
               Y2              =   1350
            End
            Begin VB.Line Line8 
               BorderColor     =   &H00C00000&
               X1              =   6741
               X2              =   6741
               Y1              =   600
               Y2              =   1365
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00C00000&
               X1              =   5790
               X2              =   5790
               Y1              =   600
               Y2              =   1365
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00C00000&
               X1              =   4839
               X2              =   4839
               Y1              =   600
               Y2              =   1365
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00C00000&
               X1              =   3888
               X2              =   3888
               Y1              =   585
               Y2              =   1350
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00C00000&
               X1              =   2937
               X2              =   2937
               Y1              =   600
               Y2              =   1365
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00C00000&
               X1              =   1986
               X2              =   1986
               Y1              =   585
               Y2              =   1350
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00C00000&
               X1              =   1035
               X2              =   1035
               Y1              =   585
               Y2              =   1350
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "9"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   8055
               TabIndex        =   60
               Top             =   285
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   7100
               TabIndex        =   59
               Top             =   285
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   6150
               TabIndex        =   58
               Top             =   285
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   5200
               TabIndex        =   57
               Top             =   285
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   4250
               TabIndex        =   56
               Top             =   285
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   3300
               TabIndex        =   55
               Top             =   285
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   2350
               TabIndex        =   54
               Top             =   285
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1400
               TabIndex        =   53
               Top             =   285
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   450
               TabIndex        =   52
               Top             =   285
               Width           =   255
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00C00000&
               BorderWidth     =   3
               X1              =   0
               X2              =   8760
               Y1              =   945
               Y2              =   945
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "S.F. Area / L.F.Perimeter"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   51
               Top             =   0
               Width           =   3420
            End
         End
         Begin VB.PictureBox picResi 
            BackColor       =   &H00C0C000&
            Height          =   1800
            Left            =   -74880
            ScaleHeight     =   1740
            ScaleWidth      =   8715
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   8775
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   11
               Left            =   7920
               TabIndex        =   103
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   10
               Left            =   7131
               TabIndex        =   102
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   11
               Left            =   7920
               TabIndex        =   123
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   10
               Left            =   7131
               TabIndex        =   122
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   1
               Left            =   75
               TabIndex        =   93
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   2
               Left            =   859
               TabIndex        =   94
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   3
               Left            =   1643
               TabIndex        =   95
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   2427
               TabIndex        =   96
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   5
               Left            =   3211
               TabIndex        =   97
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   6
               Left            =   3995
               TabIndex        =   98
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   7
               Left            =   4779
               TabIndex        =   99
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   8
               Left            =   5563
               TabIndex        =   100
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_SFArea 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   9
               Left            =   6347
               TabIndex        =   101
               Tag             =   "2S"
               Top             =   555
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   1
               Left            =   75
               TabIndex        =   106
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   2
               Left            =   859
               TabIndex        =   108
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   3
               Left            =   1643
               TabIndex        =   110
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   2427
               TabIndex        =   112
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   5
               Left            =   3211
               TabIndex        =   114
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   6
               Left            =   3995
               TabIndex        =   116
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   7
               Left            =   4779
               TabIndex        =   118
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   8
               Left            =   5563
               TabIndex        =   120
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.TextBox Resi_LFPerimeter 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   9
               Left            =   6347
               TabIndex        =   121
               Tag             =   "2S"
               Top             =   915
               Width           =   735
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00C0C000&
               Caption         =   "Area to &Bold"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   450
               Left            =   60
               TabIndex        =   125
               Top             =   1290
               Width           =   8610
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   11
                  Left            =   8175
                  TabIndex        =   138
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   10
                  Left            =   7389
                  TabIndex        =   137
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   1
                  Left            =   360
                  TabIndex        =   127
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   2
                  Left            =   1141
                  TabIndex        =   128
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   3
                  Left            =   1922
                  TabIndex        =   129
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   4
                  Left            =   2703
                  TabIndex        =   130
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   5
                  Left            =   3484
                  TabIndex        =   131
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   6
                  Left            =   4265
                  TabIndex        =   132
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   7
                  Left            =   5046
                  TabIndex        =   134
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   8
                  Left            =   5827
                  TabIndex        =   135
                  Top             =   195
                  Width           =   255
               End
               Begin VB.OptionButton optResiBold 
                  BackColor       =   &H00C0C000&
                  Caption         =   "Option3"
                  Height          =   195
                  Index           =   9
                  Left            =   6608
                  TabIndex        =   136
                  Top             =   195
                  Width           =   255
               End
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "11"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   8190
               TabIndex        =   133
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "10"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   7338
               TabIndex        =   126
               Top             =   270
               Width           =   330
            End
            Begin VB.Line Line20 
               BorderColor     =   &H00C00000&
               X1              =   7890
               X2              =   7890
               Y1              =   555
               Y2              =   1185
            End
            Begin VB.Line Line19 
               BorderColor     =   &H00C00000&
               X1              =   7095
               X2              =   7095
               Y1              =   540
               Y2              =   1185
            End
            Begin VB.Line Line18 
               BorderColor     =   &H00C00000&
               BorderWidth     =   3
               X1              =   -15
               X2              =   8745
               Y1              =   870
               Y2              =   870
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   390
               TabIndex        =   119
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   16
               Left            =   1162
               TabIndex        =   117
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   1934
               TabIndex        =   115
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   2706
               TabIndex        =   113
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   3478
               TabIndex        =   111
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   4250
               TabIndex        =   109
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   5022
               TabIndex        =   107
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   5794
               TabIndex        =   105
               Top             =   270
               Width           =   255
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "9"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   6566
               TabIndex        =   104
               Top             =   270
               Width           =   255
            End
            Begin VB.Line Line17 
               BorderColor     =   &H00C00000&
               X1              =   825
               X2              =   825
               Y1              =   540
               Y2              =   1185
            End
            Begin VB.Line Line16 
               BorderColor     =   &H00C00000&
               X1              =   1620
               X2              =   1620
               Y1              =   540
               Y2              =   1185
            End
            Begin VB.Line Line15 
               BorderColor     =   &H00C00000&
               X1              =   2415
               X2              =   2415
               Y1              =   540
               Y2              =   1185
            End
            Begin VB.Line Line14 
               BorderColor     =   &H00C00000&
               X1              =   3180
               X2              =   3180
               Y1              =   555
               Y2              =   1185
            End
            Begin VB.Line Line13 
               BorderColor     =   &H00C00000&
               X1              =   3960
               X2              =   3960
               Y1              =   570
               Y2              =   1185
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00C00000&
               X1              =   4755
               X2              =   4755
               Y1              =   525
               Y2              =   1185
            End
            Begin VB.Line Line11 
               BorderColor     =   &H00C00000&
               X1              =   5535
               X2              =   5535
               Y1              =   555
               Y2              =   1185
            End
            Begin VB.Line Line10 
               BorderColor     =   &H00C00000&
               X1              =   6315
               X2              =   6315
               Y1              =   555
               Y2              =   1185
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "S.F. Area / L.F.Perimeter"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   105
               TabIndex        =   124
               Top             =   0
               Width           =   3495
            End
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Building:"
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
            Left            =   -74520
            TabIndex        =   80
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Partition Density:"
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
            Left            =   2280
            TabIndex        =   49
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Stories:"
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
            Left            =   720
            TabIndex        =   48
            Top             =   960
            Width           =   675
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Door Density:"
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
            Left            =   4560
            TabIndex        =   47
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Architect Fees:"
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
            Left            =   6600
            TabIndex        =   46
            Top             =   1845
            Width           =   1455
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Partition Height:"
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
            Left            =   2400
            TabIndex        =   45
            Top             =   1365
            Width           =   1395
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Floor to Floor:"
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
            Left            =   120
            TabIndex        =   44
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Standard Area:"
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
            Left            =   3240
            TabIndex        =   43
            Top             =   450
            Width           =   1395
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Std Perimeter"
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
            Left            =   5400
            TabIndex        =   42
            Top             =   450
            Width           =   1275
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ext Wall Factor:"
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
            Left            =   6585
            TabIndex        =   41
            Top             =   975
            Width           =   1410
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fixture Area:"
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
            Left            =   2400
            TabIndex        =   40
            Top             =   1815
            Width           =   1275
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "OP Factor:"
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
            Left            =   480
            TabIndex        =   39
            Top             =   1815
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Graphic 1:"
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
            Left            =   4800
            TabIndex        =   38
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Graphic 2:"
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
            Left            =   4800
            TabIndex        =   37
            Top             =   1380
            Width           =   915
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No Elevators:"
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
            Left            =   6720
            TabIndex        =   36
            Top             =   1380
            Width           =   1275
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   6  'Inside Solid
            Height          =   1380
            Left            =   120
            Top             =   795
            Width           =   8790
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Building ID:"
         Height          =   255
         Left            =   60
         TabIndex        =   32
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Category:"
         Height          =   255
         Left            =   5295
         TabIndex        =   31
         Top             =   135
         Width           =   735
      End
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Partition Height:"
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   2220
      Width           =   1275
   End
End
Attribute VB_Name = "frmBuilding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As ADODB.RecordSet
Dim m_rec2 As New ADODB.RecordSet   'Common Additives grid recordset
Dim m_recArea As New ADODB.RecordSet   'Common Additives grid recordset
Dim m_recUsage As ADODB.RecordSet

Dim m_blnRecFlag As Boolean ' True if a populated RecordSet was passed, then we show data
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
Dim m_blnClone As Boolean  'Indicate if clone is in progress
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload

Dim m_lngOriginalSkey  As Long
Dim m_o_OrigValue As String     'Generic value used to store/verify changes on gotfocus/lostfocus

Dim m_strLast_bldg_id As String ' Holds last building id so we know if it changed

Dim m_objGridMap As New CComAddsMap ' Class to handle grid
'*** APEX Migration Utility Code Change ***
'Public tdbCols As TrueOleDBGrid60.Columns
Public tdbCols As TrueOleDBGrid70.Columns
'*** APEX Migration Utility Code Change ***
'Public myTDBGrid As TrueOleDBGrid60.TDBGrid
Public myTDBGrid As TrueOleDBGrid70.TDBGrid
Dim tdbOldCols As Variant
Private Sub fill_categories()
'Fill the available categories based on the type code
    Dim bRet As Boolean
    Dim rec As ADODB.RecordSet
'*** APEX Migration Utility Code Change ***
'    Dim Item As New TrueOleDBGrid60.ValueItem
    Dim Item As New TrueOleDBGrid70.ValueItem
    Dim strSelect As String
    If type_code <> "" Then
        bldg_category.Clear
        strSelect = "select bldg_category from bldg_category where type_code = '" + type_code + "' order by bldg_category"
        bRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
        If bRet Then
            If rec.RecordCount = 0 Then
                bldg_category.AddItem "(unknown)"
            Else
                While Not rec.EOF
                    bldg_category.AddItem rec.Fields("bldg_category")
                    rec.MoveNext
                Wend
            End If
        End If
        rec.Close
    End If

End Sub


Private Sub Load_Area()
Dim iAreaIndex As Integer   '1- 9 for commercial, 1-11 for resi
Dim i As Integer
Dim strSelect As String
Dim blnReturn As Boolean

On Error Resume Next
'Load the area records and fields
strSelect = "select * from bldg_area where bldg_skey = " + CStr(m_rec.Fields("bldg_skey")) + " order by bldg_area"

For i = 1 To SFArea.Count
    SFArea(i) = ""
    LFPerimeter(i) = ""
Next i

For i = 1 To Resi_SFArea.Count
    Resi_SFArea(i) = ""
    Resi_LFPerimeter(i) = ""
Next i

        ' Use DAL to perform select
    m_recArea.Close
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_recArea)
    If blnReturn Then
        If Not m_recArea.EOF Then
            iAreaIndex = 1
            Do Until m_recArea.EOF
                If type_code = "C" Then 'commercial
                    SFArea(iAreaIndex) = m_recArea.Fields("bldg_area")
                    LFPerimeter(iAreaIndex) = m_recArea.Fields("bldg_perimeter")
                ElseIf type_code = "R" Then 'residential
                    Resi_SFArea(iAreaIndex) = m_rec.Fields("bldg_area")
                    Resi_LFPerimeter(iAreaIndex) = m_rec.Fields("bldg_perimeter")
                End If
                iAreaIndex = iAreaIndex + 1
                m_recArea.MoveNext
            Loop
        End If
        Do Until m_recArea.RecordCount = 11      'Add empty records to accomodate max
            m_recArea.AddNew
            m_recArea.Fields("bldg_skey") = bldg_skey
            m_recArea.Fields("last_update_id") = 0
            m_recArea.Fields("bldg_area") = 0
            m_recArea.Fields("bldg_perimeter") = 0
            m_recArea.Update
            If type_code = "C" Then 'commercial
                SFArea(iAreaIndex) = m_recArea.Fields("bldg_area")
                LFPerimeter(iAreaIndex) = m_recArea.Fields("bldg_perimeter")
            ElseIf type_code = "R" Then 'residential
                Resi_SFArea(iAreaIndex) = m_rec.Fields("bldg_area")
                Resi_LFPerimeter(iAreaIndex) = m_rec.Fields("bldg_perimeter")
            End If
            iAreaIndex = iAreaIndex + 1
        Loop
    Else
        MsgBox "Error retrieving Building Area"
    End If
End Sub


Private Sub RebindTDBGridNow()
    Dim oldRow As Variant
    oldRow = myTDBGrid.Bookmark
    myTDBGrid.ReBind
    myTDBGrid.Bookmark = oldRow
End Sub



Private Sub m_rec_unformatfields()
'm_rec.Fields("std_equip_cost") = Format(std_equip_cost, "####0.00")
'm_rec.Fields("std_equip_cost_op") = Format(std_equip_cost_op, "####0.00")
'm_rec.Fields("std_labor_cost") = Format(std_labor_cost, "####0.00")
'm_rec.Fields("std_labor_cost_op") = Format(std_labor_cost_op, "####0.00")
'm_rec.Fields("std_mat_cost") = Format(std_mat_cost, "####0.00")
'm_rec.Fields("std_mat_cost_op") = Format(std_mat_cost_op, "####0.00")
'm_rec.Fields("std_total_cost") = Format(std_total_cost, "####0.00")
'm_rec.Fields("std_total_cost_op") = Format(std_total_cost_op, "####0.00")
'm_rec.Fields("opn_equip_cost") = Format(opn_equip_cost, "####0.00")
'm_rec.Fields("opn_equip_cost_op") = Format(opn_equip_cost_op, "####0.00")
'm_rec.Fields("opn_labor_cost") = Format(opn_labor_cost, "####0.00")
'm_rec.Fields("opn_labor_cost_op") = Format(opn_labor_cost_op, "####0.00")
'm_rec.Fields("opn_mat_cost") = Format(opn_mat_cost, "####0.00")
'm_rec.Fields("opn_mat_cost_op") = Format(opn_mat_cost_op, "####0.00")
'm_rec.Fields("opn_total_cost") = Format(opn_total_cost, "####0.00")
'm_rec.Fields("opn_total_cost_op") = Format(opn_total_cost_op, "####0.00")
'm_rec.Fields("rr_equip_cost") = Format(rr_equip_cost, "####0.00")
'm_rec.Fields("rr_equip_cost_op") = Format(rr_equip_cost_op, "####0.00")
'm_rec.Fields("rr_labor_cost") = Format(rr_labor_cost, "####0.00")
'm_rec.Fields("rr_labor_cost_op") = Format(rr_labor_cost_op, "####0.00")
'm_rec.Fields("rr_mat_cost") = Format(rr_mat_cost, "####0.00")
'm_rec.Fields("rr_mat_cost_op") = Format(rr_mat_cost_op, "####0.00")
'm_rec.Fields("rr_total_cost") = Format(rr_total_cost, "####0.00")
'm_rec.Fields("rr_total_cost_op") = Format(rr_total_cost_op, "####0.00")
'm_rec.Fields("metric_equip_cost") = Format(metric_equip_cost, "####0.00")
'm_rec.Fields("metric_equip_cost_op") = Format(metric_equip_cost_op, "##,##0.00")
'm_rec.Fields("metric_labor_cost") = Format(metric_labor_cost, "####0.00")
'm_rec.Fields("metric_labor_cost_op") = Format(metric_labor_cost_op, "####0.00")
'm_rec.Fields("metric_mat_cost") = Format(metric_mat_cost, "####0.00")
'm_rec.Fields("metric_mat_cost_op") = Format(metric_mat_cost_op, "####0.00")
'm_rec.Fields("metric_total_cost") = Format(metric_total_cost, "####0.00")
'm_rec.Fields("metric_total_cost_op") = Format(metric_total_cost_op, "####0.00")

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
    If Not m_rec.Fields("bldg_skey") = 0 Then
        m_blnRecFlag = True
    End If
End Sub



Private Sub update_area()
Dim iErrorCount As Integer
Dim strError As String
Dim blnRet As Boolean
Dim strUpdate As String
Dim lngOrigArea As Long

On Error Resume Next
m_recArea.MoveFirst
iErrorCount = 0
Do Until m_recArea.EOF
    If m_recArea.Fields("bldg_area").Value <> m_recArea.Fields("bldg_area").OriginalValue _
        Or m_recArea.Fields("bldg_perimeter").Value <> m_recArea.Fields("bldg_perimeter").OriginalValue _
        Then
        If IsNull(m_recArea.Fields("bldg_area")) Then m_recArea.Fields("bldg_area") = 0
        If IsNull(m_recArea.Fields("bldg_perimeter")) Then m_recArea.Fields("bldg_perimeter") = 0
        If IsNull(m_recArea.Fields("bldg_area").OriginalValue) Or IsEmpty(m_recArea.Fields("bldg_area").OriginalValue) Then
            lngOrigArea = 0
        Else
            lngOrigArea = m_recArea.Fields("bldg_area").OriginalValue
        End If
        strUpdate = "exec sp_update_bldg_area "
        strUpdate = strUpdate + "@bldg_skey = " + CStr(bldg_skey)
        strUpdate = strUpdate + ", @bldg_area=" + CStr(m_recArea.Fields("bldg_area"))
        strUpdate = strUpdate + ", @bldg_perimeter=" + CStr(m_recArea.Fields("bldg_perimeter"))
        strUpdate = strUpdate + ", @bldg_orig_area =" + CStr(lngOrigArea)
        strUpdate = strUpdate + ", @last_update_person='" + strUserName + "', "
        strUpdate = strUpdate + " @last_update_id=" + CStr(m_recArea.Fields("last_update_id"))
        blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)  'Update the Building Area
        If blnRet = False Then
            iErrorCount = iErrorCount + 1
        Else
            m_recArea.Fields("last_update_id") = m_recArea.Fields("last_update_id") + 1
        End If
    End If
    m_recArea.MoveNext
Loop

m_recArea.UpdateBatch

If iErrorCount > 0 Then
    MsgBox CStr(iErrorCount) + " Errors encountered updating the Area/Perimeter.  Error: " + strError
End If
End Sub

Private Sub bldg_arch_fees_GotFocus()
        bldg_arch_fees.BackColor = vbCyan
End Sub


Private Sub bldg_arch_fees_LostFocus()
    bldg_arch_fees.BackColor = &H80000005
End Sub


Private Sub bldg_desc_GotFocus()
       bldg_desc.BackColor = vbCyan
End Sub

Private Sub bldg_desc_LostFocus()
    bldg_desc.BackColor = &H80000005

End Sub

Private Sub bldg_door_density_GotFocus()
        bldg_door_density.BackColor = vbCyan

End Sub


Private Sub bldg_door_density_LostFocus()
    bldg_door_density.BackColor = &H80000005
End Sub


Private Sub bldg_elevator_no_GotFocus()
        bldg_elevator_no.BackColor = vbCyan

End Sub


Private Sub bldg_elevator_no_LostFocus()
    bldg_elevator_no.BackColor = &H80000005
End Sub


Private Sub bldg_fixture_area_GotFocus()
        bldg_fixture_area.BackColor = vbCyan
End Sub


Private Sub bldg_fixture_area_LostFocus()
    bldg_fixture_area.BackColor = &H80000005
End Sub


Private Sub bldg_part_density_GotFocus()
        bldg_part_density.BackColor = vbCyan

End Sub


Private Sub bldg_part_density_LostFocus()
    bldg_part_density.BackColor = &H80000005
End Sub


Private Sub bldg_part_hgt_GotFocus()
       bldg_part_hgt.BackColor = vbCyan

End Sub


Private Sub bldg_part_hgt_LostFocus()
  bldg_part_hgt.BackColor = &H80000005
End Sub


Private Sub bldg_stories_GotFocus()
        bldg_stories.BackColor = vbCyan
End Sub


Private Sub bldg_stories_hgt_GotFocus()
        bldg_stories_hgt.BackColor = vbCyan

End Sub


Private Sub bldg_stories_hgt_LostFocus()
  bldg_stories_hgt.BackColor = &H80000005
End Sub


Private Sub bldg_stories_LostFocus()
  bldg_stories.BackColor = &H80000005
End Sub

Private Sub bldg_wall_factor_GotFocus()
       bldg_wall_factor.BackColor = vbCyan

End Sub


Private Sub bldg_wall_factor_LostFocus()
bldg_wall_factor.BackColor = &H80000005
End Sub


Private Sub cmdAdditiveDelete_Click()
On Error Resume Next
        If TDBGrid.AddNewMode > 0 Then
            TDBGrid.ReBind
        Else
            TDBGrid.Delete
        End If

End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String

    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?  The CSI Line will be removed.  Press the Material Usage delete button to remove a material usage.", vbYesNo + vbCritical)
    If varButton = vbNo Then
        Exit Sub
    End If

    strUpdate = "exec sp_delete_unit_cost "
    strUpdate = strUpdate + "@bldg_skey=" + str(Me.Controls("bldg_skey")) + ","
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


Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim blnUpdateBldlg As Boolean
    Dim ctr As Control
    Dim fld As ADODB.Field
    Dim rec As New ADODB.RecordSet
    Dim strError As String
    Dim strPercent_flag As String
    Dim strSelect As String
    Dim strUpdate As String
    Dim strSaveUpdate As String
    Dim intStart As Integer
    Dim varSaveBookmark As Variant
    Dim i As Integer
    Dim bln_Continue As Boolean
    TDBGrid.Update
    m_blnWereErrors = False
    
    If IsNull(TDBGrid.Bookmark) Then
        bln_Continue = True
    ElseIf TDBGrid.AddNewMode = dbgAddNewCurrent Then   'Cursor in new row, no add pending
        bln_Continue = True
    ElseIf Trim(TDBGrid.Columns("ID").Text) = "" Then
        MsgBox "The ID may not be blank."
    Else
        bln_Continue = True
    End If
    If bln_Continue = True Then
         Screen.MousePointer = vbHourglass
        
        Dim recClone As ADODB.RecordSet
        Set recClone = m_rec.Clone
        recClone.AddNew
        UpdateRecordsetFromForm Me, recClone
        
        For Each fld In m_rec.Fields
            ' If the value changed
            If Not fld.Value = recClone.Fields(fld.Name).Value Or ((IsNull(fld.Value) Or fld.Value = "") Xor (recClone.Fields(fld.Name).Value = "")) Then
                Set ctr = Nothing
                Set ctr = Me.Controls(fld.Name)
                If Not ctr Is Nothing Then
                    ' See what table the field is from
                    If left(Me.Controls(fld.Name).Tag, 1) = 1 Then
                        blnUpdateBldlg = True
                    ElseIf left(Me.Controls(fld.Name).Tag, 1) = 3 Then
                        blnUpdateBldlg = True
                    End If
                End If
            End If
        Next
        
        ' Undo the changes made by the UpdateRecordsetFromForm call above
        recClone.CancelUpdate
        recClone.Close
        Set recClone = Nothing
        'Set the Cost Change flag based on any Unit Cost field changes
        
        If blnUpdateBldlg = True Or m_objGridMap.IsPendingChange Or (m_rec2.RecordCount > 0 And m_blnClone = True) Then
            strUpdate = "exec sp_update_bldg_detail "
            BuildStoredProcSQL Me, strUpdate, 1, m_rec
            strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
            If last_update_id.Text = "" Then last_update_id.Text = 0
            strUpdate = strUpdate + ", @last_update_id=" + last_update_id.Text
            ExecUpdate strUpdate
            'Retrieve new skey
            If m_blnInsert = True Then
                strSelect = "select bldg_skey from bldg_detail where bldg_id = '" + bldg_id + "'"
                rec.Close
                 blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                If blnRet = True Then
                    bldg_skey = rec.Fields("bldg_skey")
                End If
            End If
            update_area
            
            If m_blnClone = True Then       'Copy the output_usage data for the unit cost
                strUpdate = "exec sp_copy_output_usage @type = 'U', @FromSkey = '" & m_lngOriginalSkey & "', @ToSkey='" & bldg_skey.Text + "', "
                strUpdate = strUpdate + " @last_update_date='" + Format(Now(), "General Date") + "', "
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "', "
                strUpdate = strUpdate + " @last_update_id='1'"
            End If
                'Process changes or deletions
            If m_objGridMap.IsPendingChange Or (m_blnClone = True And m_rec2.RecordCount > 0) Then
                
                'If cloning, update the bldg_skey in all records in the grid.
                If m_blnClone = True Or m_blnInsert = True Then
                    'm_objGridMap.bldg_skey = bldg_skey.Text
                    If m_rec2.RecordCount > 0 Then
                        m_rec2.MoveFirst
                        Do Until m_rec2.EOF
                            m_rec2.Fields("bldg_skey") = bldg_skey.Text
                            m_rec2.MoveNext
                        Loop
                    End If
                End If
'                If m_blnClone = True Then   'Flag all rows as new
'                    m_objGridMap.SetRowState
'                End If
                blnRet = m_objGridMap.Update
                If blnRet = False Then
                    m_blnWereErrors = True
                End If
            End If
            m_blnClone = False  ' no longer cloning if we were
            If m_blnWereErrors = False Then
                ' Put latest data into source recordset
                
                UpdateRecordsetFromForm Me, m_rec
    
                If IsNull(m_rec.Fields("last_update_id").Value) Then
                    m_rec.Fields("last_update_id").Value = 1
                End If
                last_update_id.Text = m_rec.Fields("last_update_id").Value

            End If
            If m_blnWereErrors = False Then
                MsgBox "Update successful."
            End If
            RebindTDBGridNow
            varSaveBookmark = TDBGrid.Bookmark
            TDBGrid.Refresh
            TDBGrid.Bookmark = varSaveBookmark
        Else
            MsgBox "You must modify a field before updating."
        End If
        Screen.MousePointer = vbNormal
    End If
End Sub
Private Sub ExecUpdate(strUpdate As String)
Dim blnRet As Boolean
Dim strError As String
On Error Resume Next
'Update the database with the current update sql string.
'If the update fails, display a message, otherwise increment the last update Id
        blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)  'Update the Material Usage for unit cost
        If blnRet = False Then
            MsgBox strError
            m_blnWereErrors = True
        Else
            last_update_id.Text = CInt(last_update_id.Text) + 1
            m_rec.Fields("last_update_id").Value = last_update_id.Text
        End If
End Sub

Private Function ReplaceSkey(strString, strSkey As String) As String
Dim iStart As Integer
Dim iEnd As Integer
Dim strTemp As String

iStart = InStr(1, strString, "@bldg_skey=")
If iStart > 0 Then
    iEnd = InStr(iStart, strString, ",")
    strTemp = left(strString, iStart + 15) + strSkey + right(strString, Len(strString) - iEnd + 1)
    ReplaceSkey = strTemp
End If

End Function




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
    Dim iOptCount As Integer
    Move START_LEFT, START_TOP ' , 10305, 8115
    Me.Height = 7230
    Me.Width = 9255
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid

    If Not m_rec.State = adStateClosed Then
        UpdateFormFromRecordset Me, m_rec
'        format_costs
    End If

    ' If we are NOT inserting
    If m_blnInsert = False Then
        ' Lock fields that can't be changed
        bldg_id.Locked = True
        bldg_id.BackColor = LTGREY
        Me.Caption = Me.Caption + " [" + m_rec.Fields("bldg_id").Value + " - " + m_rec.Fields("bldg_desc").Value + "]"
    Else
        ' If we are inserting and not showing data
        ' Set some defaults
        bldg_skey.Text = ""
        If Not m_blnRecFlag Then
            Me.Caption = Me.Caption + " [New]"
        Else
            Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("bldg_id").Value + " - " + m_rec.Fields("bldg_desc").Value + "]"
            m_blnClone = True
            m_lngOriginalSkey = m_rec.Fields("bldg_skey").Value
        End If
    End If
    If Not m_blnClone Then
        m_strLast_bldg_id = m_rec.Fields("bldg_id").Value
    Else
        bldg_id.CausesValidation = True
    End If
    Load_Area
    If type_code = "C" Then
        optBldg_Type(0) = True
        If col_to_bold > 0 Then
            optBold(col_to_bold).Value = True
        End If
    ElseIf type_code = "R" Then
        optBldg_Type(1) = True
        If col_to_bold > 0 Then
            optResiBold(col_to_bold).Value = True
        End If
    End If
    blnReturn = LockField(Me, "bldg_area_std")
    blnReturn = LockField(Me, "bldg_perimeter_std")
    'Fill the available building types from the Commercial/Residential options
    For iOptCount = 0 To optBldg_Type.Count - 1
        If optBldg_Type(iOptCount).Value = True Then
            Select Case iOptCount
                Case 0  'Commercial
                    type_code = "C"
                Case 1  'Residential
                    type_code = "R"
            End Select
            Exit For
        End If
    Next iOptCount
    FillUsageGrid
End Sub
Private Sub Form_Resize()
If Me.Height > 6510 Then
    picGrid.Height = Me.Height - picTop.Height - picFooter.Height - 400
    fraUnitCost.Height = picGrid.Height - 200
    TDBGrid.Height = fraUnitCost.Height - 850
    cmdAdditiveDelete.top = fraUnitCost.top + fraUnitCost.Height - cmdAdditiveDelete.Height - 140
End If
fraUnitCost.Width = Me.Width - 200
TDBGrid.Width = fraUnitCost.Width - 200
ResizeForm Me
End Sub


Private Sub bldg_desc_Change()
Dim intLength As Integer
If Len(bldg_desc) > 0 Then
    If Asc(right(bldg_desc, 1)) >= 0 And Asc(right(bldg_desc, 1)) <= 31 Then
        intLength = Len(bldg_desc)
        MsgBox "Non-printable characters are not allowed in the building description."
        bldg_desc.Text = left(bldg_desc.Text, intLength - 2)
        bldg_desc.SelStart = intLength - 2
    End If
End If

End Sub




Private Sub graphic_ref_id_GotFocus()
graphic_ref_id.BackColor = vbCyan
End Sub


Private Sub graphic_ref_id_LostFocus()
graphic_ref_id.BackColor = &H80000005
End Sub


Private Sub graphic_ref_id2_GotFocus()
graphic_ref_id2.BackColor = vbCyan
End Sub


Private Sub graphic_ref_id2_LostFocus()
graphic_ref_id2.BackColor = &H80000005
End Sub


Private Sub LFPerimeter_GotFocus(Index As Integer)
    If optBold(Index).Value = True Then
        LFPerimeter(Index).BackColor = &HC0E0FF
    Else
        LFPerimeter(Index).BackColor = vbCyan
    End If
    
    m_o_OrigValue = LFPerimeter(Index)

End Sub

Private Sub LFPerimeter_KeyPress(Index As Integer, KeyAscii As Integer)
    If CheckNumericField(LFPerimeter(Index), KeyAscii, LFPerimeter(Index).SelStart, LFPerimeter(Index).SelLength, 0) = False Then
        KeyAscii = 0
    End If

End Sub

Private Sub LFPerimeter_LostFocus(Index As Integer)
If optBold(Index).Value = True Then
    LFPerimeter(Index).BackColor = &H80C0FF
Else
    LFPerimeter(Index).BackColor = &H80000005
End If

If m_o_OrigValue <> LFPerimeter(Index) Then
    m_recArea.AbsolutePosition = Index
    m_recArea.Fields("bldg_perimeter") = IIf(LFPerimeter(Index) = "", 0, LFPerimeter(Index))
End If

End Sub

Private Sub LFPerimeter_Validate(Index As Integer, Cancel As Boolean)
'Dim I As Integer
'
'If Len(LFPerimeter(Index)) > 0 And LFPerimeter(Index) <> "0" Then
'    If Index > 1 Then   'Not First element
'        If Len(LFPerimeter(Index - 1)) = 0 Or LFPerimeter(Index - 1) = "0" Then
'            MsgBox "Please enter the perimeter in ascending order."
'            Cancel = True
'        ElseIf CLng(LFPerimeter(Index)) <= CLng(LFPerimeter(Index - 1)) Then
'            MsgBox "The perimeter must be greater than the preceding perimeter."
'            Cancel = True
'        End If
'    End If
'        If Index < LFPerimeter.Count Then
'            If Len(LFPerimeter(Index + 1)) > 0 And LFPerimeter(Index + 1) <> "0" Then
'                If CLng(LFPerimeter(Index)) >= CLng(LFPerimeter(Index + 1)) Then
'                    MsgBox "The perimeter must be less then the next element."
'                    Cancel = True
'                End If
'            End If
'        End If
'Else
'    For I = Index + 1 To LFPerimeter.Count
'        If Len(LFPerimeter(I)) > 0 And LFPerimeter(I) <> "0" Then
'            MsgBox "Please delete the last perimeter first."
'            Cancel = True
'            Exit For
'        End If
'    Next I
'End If

End Sub

Private Sub op_factor_GotFocus()
op_factor.BackColor = vbCyan

End Sub


Private Sub op_factor_LostFocus()
    op_factor.BackColor = &H80000005

End Sub


Private Sub optBldg_Type_Click(Index As Integer)
If optBldg_Type(Index).Value = True Then
    Select Case Index
        Case 0  'Commercial
            type_code = "C"
        Case 1  'Residential
            type_code = "R"
    End Select
End If
End Sub

Private Sub optBold_Click(Index As Integer)
Dim i As Integer

For i = 1 To 9  'Commercial
If optBold(i).Value = True Then         'Set the std area/perimeter, highlight the selection
    LFPerimeter(i).BackColor = &H80C0FF
    SFArea(i).BackColor = &H80C0FF
    bldg_area_std = SFArea(i)
    bldg_perimeter_std = LFPerimeter(i)
Else
    LFPerimeter(i).BackColor = &H80000005
    SFArea(i).BackColor = &H80000005
End If
Next i

col_to_bold = Index

End Sub

Private Sub optResiBold_Click(Index As Integer)
Dim i As Integer
For i = 1 To 11  'Resi
If optResiBold(i).Value = True Then
    Resi_LFPerimeter(i).BackColor = &H80C0FF
    Resi_SFArea(i).BackColor = &H80C0FF
Else
    Resi_LFPerimeter(i).BackColor = &H80000005
    Resi_SFArea(i).BackColor = &H80000005
End If
Next i
col_to_bold = Index

End Sub


Private Sub Resi_LFPerimeter_GotFocus(Index As Integer)
    If optResiBold(Index).Value = True Then
        Resi_LFPerimeter(Index).BackColor = &HC0E0FF
    Else
        Resi_LFPerimeter(Index).BackColor = vbCyan
    End If
    
    m_o_OrigValue = Resi_LFPerimeter(Index)
End Sub

Private Sub Resi_LFPerimeter_KeyPress(Index As Integer, KeyAscii As Integer)
    If CheckNumericField(Resi_LFPerimeter(Index), KeyAscii, Resi_LFPerimeter(Index).SelStart, Resi_LFPerimeter(Index).SelLength, 0) = False Then
        KeyAscii = 0
    End If

End Sub

Private Sub Resi_LFPerimeter_LostFocus(Index As Integer)
If optResiBold(Index).Value = True Then
    Resi_LFPerimeter(Index).BackColor = &H80C0FF
Else
    Resi_LFPerimeter(Index).BackColor = &H80000005
End If

    If m_o_OrigValue <> Resi_LFPerimeter(Index) Then
        m_recArea.AbsolutePosition = Index
        m_recArea.Fields("bldg_perimeter") = Resi_LFPerimeter(Index)
    End If

End Sub


Private Sub Resi_LFPerimeter_Validate(Index As Integer, Cancel As Boolean)
'Dim I As Integer
'
'If Len(Resi_LFPerimeter(Index)) > 0 And Resi_LFPerimeter(Index) <> "0" Then
'    If Index > 1 Then   'Not First element
'        If Len(Resi_LFPerimeter(Index - 1)) = 0 Or Resi_LFPerimeter(Index - 1) = "0" Then
'            MsgBox "Please enter the perimeter in ascending order."
'            Cancel = True
'        ElseIf CLng(Resi_LFPerimeter(Index)) <= CLng(Resi_LFPerimeter(Index - 1)) Then
'            MsgBox "The perimeter must be greater than the preceding perimeter."
'            Cancel = True
'        End If
'    End If
'        If Index < Resi_LFPerimeter.Count Then
'            If Len(Resi_LFPerimeter(Index + 1)) > 0 And Resi_LFPerimeter(Index + 1) <> "0" Then
'                If CLng(Resi_LFPerimeter(Index)) >= CLng(Resi_LFPerimeter(Index + 1)) Then
'                    MsgBox "The perimeter must be less then the next element."
'                    Cancel = True
'                End If
'            End If
'        End If
'Else
'    For I = Index + 1 To Resi_LFPerimeter.Count
'        If Len(Resi_LFPerimeter(I)) > 0 And Resi_LFPerimeter(I) <> "0" Then
'            MsgBox "Please delete the last perimeter first."
'            Cancel = True
'            Exit For
'        End If
'    Next I
'End If
End Sub

Private Sub Resi_SFArea_GotFocus(Index As Integer)

If optResiBold(Index).Value = True Then
    Resi_SFArea(Index).BackColor = &HC0E0FF
Else
    Resi_SFArea(Index).BackColor = vbCyan
End If
m_o_OrigValue = Resi_SFArea(Index)
End Sub

Private Sub Resi_SFArea_KeyPress(Index As Integer, KeyAscii As Integer)
    If CheckNumericField(Resi_SFArea(Index), KeyAscii, Resi_SFArea(Index).SelStart, Resi_SFArea(Index).SelLength, 0) = False Then
        KeyAscii = 0
    End If

End Sub

Private Sub Resi_SFArea_LostFocus(Index As Integer)
If optResiBold(Index).Value = True Then
    Resi_SFArea(Index).BackColor = &H80C0FF
Else
    Resi_SFArea(Index).BackColor = &H80000005
End If

    If m_o_OrigValue <> Resi_SFArea(Index) Then
        m_recArea.AbsolutePosition = Index
        m_recArea.Fields("bldg_perimeter") = Resi_SFArea(Index)
    End If

End Sub



Private Sub Resi_SFArea_Validate(Index As Integer, Cancel As Boolean)
'Dim I As Integer
'
'If Len(Resi_SFArea(Index)) > 0 And Resi_SFArea(Index) <> "0" Then
'    If Index > 1 Then   'Not First element
'        If Len(Resi_SFArea(Index - 1)) = 0 Or Resi_SFArea(Index - 1) = "0" Then
'            MsgBox "Please enter the area in ascending order."
'            Cancel = True
'        ElseIf CLng(Resi_SFArea(Index)) <= CLng(Resi_SFArea(Index - 1)) Then
'            MsgBox "The area must be greater than the preceding area."
'            Cancel = True
'        End If
'    End If
'        If Index < Resi_SFArea.Count Then
'            If Len(Resi_SFArea(Index + 1)) > 0 And Resi_SFArea(Index + 1) <> "0" Then
'                If CLng(Resi_SFArea(Index)) >= CLng(Resi_SFArea(Index + 1)) Then
'                    MsgBox "The area must be less then the next element."
'                    Cancel = True
'                End If
'            End If
'        End If
'Else
'    For I = Index + 1 To Resi_SFArea.Count
'        If Len(Resi_SFArea(I)) > 0 And Resi_SFArea(I) <> "0" Then
'            MsgBox "Please delete the last area first."
'            Cancel = True
'            Exit For
'        End If
'    Next I
'End If
'
End Sub

Private Sub SFArea_GotFocus(Index As Integer)
    If optBold(Index).Value = True Then
        SFArea(Index).BackColor = &HC0E0FF
    Else
        SFArea(Index).BackColor = vbCyan
    End If
    
    m_o_OrigValue = SFArea(Index)
End Sub


Private Sub SFArea_KeyPress(Index As Integer, KeyAscii As Integer)
    If CheckNumericField(SFArea(Index), KeyAscii, SFArea(Index).SelStart, SFArea(Index).SelLength, 0) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub SFArea_LostFocus(Index As Integer)
If optBold(Index).Value = True Then
    SFArea(Index).BackColor = &H80C0FF
Else
    SFArea(Index).BackColor = &H80000005
End If

If m_o_OrigValue <> SFArea(Index) Then
    m_recArea.AbsolutePosition = Index
    If Len(SFArea(Index)) = 0 Then
        SFArea(Index) = 0
    End If
    m_recArea.Fields("bldg_area") = SFArea(Index)
End If

End Sub


Private Sub SFArea_Validate(Index As Integer, Cancel As Boolean)
'Dim I As Integer
'
'If Len(SFArea(Index)) > 0 And SFArea(Index) <> "0" Then
'    If Index > 1 Then   'Not First element
'        If Len(SFArea(Index - 1)) = 0 Or SFArea(Index - 1) = "0" Then
'            MsgBox "Please enter the area in ascending order."
'            Cancel = True
'        ElseIf CLng(SFArea(Index)) <= CLng(SFArea(Index - 1)) Then
'            MsgBox "The area must be greater than the preceding area."
'            Cancel = True
'        End If
'    End If
'        If Index < SFArea.Count Then
'            If Len(SFArea(Index + 1)) > 0 And SFArea(Index + 1) <> "0" Then
'                If CLng(SFArea(Index)) >= CLng(SFArea(Index + 1)) Then
'                    MsgBox "The area must be less then the next element."
'                    Cancel = True
'                End If
'            End If
'        End If
'Else
'    For I = Index + 1 To SFArea.Count
'        If Len(SFArea(I)) > 0 And SFArea(I) <> "0" Then
'            MsgBox "Please delete the last area first."
'            Cancel = True
'            Exit For
'        End If
'    Next I
'End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
    If optBldg_Type(0).Value = True Then 'Commercial
        picResi.Enabled = False
        picCommercialArea.Enabled = True
    Else
        picResi.Enabled = True
        picCommercialArea.Enabled = False
    End If
Else
        picResi.Enabled = False
        picCommercialArea.Enabled = False
End If


End Sub

Private Sub type_code_Change()
fill_categories
If type_code = "R" Then
    picResi.Visible = True
    picCommercialArea.Visible = False
ElseIf type_code = "C" Then
    picResi.Visible = False
    picCommercialArea.Visible = True
End If
End Sub
Private Sub bldg_id_LostFocus()
        m_strLast_bldg_id = bldg_id.Text
End Sub

Private Sub FillUsageGrid()

    On Error GoTo Error_Processing
    Dim strSelect As String
    Dim blnReturn As Boolean

        ' Check to see if the mat_id entered exists already
    If m_blnClone = True Then
        'If cloning, set skey to 0 after reading data
        strSelect = "exec sp_select_bldg_com_add_cst @bldg_skey = " + str(m_lngOriginalSkey)
'        strSelect = "Select mu.mat_skey, mu.bldg_skey, mu.unit_qty, mu.input_factor, mu.output_factor, mu.adj_factor, mu.last_update_person, mu.last_update_date, mu.last_update_id, mu.comment, m.mat_id from material_usage as mu, material as m where mu.mat_skey = m.mat_skey and mu.bldg_skey = " + str(m_lngOriginalSkey) + " order by m.mat_id"
    Else
        If bldg_skey.Text = "" Then
            strSelect = "exec sp_select_bldg_com_add_cst @bldg_skey = 0"
'            strSelect = "Select mu.mat_skey, mu.bldg_skey as bldg_skey, mu.unit_qty, mu.input_factor, mu.output_factor, mu.adj_factor, mu.last_update_person, mu.last_update_date, mu.last_update_id, mu.comment, m.mat_id from material_usage as mu, material as m where mu.mat_skey = m.mat_skey and mu.bldg_skey = 0 order by m.mat_id"
        Else
            strSelect = "exec sp_select_bldg_com_add_cst @bldg_skey = " + bldg_skey.Text
'            strSelect = "Select mu.mat_skey, mu.bldg_skey, mu.unit_qty, mu.input_factor, mu.output_factor, mu.adj_factor, mu.last_update_person, mu.last_update_date, mu.last_update_id, mu.comment, m.mat_id from material_usage as mu, material as m where mu.mat_skey = m.mat_skey and mu.bldg_skey = " + bldg_skey.Text + " order by m.mat_id"
        End If
    End If

    ' Use DAL to perform select
    m_rec2.Close
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec2)
    If Not IsNumeric(bldg_skey.Text) Then
        bldg_skey.Text = 0
    End If
'    m_objGridMap.bldgSKey = CLng(bldg_skey.Text)
    If bldg_skey.Text = "" And m_rec2.RecordCount > 0 Then
        m_rec2.MoveFirst
        Do Until m_rec2.EOF
            m_rec2.Fields("bldg_skey") = 0
            m_rec2.MoveNext
        Loop
    End If

    m_objGridMap.RecordSet = m_rec2
    If m_blnClone = True Then
        blnReturn = m_objGridMap.SetRowStateNew
    Else
        blnReturn = m_objGridMap.SetRowStateNone
    End If
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec2.RecordCount
    m_objGridMap.bldgSKey = bldg_skey
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
    Dim bln_New As Boolean

    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        blnPendingChange = IsControlChanged(Me, m_rec)
        If blnPendingChange = True Or m_objGridMap.IsPendingChange Then
            Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
            If Button = vbYes Then
                m_blnWereErrors = False
                If m_blnInsert Or m_blnClone Then
                    bln_New = True
                End If
                If m_blnWereErrors Then
                    Cancel = True
                Else
                    cmdUpdate_Click
                    ' If there were errors, cancel the close
                    If m_blnWereErrors Then
                        Cancel = True
                    Else
                        RestoreGridValues
                    End If
                End If
            ElseIf Button = vbCancel Then
                Cancel = True
                Exit Sub
            ElseIf m_blnInsert = True Then
                m_rec.Delete
            End If
        End If
    End If
End Sub

Private Function LockField(frm As Form, fld As String) As Boolean

        frm.Controls(fld).Enabled = False
        frm.Controls(fld).Locked = True
        frm.Controls(fld).ForeColor = LTGREY

End Function



Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Dim strErrorMsg As String
        strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
        If Len(strErrorMsg) > 0 Then
            MsgBox strErrorMsg
        End If
    End If
End Sub
Private Sub bldg_id_Validate(Cancel As Boolean)
Dim bln_New As Boolean
Dim rec As ADODB.RecordSet
Dim bln_result As Boolean
Dim strSelect As String

If m_strLast_bldg_id <> bldg_id.Text Or bldg_id.Text = "" Then
    If m_blnInsert Or m_blnClone Then
        bln_New = True
    End If
'    If Invalid_id_Format(bldg_id, "bldg_id", m_rec, bln_New, "bldg_detail", False) = True Then
'        Cancel = True
'        Else
            strSelect = "select bldg_skey from bldg_detail where bldg_id = '" + bldg_id + "'"
            bln_result = g_objDAL.GetRecordset(g_cnShared, strSelect, rec)
            m_objGridMap.bldgSKey = bldg_skey
'    End If
End If
End Sub
