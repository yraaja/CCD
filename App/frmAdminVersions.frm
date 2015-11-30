VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAdminVersions 
   Caption         =   "Version Administration"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdminVersions.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   7635
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   450
      Left            =   6360
      TabIndex        =   3
      Top             =   4080
      Width           =   1110
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   450
      Left            =   5040
      TabIndex        =   2
      Top             =   4080
      Width           =   1110
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Version Table"
      TabPicture(0)   =   "frmAdminVersions.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Status Messages"
      TabPicture(1)   =   "frmAdminVersions.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TDBGrid2"
      Tab(1).ControlCount=   1
      Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
         Height          =   2175
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Version"
         Columns(0).DataField=   "version_string"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Major"
         Columns(1).DataField=   "major"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Minor"
         Columns(2).DataField=   "minor"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Build"
         Columns(3).DataField=   "build"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Release Date"
         Columns(4).DataField=   "release_date"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   48
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Status"
         Columns(5).DataField=   "status"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Setup MSI Path"
         Columns(6).DataField=   "setup_uri"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Upgrade Folder Path"
         Columns(7).DataField=   "upgrade_uri"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Last Updated"
         Columns(8).DataField=   "last_update_date"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Last Updated"
         Columns(9).DataField=   "last_update_person"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1905"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1826"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=926"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=847"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1032"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=953"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=979"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=900"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=3493"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3413"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2831"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2752"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=4683"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=4604"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=5318"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=5239"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=3096"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3016"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=1958"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1879"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowAddNew     =   -1  'True
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Named:id=33:Normal"
         _StyleDefs(77)  =   ":id=33,.parent=0"
         _StyleDefs(78)  =   "Named:id=34:Heading"
         _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   ":id=34,.wraptext=-1"
         _StyleDefs(81)  =   "Named:id=35:Footing"
         _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   "Named:id=36:Selected"
         _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=37:Caption"
         _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(87)  =   "Named:id=38:HighlightRow"
         _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(89)  =   "Named:id=39:EvenRow"
         _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(91)  =   "Named:id=40:OddRow"
         _StyleDefs(92)  =   ":id=40,.parent=33"
         _StyleDefs(93)  =   "Named:id=41:RecordSelector"
         _StyleDefs(94)  =   ":id=41,.parent=34"
         _StyleDefs(95)  =   "Named:id=42:FilterBar"
         _StyleDefs(96)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Status"
         Columns(0).DataField=   "version_status"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Description"
         Columns(1).DataField=   "status_description"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Header"
         Columns(2).DataField=   "status_header"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Message Text"
         Columns(3).DataField=   "status_text"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=926"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=3175"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3096"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3228"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3149"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=10504"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=10425"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Label lblNotesLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   4080
      Width           =   4455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version Administration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11100
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmAdminVersions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const USER_TABLE = "USER_NAMES"
Const ALL_TABLES = "All"

Private m_version_rec As New ADODB.RecordSet   ' Recordset to hold parameters
Private m_status_rec As New ADODB.RecordSet   ' Recordset to hold parameters

Dim m_sngYCoord As Single
'
'   Keeps up with the field that last had focus when form
'   is deactivate, so when activated can set focus.
Dim m_strCurrentFormControl As String
'
'   Notifies that it wants to see changes.
Dim sEventSubscriberID As String
'

Public Function UpdateTable() As Boolean
    Const TABLE_NAME = "VERSIONS_STATUS"
    Const KEY_FIELD = "version_status"
    Dim strUpdate As String     ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateRow As Boolean
    Dim fld As ADODB.Field
    Dim strError As String      ' Error string returned from DAL
    Dim intErrors As Integer    ' Tracks if any errors have occurred
    Dim intSuccess As Integer   ' Tracks successful updates
    Dim m_rec As ADODB.RecordSet

    'On Error Resume Next
    UpdateTable = True
    intErrors = 0
    intSuccess = 0
    
    Set m_rec = m_status_rec
    m_rec.MoveFirst
    ' Loop through all grid records
    Do While Not m_rec.EOF
        ' Skip the record if it didn't change
            blnReturn = False
            blnUpdateRow = False
            ' Loop through the fields to look for changes
            For Each fld In m_rec.Fields
                ' If the value changed
                If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
                    blnUpdateRow = True
                    Exit For
                End If
            Next
            If blnUpdateRow Then
                strUpdate = "UPDATE " & TABLE_NAME & " SET "
                For Each fld In m_rec.Fields
                    If fld.Name <> KEY_FIELD Then
                        If fld.Name = "last_update_person" Then
                            fld.Value = strUserName
                        End If
                        strUpdate = strUpdate & " " & fld.Name & "='" & SQLFixString(fld.Value) & "',"
                    End If
                Next
                If Right(strUpdate, 1) = "," Then
                    strUpdate = Left(strUpdate, Len(strUpdate) - 1)
                End If
                strUpdate = strUpdate & " WHERE"
                strUpdate = strUpdate & " " & KEY_FIELD & "='" & m_rec.Fields(KEY_FIELD) & "'"
                blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                ' Reset on success
                If blnReturn = True Then
                    intSuccess = intSuccess + 1
                Else
                    intErrors = intErrors + 1
                End If
            End If
        m_rec.MoveNext
    Loop
    m_rec.UpdateBatch

    Dim strMsg As String
    strMsg = ""
    If intSuccess > 0 Then
        strMsg = str(intSuccess) + " rows updated successfully." + vbCrLf
    End If
    If intErrors > 0 Then
        strMsg = strMsg + str(intErrors) + " errors occurred."
        ' Return value will be False
        UpdateTable = False
    End If
    If Len(strMsg) > 0 Then
        'MsgBox strMsg, vbInformation + vbOKOnly
    End If
    
End Function

Public Function UpdateVersionTable() As Boolean
    Dim strUpdate As String     ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateRow As Boolean
    Dim fld As ADODB.Field
    Dim strError As String      ' Error string returned from DAL
    Dim intErrors As Integer    ' Tracks if any errors have occurred
    Dim intSuccess As Integer   ' Tracks successful updates
    Dim m_rec As ADODB.RecordSet
    
    On Error Resume Next
    UpdateVersionTable = True
    intErrors = 0
    intSuccess = 0
    
    Set m_rec = m_version_rec
    m_rec.MoveFirst
    ' Loop through all grid records
    Do While Not m_rec.EOF
        ' Skip the record if it didn't change
            blnReturn = False
            blnUpdateRow = False
            ' Loop through the fields to look for changes
            For Each fld In m_rec.Fields
                ' If the value changed
                If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor IsNull(fld.Value)) Then
                    blnUpdateRow = True
                    Exit For
                End If
            Next
            If blnUpdateRow Then
                strUpdate = "exec usp_update_versions "
                For Each fld In m_rec.Fields
                    If fld.Name = "last_update_person" Then
                        fld.Value = strUserName
                    End If
                    strUpdate = strUpdate & " @" & fld.Name & "='" & fld.Value & "',"
                Next
                If Right(strUpdate, 1) = "," Then
                    strUpdate = Left(strUpdate, Len(strUpdate) - 1)
                End If
                blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                ' Reset on success
                If blnReturn = True Then
                    intSuccess = intSuccess + 1
                Else
                    intErrors = intErrors + 1
                End If
            End If
        m_rec.MoveNext
    Loop
    m_rec.UpdateBatch

    Dim strMsg As String
    strMsg = ""
    If intSuccess > 0 Then
        strMsg = str(intSuccess) + " rows updated successfully." + vbCrLf
    End If
    If intErrors > 0 Then
        strMsg = strMsg + str(intErrors) + " errors occurred."
        ' Return value will be False
        UpdateVersionTable = False
    End If
    If Len(strMsg) > 0 Then
        MsgBox strMsg, vbInformation + vbOKOnly
    End If
    
    ' REFRESH GRIDS
    LoadGrids

End Function


Private Sub LoadGrids()
    Dim strSELECT As String
    Dim blnReturn As Boolean
    Dim Item As New TrueOleDBGrid80.ValueItem
    
    ' SET VERSION TABLE GRID
    strSELECT = "SELECT * FROM VERSIONS"
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, m_version_rec)
    TDBGrid1.DataSource = m_version_rec
    TDBGrid1.ReBind
    
    ' SET STATUS TABLE GRID
    strSELECT = "SELECT * FROM VERSIONS_STATUS"
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, m_status_rec)
    TDBGrid2.DataSource = m_status_rec
    TDBGrid2.ReBind
    
    ' SET STATUS FIELD IN VERSION TABLE GRID DROP-DOWN VALUES
    With TDBGrid1.Columns("Status").ValueItems
        .Clear
        Do While Not m_status_rec.EOF
            Item.Value = m_status_rec.Fields("version_status")
            Item.DisplayValue = m_status_rec.Fields("version_status") & ": " & m_status_rec.Fields("status_description")
            .Add Item
            m_status_rec.MoveNext
        Loop
        .Translate = True
        .Validate = True
        .Presentation = dbgComboBox
    End With
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Set frmAdminVersions = Nothing
End Sub

Private Sub cmdUpdate_Click()
    UpdateTable
    UpdateVersionTable
End Sub

Private Sub Form_Activate()
    Dim ctl As Control
    
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
        'ShowToolbarIcons True
    End If
End Sub

Private Sub Form_Deactivate()
    m_strCurrentFormControl = Me.ActiveControl.Name
    'ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()
    
    Status ("Loading Version Admin Control Panel...")
    Screen.MousePointer = vbHourglass
    sEventSubscriberID = EventSubscriberAdd(Me)
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim I As Long
    Dim aLocks As Variant
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    TDBGrid1.AlternatingRowStyle = True
    TDBGrid1.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid1.EvenRowStyle.BackColor = g_intAlternateRowColor
    TDBGrid2.AlternatingRowStyle = True
    TDBGrid2.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid2.EvenRowStyle.BackColor = g_intAlternateRowColor
    
    ' SET READ ONLY COLUMNS
    aLocks = Array(1, 2, 3, 8, 9)
    For I = LBound(aLocks) To UBound(aLocks)
        TDBGrid1.Columns(aLocks(I)).Locked = True
        TDBGrid1.Columns(aLocks(I)).ForeColor = vbGrayText
    Next
    
    LoadGrids
    SSTab1_Click (0)
    
    Status ("")
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_Resize()
    Dim iDlgUnit As Long
    On Error Resume Next
    '
    '   Need to place in common routine for all forms.
    '   Possibly place all buttons in a frame like frame1 with
    '   common name and can just place it.
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 10500 Then
            iDlgUnit = SSTab1.Left
            Line2.X2 = Me.Width - 210
            cmdClose.Left = Me.Width - cmdClose.Width - iDlgUnit * 2
            cmdClose.Top = Me.Height - cmdClose.Height - iDlgUnit * 3
            cmdUpdate.Left = cmdClose.Left - cmdUpdate.Width - iDlgUnit
            cmdUpdate.Top = cmdClose.Top
            SSTab1.Width = Me.Width - iDlgUnit * 3
            SSTab1.Height = cmdClose.Top - cmdClose.Height - SSTab1.Top
            TDBGrid1.Width = SSTab1.Width - iDlgUnit * 2
            TDBGrid1.Height = SSTab1.Height - iDlgUnit * 4
            TDBGrid2.Width = TDBGrid1.Width
            TDBGrid2.Height = TDBGrid1.Height
            lblNotes.Top = cmdClose.Top
            lblNotesLabel.Top = cmdClose.Top
        Else
            Me.Width = 10500
        End If
        If Me.Height >= 6135 Then
        Else
            Me.Height = 6135
        End If
    Else
        ShowMinimizedForms
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'ShowToolbarIcons False
    EventSubscriberRemove sEventSubscriberID
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Select Case SSTab1.Tab
    Case 0
        lblNotes.Caption = "Setup = Full Path to the MSI file (include MSI file)" & vbCrLf & _
                            "Upgrade = Folder containing the upgrade files"
    Case 1
        lblNotes.Caption = "The header and message are displayed in the Launcher's upgrade notification window."
    End Select
    
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    Dim sVersion As String
    Dim a As Variant
    
    On Error Resume Next
    If ColIndex = TDBGrid1.Columns("Version").ColIndex Then
        ' PARSE VERSION STRING INTO MAJOR/MINOR/BUILD VALUES
        TDBGrid1.Columns("Major").Value = 0
        TDBGrid1.Columns("Minor").Value = 0
        TDBGrid1.Columns("Build").Value = 0
        sVersion = TDBGrid1.Columns("Version").Text
        a = Split(sVersion, ".")
        TDBGrid1.Columns("Major").Value = a(0)
        TDBGrid1.Columns("Minor").Value = a(1)
        TDBGrid1.Columns("Build").Value = a(2)
    End If
End Sub

