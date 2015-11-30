VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form dlgOutput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Output"
   ClientHeight    =   8490
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   14355
   Icon            =   "dlgOutput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   8040
      Width           =   1335
   End
   Begin VB.PictureBox picControl1 
      Height          =   375
      Left            =   105
      Picture         =   "dlgOutput.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   270
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picControl0 
      Height          =   375
      Left            =   105
      Picture         =   "dlgOutput.frx":052C
      ScaleHeight     =   315
      ScaleWidth      =   270
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picControl2 
      Height          =   375
      Left            =   105
      Picture         =   "dlgOutput.frx":0616
      ScaleHeight     =   315
      ScaleWidth      =   270
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CheckBox optGroupItems 
      Caption         =   "All Misc"
      Height          =   255
      Index           =   4
      Left            =   105
      TabIndex        =   9
      Tag             =   "9"
      Top             =   5925
      Width           =   1215
   End
   Begin VB.CheckBox optGroupItems 
      Caption         =   "All Metric"
      Height          =   255
      Index           =   3
      Left            =   105
      TabIndex        =   8
      Tag             =   "4"
      Top             =   4820
      Width           =   1215
   End
   Begin VB.CheckBox optGroupItems 
      Caption         =   "All R&&R"
      Height          =   255
      Index           =   2
      Left            =   105
      TabIndex        =   7
      Tag             =   "3"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox optGroupItems 
      Caption         =   "All Open"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Tag             =   "2"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox optGroupItems 
      Caption         =   "All Standard"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Tag             =   "1"
      Top             =   120
      Width           =   1215
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridOutput1 
      Height          =   2385
      Left            =   1365
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   12100
      _ExtentX        =   21352
      _ExtentY        =   4207
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      Appearance      =   3
      DataMode        =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(45)  =   ":id=29,.parent=0,.bgcolor=&HFFFFFF&"
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
   Begin TrueOleDBGrid80.TDBGrid TDBGridOutput2 
      Height          =   1000
      Left            =   1365
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2600
      Width           =   12100
      _ExtentX        =   21352
      _ExtentY        =   1773
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      Appearance      =   3
      DataMode        =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(45)  =   ":id=29,.parent=0,.bgcolor=&HFFFFFF&"
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
   Begin TrueOleDBGrid80.TDBGrid TDBGridOutput3 
      Height          =   1005
      Left            =   1365
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Width           =   12100
      _ExtentX        =   21352
      _ExtentY        =   1773
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      Appearance      =   3
      DataMode        =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(45)  =   ":id=29,.parent=0,.bgcolor=&HFFFFFF&"
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
   Begin TrueOleDBGrid80.TDBGrid TDBGridOutput4 
      Height          =   1005
      Left            =   1365
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4820
      Width           =   12100
      _ExtentX        =   21352
      _ExtentY        =   1773
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      Appearance      =   3
      DataMode        =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(45)  =   ":id=29,.parent=0,.bgcolor=&HFFFFFF&"
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
   Begin TrueOleDBGrid80.TDBGrid TDBGridOutput5 
      Height          =   1980
      Left            =   1365
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5925
      Width           =   12100
      _ExtentX        =   21352
      _ExtentY        =   3493
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
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      Appearance      =   3
      DataMode        =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(45)  =   ":id=29,.parent=0,.bgcolor=&HFFFFFF&"
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
Attribute VB_Name = "dlgOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Class to handle grid
Dim m_objGridMap() As New COutputMap
' Recordset to hold query results
Dim m_recOutPut() As New ADODB.RecordSet
'
' Recordset to hold query results
Dim m_recOutputSkeys As New ADODB.RecordSet
'
'   Recordset to hold query results
Dim m_rec As New ADODB.RecordSet

Dim m_Dialog As New CCommonDialog
Dim m_SKey As String
Public m_strKeyType As String          'rlh 7/14/2008 (moved dim to  mainModule.bas

Dim b1stLoad As Boolean
Dim bln_Save As Boolean
Dim blnCheckGroup As Boolean
Dim m_iSkeyCount As Integer

Public bShowAllFields As Boolean
Public m_objOutput As New CCDdal.CRSMDataAccess ' Global DAL object

'Property to Get/Set Allowable Output Usage Formats
Public Property Get OutputUsageFormat() As OUTPUT_USAGE_FORMAT
    OutputUsageFormat = m_objGridMap(0).OutputUsageFormat
End Property
Public Property Let OutputUsageFormat(NewValue As OUTPUT_USAGE_FORMAT)
    Dim I As Long
    'Set all 5 grids to new OutputUsageFormat value
    For I = 0 To 4
        m_objGridMap(I).OutputUsageFormat = NewValue
    Next
End Property

Private Sub cmdCancel_Click()

    Unload Me
    Set dlgOutput = Nothing

End Sub

Private Sub Form_Initialize()
    
    Me.Top = Forms(0).Top + 240
    ' Cache a global connection
    m_objOutput.CacheConnection (CONNECT)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim I As Integer

    If bln_Save = False Then
        ' Check if there are pending changes
        On Error Resume Next
        ' CODE BLOCK COMMENTED 8/9/2005 RTD
        '       THIS IS INCORRECT USAGE OF DATACHANGED PROPERTY,
        '       WHICH INDICATES A ROW IS CURRENTLY BEING EDITED.
        'If TDBGridOutput1.DataChanged = True Then
        '    bln_Save = True
        'ElseIf TDBGridOutput2.DataChanged = True Then
        '    bln_Save = True
        'ElseIf TDBGridOutput3.DataChanged = True Then
        '    bln_Save = True
        'ElseIf TDBGridOutput4.DataChanged = True Then
        '    bln_Save = True
        'ElseIf TDBGridOutput5.DataChanged = True Then
        '    bln_Save = True
        'End If
        ' MODIFIED 8/9/2005 RTD
        ' TO CORRECT PROBLEM WITH FORM UNLOAD NOT PROMPTING TO SAVE CHANGES
        bln_Save = m_objGridMap(0).IsPendingChange Or _
                   m_objGridMap(1).IsPendingChange Or _
                   m_objGridMap(2).IsPendingChange Or _
                   m_objGridMap(3).IsPendingChange Or _
                   m_objGridMap(4).IsPendingChange
    End If

    If bln_Save Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel + vbQuestion)
        If Button = vbYes Then
            SaveButton_Click
            ' If there were errors, cancel the close
'            If m_blnWereErrors Then
'                Cancel = True
'            End If
        ElseIf Button = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim I As Integer

    On Error Resume Next
    For I = 0 To 4
        m_recOutPut(I).Close
    Next I

End Sub

Private Sub Form_Load()
    
    Me.Visible = False
    DoEvents
    Screen.MousePointer = vbHourglass
    'Me.Top = 240
    'Me.Left = 4400
    CenterFormInParent Me, fMainForm
    
    '   Initialize grid.
    ReDim Preserve m_objGridMap(5)
    ReDim Preserve m_recOutPut(5)
    b1stLoad = True

    ' Turn on the TopMost attribute.
    'SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, conSwpNoActivate Or conSwpShowWindow Or conSwpNoMove Or conSwpNoSize

End Sub

Public Sub FillData(Optional ByRef bNewSelection As Boolean = True)
    Dim rec As New ADODB.RecordSet
    ' Get the data for the selected key, if only one
    Dim blnRet As Boolean
    Dim strSELECT As String
    Dim I As Integer
    
    'On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    'Select output for all currently selected skeys in calling form's grid
    'Temp table will hold skeys
    If b1stLoad = True Then
        '
        '   Get the skeys & skey_type they selected
        strSELECT = "SELECT * FROM ##output_skeys"
        blnRet = m_objOutput.GetRecordset(vbNullString, strSELECT, m_recOutputSkeys)
        '
        '   Seem to get repeated errors here but if search again immediately it works.
        If blnRet = False Then
            blnRet = m_objOutput.GetRecordset(vbNullString, strSELECT, m_recOutputSkeys)
        End If
        If blnRet = False Or m_recOutputSkeys.RecordCount = 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching.", vbExclamation
            Exit Sub
        Else
            m_iSkeyCount = m_recOutputSkeys.RecordCount
            If m_iSkeyCount = 0 Then
                Screen.MousePointer = vbNormal
                MsgBox "Please select record(s) before using Output.", vbInformation
                Exit Sub
            End If
        End If
        
        With m_objGridMap(0)
            .SetGrid TDBGridOutput1
            .InitGrid picControl0, picControl1, picControl2, m_iSkeyCount
            '.DataChanged = False
        End With
        With m_objGridMap(1)
            .SetGrid TDBGridOutput2
            .InitGrid picControl0, picControl1, picControl2, m_iSkeyCount
            '.DataChanged = False
        End With
        With m_objGridMap(2)
            .SetGrid TDBGridOutput3
            .InitGrid picControl0, picControl1, picControl2, m_iSkeyCount
            '.DataChanged = False
        End With
        With m_objGridMap(3)
            .SetGrid TDBGridOutput4
            .InitGrid picControl0, picControl1, picControl2, m_iSkeyCount
            '.DataChanged = False
        End With
        With m_objGridMap(4)
            .SetGrid TDBGridOutput5
            .InitGrid picControl0, picControl1, picControl2, m_iSkeyCount
            '.DataChanged = False
        End With
        b1stLoad = False
    ElseIf bNewSelection = True Then
        '
        '   Check to make sure our skeys were not overwritten
        '   from the temp table
        '
        '   Get the skeys & skey_type they selected
        strSELECT = "SELECT * FROM ##output_skeys"
        blnRet = m_objOutput.GetRecordset(vbNullString, strSELECT, m_recOutputSkeys)
        '
        '   Seem to get repeated errors here but if search again immediately it works.
        If blnRet = False Then
            blnRet = m_objOutput.GetRecordset(vbNullString, strSELECT, m_recOutputSkeys)
        End If
        If blnRet = False Or m_recOutputSkeys.RecordCount = 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching.", vbExclamation
            Exit Sub
        Else
            m_iSkeyCount = m_recOutputSkeys.RecordCount
            If m_iSkeyCount = 0 Then
                Screen.MousePointer = vbNormal
                MsgBox "Please select record(s) before using Output.", vbInformation
                Exit Sub
            End If
        End If
        '
        '   Refresh grid value items so can display appropriate checkboxes
        For I = 0 To 4
            m_objGridMap(I).RefreshValueItems picControl0, picControl1, picControl2, m_iSkeyCount
        Next I
    Else
        m_recOutputSkeys.MoveFirst
        strSELECT = "SELECT * FROM ##output_skeys WHERE skey_type = '" & Trim(m_recOutputSkeys.Fields("skey_type").Value) & "'"
        blnRet = m_objOutput.GetRecordset(vbNullString, strSELECT, m_rec)
        If blnRet = False Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching, please close and reopen Output.", vbExclamation
            Exit Sub
        ElseIf m_rec.RecordCount = 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "Unable to refresh Output screen, please close and reopen Output.", vbExclamation
            Exit Sub
        Else
            m_iSkeyCount = m_recOutputSkeys.RecordCount
            If m_iSkeyCount = 0 Then
                Screen.MousePointer = vbNormal
                MsgBox "Please select records before using output.", vbInformation
                Exit Sub
            End If
        End If
        m_rec.Close
    End If
    
    For I = 0 To 4
        Set m_recOutPut(I) = Nothing
        Set m_recOutPut(I) = New ADODB.RecordSet
        '
        'Select Summary information for Output by group, output id
        strSELECT = "exec sp_select_temp_output_usage @output_group_id=" + CStr(optGroupItems(I).Tag)
        'Stop 'RLH - TEMPORARY FOR DEBUG ONLY!!!
        blnRet = m_objOutput.GetRecordset(vbNullString, strSELECT, m_recOutPut(I))
        If blnRet = False Or m_recOutPut(I).RecordCount = 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching.", vbExclamation
            Exit Sub
        End If
        '
        'Update summary table with multi use indicators for each value
        'Count the number of detail records assigned for one output
        'If 0, the check box will be unchecked
        'If >0 but less than the total skeys selected, the check box will be grayed
        'If = to the total skeys, the check box will be checked
        'The same process will be used for each value being displayed:
        '
        'If iCountSkeysSelected > 0 And iCountSkeysSelected < m_iSkeyCount Then
        '    m_recOutPut(i).Fields("selected").Value = 2 'Some, not all
        'ElseIf iCountSkeysSelected = m_iSkeyCount Then
        '    m_recOutPut(i).Fields("selected").Value = 1 'all (dft = 0, none)
        'End If
        '
        '   Reset the grid contents
        Select Case I
            Case 0
                '
                '   Pass recordset to handler class.
                m_objGridMap(I).RecordSet = m_recOutPut(I)
            
                With TDBGridOutput1
                    .Bookmark = Null
                    .ReBind
                    .ApproxCount = m_recOutPut(I).RecordCount
                End With
            Case 1
                '
                '   Pass recordset to handler class.
                m_objGridMap(I).RecordSet = m_recOutPut(I)

                With TDBGridOutput2
                    .Bookmark = Null
                    .ReBind
                    .ApproxCount = m_recOutPut(I).RecordCount
                End With
            Case 2
                '
                '   Pass recordset to handler class.
                m_objGridMap(I).RecordSet = m_recOutPut(I)

                With TDBGridOutput3
                    .Bookmark = Null
                    .ReBind
                    .ApproxCount = m_recOutPut(I).RecordCount
                End With
            Case 3
                '
                '   Pass recordset to handler class.
                m_objGridMap(I).RecordSet = m_recOutPut(I)

                With TDBGridOutput4
                    .Bookmark = Null
                    .ReBind
                    .ApproxCount = m_recOutPut(I).RecordCount
                End With
            Case 4
                '
                '   Pass recordset to handler class.
                m_objGridMap(I).RecordSet = m_recOutPut(I)

                With TDBGridOutput5
                    .Bookmark = Null
                    .ReBind
                    .ApproxCount = m_recOutPut(I).RecordCount
                End With
        End Select
        CheckGroupValue I
    Next I
    Me.Visible = True
    Me.ZOrder 1
    Screen.MousePointer = vbNormal
End Sub

Private Sub CheckGroupValue(index As Integer)
    Dim blnTrueValue As Boolean
    Dim blnFalseValue As Boolean
    Dim rsCheckRecords As ADODB.RecordSet

    'Set the value of the group options
    Set rsCheckRecords = m_recOutPut(index).Clone
    blnCheckGroup = True
    If rsCheckRecords.RecordCount > 0 Then
        rsCheckRecords.MoveFirst
        Do Until rsCheckRecords.EOF
            Select Case rsCheckRecords![Selected]
            Case 0  'unchecked
                blnFalseValue = True
                If blnTrueValue = True Then
                    Exit Do
                End If
            Case 1  'checked
                blnTrueValue = True
                If blnFalseValue = True Then
                    Exit Do
                End If
            Case 2  'grayed
                blnFalseValue = True
                blnTrueValue = True
                Exit Do
            End Select
            rsCheckRecords.MoveNext
        Loop
    Else
        optGroupItems(index).Value = 0  'unchecked
    End If

    If blnFalseValue = True And blnTrueValue = True Then
        optGroupItems(index).Value = 2      'grayed
    Else
        If blnFalseValue = True Then
            optGroupItems(index).Value = 0  'unchecked
        Else
            optGroupItems(index).Value = 1  'checked
        End If
    End If
    rsCheckRecords.Close
    Set rsCheckRecords = Nothing
    blnCheckGroup = False
End Sub

Public Function ValidateOutputFormat() As Boolean
'ADDED 8/15/2005 RTD
'VALIDATE THAT THE 'EXT_INDICATOR' FIELD IS NOT BLANK FOR ALL SELECTED BOOKS
    Dim I As Long
    Dim bResult As Boolean
    
    bResult = True
    If m_strKeyType2 = "U" Then
        'UPDATED 8/23/2005 RTD
        'EXT_INDICATOR IS ONLY REQUIRED FOR 'U' TYPE RECORDS
        For I = 0 To 4
            m_recOutPut(I).MoveFirst
            Do While Not m_recOutPut(I).EOF
                If m_recOutPut(I).Fields("selected").Value Then
                    'rlh 07/10/2008
                    Select Case m_recOutPut(I).Fields("ext_indicator").Value
                    Case 0, 1, 2, 3
                    Case Else
                        'MsgBox ("Found a selected row that needs a MSTFMT to be specified")
                        bResult = False
                        Exit For
                    End Select
                    
                    End If
                
                m_recOutPut(I).MoveNext
            Loop
        Next
        End If
    
    
    ValidateOutputFormat = bResult
    
End Function

Private Sub SaveButton_Click()
'    On Error Resume Next
    Dim lstOutput As Object
    Dim I As Integer
    Dim strUpdate As String
    Dim strError As String
    Dim blnRet As Boolean
    Dim blnErrors As Boolean
    Dim strFind As String
    Dim blnUpdate As Boolean
    
    If Not ValidateOutputFormat Then
        MsgBox "You must choose a MasterFormat output selection for each enabled Book.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    For I = 0 To 4      'For each grid displayed
        m_recOutPut(I).MoveFirst
        If Not (m_recOutPut(I).BOF And m_recOutPut(I).EOF) Then
            TDBGridOutput1.Update
            TDBGridOutput2.Update
            TDBGridOutput3.Update
            TDBGridOutput4.Update
            TDBGridOutput5.Update
            Do Until m_recOutPut(I).EOF = True
                If m_recOutPut(I).Fields("selected") <> IIf(IsNull(m_recOutPut(I).Fields("selected").OriginalValue), "", m_recOutPut(I).Fields("selected").OriginalValue) _
                    Or m_recOutPut(I).Fields("ext_indicator") <> IIf(IsNull(m_recOutPut(I).Fields("ext_indicator").OriginalValue), "", m_recOutPut(I).Fields("ext_indicator").OriginalValue) _
                    Or m_recOutPut(I).Fields("graphic_ref_id") <> IIf(IsNull(m_recOutPut(I).Fields("graphic_ref_id").OriginalValue), "", m_recOutPut(I).Fields("graphic_ref_id").OriginalValue) _
                    Or m_recOutPut(I).Fields("table_ref_id") <> IIf(IsNull(m_recOutPut(I).Fields("table_ref_id").OriginalValue), "", m_recOutPut(I).Fields("table_ref_id").OriginalValue) _
                    Or m_recOutPut(I).Fields("ext_graphic_ref_id") <> IIf(IsNull(m_recOutPut(I).Fields("ext_graphic_ref_id").OriginalValue), "", m_recOutPut(I).Fields("ext_graphic_ref_id").OriginalValue) _
                    Or m_recOutPut(I).Fields("tag_code") <> IIf(IsNull(m_recOutPut(I).Fields("tag_code").OriginalValue), "", m_recOutPut(I).Fields("tag_code").OriginalValue) _
                    Or m_recOutPut(I).Fields("ext_table_ref_id") <> IIf(IsNull(m_recOutPut(I).Fields("ext_table_ref_id").OriginalValue), "", m_recOutPut(I).Fields("ext_table_ref_id").OriginalValue) _
                    Or m_recOutPut(I).Fields("format_characters") <> IIf(IsNull(m_recOutPut(I).Fields("format_characters").OriginalValue), "", m_recOutPut(I).Fields("format_characters").OriginalValue) _
                    Or m_recOutPut(I).Fields("indent_code") <> IIf(IsNull(m_recOutPut(I).Fields("indent_code").OriginalValue), "", m_recOutPut(I).Fields("indent_code").OriginalValue) _
                    Or m_recOutPut(I).Fields("format_code") <> IIf(IsNull(m_recOutPut(I).Fields("format_code").OriginalValue), "", m_recOutPut(I).Fields("format_code").OriginalValue) Then
                    'Update the output based on the grid
                    'The detail recordset contains the original values
                    'If selected = 0 (none) output will be deleted for all the detail
                    'If selected = 1 (all) the detail output will be compared.  If no value exists for the output_id, it will be added.
                    '   If it exists and one of the values have been changed, it will be updated.
                    '   If it exists and the values all match, it will be bypassed.
                    'If selected = 2 (some) no update will be performed.
                    If m_recOutPut(I).Fields("selected") = 0 Then   'Not selected - delete all
                        m_recOutputSkeys.MoveFirst
                        Do Until m_recOutputSkeys.EOF
                            Status "Deleting ID " & m_recOutPut(I).Fields("output_id") & "..."
                            DoEvents
                            strUpdate = "exec sp_delete_output_usage @output_id=" + CStr(m_recOutPut(I).Fields("output_id")) + ", @skey=" + CStr(m_recOutputSkeys.Fields("skey").Value) + ", @skey_type='" + CStr(Trim(m_recOutputSkeys.Fields("skey_type").Value)) + "'"
                            blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                            If Not blnRet Then
                                Screen.MousePointer = vbNormal
                                MsgBox strError, vbExclamation
                                blnErrors = True
                                Exit Do
                            End If
                            m_recOutputSkeys.MoveNext
                        Loop
                        
                    ElseIf m_recOutPut(I).Fields("selected") = 1 Then   'All selected, either update or add all selected skeys
                        ' Update changes for all selected skeys.
                        m_recOutputSkeys.MoveFirst
                        Do Until m_recOutputSkeys.EOF
                            ' 8/10/2005 RTD - DISPLAY PROGRESS IN STATUS BAR
                            Status "Updating " & m_recOutPut(I).Fields("output_desc").Value & "..."
                            DoEvents
                            strUpdate = "exec sp_insert_output_usage @output_id=" + CStr(m_recOutPut(I).Fields("output_id")) + _
                                ", @skey_type='" + CStr(Trim(m_recOutputSkeys.Fields("skey_type").Value)) + "', @skey=" + CStr(m_recOutputSkeys.Fields("skey").Value) + _
                                ", @output_group_id=" + CStr(m_recOutPut(I).Fields("output_group_id").Value) + ", "
                            
                            If IsNull(m_recOutPut(I).Fields("graphic_ref_id").Value) Then
                                strUpdate = strUpdate + "@graphic_ref_id='', "
                            Else
                                strUpdate = strUpdate + "@graphic_ref_id='" + m_recOutPut(I).Fields("graphic_ref_id").Value + "', "
                            End If
                            If IsNull(m_recOutPut(I).Fields("table_ref_id").Value) Then
                                strUpdate = strUpdate + "@table_ref_id='',"
                            Else
                                strUpdate = strUpdate + "@table_ref_id='" + m_recOutPut(I).Fields("table_ref_id").Value + "', "
                            End If
                            ' 8/10/2005 RTD - ADDED TO SUPPORT MASTERFORMAT 2004
                            If IsNull(m_recOutPut(I).Fields("ext_indicator").Value) Then
                                strUpdate = strUpdate + "@ext_indicator='0', "
                            Else
                                strUpdate = strUpdate + "@ext_indicator='" & Abs(m_recOutPut(I).Fields("ext_indicator").Value) & "', "
                            End If
                            ' 8/10/2005 RTD - ADDED TO SUPPORT MASTERFORMAT 2004
                            If IsNull(m_recOutPut(I).Fields("ext_graphic_ref_id").Value) Then
                                strUpdate = strUpdate + "@ext_graphic_ref_id='', "
                            Else
                                strUpdate = strUpdate + "@ext_graphic_ref_id='" + m_recOutPut(I).Fields("ext_graphic_ref_id").Value + "', "
                            End If
                            ' 06/11/2008 RLH - ADDED TO SUPPORT "GREEN" TAG
                            If IsNull(m_recOutPut(I).Fields("tag_code").Value) Then
                                strUpdate = strUpdate + "@tag_code='', "
                            Else
                                strUpdate = strUpdate + "@tag_code='" + m_recOutPut(I).Fields("tag_code").Value + "', "
                            End If
                            ' 8/10/2005 RTD - ADDED TO SUPPORT MASTERFORMAT 2004
                            If IsNull(m_recOutPut(I).Fields("ext_table_ref_id").Value) Then
                                strUpdate = strUpdate + "@ext_table_ref_id='',"
                            Else
                                strUpdate = strUpdate + "@ext_table_ref_id='" + m_recOutPut(I).Fields("ext_table_ref_id").Value + "', "
                            End If
                            If IsNull(m_recOutPut(I).Fields("indent_code").Value) Then
                                strUpdate = strUpdate + "@indent_code=0,"
                            Else
                                strUpdate = strUpdate + "@indent_code=" + CStr(m_recOutPut(I).Fields("indent_code").Value) + ", "
                            End If
                            If IsNull(m_recOutPut(I).Fields("format_code").Value) Then
                                strUpdate = strUpdate + "@format_code='',"
                            Else
                                strUpdate = strUpdate + "@format_code='" + m_recOutPut(I).Fields("format_code").Value + "', "
                            End If
                            If IsNull(m_recOutPut(I).Fields("format_characters").Value) Then
                                strUpdate = strUpdate + "@format_characters=0,"
                            Else
                                strUpdate = strUpdate + "@format_characters=" + CStr(m_recOutPut(I).Fields("format_characters").Value) + ", "
                            End If
                            strUpdate = strUpdate + " @last_update_person='" + strUserName
    
                            If CStr(Trim(m_recOutPut(I).Fields("last_update_id"))) = "" Then
                                strUpdate = strUpdate + "', @last_update_id=1"
                            Else
                                strUpdate = strUpdate + "', @last_update_id=" + CStr(m_recOutPut(I).Fields("last_update_id"))
                            End If
                            blnRet = g_objDAL.ExecQuery(strConnect, strUpdate, strError)
                            If Not blnRet Then
                                Screen.MousePointer = vbNormal
                                MsgBox strError, vbExclamation
                                blnErrors = True
                                Exit Do
                            End If
                            m_recOutputSkeys.MoveNext
                        Loop
                    End If
                End If
                m_recOutPut(I).MoveNext
            Loop
        End If
    Next I

    Status ""
    If Not blnErrors Then
        Screen.MousePointer = vbNormal
        m_Dialog.Message = "Update successful."
        m_Dialog.X = Me.Left + (Me.Width / 2)
        m_Dialog.Y = Me.Top + (Me.Height / 2)
        m_Dialog.ShowMessage
    End If

    FillData False
    Screen.MousePointer = vbNormal

End Sub

Private Sub optGroupItems_Click(index As Integer)
    Dim intGroupId As Integer
    Dim I As Integer
    
    If blnCheckGroup = True Then Exit Sub
    ' Skip 0 as it is the initial, unused item
    m_recOutPut(index).MoveFirst
    If Not (m_recOutPut(index).BOF And m_recOutPut(index).EOF) Then
        Do Until m_recOutPut(index).EOF
            m_recOutPut(index)![Selected] = optGroupItems(index).Value
            m_recOutPut(index).Update
            m_recOutPut(index).MoveNext
        Loop
    End If
    Select Case index
        Case 0
            TDBGridOutput1.ReBind
            TDBGridOutput1.Update
        
        Case 1
            TDBGridOutput2.ReBind
            TDBGridOutput2.Update
        
        Case 2
            TDBGridOutput3.ReBind
            TDBGridOutput3.Update
        
        Case 3
            TDBGridOutput4.ReBind
            TDBGridOutput4.Update
        
        Case 4
            TDBGridOutput5.ReBind
            TDBGridOutput5.Update
    End Select
        
    intGroupId = optGroupItems(index).Tag
End Sub

Public Sub SetKeys(sKey As String, strType As String)
    m_SKey = sKey
    m_strKeyType = strType
    FillData True
End Sub

Public Sub GetKey(sKey As String)
    sKey = m_SKey
End Sub

Private Sub TDBGrid1_AfterUpdate(index As Integer)
    CheckGroupValue index
End Sub

Private Sub TDBGrid2_AfterUpdate(index As Integer)
    CheckGroupValue index
End Sub

Private Sub TDBGrid3_AfterUpdate(index As Integer)
    CheckGroupValue index
End Sub

Private Sub TDBGrid4_AfterUpdate(index As Integer)
    CheckGroupValue index
End Sub

Private Sub TDBGrid5_AfterUpdate(index As Integer)
    CheckGroupValue index
End Sub

