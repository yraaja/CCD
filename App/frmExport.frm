VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmExport 
   Caption         =   "Export"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5741
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
      Splits(0).DividerColor=   13160660
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
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2778
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
      Splits(0).DividerColor=   13160660
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
      AllowUpdate     =   0   'False
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridFilter 
      Height          =   1335
      Left            =   5640
      TabIndex        =   6
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2355
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Field"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
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
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483636
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
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin VB.Label lblFilters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filters:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   7
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label lblResultset 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resultset:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select columns to export:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As New ADODB.RecordSet
Dim m_original_rec As New ADODB.RecordSet
Dim m_data_rec As New ADODB.RecordSet
Dim m_filter_rec As New ADODB.RecordSet

Dim m_Title As String

'THE TITLE TO DISPLAY
Public Property Get Title() As String
    Title = m_Title
End Property
Public Property Let Title(NewValue As String)
    m_Title = NewValue
    Me.Caption = "Export - " & NewValue
End Property

'SET THE DATABASE RECORDSET AND POPULATE THE SETUP GRID
Public Sub SetRow(TDBG As TDBGrid, rec As ADODB.RecordSet)
    Dim Col As TrueOleDBGrid80.Column
    Dim bVisible As Boolean
    
    Set m_data_rec = rec.Clone
    Set m_original_rec = rec.Clone
    
    ' Add Grid columns to Export Setup Grid
    On Error Resume Next
    m_data_rec.MoveFirst
    m_rec.Open
    'm_rec.Delete adAffectAll
    For Each Col In TDBG.Splits(0).Columns
        bVisible = Col.Visible
        If TDBG.Splits.Count > 1 Then
            bVisible = bVisible Or TDBG.Splits(1).Columns(Col.Caption).Visible
        End If
        m_rec.AddNew
        m_rec.Fields("Export") = bVisible
        m_rec.Fields("Name") = Col.Caption
        m_rec.Fields("Field") = Col.DataField
        m_rec.Fields("Alignment") = Col.Alignment
        If Not rec.EOF Then
            m_rec.Fields("Value") = m_data_rec.Fields(Col.DataField).Value
            m_rec.Fields("DataType") = m_data_rec.Fields(Col.DataField).Type
        End If
        m_rec.Update
    Next

    Set Col = TDBGrid.Columns(0)
    Col.Caption = "Export"
    Col.ValueItems.Presentation = dbgCheckBox
    Col.Alignment = dbgCenter

    Set Col = TDBGrid.Columns(1)
    Col.Caption = "Column Name"
    Col.Locked = True
    Col.AutoSize

    Set Col = TDBGrid.Columns(2)
    Col.Caption = "Field Name"
    Col.Locked = True
    Col.AutoSize
    
    Set Col = TDBGrid.Columns(3)
    Col.Caption = "Data Type"
    Col.Locked = True
    Col.AutoSize
    Col.Visible = False
    
    Set Col = TDBGrid.Columns(4)
    Col.Caption = "Alignment"
    Col.Locked = True
    Col.AutoSize
    Col.Visible = False
    
    Set Col = TDBGrid.Columns(5)
    Col.Caption = "Sample Value"
    Col.Locked = True
    Col.AutoSize

    'TDBGrid.HoldFields
    'InitGrid
    
    lblResultset.Caption = "Resultset (" & m_data_rec.RecordCount & " records):"
    
End Sub

Private Sub InitGrid()
' Initialize grid
    Dim Col As TrueOleDBGrid80.Column
   
    TDBGrid.Appearance = dbgXPTheme ' Confirm to themes on XP/standard 3D on 2000
    TDBGrid.DeadAreaBackColor = vbApplicationWorkspace
    TDBGrid.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid.AlternatingRowStyle = True
    TDBGrid.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid.EvenRowStyle.BackColor = g_intAlternateRowColor ' Make even rows 'gray'
    TDBGrid.ScrollBars = dbgAutomatic
    TDBGrid.TabAcrossSplits = True
    TDBGrid.TabAction = dbgGridNavigation ' Tab moves from column to column
    TDBGrid.WrapCellPointer = True ' Wrap from end of row to beginning of next row
    TDBGrid.AnchorRightColumn = True
    TDBGrid.FetchRowStyle = True
    TDBGrid.AllowColMove = False
    TDBGrid.AllowAddNew = False
    
    TDBGrid.DataSource = m_rec
    TDBGrid.HoldFields
    TDBGrid.ReBind
    'TDBGrid.Bookmark = Null

    TDBGrid1.AlternatingRowStyle = True
    TDBGrid1.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid1.EvenRowStyle.BackColor = g_intAlternateRowColor ' Make even rows 'gray'
    TDBGridFilter.AlternatingRowStyle = True
    TDBGridFilter.OddRowStyle.BackColor = vbWindowBackground
    TDBGridFilter.EvenRowStyle.BackColor = g_intAlternateRowColor ' Make even rows 'gray'

    TDBGridFilter.DataSource = m_filter_rec
    TDBGridFilter.HoldFields
    TDBGridFilter.ReBind
    TDBGrid1.DataSource = m_data_rec
    TDBGrid1.ReBind
    
    InitFilterGrid
    UpdateFilterset

End Sub

Private Sub InitFilterGrid()
    Const OPERATORS = "=,<,>,<=,>=,<>,LIKE"
    Dim Col As TrueOleDBGrid80.Column
    Dim Item As New TrueOleDBGrid80.ValueItem
    Dim aOperators As Variant
    Dim I As Long
    
    If TDBGridFilter.Columns.Count = 1 Then
        Set Col = TDBGridFilter.Columns(0)
        Col.DataField = "Field"
        Col.Caption = "Database Field"
        Col.Visible = True
        Set Col = TDBGridFilter.Columns.Add(1)
        Col.DataField = "Operator"
        Col.Caption = "Operator"
        Col.Width = 700
        Col.DefaultValue = "="
        Col.Visible = True
        Set Col = TDBGridFilter.Columns.Add(2)
        Col.DataField = "Value"
        Col.Caption = "Value"
        Col.Visible = True
    End If
    
    'Populate Field drop-down box with the list of database fields
    TDBGridFilter.Columns("Field").ValueItems.Clear
    TDBGrid.MoveFirst
    Do While Not TDBGrid.EOF
        Item.Value = TDBGrid.Columns("Field").Value
        TDBGridFilter.Columns("Field").ValueItems.Add Item
        TDBGrid.MoveNext
    Loop
    TDBGridFilter.Columns("Field").ValueItems.Presentation = dbgComboBox
    TDBGridFilter.Columns("Field").ValueItems.Validate = True
    TDBGridFilter.Columns("Field").AutoDropDown = True
    
    'Populate Operator drop-down box with the list of operators
    TDBGridFilter.Columns("Operator").ValueItems.Clear
    aOperators = Split(OPERATORS, ",")
    For I = LBound(aOperators) To UBound(aOperators)
        Item.Value = aOperators(I)
        TDBGridFilter.Columns("Operator").ValueItems.Add Item
    Next
    TDBGridFilter.Columns("Operator").ValueItems.Presentation = dbgComboBox
    TDBGridFilter.Columns("Operator").ValueItems.Validate = True
    TDBGridFilter.Columns("Operator").AutoDropDown = True
    
End Sub

Private Sub UpdateFilterset()
    Dim rsTemp As New ADODB.RecordSet
    Dim fld As ADODB.Field
    Dim Col As TrueOleDBGrid80.Column
    Dim sFilter As String
    Dim sField As String
    Dim sValue As String
    Dim sDelimiter As String
    Dim bIsString As Boolean
    Dim bIsDate As Boolean
    Dim aTokens As Variant
    Dim I As Long
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    TDBGridFilter.Update
    Set m_data_rec = m_original_rec
    
    'Build Filter String
    TDBGridFilter.MoveFirst
    Do While Not TDBGridFilter.EOF
        sFilter = ""
        sField = TDBGridFilter.Columns("Field")
        sValue = Replace(TDBGridFilter.Columns("Value"), "'", "''")
        If TDBGridFilter.Columns("Operator") = "" Then
            TDBGridFilter.Columns("Operator") = "="
        End If
        sDelimiter = ""
        bIsDate = (m_data_rec.Fields(sField).Type = adDate) Or _
            (m_data_rec.Fields(sField).Type = adDBDate) Or _
            (m_data_rec.Fields(sField).Type = adDBTime) Or _
            (m_data_rec.Fields(sField).Type = adDBTimeStamp)
        bIsString = (m_data_rec.Fields(sField).Type = adChar) Or _
            (m_data_rec.Fields(sField).Type = adVarChar) Or _
            (m_data_rec.Fields(sField).Type = adVarWChar) Or _
            (m_data_rec.Fields(sField).Type = adWChar)
        If bIsDate Then sDelimiter = "#"
        If bIsString Then sDelimiter = "'"
        If bIsString And InStr(sValue, ",") > 0 Then
            aTokens = Split(sValue, ",")
            For I = LBound(aTokens) To UBound(aTokens)
                If I > LBound(aTokens) Then sFilter = sFilter & " OR "
                sValue = aTokens(I)
                sFilter = sFilter & sField & " " & TDBGridFilter.Columns("Operator") & " "
                sFilter = sFilter & sDelimiter & sValue & sDelimiter
            Next
        Else
            If sDelimiter = "" And IsNumeric(sValue) And InStr(sValue, ",") > 0 Then
                sValue = Replace(sValue, ",", "")
            End If
            sFilter = sFilter & sField & " " & TDBGridFilter.Columns("Operator") & " "
            sFilter = sFilter & sDelimiter & sValue & sDelimiter
        End If
        'Debug.Print m_data_rec.RecordCount & ">" & sFilter
        m_data_rec.Filter = sFilter
        'rsTemp.Delete adAffectAll
        Set rsTemp = New ADODB.RecordSet
        CopyRSFields rsTemp, m_data_rec     ' Copy the recordset schema to rsTemp
        rsTemp.Delete       ' CopyRSFields adds a blank record -- delete it
        Do While Not m_data_rec.EOF
            rsTemp.AddNew
            For Each fld In m_data_rec.Fields
                rsTemp.Fields(fld.Name).Value = fld.Value
            Next
            rsTemp.Update
            m_data_rec.MoveNext
        Loop
        Set m_data_rec = rsTemp.Clone
        rsTemp.Close
        TDBGridFilter.MoveNext
    Loop
    
    'Refresh the Resultset Grid
    'Debug.Print m_data_rec.RecordCount & ">" & m_data_rec.Filter
    m_data_rec.Filter = ""
    TDBGrid1.DataSource = m_data_rec
    TDBGrid1.ReBind
    TDBGrid1.ApproxCount = m_data_rec.RecordCount
    lblResultset.Caption = "Resultset (" & m_data_rec.RecordCount & " records):"
    For Each Col In TDBGrid1.Columns
        Col.Visible = False
    Next
    
    TDBGrid.MoveFirst
    Do While Not TDBGrid.EOF
        If TDBGrid.Columns("Export") Then
            TDBGrid1.Columns(TDBGrid.Columns("Field").Value).Visible = True
            TDBGrid1.Columns(TDBGrid.Columns("Field").Value).Order = TDBGrid1.Columns.Count
        End If
        TDBGrid.MoveNext
    Loop
    
    TDBGrid1.MoveFirst
    TDBGrid1.LeftCol = 0
    'Set rsTemp = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Err_Handler:
    'Set rsTemp = Nothing
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation
    Exit Sub
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Set frmExport = Nothing
End Sub

Private Sub cmdExport_Click()
    SaveFile
End Sub

Private Sub cmdRefresh_Click()
    UpdateFilterset
End Sub

Private Sub Form_Activate()
    ShowToolbarIcons True
End Sub

Private Sub Form_Deactivate()
    ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()
    ' Setup Export Setup Grid recordset
    Set m_rec = New ADODB.RecordSet
    If m_rec.Fields.Count = 0 Then
        m_rec.Fields.Append "Export", adBoolean
        m_rec.Fields.Append "Name", adVarChar, 100
        m_rec.Fields.Append "Field", adVarChar, 100
        m_rec.Fields.Append "DataType", adInteger
        m_rec.Fields.Append "Alignment", adInteger
        m_rec.Fields.Append "Value", adVarChar, 100
    End If
    
    ' Setup Filters Grid recordset
    Set m_filter_rec = New ADODB.RecordSet
    If m_filter_rec.Fields.Count = 0 Then
        'm_filter_rec.Fields.Append "And", adBoolean
        'm_filter_rec.Fields.Append "Name", adVarChar, 100
        m_filter_rec.Fields.Append "Field", adVarChar, 100
        m_filter_rec.Fields.Append "Operator", adVarChar, 5
        m_filter_rec.Fields.Append "Value", adVarChar, 255
        m_filter_rec.Open
    End If
End Sub

Private Sub Form_Load()
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    
    InitGrid
    InitFilterGrid
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'TDBGridFilter.Top = TDBGridFilter.Top
    TDBGridFilter.Height = Me.Height - TDBGridFilter.Top - TDBGrid1.Height - (TDBGrid.Left * 5)
    TDBGridFilter.Left = Me.Width - TDBGridFilter.Width - TDBGrid.Left
    'lblFilters.Top = TDBGridFilter.Top - TDBGrid.Left
    lblFilters.Left = TDBGridFilter.Left
    
    TDBGrid.Height = Me.Height - TDBGrid.Top - TDBGrid1.Height - (TDBGrid.Left * 5)
    TDBGrid.Width = cmdExport.Left - (TDBGrid.Left * 3)
    
    TDBGrid1.Top = Me.Height - TDBGrid1.Height - (TDBGrid1.Left * 3)
    TDBGrid1.Width = Me.Width - (TDBGrid1.Left * 2)
    lblResultset.Top = TDBGrid1.Top - TDBGrid.Left
    
    cmdRefresh.Top = TDBGridFilter.Top + TDBGridFilter.Height + 75
    cmdRefresh.Left = Me.Width - cmdRefresh.Width - TDBGrid.Left
    
    cmdExport.Left = TDBGridFilter.Left ' Me.Width - cmdExport.Width - (TDBGrid.Left * 2)
    cmdClose.Left = cmdExport.Left
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowToolbarIcons False
End Sub

Private Sub ShowToolbarIcons(bShowIcons As Boolean)

    fMainForm.tbToolBar.Buttons.Item(tbrSAVE).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrSAVE).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrEXPORT).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrEXPORT).Visible = False
    fMainForm.mnuFilePageSetup.Enabled = bShowIcons
    fMainForm.mnuFilePrint.Enabled = bShowIcons
    fMainForm.mnuFilePrintPreview.Enabled = bShowIcons

End Sub

Private Function GetFileExtension(sFilename As String) As String
    Dim P As Long
    
    P = InStrRev(sFilename, ".")
    If P > 0 Then
        GetFileExtension = Mid(sFilename, P + 1)
    Else
        GetFileExtension = ""
    End If

End Function

Private Function ExportToCsv(sFilename As String) As Boolean
    Dim f As Long
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    f = FreeFile
    Open sFilename For Output As #f
    m_rec.MoveFirst
    Do While Not m_rec.EOF
        If m_rec.Fields("Export") Then
            Write #f, m_rec.Fields("Name").Value;
        End If
        m_rec.MoveNext
    Loop
    Write #f,
    m_data_rec.MoveFirst
    Do While Not m_data_rec.EOF
        m_rec.MoveFirst
        Do While Not m_rec.EOF
            If m_rec.Fields("Export") Then
                Write #f, m_data_rec.Fields(m_rec.Fields("Field").Value).Value & "";
            End If
            m_rec.MoveNext
        Loop
        Write #f,
        m_data_rec.MoveNext
    Loop
    Close #f
    Screen.MousePointer = vbDefault
    ExportToCsv = True
    Exit Function
    
Err_Handler:
    ExportToCsv = False
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation
    Exit Function
    
End Function

Private Function ExportToHtml(sFilename As String) As Boolean
    Dim f As Long
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    f = FreeFile
    Open sFilename For Output As #f
    Print #f, "<html>"
    Print #f, "<head>"
    Print #f, "   <title>" & m_Title & "</title>"
    Print #f, "</head>"
    Print #f, "<body>"
    Print #f, "   <table border=""1"" style=""font-family: Tahoma; font-size: 8pt"">"
    Print #f, "   <tr bgcolor=""silver"">"
    m_rec.MoveFirst
    Do While Not m_rec.EOF
        If m_rec.Fields("Export") Then
            Print #f, "      <td>" & m_rec.Fields("Name").Value & "</td>"
        End If
        m_rec.MoveNext
    Loop
    Print #f, "   </tr>"
    m_data_rec.MoveFirst
    Do While Not m_data_rec.EOF
        m_rec.MoveFirst
        Print #f, "   <tr>"
        Do While Not m_rec.EOF
            If m_rec.Fields("Export") Then
                If m_rec.Fields("Alignment") = 1 Then
                    Print #f, "      <td align=""right"">";
                Else
                    Print #f, "      <td>";
                End If
                Print #f, m_data_rec.Fields(m_rec.Fields("Field").Value).Value & "</td>"
            End If
            m_rec.MoveNext
        Loop
        Print #f, "   </tr>"
        m_data_rec.MoveNext
    Loop
    Print #f, "   </table>"
    Print #f, "</body>"
    Print #f, "</html>"
    Close #f
    Screen.MousePointer = vbDefault
    ExportToHtml = True
    Exit Function
    
Err_Handler:
    ExportToHtml = False
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation
    Exit Function
    
End Function

Private Function ExportToXml(sFilename As String) As Boolean
    Dim f As Long
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    f = FreeFile
    Open sFilename For Output As #f
    Print #f, "<?xml version=""1.0"" ?>"
    Print #f, "<workbook>"
    Print #f, "   <worksheet>"
    m_data_rec.MoveFirst
    Do While Not m_data_rec.EOF
        m_rec.MoveFirst
        Print #f, "   <row>"
        Do While Not m_rec.EOF
            If m_rec.Fields("Export") Then
                Print #f, "      <" & m_rec.Fields("Field").Value & ">";
                Print #f, "" & Replace(m_data_rec.Fields(m_rec.Fields("Field").Value).Value & "", "&", "&amp;") & "";
                Print #f, "</" & m_rec.Fields("Field").Value & ">"
            End If
            m_rec.MoveNext
        Loop
        Print #f, "   </row>"
        m_data_rec.MoveNext
    Loop
    Print #f, "   </worksheet>"
    Print #f, "</workbook>"
    Close #f
    Screen.MousePointer = vbDefault
    ExportToXml = True
    Exit Function
    
Err_Handler:
    ExportToXml = False
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation
    Exit Function
    
End Function

Public Function ExportToExcel(sFilename As String) As Boolean
    Dim f As Long
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    f = FreeFile
    Open sFilename For Output As #f
    m_rec.MoveFirst
    Do While Not m_rec.EOF
        If m_rec.Fields("Export") Then
            Print #f, m_rec.Fields("Name").Value & vbTab;
        End If
        m_rec.MoveNext
    Loop
    Print #f,
    m_data_rec.MoveFirst
    Do While Not m_data_rec.EOF
        m_rec.MoveFirst
        Do While Not m_rec.EOF
            If m_rec.Fields("Export") Then
                Print #f, m_data_rec.Fields(m_rec.Fields("Field").Value).Value & "" & vbTab;
            End If
            m_rec.MoveNext
        Loop
        Print #f,
        m_data_rec.MoveNext
    Loop
    Close #f

    Screen.MousePointer = vbDefault
    ExportToExcel = True
    Exit Function
    
Err_Handler:
    ExportToExcel = False
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation
    Exit Function
    
End Function

Public Sub SaveFile()
    Dim sFilename As String
    Dim sFileExtension As String
    
    On Error Resume Next
    With CommonDialog1
        .Filter = "CSV File (*.csv)|*.csv|Excel File (*.xls)|*.xls|HTML File (*.htm)|*.htm|XML File (*.xml)|*.xml"
        .FilterIndex = 2
        .DefaultExt = ".htm"
        .CancelError = True
        .FileName = m_Title
        .DialogTitle = "Export data to file..."
        .ShowSave
    End With
    If (CommonDialog1.FileName <> "") And (Err.Number = 0) Then
        sFilename = CommonDialog1.FileName
        sFileExtension = GetFileExtension(sFilename)
        Select Case sFileExtension
        Case "csv"
            If ExportToCsv(sFilename) Then
                MsgBox "Data was successfully exported to file:" & vbCrLf & sFilename, vbInformation
            End If
        Case "htm", "html"
            If ExportToHtml(sFilename) Then
                MsgBox "Data was successfully exported to file:" & vbCrLf & sFilename, vbInformation
            End If
        Case "xml"
            If ExportToXml(sFilename) Then
                MsgBox "Data was successfully exported to file:" & vbCrLf & sFilename, vbInformation
            End If
        Case "xls"
            If ExportToExcel(sFilename) Then
                MsgBox "Data was successfully exported to file:" & vbCrLf & sFilename, vbInformation
            End If
        Case ""
            
        Case Else
            MsgBox "Unsupported export type '" & sFileExtension & "'", vbExclamation
        End Select
    End If
    
End Sub

Private Sub TDBGrid_AfterColUpdate(ByVal ColIndex As Integer)
    
    If ColIndex = 0 Then
        ' if Export checkbox changed, then update filterset
        UpdateFilterset
    End If
    
End Sub

Private Sub TDBGridFilter_AfterUpdate()
    UpdateFilterset
End Sub

