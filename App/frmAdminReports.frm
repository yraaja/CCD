VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmAdminReports 
   Caption         =   "Report Manager"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdminReports.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   12450
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   450
      Left            =   7080
      TabIndex        =   18
      Top             =   6000
      Width           =   1110
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   450
      Left            =   5760
      TabIndex        =   17
      Top             =   6000
      Width           =   1110
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   450
      Left            =   8400
      TabIndex        =   19
      Top             =   6000
      Width           =   1110
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Setup"
      Height          =   2295
      Left            =   4680
      TabIndex        =   1
      Top             =   960
      Width           =   7815
      Begin VB.CommandButton cmdCopyParameters 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         MaskColor       =   &H00FF00FF&
         Picture         =   "frmAdminReports.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Copy Parameters"
         Top             =   1800
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox txtReportID 
         Height          =   285
         Left            =   6960
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdReportBrowser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         MaskColor       =   &H00FF00FF&
         Picture         =   "frmAdminReports.frx":0984
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Browse for File"
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox txtReportCategory 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox cboReportStoredProc 
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   1800
         Width           =   4095
      End
      Begin VB.ComboBox cboReportDefName 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtReportFile 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtReportName 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report ID"
         Height          =   195
         Left            =   6120
         TabIndex        =   4
         Top             =   400
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report &Category"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   760
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stored &Procedure"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   1840
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report &Definition"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report &File"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1120
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report &Name"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   400
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stored Procedure Parameters"
      Height          =   2415
      Left            =   4680
      TabIndex        =   15
      Top             =   3360
      Width           =   6135
      Begin TrueOleDBGrid80.TDBGrid TDBGrid 
         Height          =   1815
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3201
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Prm#"
         Columns(0).DataField=   "parameter_number"
         Columns(0).DataWidth=   3
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Parameter Name"
         Columns(1).DataField=   "parameter_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Parameter Label"
         Columns(2).DataField=   "parameter_label"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Default"
         Columns(3).DataField=   "parameter_default_value"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   1
         Columns(4)._MaxComboItems=   5
         Columns(4).ValueItems(0)._DefaultItem=   0
         Columns(4).ValueItems(0).Value=   "string"
         Columns(4).ValueItems(0).Value.vt=   8
         Columns(4).ValueItems(0).DisplayValue=   "string"
         Columns(4).ValueItems(0).DisplayValue.vt=   8
         Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems(1)._DefaultItem=   0
         Columns(4).ValueItems(1).Value=   "number"
         Columns(4).ValueItems(1).Value.vt=   8
         Columns(4).ValueItems(1).DisplayValue=   "number"
         Columns(4).ValueItems(1).DisplayValue.vt=   8
         Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems.Count=   2
         Columns(4).Caption=   "DataType"
         Columns(4).DataField=   "parameter_data_type"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   17
         Columns(5)._MaxComboItems=   5
         Columns(5).ValueItems(0)._DefaultItem=   0
         Columns(5).ValueItems(0).Value=   "1"
         Columns(5).ValueItems(0).Value.vt=   8
         Columns(5).ValueItems(0).DisplayValue=   "Text Box"
         Columns(5).ValueItems(0).DisplayValue.vt=   8
         Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(1)._DefaultItem=   0
         Columns(5).ValueItems(1).Value=   "2"
         Columns(5).ValueItems(1).Value.vt=   8
         Columns(5).ValueItems(1).DisplayValue=   "Check Box"
         Columns(5).ValueItems(1).DisplayValue.vt=   8
         Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(2)._DefaultItem=   0
         Columns(5).ValueItems(2).Value=   "3"
         Columns(5).ValueItems(2).Value.vt=   8
         Columns(5).ValueItems(2).DisplayValue=   "Date Picker"
         Columns(5).ValueItems(2).DisplayValue.vt=   8
         Columns(5).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(3)._DefaultItem=   0
         Columns(5).ValueItems(3).Value=   "12"
         Columns(5).ValueItems(3).Value.vt=   8
         Columns(5).ValueItems(3).DisplayValue=   "MasterFormat"
         Columns(5).ValueItems(3).DisplayValue.vt=   8
         Columns(5).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(4)._DefaultItem=   0
         Columns(5).ValueItems(4).Value=   "10"
         Columns(5).ValueItems(4).Value.vt=   8
         Columns(5).ValueItems(4).DisplayValue=   "Class System"
         Columns(5).ValueItems(4).DisplayValue.vt=   8
         Columns(5).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(5)._DefaultItem=   0
         Columns(5).ValueItems(5).Value=   "20"
         Columns(5).ValueItems(5).Value.vt=   8
         Columns(5).ValueItems(5).DisplayValue=   "Quarters"
         Columns(5).ValueItems(5).DisplayValue.vt=   8
         Columns(5).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(6)._DefaultItem=   0
         Columns(5).ValueItems(6).Value=   "21"
         Columns(5).ValueItems(6).Value.vt=   8
         Columns(5).ValueItems(6).DisplayValue=   "Countries"
         Columns(5).ValueItems(6).DisplayValue.vt=   8
         Columns(5).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(7)._DefaultItem=   0
         Columns(5).ValueItems(7).Value=   "22"
         Columns(5).ValueItems(7).Value.vt=   8
         Columns(5).ValueItems(7).DisplayValue=   "States"
         Columns(5).ValueItems(7).DisplayValue.vt=   8
         Columns(5).ValueItems(7)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(8)._DefaultItem=   0
         Columns(5).ValueItems(8).Value=   "24"
         Columns(5).ValueItems(8).Value.vt=   8
         Columns(5).ValueItems(8).DisplayValue=   "Cities"
         Columns(5).ValueItems(8).DisplayValue.vt=   8
         Columns(5).ValueItems(8)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(9)._DefaultItem=   0
         Columns(5).ValueItems(9).Value=   "11"
         Columns(5).ValueItems(9).Value.vt=   8
         Columns(5).ValueItems(9).DisplayValue=   "CCI Geo System"
         Columns(5).ValueItems(9).DisplayValue.vt=   8
         Columns(5).ValueItems(9)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(10)._DefaultItem=   0
         Columns(5).ValueItems(10).Value=   "23"
         Columns(5).ValueItems(10).Value.vt=   8
         Columns(5).ValueItems(10).DisplayValue=   "CCI Index Classes"
         Columns(5).ValueItems(10).DisplayValue.vt=   8
         Columns(5).ValueItems(10)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(11)._DefaultItem=   0
         Columns(5).ValueItems(11).Value=   "25"
         Columns(5).ValueItems(11).Value.vt=   8
         Columns(5).ValueItems(11).DisplayValue=   "Labor Trade IDs"
         Columns(5).ValueItems(11).DisplayValue.vt=   8
         Columns(5).ValueItems(11)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems.Count=   12
         Columns(5).Caption=   "Appearance"
         Columns(5).DataField=   "parameter_appearance"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   4
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Required"
         Columns(6).DataField=   "parameter_required"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Visible"
         Columns(7).DataField=   "parameter_visible"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=953"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2963"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2884"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1138"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1058"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2117"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2037"
         Splits(0)._ColumnProps(20)=   "Column(4).Button=1"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(5).Button=1"
         Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(27)=   "Column(6).Width=1323"
         Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=1244"
         Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(31)=   "Column(7).Width=1164"
         Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=1085"
         Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
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
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   450
      Left            =   9720
      TabIndex        =   20
      Top             =   6000
      Width           =   1110
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminReports.frx":0EC6
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminReports.frx":12AE
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminReports.frx":169A
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminReports.frx":1A63
            Key             =   "Leaf"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTreeCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Report:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   840
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Manager"
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
      TabIndex        =   21
      Top             =   120
      Width           =   2265
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11100
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmAdminReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iReportID As Long
Private m_parm_rec As New ADODB.RecordSet   ' Recordset to hold parameters

Dim m_sngYCoord As Single
'
'   Keeps up with the field that last had focus when form
'   is deactivate, so when activated can set focus.
Dim m_strCurrentFormControl As String
'
'   Notifies that it wants to see changes.
Dim sEventSubscriberID As String
'

Public Function UpdateMasterTable() As Boolean
    Dim strUpdate As String     ' SQL string
    Dim strError As String      ' Error string returned from DAL
    Dim intErrors As Integer    ' Tracks if any errors have occurred
    Dim intSuccess As Integer   ' Tracks successful updates
    Dim blnReturn As Boolean
    Dim rec As ADODB.RecordSet
    
    If m_iReportID = 0 Then
        strUpdate = "INSERT INTO REPORT_MASTER "
        strUpdate = strUpdate & " report_name = 'NEWREPORTNAME'"
        blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
        If blnReturn = True Then
            intSuccess = intSuccess + 1
            strUpdate = "SELECT report_id FROM REPORT_MASTER "
            strUpdate = strUpdate & " WHERE report_name = 'NEWREPORTNAME'"
            blnReturn = g_objDAL.GetRecordset(vbNullString, strUpdate, rec)
            If blnReturn Then
                m_iReportID = rec.Fields("report_id")
            End If
        Else
            intErrors = intErrors + 1
            Exit Function
        End If
    End If
    strUpdate = "UPDATE REPORT_MASTER SET "
    strUpdate = strUpdate & " report_name = '" & SQLFixString(Me.txtReportName.Text) & "',"
    strUpdate = strUpdate & " report_file_name = '" & SQLFixString(Me.txtReportFile.Text) & "',"
    strUpdate = strUpdate & " report_file_def_name = '" & SQLFixString(Me.cboReportDefName.Text) & "',"
    strUpdate = strUpdate & " report_stored_proc = '" & SQLFixString(Me.cboReportStoredProc.Text) & "',"
    strUpdate = strUpdate & " report_category = '" & SQLFixString(Me.txtReportCategory.Text) & "'"
    strUpdate = strUpdate & " WHERE report_id = " & m_iReportID & ""
    blnReturn = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
    If blnReturn = True Then
        intSuccess = intSuccess + 1
    Else
        intErrors = intErrors + 1
    End If
    
    UpdateMasterTable = (intErrors = 0)

End Function

Public Function UpdateTable() As Boolean
    Const TABLE_NAME = "REPORT_PARAMETERS"
    Const KEY_FIELD = "report_parameter_id"
    Dim strUpdate As String     ' SQL string
    Dim blnReturn As Boolean
    Dim blnUpdateRow As Boolean
    Dim fld As ADODB.Field
    Dim strError As String      ' Error string returned from DAL
    Dim intErrors As Integer    ' Tracks if any errors have occurred
    Dim intSuccess As Integer   ' Tracks successful updates
    Dim m_rec As ADODB.RecordSet
    Dim sValue As String
    
    'On Error Resume Next
    UpdateTable = True
    intErrors = 0
    intSuccess = 0
    
    Set m_rec = m_parm_rec
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
                    If LCase(fld.Name) <> LCase(KEY_FIELD) Then
                        If fld.Name = "last_update_person" Then
                            fld.Value = strUserName
                        End If
                        sValue = SQLFixString(fld.Value & "")
                        If fld.Type = adBoolean Then
                            If fld.Value Then
                                sValue = 1
                            Else
                                sValue = 0
                            End If
                        End If
                        strUpdate = strUpdate & " " & fld.Name & "='" & sValue & "',"
                    End If
                    If fld.Name = "Parameter_Visible" Then Exit For
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
        MsgBox strMsg, vbInformation + vbOKOnly
    End If
    
End Function


Private Function AddNewReport() As Long

    m_iReportID = 0
    RemoveReportParameters
    UnLockField Me, "txtReportName"
    UnLockField Me, "txtReportCategory"
    UnLockField Me, "txtReportFile"
    UnLockField Me, "cboReportDefName"
    UnLockField Me, "cboReportStoredProc"
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdReportBrowser.Enabled = True
    cmdCopyParameters.Enabled = True
    txtReportName.SetFocus
    
End Function

Private Function LoadStoredProcedureCombo() As Boolean
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    Dim strSELECT As String
    
    On Error GoTo Err_Handler
    cboReportStoredProc.Clear
    strSELECT = "SELECT name FROM sysobjects WHERE type = 'P' AND category = 0 ORDER BY name"
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
    If blnReturn Then
        Do While Not rec.EOF
            cboReportStoredProc.AddItem rec.Fields("name")
            rec.MoveNext
        Loop
    End If
    rec.Close
    Set rec = Nothing
    LoadStoredProcedureCombo = True
    Exit Function
    
Err_Handler:
    MsgBox "An error occurred while loading Stored Procedure combo:" & vbCrLf & Err.Description, vbCritical
    LoadStoredProcedureCombo = False
    Exit Function
    
End Function

Private Function GetStoredProcParameters()
    Dim sStoredProc As String
    Dim iReportID As Long
    Dim iParamNumber As Long
    Dim sSQL As String
    Dim sPrmLabel As String
    Dim sPrmType As String
    Dim iPrmAppearance As Integer
    
    iReportID = m_iReportID
    If m_iReportID = 0 Then
        MsgBox "You can not perform this function until a Report ID has been generated.", vbExclamation + vbOKOnly
        Exit Function
    End If
    
    sStoredProc = cboReportStoredProc.Text
    If sStoredProc = "" Then Exit Function
    
    If m_parm_rec.RecordCount > 0 Then
        If MsgBox("Are you sure you want to load the Parameters? This will clear any customizations.", vbQuestion + vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim prm As ADODB.Parameter
    
    con.Open CONNECT
    ' DELETE EXISTING PARAMETER RECORDS
    con.Execute "DELETE FROM REPORT_PARAMETERS WHERE (report_id=" & iReportID & ")"
    
    cmd.ActiveConnection = con
    cmd.CommandText = sStoredProc
    cmd.CommandType = adCmdStoredProc
    ' INSERT STORED PROC PARAMETER RECORDS
    For Each prm In cmd.Parameters
        If Left(prm.Name, 1) = "@" Then
            iParamNumber = iParamNumber + 1
            Select Case prm.Type
            Case adBoolean 'boolean
                iPrmAppearance = 2
                sPrmType = "number"
            Case adDate
                iPrmAppearance = 3
                sPrmType = "string"
            Case adChar, adVarChar, adWChar, adVarWChar, adLongVarChar, adLongVarWChar
                iPrmAppearance = 1
                sPrmType = "string"
            Case Else
                iPrmAppearance = 1
                sPrmType = "number"
            End Select
            sPrmLabel = Mid(prm.Name, 2)
            sPrmLabel = StrConv(Replace(sPrmLabel, "_", " "), vbProperCase)
            sSQL = "INSERT INTO REPORT_PARAMETERS"
            sSQL = sSQL & " (report_id, parameter_number, parameter_name, parameter_label, parameter_data_type, parameter_appearance, parameter_default_value, parameter_visible)"
            sSQL = sSQL & " VALUES (" & iReportID & ", " & iParamNumber & ","
            sSQL = sSQL & " '" & prm.Name & "',"
            sSQL = sSQL & " '" & sPrmLabel & "',"
            sSQL = sSQL & " '" & sPrmType & "',"
            sSQL = sSQL & " '" & iPrmAppearance & "',"
            sSQL = sSQL & " '', 1)"
            con.Execute sSQL
        End If
    Next
    
    con.Close
    Set prm = Nothing
    Set cmd = Nothing
    Set con = Nothing
    'REQUERY PARAMETER GRID
    If g_objDAL.GetRecordset(vbNullString, "exec usp_select_report_parameters @report_id = " & iReportID, m_parm_rec) Then
        TDBGrid.DataSource = m_parm_rec
        TDBGrid.ReBind
    End If
    Screen.MousePointer = vbDefault
    Exit Function
    
Err_Handler:
    Screen.MousePointer = vbDefault

End Function

Private Function GetReportsFromXml() As Long
    Dim sPath As String
    Dim sFile As String
    Dim objDom As Object
    Dim objNodeList As Variant
    Dim Node As Variant
    Dim iReportCount As Long
    
    sPath = App.Path
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    sFile = Me.txtReportFile.Text
    On Error GoTo Err_Handler
    Set objDom = CreateObject("MSXML2.DOMDocument")
    objDom.Async = False
    objDom.Load sPath & sFile
    Set objNodeList = objDom.selectNodes("//Report/Name")
    iReportCount = objNodeList.Length
    If iReportCount > 0 Then
        For Each Node In objNodeList
            cboReportDefName.AddItem Node.Text
        Next
    End If
    GetReportsFromXml = iReportCount
    Set Node = Nothing
    Set objNodeList = Nothing
    Set objDom = Nothing
    Exit Function

Err_Handler:
    Exit Function

End Function

Private Sub LoadReportParameters(iReportID As Long)
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    Dim strSELECT As String
    
    Screen.MousePointer = vbHourglass
    strSELECT = "SELECT * FROM REPORT_MASTER WHERE Report_ID = " & iReportID
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
    If blnReturn Then
        ' Populate Settings
        Me.txtReportID.Text = rec.Fields("report_id")
        Me.txtReportName.Text = rec.Fields("report_name")
        Me.txtReportCategory.Text = rec.Fields("report_category")
        Me.txtReportFile.Text = rec.Fields("report_file_name")
        Me.cboReportDefName.Clear
        Me.cboReportDefName.Text = rec.Fields("report_file_def_name")
        Me.cboReportStoredProc.Text = rec.Fields("report_stored_proc")
        ' Populate Parameters Grid
        strSELECT = "exec usp_select_report_parameters @report_id = " & iReportID
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, m_parm_rec)
        TDBGrid.DataSource = m_parm_rec
        TDBGrid.ReBind
        GetReportsFromXml
        ' Set Controls
        Frame1.Enabled = True
        Frame2.Enabled = True
        UnLockField Me, "txtReportName"
        UnLockField Me, "txtReportCategory"
        UnLockField Me, "txtReportFile"
        UnLockField Me, "cboReportDefName"
        UnLockField Me, "cboReportStoredProc"
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
        cmdReportBrowser.Enabled = True
        cmdCopyParameters.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub RemoveReportParameters()
    
    Me.txtReportID.Text = ""
    Me.txtReportName.Text = ""
    Me.txtReportCategory.Text = ""
    Me.txtReportFile.Text = ""
    Me.cboReportDefName.Clear
    Me.cboReportDefName.Text = ""
    Me.cboReportStoredProc.Text = ""
    
    TDBGrid.DataSource = Nothing
    TDBGrid.ReBind
    
    LockField Me, "txtReportName"
    LockField Me, "txtReportCategory"
    LockField Me, "txtReportFile"
    LockField Me, "cboReportDefName"
    LockField Me, "cboReportStoredProc"
    Frame1.Enabled = False
    Frame2.Enabled = False
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdReportBrowser.Enabled = False
    cmdCopyParameters.Enabled = False
        
End Sub

Private Sub LoadReportTree()
    Dim blnReturn As Boolean
    Dim strSELECT As String
    Dim rsTree As New ADODB.RecordSet
    Dim strCategory As String
    Dim strParent As String
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    With TreeView1
        ' Reset the Tree; Add Root Node [K0]
        .Nodes.Clear
        .Nodes.Add , , "K0", "Reports", "Closed", "Open"
        strParent = "K0"
        ' Load the Tree
        strSELECT = "SELECT * FROM REPORT_MASTER" & _
                    " WHERE (Report_Category IS NOT NULL)" & _
                    " ORDER BY Report_Category, Report_Name"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rsTree)
        If blnReturn = False Then
            MsgBox "An error occurred while searching:" & vbCrLf & g_objDAL.LastErrorDescription, vbCritical
            Screen.MousePointer = vbNormal
            Exit Sub
        Else
            strCategory = ""
            Do While Not rsTree.EOF
                If rsTree.Fields("report_category") <> strCategory Then
                    'Create a "folder" node for the new category
                    strParent = "K" & Replace(rsTree.Fields("report_category"), " ", "_")
                    .Nodes.Add "K0", tvwChild, strParent, rsTree.Fields("report_category"), "Closed", "Open"
                End If
                .Nodes.Add strParent, tvwChild, "K" & rsTree.Fields("report_id"), rsTree.Fields("report_name"), "Leaf"
                strCategory = rsTree.Fields("report_category") & ""
                rsTree.MoveNext
            Loop
        End If
        ' Finish Up
        rsTree.Close
        .Nodes("K0").Expanded = True
        .Nodes("K0").EnsureVisible
    End With
    Screen.MousePointer = vbNormal
    Exit Sub
    
Err_Handler:
    Screen.MousePointer = vbNormal
    MsgBox "An error occurred while loading Report Tree:" & vbCrLf & Err.Description, vbCritical
    Exit Sub
    
End Sub


Private Sub cboReportStoredProc_Change()
    cmdCopyParameters.Enabled = (cboReportStoredProc.Text <> "")
End Sub

Private Sub cmdAdd_Click()
    AddNewReport
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Set frmAdminReports = Nothing
End Sub

Private Sub cmdCopyParameters_Click()
    GetStoredProcParameters
End Sub

Private Sub cmdDelete_Click()
    Dim iReportID As Long
    Dim blnReturn As Boolean
    
    iReportID = m_iReportID
    If iReportID = 0 Then
        'CANCEL NEW REPORT
        RemoveReportParameters
    Else
        'DELETE REPORT
        If MsgBox("Are you sure you want to permanently delete this report?", vbQuestion + vbYesNo, "Delete Report " & iReportID) = vbYes Then
            blnReturn = g_objDAL.ExecQuery(CONNECT, "DELETE FROM REPORT_MASTER WHERE (report_id=" & iReportID & ")")
            blnReturn = g_objDAL.ExecQuery(CONNECT, "DELETE FROM REPORT_PARAMETERS WHERE (report_id=" & iReportID & ")")
            LoadReportTree
        End If
    End If
    
End Sub

Private Sub cmdReportBrowser_Click()
    Dim sFile As String
    
    On Error GoTo Err_Handler
    CommonDialog1.DialogTitle = "Select a Report Definition file"
    CommonDialog1.CancelError = True
    CommonDialog1.DefaultExt = "xml"
    CommonDialog1.Filter = "Report Definition Files (.xml)|*.xml|"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen
    sFile = CommonDialog1.FileTitle
    Me.txtReportFile.Text = sFile
    Me.cboReportDefName.Clear
    GetReportsFromXml
    Exit Sub
    
Err_Handler:
    Exit Sub
    
End Sub

Private Sub cmdUpdate_Click()
    If UpdateMasterTable Then
        UpdateTable
        LoadReportTree
        TreeView1.Nodes("K" & m_iReportID).Selected = True
        SelectNode (m_iReportID)
    End If
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
    
    Status ("Loading Reports Admin Control Panel...")
    Screen.MousePointer = vbHourglass
    sEventSubscriberID = EventSubscriberAdd(Me)
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim blnReturn As Boolean
    
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    LoadStoredProcedureCombo
    LoadReportTree
    RemoveReportParameters
    LockField Me, "txtReportID"
    TDBGrid.AlternatingRowStyle = True
    TDBGrid.OddRowStyle.BackColor = vbWindowBackground
    TDBGrid.EvenRowStyle.BackColor = g_intAlternateRowColor
    
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
        If Me.Width >= 11070 Then
            Line2.X2 = Me.Width - 210
            iDlgUnit = TreeView1.Left
            cmdClose.Left = Me.Width - cmdClose.Width - iDlgUnit '* 2
            cmdClose.Top = Me.Height - cmdClose.Height - iDlgUnit * 3
            cmdUpdate.Left = cmdClose.Left - cmdUpdate.Width - iDlgUnit
            cmdUpdate.Top = cmdClose.Top
            cmdDelete.Left = cmdUpdate.Left - cmdDelete.Width - iDlgUnit
            cmdDelete.Top = cmdClose.Top
            cmdAdd.Left = cmdDelete.Left - cmdAdd.Width - iDlgUnit
            cmdAdd.Top = cmdClose.Top
            Frame1.Left = TreeView1.Left + TreeView1.Width + iDlgUnit
            Frame1.Width = Me.Width - Frame1.Left - iDlgUnit '* 2
            Frame1.Height = cmdClose.Top - Frame1.Top - iDlgUnit
            Frame2.Left = TreeView1.Left + TreeView1.Width + iDlgUnit
            Frame2.Width = Me.Width - Frame1.Left - iDlgUnit '* 2
            TreeView1.Height = cmdClose.Top + cmdClose.Height - TreeView1.Top
            TDBGrid.Width = Frame1.Width - iDlgUnit * 2
            TDBGrid.Height = Frame1.Height - iDlgUnit * 3
        Else
            Me.Width = 11070
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

Private Sub SelectNode(iReportID As Long)
    m_iReportID = iReportID
    If (iReportID > 0) Then
        LoadReportParameters iReportID
    Else
        RemoveReportParameters
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'A TREE NODE HAS BEEN SELECTED, UPDATE DETAILS FRAMES
    Dim iReportID As Long
    Dim sNodeKey As String
    
    sNodeKey = Node.Key
    If IsNumeric(Mid(sNodeKey, 2)) Then
        iReportID = Mid(sNodeKey, 2)
        SelectNode (iReportID)
    Else
        m_iReportID = 0
        RemoveReportParameters
        'cmdPreview.Enabled = False
    End If
    
End Sub
