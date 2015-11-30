VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Construction Cost Database"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9525
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   840
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   41
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0794
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AE6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E38
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":118A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14DC
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":182E
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B80
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ED2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2224
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2576
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":336C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A10
            Key             =   "Fax"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D62
            Key             =   "Info"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4406
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4758
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":514E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":653A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":688C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7282
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7926
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":831C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":866E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":89C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Export"
            Object.ToolTipText     =   "Save to PDF"
            ImageIndex      =   36
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Fax"
            Object.ToolTipText     =   "Fax Report"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "EMail"
            Object.ToolTipText     =   "E-Mail Report"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Excel"
            Object.ToolTipText     =   "Export Data"
            ImageIndex      =   37
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Sort Ascending"
            Object.ToolTipText     =   "Sort Ascending"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Sort Descending"
            Object.ToolTipText     =   "Sort Descending"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrintScreen"
            Object.ToolTipText     =   "Print Screen"
            ImageIndex      =   41
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Database"
            Style           =   4
            Object.Width           =   2500
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "LineBreak"
            Style           =   4
            Object.Width           =   350
         EndProperty
      EndProperty
      Begin VB.TextBox txtDatabase 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Current server and database"
         Top             =   60
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         Caption         =   "lblDatabase"
         Height          =   195
         Left            =   5460
         TabIndex        =   2
         Top             =   60
         Width           =   840
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   5445
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10266
            MinWidth        =   1693
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Current user, server, and database"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   1693
            TextSave        =   "2/17/2012"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   1693
            TextSave        =   "10:42 AM"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
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
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save Report as PDF..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileFax 
         Caption         =   "&Fax Report..."
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintScreen 
         Caption         =   "Print Screen"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewNavTree 
         Caption         =   "&Nav Tree"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewNavMap 
         Caption         =   "Nav &Map"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "&Web Browser"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "F&unctions"
      Begin VB.Menu mnuFunctionsMaterial 
         Caption         =   "&Material"
         Begin VB.Menu mnuFunctionsMaterialPrice 
            Caption         =   "&Price"
         End
         Begin VB.Menu mnuFunctionsMaterialMaintenance 
            Caption         =   "&Maintenance"
         End
         Begin VB.Menu mnuFunctionsMaterialManufacturer 
            Caption         =   "Manu&facturer"
         End
         Begin VB.Menu mnuFunctionsMaterialUsage 
            Caption         =   "Material &Usage"
         End
         Begin VB.Menu mnuFunctionsPubMatRollup 
            Caption         =   "Published Material &Rollup"
         End
         Begin VB.Menu mnuFunctionsMaterialWasIs 
            Caption         =   "&Was/Is"
         End
      End
      Begin VB.Menu mnuFunctionsLabor 
         Caption         =   "&Labor"
         Begin VB.Menu mnuFunctionsLaborRates 
            Caption         =   "&Rates"
         End
         Begin VB.Menu mnuFunctionsTradeGroups 
            Caption         =   "Trade &Groups"
         End
      End
      Begin VB.Menu mnuFunctionsEquipment 
         Caption         =   "&Equipment"
         Begin VB.Menu mnuFunctionsEquipmentRate 
            Caption         =   "&Rate"
         End
         Begin VB.Menu mnuFunctionsEquipmentMaintenance 
            Caption         =   "&Maintenance"
         End
      End
      Begin VB.Menu mnuFunctionsCrews 
         Caption         =   "&Crews"
         Begin VB.Menu mnuFunctionsCrewMaintenance 
            Caption         =   "&Maintenance"
         End
      End
      Begin VB.Menu mnuFunctionsUnitCost 
         Caption         =   "&Unit Cost"
         Begin VB.Menu mnuFunctionsUnitCostMaintenance 
            Caption         =   "&Maintenance"
         End
         Begin VB.Menu mnuFunctionsUnitCostUsage 
            Caption         =   "Unit &Cost Usage"
         End
         Begin VB.Menu mnuFunctionsUnitCostMatUsage 
            Caption         =   "Material &Usage"
         End
         Begin VB.Menu mnuFunctionsUnitCostWasIs 
            Caption         =   "&Was/Is"
         End
         Begin VB.Menu mnuFunctionsUnitCostLongDesc 
            Caption         =   "&Long Descriptions"
         End
      End
      Begin VB.Menu mnuFunctionAssemblies 
         Caption         =   "&Assembly"
         Begin VB.Menu mnuFunctionsAssemblyMaint 
            Caption         =   "&Maintenance"
         End
         Begin VB.Menu mnuFunctionsAssemblyUsage 
            Caption         =   "&Unit Cost Usage"
         End
         Begin VB.Menu mnuFunctionsAssemblyBkDtl 
            Caption         =   "&Book Detail"
         End
      End
      Begin VB.Menu mnuFunctionModels 
         Caption         =   "&Models"
         Begin VB.Menu mnuFunctionBuildingMaint 
            Caption         =   "&Building"
         End
         Begin VB.Menu mnuFunctionModelMaint 
            Caption         =   "&Models"
         End
      End
      Begin VB.Menu mnuFunctionProjects 
         Caption         =   "&Projects"
         Begin VB.Menu mnuFunctionProjectMaint 
            Caption         =   "&Maintenance"
         End
         Begin VB.Menu mnuFunctionProjectAnalysis 
            Caption         =   "&Analysis"
         End
      End
      Begin VB.Menu mnuFunctionCCI 
         Caption         =   "City Cost Inde&x"
         Begin VB.Menu mnuFunctionCCIMatPrice 
            Caption         =   "&Material Price"
         End
         Begin VB.Menu mnuFunctionCCIEquipmentRate 
            Caption         =   "&Equipment Rate"
         End
         Begin VB.Menu mnuFunctionCCILaborRate 
            Caption         =   "&Labor Rate"
         End
         Begin VB.Menu mnuFunctionCCIMatEquRpt 
            Caption         =   "Mat/E&qu Exception"
         End
         Begin VB.Menu mnuFunctionCCILaborExc 
            Caption         =   "Labo&r Exception"
         End
         Begin VB.Menu mnuFunctionCCIDetail 
            Caption         =   "City &Detail"
         End
         Begin VB.Menu mnuFunctionCCIIndexDetail 
            Caption         =   "&Index Detail"
         End
         Begin VB.Menu mnuFunctionCCIDollarList 
            Caption         =   "Dolla&r Listing"
         End
         Begin VB.Menu mnuFunctionCCIIdxDtlExc 
            Caption         =   "Index Detail E&xception"
         End
         Begin VB.Menu mnuFunctionCCIComponentUsage 
            Caption         =   "&Component Usage"
         End
      End
      Begin VB.Menu mnuFunctionCCIAdmin 
         Caption         =   "CCI Administration"
         Begin VB.Menu mnuFunctionCCIAdmin1 
            Caption         =   "Clone Qtr Mat/Equ Prices"
         End
         Begin VB.Menu mnuFunctionCCIAdmin2 
            Caption         =   "Publish Qtr Labor Rates"
         End
         Begin VB.Menu mnuFunctionCCIAdmin3 
            Caption         =   "Report Qtr Mat/Equ Prices"
         End
         Begin VB.Menu mnuFunctionCCIAdmin4 
            Caption         =   "Generate MasterFormat Index "
         End
         Begin VB.Menu mnuFunctionCCIAdmin5 
            Caption         =   "Generate UNIFormat Index"
         End
         Begin VB.Menu mnuFunctionCCIAdmin6 
            Caption         =   "Generate Residential Index"
         End
         Begin VB.Menu mnuFunctionCCIAdmin7 
            Caption         =   "Generate Dollar Listing"
         End
         Begin VB.Menu mnuFunctionCCIAdmin8 
            Caption         =   "Generate MF  Exception"
         End
         Begin VB.Menu mnuFunctionCCIAdmin9 
            Caption         =   "Generate Mailing Labels"
         End
      End
      Begin VB.Menu mnuFunctionsInformationSources 
         Caption         =   "&Information Sources"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsMenu 
         Caption         =   "&Reports Menu..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuReportsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportsMaterial 
         Caption         =   "&Material"
         Enabled         =   0   'False
         Begin VB.Menu mnuReportsMaterialPriceDiv1_14 
            Caption         =   "&Material Price Div 1-14"
         End
         Begin VB.Menu mnuReportsMaterialPriceDiv15_16 
            Caption         =   "&Material Price Div 15-16"
         End
      End
      Begin VB.Menu mnuReportsLabor 
         Caption         =   "&Labor"
         Begin VB.Menu mnuReportsLaborParentHelperEdit 
            Caption         =   "&Labor Rate Parent/Helper Edit"
         End
         Begin VB.Menu mnuReportsLaborOutOfDate 
            Caption         =   "&Out of Date"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuReportsEquipment 
         Caption         =   "&Equipment"
         Enabled         =   0   'False
         Begin VB.Menu mnuReportsEquipmentMaster 
            Caption         =   "&Equipment Master"
         End
      End
      Begin VB.Menu mnuReportsCCI 
         Caption         =   "&CCI"
         Enabled         =   0   'False
         Begin VB.Menu mnuReportsCCIMaterial 
            Caption         =   "CCI &Material"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuReportsCCIComparison 
            Caption         =   "CCI &Comparison"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuReportsInfoSources 
         Caption         =   "&Information Sources"
         Enabled         =   0   'False
         Begin VB.Menu mnuReportsInformationSourcesUpdateStsByDiv 
            Caption         =   "&Update Status by Div"
         End
         Begin VB.Menu mnuReportsInformationSourcesContactsWithManubyContact 
            Caption         =   "&Contacts with Manufacturers by Contact Num"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOutput 
         Caption         =   "&Output"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsGridPreferences 
         Caption         =   "&Grid Preferences"
         Begin VB.Menu mnuGridPreferencesMat 
            Caption         =   "&Material"
            Begin VB.Menu mnuGridPreferencesMaterial 
               Caption         =   "&Maintenance"
            End
            Begin VB.Menu mnuGridPreferencesMatPrice 
               Caption         =   "&Price"
            End
            Begin VB.Menu mnuGridPreferencesMatHistory 
               Caption         =   "&History"
            End
            Begin VB.Menu mnuGridPreferencesMatManufacturer 
               Caption         =   "Manu&facturer"
            End
            Begin VB.Menu mnuGridPreferencesMatUsage 
               Caption         =   "&Usage"
            End
         End
         Begin VB.Menu mnuGridPreferencesLab 
            Caption         =   "&Labor"
            Begin VB.Menu mnuGridPreferencesLaborRate 
               Caption         =   "&Labor Rate"
            End
            Begin VB.Menu mnuGridPreferencesLaborTradeGroups 
               Caption         =   "&Trade Groups"
            End
         End
         Begin VB.Menu mnuGridPreferencesEquip 
            Caption         =   "&Equipment"
            Begin VB.Menu mnuGridPreferencesEquipment 
               Caption         =   "Maint&enance"
            End
            Begin VB.Menu mnuGridPreferencesEquipRate 
               Caption         =   "R&ate"
            End
            Begin VB.Menu mnuGridPreferencesEquipHistory 
               Caption         =   "Hi&story"
            End
         End
         Begin VB.Menu mnuGridPreferencesCrew 
            Caption         =   "&Crews"
            Begin VB.Menu mnuGridPreferencesCrews 
               Caption         =   "&Maintenance"
            End
         End
         Begin VB.Menu mnuGridPreferencesUnitC 
            Caption         =   "&Unit Cost"
            Begin VB.Menu mnuGridPreferencesUnitCost 
               Caption         =   "&Maintenance"
            End
            Begin VB.Menu mnuGridPreferencesUnitCostUsage 
               Caption         =   "Usa&ge"
            End
            Begin VB.Menu mnuGridPreferencesUnitCostHistory 
               Caption         =   "His&tory"
            End
         End
         Begin VB.Menu mnuGridPreferencesAssembly 
            Caption         =   "&Assembly"
            Begin VB.Menu mnuGridPreferencesAssemblyMaintenance 
               Caption         =   "&Maintenance"
            End
            Begin VB.Menu mnuGridPreferencesAsblyBk 
               Caption         =   "&Book Detail"
            End
            Begin VB.Menu mnuGridPreferencesAsblyUCUsage 
               Caption         =   "&Unit Cost Usage"
            End
            Begin VB.Menu mnuGridPreferencesAsblyHs 
               Caption         =   "&History"
            End
         End
         Begin VB.Menu mnuGridPreferencesModels 
            Caption         =   "&Models"
            Begin VB.Menu mnuGridPreferencesBuildingMaintenance 
               Caption         =   "&Buildings"
            End
            Begin VB.Menu mnuGridPreferencesModelMaintenance 
               Caption         =   "&Models"
            End
            Begin VB.Menu mnuGridPreferencesCommonAdditives 
               Caption         =   "&Common Additives"
            End
            Begin VB.Menu mnuGridPreferencesModelAssemblies 
               Caption         =   "Model &Assemblies"
            End
            Begin VB.Menu mnuGridPreferencesSummaryEstimate 
               Caption         =   "&Summary Estimate"
            End
         End
         Begin VB.Menu mnuGridPreferencesCCI 
            Caption         =   "City C&ost Index"
            Begin VB.Menu mnuGridPreferencesCCIMatPrice 
               Caption         =   "&Material Price"
            End
            Begin VB.Menu mnuGridPreferencesCCIEquipRate 
               Caption         =   "&Equipment Rate"
            End
            Begin VB.Menu mnuGridPreferencesCCILaborRate 
               Caption         =   "&Labor Rate"
            End
            Begin VB.Menu mnuGridPreferencesCCIQtrlyEstRpt 
               Caption         =   "Mat/&Equ Exception"
            End
            Begin VB.Menu mnuGridPreferencesCCILaborExcRpt 
               Caption         =   "Labo&r Exception"
            End
            Begin VB.Menu mnuGridPreferencesCCIDetail 
               Caption         =   "City &Detail"
            End
            Begin VB.Menu mnuGridPreferencesCCIIndexDetail 
               Caption         =   "&Index Detail"
            End
            Begin VB.Menu mnuGridPreferencesCCIDollarListing 
               Caption         =   "Do&llar Listing"
            End
            Begin VB.Menu mnuGridPreferencesCCIIdxDtlExc 
               Caption         =   "Inde&x Detail Exception"
            End
            Begin VB.Menu mnuGridPreferencesCCICompUsage 
               Caption         =   "&Component Usage"
            End
         End
         Begin VB.Menu mnuGridPreferencesProjectCosts 
            Caption         =   "&Project Costs"
            Begin VB.Menu mnuGridPreferencesProjectGrid 
               Caption         =   "Project &Grid"
            End
            Begin VB.Menu mnuGridPreferencesProjectAnalysis 
               Caption         =   "Project &Analysis"
            End
         End
         Begin VB.Menu mnuGridPreferencesInfoSource 
            Caption         =   "&Information Source"
         End
         Begin VB.Menu mnuGridPreferencesOutput 
            Caption         =   "&Output Usage"
         End
      End
      Begin VB.Menu mnuToolsSettings 
         Caption         =   "&Settings"
         Begin VB.Menu mnuSuperUserMode 
            Caption         =   "Super User Mode"
         End
         Begin VB.Menu mnuToolsSettingsCurrentQuarter 
            Caption         =   "Current &Quarter..."
         End
         Begin VB.Menu mnuToolsSettingsUser 
            Caption         =   "&User Information..."
         End
         Begin VB.Menu mnuToolsSettingsSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuToolsSettingsDatabaseOptions 
            Caption         =   "Database Options"
            Begin VB.Menu mnuToolsSettingsDatabaseOptionsDatabase 
               Caption         =   "Default &Database"
            End
            Begin VB.Menu mnuToolsSettingsDatabaseOptionsServer 
               Caption         =   "Default &Server"
            End
            Begin VB.Menu mnuToolsSettingsDatabaseOptionsDefaultUser 
               Caption         =   "Default &User"
            End
            Begin VB.Menu DefaultUserPassword 
               Caption         =   "Default User &Password"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuToolsSettingsDatabaseOptionsMaxRecords 
               Caption         =   "&Max Records"
            End
         End
         Begin VB.Menu mnuToolsSettingsMF 
            Caption         =   "Default MasterFormat"
            Begin VB.Menu mnuToolsSettingsMF1995 
               Caption         =   "MasterFormat 1995"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuToolsSettingsMF2004 
               Caption         =   "MasterFormat 2004"
            End
         End
         Begin VB.Menu mnuToolsSettingsFlatToolbar 
            Caption         =   "Use Flat &Toolbar Style"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuToolsSettingsAltDisabledColor 
            Caption         =   "Use &Alternate Locked Color"
         End
         Begin VB.Menu mnuToolsSettingsWhiteBackground 
            Caption         =   "Use &White Grid Background"
         End
         Begin VB.Menu mnuToolsSettingsMaximizeGrids 
            Caption         =   "&Maximize Grid Forms"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuHierarchy 
         Caption         =   "&Hierarchy"
         Begin VB.Menu mnuInsertHierarchy 
            Caption         =   "&Insert Hierarchy"
         End
         Begin VB.Menu UpdateHierarchyTotals 
            Caption         =   "&Update Hierarchy Totals"
         End
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsLogViewer 
         Caption         =   "Log Viewer..."
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpFAQ 
         Caption         =   "&Frequently Asked Questions..."
      End
      Begin VB.Menu mnuHelpReleaseNotes 
         Caption         =   "&Release Notes..."
      End
      Begin VB.Menu mnuToolsSettingsVersionCheck 
         Caption         =   "Check for &Updates..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <modulename> frmMain</modulename>
''' <functionname>General (Main) </functionname>
'''
''' <summary>
''' This is the parent MDI window/form from which all child forms eg. Unit Cost, Material , Assembly Maintenance, Sq Foot Models etc. are built:
'''
'''Heavy use of the REGISTRY is made in the placement and sizing of  the "Navigation Map" (i.e. "roadmap") and the navigation tree or "CCD Functions" to the left of it.
'''
'''There are (3) ways to get into CCD functionalit:
'''"   Main Window (frmMain) menu bars
'''"   Main Window "Roadmap"                   (frmNavMap)
'''"   Main Window "Tree" (in left margin)         (frmNavTree)
'''
'''HELPER CLASS: N/A
'''
'''The following menus, toolbar menu items, dialogs and forms are built and launched from here:
'''"   frmNavTree  - manages how menu tree item functionality is launched
'''"   frmNavMap  - displays the CCD "roadmap" in graphical block form.  Each graphic rectangle is a "hotspot" from which the indicated functionality can be double clicked upon and launched
'''"   Menu Editor    (File, Edit, View, Functions, Report, Tools, Window, Help and all associated menu items)
'''"   Toolbar Settings
'''
'''GLOBAL FUNCTIONS SETUP:
'''
'''"   PrintScreen
'''"   Refresh
'''
'''</summary>
'''
'''<seealso> frmNavTree </seealso>
'''<seealso>frmNavMap</seealso>
'''
''' <datastruct>m_objGridMap</datastruct>
'''<datastruct>m_rec</datastruct>
'''
''' <storedprocedurename> N/A </storedprocedurename>
'''
'''<returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''WINDOWS API Definitions:
'''
'''Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'''Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
'''Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'''Private Declare Function GetDesktopWindow Lib "user32" () As Long
'''Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'''Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
'''Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'''Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
'''<code>
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  REGISTRY SETTINGS
'''2.  WINDOWS APIs
'''3.  CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''</dependson>


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Const SRCCOPY = &HCC0020
Const EM_UNDO = &HC7

Dim frmNavTree As frmNavTree
Dim frmNavMap As frmNavMap
Dim linBreak As Line
Dim intLastWindowState As Integer
Dim m_strGridTypes() As String

Const ASSEMBLY_BOOK_DETAIL = 0
Const ASSEMBLY_HISTORY = 1
Const ASSEMBLY_MAINTENANCE = 2
Const BUILDING = 3
Const CREWS = 4
Const EQUIPMENT_HISTORY = 5
Const EQUIPMENT = 6
Const EQUIPMENT_RATE = 7
Const INFORMATION_SOURCE = 8
Const LABOR_RATE = 9
Const TRADE_GROUP = 10
Const MATERIAL = 11
Const MATERIAL_HISTORY = 12
Const MATERIAL_MANUFACTURER = 13
Const material_price = 14
Const MATERIAL_USAGE = 15
Const MODEL = 16
Const UNIT_COST = 17
Const UNIT_COST_HISTORY = 18
Const UNIT_COST_USAGE = 19

Const RegistryNameIndex = 0
Const ValueCheckedIndex = 1

Private Function GetRegistryValue(strKeyParm As String) As Variant
    Dim vValue As Variant
    Dim lSize As Long
    Dim strKey As String
    Dim hKey As Long
    Dim lRet As Long
    
    lSize = 4
    strKey = CCD_KEY + "\" + strKeyParm
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    ' Test to see if the Maximized Grid  Value is there
    If lRet = ERROR_NONE Then
        lRet = RegQueryValueExLong(hKey, "Background", 0&, REG_DWORD, vValue, lSize)
        GetRegistryValue = vValue
    End If
    RegCloseKey (hKey)

End Function

Private Sub LoadWhiteBackground()
    Dim iCurrentGrid As Integer
    Dim lSize As Long
    Dim strKey As String
    Dim hKey As Long
    Dim lRet As Long
    
    ReDim m_strGridTypes(0 To 19, 0 To 1)
    
    m_strGridTypes(ASSEMBLY_BOOK_DETAIL, RegistryNameIndex) = "AssemblyBook"
    m_strGridTypes(ASSEMBLY_HISTORY, RegistryNameIndex) = "AssemblyHistory"
    m_strGridTypes(ASSEMBLY_MAINTENANCE, RegistryNameIndex) = "Assembly"
    m_strGridTypes(BUILDING, RegistryNameIndex) = "Building"
    m_strGridTypes(CREWS, RegistryNameIndex) = "CrewUsage"
    m_strGridTypes(EQUIPMENT_HISTORY, RegistryNameIndex) = "EquipmentHistory"
    m_strGridTypes(EQUIPMENT, RegistryNameIndex) = "Equipment"
    m_strGridTypes(EQUIPMENT_RATE, RegistryNameIndex) = "EquipmentRate"
    m_strGridTypes(INFORMATION_SOURCE, RegistryNameIndex) = "Information Sources"
    m_strGridTypes(LABOR_RATE, RegistryNameIndex) = "Labor"
    m_strGridTypes(TRADE_GROUP, RegistryNameIndex) = "Trade Group"
    m_strGridTypes(MATERIAL, RegistryNameIndex) = "Material"
    m_strGridTypes(MATERIAL_HISTORY, RegistryNameIndex) = "MaterialHistory"
    m_strGridTypes(MATERIAL_MANUFACTURER, RegistryNameIndex) = "MaterialManufacturer"
    m_strGridTypes(material_price, RegistryNameIndex) = "MaterialPrice"
    m_strGridTypes(MATERIAL_USAGE, RegistryNameIndex) = "MaterialUsage"
    m_strGridTypes(MODEL, RegistryNameIndex) = "Model"
    m_strGridTypes(UNIT_COST, RegistryNameIndex) = "UnitCost"
    m_strGridTypes(UNIT_COST_HISTORY, RegistryNameIndex) = "UnitCostHistory"
    m_strGridTypes(UNIT_COST_USAGE, RegistryNameIndex) = "UnitCostUsage"
    
    For iCurrentGrid = 0 To UBound(m_strGridTypes)
        m_strGridTypes(iCurrentGrid, ValueCheckedIndex) = GetRegistryValue(m_strGridTypes(iCurrentGrid, RegistryNameIndex))
    Next iCurrentGrid
    For iCurrentGrid = 0 To UBound(m_strGridTypes)       'All must be on for global white background to be set
        If m_strGridTypes(iCurrentGrid, ValueCheckedIndex) = "0" Then
            mnuToolsSettingsWhiteBackground.Checked = False
            g_blnWhiteBackground = False
            lSize = 4
            strKey = CCD_KEY + "\Defaults\WhiteBackground"
            lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
            lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, 0, lSize)
            RegCloseKey (hKey)
            Exit For
        End If
    Next iCurrentGrid

End Sub

Private Function SaveRegistryValue(strKeyParm As String, blnSetting As Boolean) As Integer
    'Save the background value for the specified module
    Dim lSize As Long
    Dim strKey As String
    Dim hKey As Long
    Dim lRet As Long

    strKey = CCD_KEY + "\" + strKeyParm
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    If lRet = 0 Then
        lRet = RegSetValueExLong(hKey, "Background", 0&, REG_DWORD, IIf(blnSetting = True, 1, 0), 4)
        SaveRegistryValue = IIf(blnSetting = True, 1, 0)
    Else
        SaveRegistryValue = vbEmpty
    End If

End Function

Private Sub mnuFileFax_Click()
    
    On Error Resume Next
    If Me.ActiveForm Is Nothing Then Exit Sub
    Me.ActiveForm.FaxReport
    
End Sub

Private Sub mnuFilePrintScreen_Click()
    
    'Me.ActiveForm.PrintForm
    PrintScreen
    
End Sub

Private Sub mnuFunctionBuildingMaint_Click()
    Dim frm As frmBuildingGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmBuildingGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionCCIAdmin1_Click()
CCI_Admin 1
End Sub

Private Sub mnuFunctionCCIAdmin2_Click()
CCI_Admin 2
End Sub

Private Sub mnuFunctionCCIAdmin3_Click()
CCI_Admin 3
End Sub

Private Sub mnuFunctionCCIAdmin4_Click()
CCI_Admin 4
End Sub

Private Sub mnuFunctionCCIAdmin5_Click()
CCI_Admin 5
End Sub

Private Sub mnuFunctionCCIAdmin6_Click()
CCI_Admin 6
End Sub

Private Sub mnuFunctionCCIAdmin7_Click()
CCI_Admin 7
End Sub

Private Sub mnuFunctionCCIAdmin8_Click()
CCI_Admin 8
End Sub
Private Sub mnuFunctionCCIAdmin9_Click()
CCI_Admin 9
End Sub

Private Sub mnuFunctionCCIComponentUsage_Click()
    Dim frm As frmCCICompUsageGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCICompUsageGrid
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuFunctionCCIDetail_Click()
    Dim frm As frmCCIDetailGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCIDetailGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionCCIDollarList_Click()
    Dim frm As frmCCICSIFmtSumRptGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCICSIFmtSumRptGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionCCIEquipmentRate_Click()
    Dim frm As frmCCIEquipRateGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCIEquipRateGrid
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuFunctionCCIIdxDtlExc_Click()
    Dim frm As frmCCIIdxDtlExcGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCIIdxDtlExcGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionCCIIndexDetail_Click()
    Dim frm As frmCCIIndexDetailGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCIIndexDetailGrid
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuFunctionCCILaborExc_Click()
    Dim frm As frmCCILabExcGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCILabExcGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionCCILaborRate_Click()
    Dim frm As frmCCILaborRateGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCILaborRateGrid
    frm.Show
    Screen.MousePointer = vbNormal

End Sub


Private Sub mnuFunctionCCIMatEquRpt_Click()
    Dim frm As frmCCIMatEquRptGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCIMatEquRptGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub
Private Sub mnuFunctionCCIMatPrice_Click()
    Dim frm As frmCCIMatPriceGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCCIMatPriceGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionModelMaint_Click()
    Dim frm As frmModelGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmModelGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionProjectAnalysis_Click()
    Dim frm As frmProjectAnalysis
    Screen.MousePointer = vbHourglass
    Set frm = New frmProjectAnalysis
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionProjectMaint_Click()
    Dim frm As frmProjectGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmProjectGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsUnitCostLongDesc_Click()
    Dim frm As frmLongDescriptionGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmLongDescriptionGrid
    frm.ShowMasterFormatTree True
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesCCICompUsage_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Component Usage")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuGridPreferencesCCIDetail_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Detail")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesOutput_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Output Usage")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesProjectAnalysis_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Project Analysis")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuGridPreferencesProjectGrid_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Project Grid")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesCCIDollarListing_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Dollar Report")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuGridPreferencesCCIEquipRate_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Equipment Rate")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuGridPreferencesCCIIdxDtlExc_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Index Detail Exception Report")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesCCIIndexDetail_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Index Detail")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub


Private Sub mnuGridPreferencesCCILaborExcRpt_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Labor Exception")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesCCILaborRate_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Labor Rate")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuGridPreferencesCCIMatPrice_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Material Price")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuGridPreferencesCCIQtrlyEstRpt_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("CCI Material/Equipment Exception")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuHelpFAQ_Click()
'LAUNCH THE FAQ IN THE SYSTEM'S WEB BROWSER
    Dim sDefaultURL As String
    Dim sURL As String
    
    sDefaultURL = App.Path & "\rsmeans-ccd-faq.htm"
    'CHECK IF A URL HAS BEEN LOADED TO THE REGISTRY
    sURL = QueryRegistryKey(HKEY_CURRENT_USER, CCD_KEY & "\Defaults\FAQ", "URL", sDefaultURL)
    If LaunchBrowser(sURL) Then
        'browser launched successfully
    Else
        MsgBox "The web page failed to start.", vbCritical + vbOKOnly
    End If
    
End Sub

Private Sub mnuPriorityUser_Click()

End Sub

Private Sub mnuInsertHierarchy_Click()
    
    If (modCommon.CheckUserAuth()) Then
        Dim frm As frmHierarchyTree
        Set frm = New frmHierarchyTree
        frm.InsertMode (True)
        frm.Show vbModal
    Else
        MsgBox "Sorry you are not authorized to Change the Hierarchy"
    End If
        
End Sub

Private Sub mnuReportsLaborParentHelperEdit_Click()
    LaborRateParentHelperEditPrintPreview
End Sub

Private Sub mnuReportsMaterialPriceDiv1_14_Click()
    MatPriceDiv1_14PrintPreview
End Sub

Private Sub mnuReportsMaterialPriceDiv15_16_Click()
    MatPriceDiv15_16PrintPreview
End Sub

Private Sub mnuReportsMenu_Click()
' 10/12/2005 RTD - ADDED REPORTS MENU ITEM; HOTKEY = CTRL-R
    Dim frm As New frmReportMenu
    frm.Show
    
End Sub

Private Sub mnuSuperUserMode_Click()
 'MsgBox ("You user name is: " & strUserName)
 
 Select Case UCase(strUserName)
    Case UCase("hancockrl"), UCase("krodriguez"), UCase("rodriguezks"), UCase("ghoitt"), _
    UCase("hoittgl"), UCase("willsjn"), UCase("mykulowyczv")
        Dim ans As Variant
        ans = MsgBox("Turn on both MF95/MF04 ?", vbYesNo, "SUPER USER PROMPT")
        Select Case ans
            Case vbYes
                MF95_ENABLED = True
                SUPER_USER_SUPPORT = True
            Case vbNo
                 MF95_ENABLED = False
        End Select
    Case Else
    End Select
  
End Sub

Private Sub mnuToolsLogViewer_Click()
    Dim f As New frmLogViewer
    
    f.Show
    
End Sub

Private Sub mnuToolsSettingsAltDisabledColor_Click()
    Dim strKey As String
    Dim hKey As Long
    Dim lSize As Long
    Dim lRet As Long
    Dim vValue As Variant
    
    g_blnUseAlternateDisabledColor = Not g_blnUseAlternateDisabledColor
    fMainForm.mnuToolsSettingsAltDisabledColor.Checked = g_blnUseAlternateDisabledColor
    lSize = 4
    strKey = CCD_KEY + "\Defaults\UseAlternateDisabledColor"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    If lRet <> ERROR_NONE Then
        ' create key
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
    End If
    lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, Abs(g_blnUseAlternateDisabledColor), lSize)
    RegCloseKey (hKey)
    
    On Error Resume Next
    Screen.ActiveForm.Refresh

End Sub

Private Sub mnuToolsSettingsCurrentQuarter_Click()
    Dim ListQuarters As New cdlgLstSel
    Dim strSelectedQtr As String
    Dim strKey As String
    Dim lSize As Long
    Dim hKey As Long
    Dim lRet As Long
    
    strKey = CCD_KEY + "\Defaults\CurrentQuarter"
    strSelectedQtr = GetQuarterID(ListQuarters, "Current Quarter:")
    If strSelectedQtr <> "-1" Then
        lSize = Len(strSelectedQtr)
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strSelectedQtr, lSize)
        RegCloseKey (hKey)
        g_sQuarterID = strSelectedQtr
        MsgBox "New value stored.", vbInformation, "CCD Default Quarter"
    End If
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuToolsSettingsFlatToolbar_Click()
    Dim strKey As String
    Dim hKey As Long
    Dim lSize As Long
    Dim lRet As Long
    Dim vValue As Variant
    
    g_blnFlatToolbar = Not g_blnFlatToolbar
    fMainForm.mnuToolsSettingsFlatToolbar.Checked = g_blnFlatToolbar
    If g_blnFlatToolbar Then
        fMainForm.tbToolBar.Style = tbrFlat
    Else
        fMainForm.tbToolBar.Style = tbrStandard
    End If
    lSize = 4
    strKey = CCD_KEY + "\Defaults\FlatToolbar"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    If lRet <> ERROR_NONE Then
        ' create key
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
    End If
    lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, Abs(g_blnFlatToolbar), lSize)
    RegCloseKey (hKey)
End Sub

Private Sub mnuToolsSettingsMF1995_Click()
    Dim strKey As String
    Dim hKey As Long
    Dim lSize As Long
    Dim lRet As Long
    Dim vValue As Variant
    
    g_intMasterFormat = 1995
    fMainForm.mnuToolsSettingsMF1995.Checked = Not (g_intMasterFormat = 2004)
    fMainForm.mnuToolsSettingsMF2004.Checked = (g_intMasterFormat = 2004)
    lSize = 4
    strKey = CCD_KEY + "\Defaults\MasterFormat"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    If lRet <> ERROR_NONE Then
        ' create key
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
    End If
    lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, g_intMasterFormat, lSize)
    RegCloseKey (hKey)
    
End Sub

Private Sub mnuToolsSettingsMF2004_Click()
    Dim strKey As String
    Dim hKey As Long
    Dim lSize As Long
    Dim lRet As Long
    Dim vValue As Variant
    
    g_intMasterFormat = 2004
    fMainForm.mnuToolsSettingsMF1995.Checked = Not (g_intMasterFormat = 2004)
    fMainForm.mnuToolsSettingsMF2004.Checked = (g_intMasterFormat = 2004)
    lSize = 4
    strKey = CCD_KEY + "\Defaults\MasterFormat"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    If lRet <> ERROR_NONE Then
        ' create key
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
    End If
    lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, g_intMasterFormat, lSize)
    RegCloseKey (hKey)
    
End Sub

Private Sub mnuToolsSettingsUser_Click()
    Dim fInfo As New dlgUserInfo
    
    fInfo.ShowExtendedInfo = False
    fInfo.ReadOnly = False
    fInfo.Show vbModal
    Set fInfo = Nothing

End Sub

Private Sub mnuToolsSettingsVersionCheck_Click()
' 10/25/2005 RTD - OPEN THE LAUNCHER APPLET WITH /CHECKONLY FLAG
'                  TO DISPLAY VERSION/RELEASE INFORMATION
    Dim res As Long
    Dim sAppPath As String
    Dim sCommandLine As String
    
    sAppPath = App.Path
    If Right(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
    sCommandLine = "/CHECKONLY"
    sCommandLine = sCommandLine & " /CONNECT_DB=" & strConnectDatabase
    res = ShellExecute(Me.hWnd, "open", sAppPath & "CCDLaunch.exe", sCommandLine, sAppPath, vbNormalFocus)
    If res > 32 Then
        Call BringWindowToTop(res)
    End If
    
End Sub

Private Sub mnuToolsSettingsWhiteBackground_Click()
    Dim strKey As String
    Dim lSize As Long
    Dim hKey As Long
    Dim lRet As Long
    Dim iCurrentGrid As Integer
    
    mnuToolsSettingsWhiteBackground.Checked = Not mnuToolsSettingsWhiteBackground.Checked
    g_blnWhiteBackground = mnuToolsSettingsWhiteBackground.Checked
    lSize = 4
    strKey = CCD_KEY + "\Defaults\WhiteBackground"
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
    lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, IIf(g_blnWhiteBackground = True, 1, 0), lSize)
    RegCloseKey (hKey)

    For iCurrentGrid = 0 To UBound(m_strGridTypes)
        m_strGridTypes(iCurrentGrid, ValueCheckedIndex) = SaveRegistryValue(m_strGridTypes(iCurrentGrid, RegistryNameIndex), g_blnWhiteBackground)
    Next iCurrentGrid

End Sub
Private Sub mnuToolsSettingsMaximizeGrids_Click()
    Dim strKey As String
    Dim lSize As Long
    Dim hKey As Long
    Dim lRet As Long

    mnuToolsSettingsMaximizeGrids.Checked = Not mnuToolsSettingsMaximizeGrids.Checked
    g_blnMaximize = mnuToolsSettingsMaximizeGrids.Checked

'    Dim vValue As Variant
'    strTemp = IIf((g_blnMaximize = 0), "No", "Yes")
'    strTemp = InputBox("Do you want to maximize all grid forms when they open?", "CCD Maximize Grid Forms", strTemp)
'
'    Select Case UCase(strTemp)
'    Case "NO", "YES"
'        g_blnMaximize = IIf((UCase(strTemp) = "NO"), True, False)
        
        lSize = 4
        strKey = CCD_KEY + "\Defaults\MaximizeGridForms"
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, IIf(g_blnMaximize, 1, 0), lSize)
        RegCloseKey (hKey)

'    Case ""
'        Exit Sub
'    Case Else
'        MsgBox "You've specified an invalid choice. Please use 'Yes' or 'No'.", vbInformation, "CCD Maximize Grid Forms"
'        Call mnuToolsSettingsMaximizeGrids_Click
'    End Select

End Sub
Private Sub mnuToolsSettingsDatabaseOptionsDatabase_Click()

Dim strTemp As String
Dim strKey As String
Dim lSize As Long
Dim hKey As Long
Dim lRet As Long

    lSize = 1000
    
    strKey = CCD_KEY + "\Defaults\DBase"

    strTemp = InputBox("Specify a Default Database.", "CCD Default Database", strConnectDatabase)
    
    Select Case strTemp
    Case ""                 ' The user clicked on CANCEL
        Exit Sub
    Case Else               ' The user entered in some Text
        lSize = Len(strTemp)
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strTemp, lSize)
        RegCloseKey (hKey)
        
        MsgBox "New value stored." & vbCrLf & "Please note that the application must be restarted for this change to take effect.", vbInformation, "CCD Default Server"
    End Select

End Sub

Private Sub mnuToolsSettingsDatabaseOptionsDefaultUser_Click()

Dim strTemp As String
Dim strKey As String
Dim hKey As Long
Dim lSize As Long
Dim lRet As Long

    lSize = 1000
    
    strKey = CCD_KEY + "\Defaults\DBUser"

    strTemp = InputBox("Specify the default CCD User ID.", "Default User ID", strUserName)
    
    Select Case strTemp
    Case ""                 ' The user clicked on CANCEL
        Exit Sub
    Case Else               ' The user entered in some Text
        lSize = Len(strTemp)
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strTemp, lSize)
        RegCloseKey (hKey)
        MsgBox "New value stored." & vbCrLf & "Please note that the application must be restarted for this change to take effect.", vbInformation, "CCD Default Server"
    End Select

End Sub

Private Sub mnuToolsSettingsDatabaseOptionsMaxRecords_Click()

Dim strTemp As String
Dim strKey As String
Dim hKey As Long
Dim lSize As Long
Dim lRet As Long
Dim vValue As Variant

    On Error GoTo ErrHandler

    lSize = 4
    
    strKey = CCD_KEY + "\Defaults\MaxRecords"

    strTemp = InputBox("What is the Maximum number of Records you'd like to return?", "CCD Maximum Records", MAX_RECORDS)
    
    If strTemp = "" Then Exit Sub
    
    Select Case CInt(strTemp)
    Case 0 To 32000
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        lRet = RegSetValueExLong(hKey, "Value", 0&, REG_DWORD, CInt(strTemp), lSize)
        RegCloseKey (hKey)
        
        MAX_RECORDS = CInt(strTemp)
    Case Else
        MsgBox "You've specified an invalid amount. Please select an amount between 0 and 32000.", vbInformation, "CCD Maximum Records"
        Call mnuToolsSettingsDatabaseOptionsMaxRecords_Click
    End Select
    
    Exit Sub

ErrHandler:
    Select Case Err.Number
    Case 6, 13
        MsgBox "You've specified an invalid amount. Please select an amount between 0 and 32000.", vbInformation, "CCD Maximum Records"
        Call mnuToolsSettingsDatabaseOptionsMaxRecords_Click
    End Select

End Sub

Private Sub DefaultUserPassword_Click()

Dim strTemp As String
Dim strKey As String
Dim hKey As Long
Dim lSize As Long
Dim lRet As Long

    lSize = 1000
    
    strKey = CCD_KEY + "\Defaults\DBUserPW"

    strTemp = InputBox("Specify the Default User Password.", "CCD Default Server", CONN_USERPW)
    
    Select Case strTemp
    Case ""                 ' The user clicked on CANCEL
        Exit Sub
    Case Else               ' The user entered in some Text
        lSize = Len(strTemp)
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strTemp, lSize)
        RegCloseKey (hKey)
        
        MsgBox "New value stored." & vbCrLf & "Please note that the application must be restarted for this change to take effect.", vbInformation, "CCD Default Server"
    End Select

End Sub


Private Sub mnuToolsSettingsDatabaseOptionsServer_Click()

Dim strTemp As String
Dim strKey As String
Dim hKey As Long
Dim lSize As Long
Dim lRet As Long

    lSize = 1000
    
    strKey = CCD_KEY + "\Defaults\DBServer"

    strTemp = InputBox("Specify a Default Server.", "CCD Default Server", strConnectServer)
    
    Select Case strTemp
    Case ""                 ' The user clicked on CANCEL
        Exit Sub
    Case Else               ' The user entered in some Text
        lSize = Len(strTemp)
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0&, KEY_ALL_ACCESS, hKey)
        lRet = RegSetValueExString(hKey, "Value", 0&, REG_SZ, strTemp, lSize)
        RegCloseKey (hKey)
        
        MsgBox "New value stored." & vbCrLf & "Please note that the application must be restarted for this change to take effect.", vbInformation, "CCD Default Server"
    End Select

End Sub

Private Sub MDIForm_Activate()
    'OutputView False
End Sub

Private Sub MDIForm_Load()
'    Me.Left = GetSetting(App.title, "Settings", "MainLeft", 480)
'    Me.Top = GetSetting(App.title, "Settings", "MainTop", 300)
'    Me.Width = GetSetting(App.title, "Settings", "MainWidth", 13530)
'    Me.Height = GetSetting(App.title, "Settings", "MainHeight", 10155)
    Dim lRet As Long
    Dim hKey As Long
    Dim lValue As Long
    Dim lSize As Long
    lSize = 4
    
    'code added by Mohan on Jan 18, 2012: for version info - to be deleted before going live
    Dim strVersion As String
    strVersion = " Version " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Caption = Me.Caption + strVersion
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    Dim strSelect As String
    strSelect = "select domain_value from domain_tbl where domain_name = 'PUB_MAT_OPTION'"
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
    If blnReturn = True Then
        g_intRollupOption = rec.Fields("domain_value")
        rec.Close
    End If
    Set rec = Nothing
    ' Build key to open
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, CCD_KEY, 0&, KEY_ALL_ACCESS, hKey)
    ' If the key doesn't exist
    If lRet <> 0 Then
        ' Create it
        lRet = RegCreateKeyEx(HKEY_CURRENT_USER, CCD_KEY, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRet)
    End If
    ' Test to see if the values are there
    lRet = RegQueryValueExLong(hKey, "MainLeft", 0&, REG_DWORD, lValue, lSize)
    If lRet <> 0 Then
        ' Populate the values
        lRet = RegSetValueExLong(hKey, "MainLeft", 0&, REG_DWORD, 480, 4)
        lRet = RegSetValueExLong(hKey, "MainTop", 0&, REG_DWORD, 300, 4)
        lRet = RegSetValueExLong(hKey, "MainWidth", 0&, REG_DWORD, 13530, 4)
        lRet = RegSetValueExLong(hKey, "MainHeight", 0&, REG_DWORD, 10155, 4)
    End If
    
    ' Get all of the values
    lRet = RegQueryValueExLong(hKey, "MainLeft", 0&, REG_DWORD, lValue, lSize)
    Me.Left = lValue
    lRet = RegQueryValueExLong(hKey, "MainTop", 0&, REG_DWORD, lValue, lSize)
    Me.Top = lValue
    lRet = RegQueryValueExLong(hKey, "MainWidth", 0&, REG_DWORD, lValue, lSize)
    Me.Width = lValue
    lRet = RegQueryValueExLong(hKey, "MainHeight", 0&, REG_DWORD, lValue, lSize)
    Me.Height = lValue
    txtDatabase.Text = ""
    ' lblDatabase is only used to find out how wide to make txtDatabase.
    ' Labels cannot be placed on toolbars, which is why we don't just use the label
    'lblDatabase.Caption = strUserName + " on " + strConnectServer + ":" + CONN_DATABASE
    'txtDatabase.Width = lblDatabase.Width
    sbStatusBar.Panels(2).Text = strUserName + " on " + strConnectServer + ":" + strConnectDatabase & " "
    
    
    
    'rlh 03/30/2009  CCD 8.4
    If MF95_ENABLED = False Then
        mnuToolsSettingsMF1995.Enabled = False
    End If

    'disable
    If Not (modCommon.CheckUserAuth()) Then 'only users with 128 will have access to this
        mnuHierarchy.Enabled = False
    End If
    
    RegCloseKey (hKey)
    LoadWhiteBackground
    LoadNewDoc
    
    
    
    
End Sub

Private Sub LoadNewDoc()
    Set frmNavTree = New frmNavTree
    frmNavTree.Show
    Set frmNavMap = New frmNavMap
    frmNavMap.Show
End Sub

Private Sub MDIForm_Resize()
    If Not Me.WindowState = vbMinimized And Not intLastWindowState = vbMinimized And Not Me.ScaleHeight = 0 Then
        If Not frmNavTree Is Nothing Then
            frmNavTree.SizeIt
        End If
        If Not frmNavMap Is Nothing Then
            frmNavMap.SizeIt
        End If
    End If
    intLastWindowState = Me.WindowState
    With tbToolBar.Buttons("Database")
        txtDatabase.Move tbToolBar.Width - txtDatabase.Width, .Top + 50 '.Left, .Top + 50, .Width
        txtDatabase.ZOrder 0
    End With
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
'        SaveSetting App.title, "Settings", "MainLeft", Me.Left
'        SaveSetting App.title, "Settings", "MainTop", Me.Top
'        SaveSetting App.title, "Settings", "MainWidth", Me.Width
'        SaveSetting App.title, "Settings", "MainHeight", Me.Height
        Dim lRet As Long
        Dim hKey As Long
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, CCD_KEY, 0&, KEY_ALL_ACCESS, hKey)
        lRet = RegSetValueExLong(hKey, "MainLeft", 0&, REG_DWORD, Me.Left, 4)
        lRet = RegSetValueExLong(hKey, "MainTop", 0&, REG_DWORD, Me.Top, 4)
        lRet = RegSetValueExLong(hKey, "MainWidth", 0&, REG_DWORD, Me.Width, 4)
        lRet = RegSetValueExLong(hKey, "MainHeight", 0&, REG_DWORD, Me.Height, 4)
        RegCloseKey (hKey)
    End If
    
    ' Release the global connections
    g_objDAL.CacheConnection
    g_cnShared.Close
    g_cnSharedLong.Close
    Set g_objDAL = Nothing
    Set g_cnShared = Nothing
    Set g_cnSharedLong = Nothing
        
End Sub

Private Sub mnuEditSelectAll_Click()
    On Error Resume Next
    Me.ActiveForm.SelectAllRows
End Sub

Private Sub mnuFunctionsAssemblyBkDtl_Click()
    Dim frm As frmAssemblyBookGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmAssemblyBookGrid
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuFunctionsAssemblyMaint_Click()
    Dim frm As frmAssemblyGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmAssemblyGrid
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuFunctionsAssemblyUsage_Click()
    Dim frm As frmUCostUsageGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmUCostUsageGrid
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuFunctionsCrewMaintenance_Click()
    Dim frm As frmCrewGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCrewGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub
Private Sub mnuFunctionsInformationSources_Click()
    Dim frm As frmInfoSourceGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmInfoSourceGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsLaborRates_Click()

'MsgBox "The Labor form is not functional."
'Exit Sub

    Dim frm1 As frmLaborRateGrid
    Screen.MousePointer = vbHourglass
    Set frm1 = New frmLaborRateGrid
    frm1.Show
    Screen.MousePointer = vbNormal

End Sub
Private Sub mnuFunctionsMaterialMaintenance_Click()
    Dim frm As frmMaterialGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmMaterialGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsMaterialManufacturer_Click()
    Dim frm As frmMatManufacturerGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmMatManufacturerGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsMaterialPrice_Click()
    Dim frm1 As frmMatPriceGrid
    Screen.MousePointer = vbHourglass
    Set frm1 = New frmMatPriceGrid
    frm1.Show
    Screen.MousePointer = vbNormal
End Sub
Private Sub mnuFunctionsMaterialUsage_Click()
    Dim frm As frmMatUsageGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmMatUsageGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsEquipmentMaintenance_Click()
    Dim frm As frmEquipmentGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmEquipmentGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsEquipmentRate_Click()
    Dim frm1 As frmEquipRateGrid
    Screen.MousePointer = vbHourglass
    Set frm1 = New frmEquipRateGrid
    frm1.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsPubMatRollup_Click()
    Dim frm As frmMatPubRollupGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmMatPubRollupGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsTradeGroups_Click()
    Dim frm1 As frmTradeGroupGrid
    Screen.MousePointer = vbHourglass
    Set frm1 = New frmTradeGroupGrid
    frm1.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuFunctionsUnitCostMaintenance_Click()
    Dim frm As frmUnitCostGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmUnitCostGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFunctionsUnitCostUsage_Click()
    Dim frm As frmUCostUsageGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmUCostUsageGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub
Private Sub mnuFunctionsUnitCostMatUsage_Click()
    Dim frm As frmMatUsageGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmMatUsageGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesAsblyBk_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Assembly Book Detail")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesAsblyUCUsage_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Assembly Unit Cost Usage")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesAsblyHs_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Assembly History")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesAssemblyMaintenance_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Assembly Maintenance")
    frm.Show
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuGridPreferencesBuildingMaintenance_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Building")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesModelMaintenance_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Model")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesCommonAdditives_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Common Additives")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesModelAssemblies_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Model Assemblies")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesSummaryEstimate_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Summary Estimate")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesCrews_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Crews")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesEquipHistory_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Equipment History")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesEquipment_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Equipment")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesEquipRate_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Equipment Rate")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesInfoSource_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Information Source")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesLaborTradeGroups_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Trade Group")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub
Private Sub mnuGridPreferencesMaterial_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Material")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesMatHistory_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Material History")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesMatManufacturer_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Material Manufacturer")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesMatPrice_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Material Price")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesLaborRate_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Labor Rate")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesMatUsage_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Material Usage")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesUnitCost_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Unit Cost")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesUnitCostUsage_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Unit Cost Usage")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuGridPreferencesUnitCostHistory_Click()
    Dim frm As frmGridPreference
    Screen.MousePointer = vbHourglass
    Set frm = New frmGridPreference
    frm.SetType ("Unit Cost History")
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuHelpReleaseNotes_Click()
    frmReleaseNotes.Show
    CenterFormInParent frmReleaseNotes, Me
End Sub

Private Sub mnuToolsOutput_Click()
'    Dim col As New Collection
    On Error Resume Next
'    col.Add 1244
'    dlgOutput.SetKeys col, "E"
'    dlgOutput.Show vbModeless, Me
    ActiveForm.DoOutput
End Sub

Private Sub mnuViewNavMap_Click()
    mnuViewNavMap.Checked = Not mnuViewNavMap.Checked
    frmNavMap.Visible = mnuViewNavMap.Checked
End Sub

Private Sub mnuViewNavTree_Click()
    mnuViewNavTree.Checked = Not mnuViewNavTree.Checked
    frmNavTree.Visible = mnuViewNavTree.Checked
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePrintPreview_Click
        Case "Export"
            mnuFileSaveAs_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Fax"
            Screen.ActiveForm.FaxReport
        Case "EMail"
            Screen.ActiveForm.MailReport
        Case "Delete"
            Screen.ActiveControl.SelText = ""
        Case "PrintScreen"
            PrintScreen
        Case "Undo"
            'ToDo: Add 'Undo' button code.
'            MsgBox "Add 'Undo' button code."
        Case "Find"
            'ToDo: Add 'Find' button code.
'            MsgBox "Add 'Find' button code."
        Case "Sort Ascending"
            Screen.ActiveForm.Sort (SORT_ASCENDING)
        Case "Sort Descending"
            Screen.ActiveForm.Sort (SORT_DESCENDING)
        Case "Excel"
            Screen.ActiveForm.ExportData
        Case "Help"
            mnuHelpContents_Click
    End Select
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
'8/26/2005 RTD
'LAUNCH THE HELP PAGE IN THE SYSTEM'S WEB BROWSER
    Dim sDefaultURL As String
    Dim sURL As String
    
    sDefaultURL = App.Path & "\rsmeans-ccd-faq.htm"
    'CHECK IF A URL HAS BEEN LOADED TO THE REGISTRY
    sURL = QueryRegistryKey(HKEY_CURRENT_USER, CCD_KEY & "\Defaults\Help", "URL", sDefaultURL)
    If LaunchBrowser(sURL) Then
        'browser launched successfully
    Else
        MsgBox "The web page failed to start.", vbCritical + vbOKOnly
    End If

End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewWebBrowser_Click()
    'ToDo: Add 'mnuViewWebBrowser_Click' code.
    MsgBox "Add 'mnuViewWebBrowser_Click' code."
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
'Dim strActiveForm As String
'Dim intDigits As Integer
'
'strActiveForm = Screen.ActiveForm.Name
'intDigits = Len(Screen.ActiveForm.Name) - 3
'
'
'strActiveForm = right(strActiveForm, intDigits)
'MsgBox (strActiveForm & " is the active form")
    Refresh
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
'    ActiveForm.rtfText.SelRTF = Clipboard.GetText
    Screen.ActiveControl.SelText = Clipboard.GetText()
    Screen.ActiveControl.SetFocus
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
'    Clipboard.SetText ActiveForm.rtfText.SelRTF
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
'    Clipboard.SetText ActiveForm.rtfText.SelRTF
'    ActiveForm.rtfText.SelText = vbNullString
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
    Screen.ActiveControl.SelText = ""
End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    thank_you
    'End
    Unload Me
    Set frmMain = Nothing
    
End Sub


Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    
    Screen.ActiveForm.PrintReport

'    With dlgCommonDialog
'        .DialogTitle = "Print"
'        .CancelError = True
'        .flags = cdlPDReturnDC + cdlPDNoPageNums
'        If ActiveForm.rtfText.SelLength = 0 Then
'            .flags = .flags + cdlPDAllPages
'        Else
'            .flags = .flags + cdlPDSelection
'        End If
'        .ShowPrinter
'        If Err <> MSComDlg.cdlCancel Then
'            ActiveForm.rtfText.SelPrint .hDC
'        End If
'    End With
'
End Sub

Private Sub mnuFilePrintPreview_Click()
    
    If ActiveForm Is Nothing Then Exit Sub
    Screen.ActiveForm.PreviewReport

End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileSaveAs_Click()
    
    On Error Resume Next
    If Me.ActiveForm Is Nothing Then Exit Sub
    Me.ActiveForm.ExportReport

End Sub

Private Sub mnuFileSave_Click()

    On Error Resume Next
    If Me.ActiveForm Is Nothing Then Exit Sub
    Me.ActiveForm.SaveFile

End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String

    If ActiveForm Is Nothing Then LoadNewDoc
    
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.RTFText.LoadFile sFile
    ActiveForm.Caption = sFile

End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

Public Sub Refresh()

    Dim strActiveForm As String
    If Not (Me.ActiveForm Is Nothing) Then
        Me.ActiveForm.Refresh
    End If
    'strActiveForm = MDI Form.ActiveForm.Name.TDBGrid.ReBind
    'strActiveForm.TDBGrid.ReBind
    
End Sub

Public Sub thank_you()
    Dim SoundName As String
    Dim Result As Long
    
    SoundName = "c:\thank_you.wav"
    Result = sndPlaySound(SoundName$, cSndASYNC Or cSndNODEFAULT)
    
End Sub

Public Function PrintScreen()
' PRINT MAIN FORM WINDOW TO THE PRINTER
' ADDED 5/31/2005 RTD
    Dim hWnd As Long
    Dim hdc As Long
    Dim Picture1 As PictureBox
    Dim iOrientation As Long
    Dim iPrtTop As Long
    Dim iPrtLeft As Long
    
    'hWnd = GetDesktopWindow()
    'hdc = GetDC(hWnd)          ' context to Window client area
    hWnd = Me.hWnd
    hdc = GetWindowDC(hWnd)     ' context to entire Window area
    
    Set Picture1 = Me.Controls.Add("VB.PictureBox", "Picture1")
    Picture1.AutoRedraw = True 'this is required otherwise we will not be able to save the picture
    Picture1.Width = Me.Width + (Screen.TwipsPerPixelX * 2)
    Picture1.Height = Me.Height + (Screen.TwipsPerPixelY * 2)
    Picture1.AutoSize = True
    
    'BitBlt Picture1.hdc, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, hdc, 0, 0, SRCCOPY
    BitBlt Picture1.hdc, 0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, hdc, 0, 0, SRCCOPY
    Call ReleaseDC(hWnd, hdc)
    
    'SavePicture Picture1.Image, "c:\test.bmp"
    iPrtTop = 0
    iPrtLeft = 0
    iOrientation = Printer.Orientation
    Printer.Orientation = vbPRORLandscape
    If Picture1.Width < Printer.Width Then
        iPrtLeft = (Printer.Width - Picture1.Width) / 2
    End If
    If Picture1.Height < Printer.Height Then
        iPrtTop = (Printer.Height - Picture1.Height) / 2
    End If
    Printer.PaintPicture Picture1.Image, iPrtTop, iPrtLeft
    Printer.EndDoc
    Printer.Orientation = iOrientation
    Me.Controls.Remove "Picture1"
    Set Picture1 = Nothing
    
End Function

Private Sub UpdateHierarchyTotals_Click()
    If (modCommon.CheckUserAuth()) Then
        Dim retCheck As Integer
        retCheck = MsgBox("Are you sure you want to 'Update Hierarchy Totals', it takes a few seconds and locks up the MASTERFORMAT04_ID_HIERARCHY for that time", vbYesNo + vbQuestion)
        If retCheck = vbYes Then
            Dim retBlnForUpdate As Boolean
            retBlnForUpdate = False
            retBlnForUpdate = MainModule.Update_MasterFormat04_ID_Hierarchy_Totals_Only()
            If retBlnForUpdate = True Then
                MsgBox "Successfully updated the Hierarchy Totals", vbInformation
            Else
                MsgBox "Failed to update the Hierarchy Totals", vbExclamation
            End If
        End If
    Else
        MsgBox "Sorry you are not authorized to Update Hierarchy Totals", vbExclamation
    End If
    
End Sub
