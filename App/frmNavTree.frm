VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNavTree 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CCD Functions"
   ClientHeight    =   6975
   ClientLeft      =   1320
   ClientTop       =   1155
   ClientWidth     =   2295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavTree.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6855
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   12091
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmNavTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    ShowMinimizedForms
    SizeIt
    OutputView False
End Sub

Private Sub Form_Load()
    Dim NewNode As Object

    Set TreeView1.ImageList = ImageList1
    
    Set NewNode = TreeView1.Nodes.Add(, , "n01", "Material", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n02", "Labor", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n03", "Equipment", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n04", "Crews", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n05", "Unit Cost", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n06", "Assembly", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n07", "Models", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n08", "Projects", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n09", "City Cost Index", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n10", "CCI Administration", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n11", "Information Sources", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n12", "System Administration", 1)
    Set NewNode = TreeView1.Nodes.Add(, , "n13", "Reports", 1)
    
    ' Material Nodes
    Set NewNode = TreeView1.Nodes.Add("n01", tvwChild, "n0101", "Price", 1)
    Set NewNode = TreeView1.Nodes.Add("n01", tvwChild, "n0102", "Usage", 1)
    Set NewNode = TreeView1.Nodes.Add("n01", tvwChild, "n0104", "Maintenance", 1)
    Set NewNode = TreeView1.Nodes.Add("n01", tvwChild, "n0105", "Manufacturers", 1)
    Set NewNode = TreeView1.Nodes.Add("n01", tvwChild, "n0106", "Rollup", 1)
    Set NewNode = TreeView1.Nodes.Add("n01", tvwChild, "n0107", "Was/Is", 1)
    
    ' Labor Nodes
    Set NewNode = TreeView1.Nodes.Add("n02", tvwChild, "n0201", "Rates", 1)
    Set NewNode = TreeView1.Nodes.Add("n02", tvwChild, "n0202", "Trade Groups", 1)
    If (DEBUGON) Then
        Stop    'rlh 03/06/2010
    End If
    Set NewNode = TreeView1.Nodes.Add("n02", tvwChild, "n0203", "Out-of-Date Report", 1)
    Set NewNode = TreeView1.Nodes.Add("n02", tvwChild, "n0204", "Extend Term Date", 1)
    ' Equipment Nodes
    Set NewNode = TreeView1.Nodes.Add("n03", tvwChild, "n0301", "Rate", 1)
    Set NewNode = TreeView1.Nodes.Add("n03", tvwChild, "n0302", "Maintenance", 1)
    ' Crew Nodes
    Set NewNode = TreeView1.Nodes.Add("n04", tvwChild, "n0401", "Maintenance", 1)
    ' Unit Cost Nodes
    Set NewNode = TreeView1.Nodes.Add("n05", tvwChild, "n0501", "Maintenance", 1)
    Set NewNode = TreeView1.Nodes.Add("n05", tvwChild, "n0502", "Unit Cost Usage", 1)
    Set NewNode = TreeView1.Nodes.Add("n05", tvwChild, "n0503", "Material Usage", 1)
    Set NewNode = TreeView1.Nodes.Add("n05", tvwChild, "n0504", "Was/Is", 1)
    Set NewNode = TreeView1.Nodes.Add("n05", tvwChild, "n0505", "Long Descriptions", 1)
    ' Assembly Nodes
    Set NewNode = TreeView1.Nodes.Add("n06", tvwChild, "n0601", "Maintenance", 1)
    Set NewNode = TreeView1.Nodes.Add("n06", tvwChild, "n0602", "Unit Cost Usage", 1)
    Set NewNode = TreeView1.Nodes.Add("n06", tvwChild, "n0603", "Book Detail", 1)
    ' Model Nodes
    Set NewNode = TreeView1.Nodes.Add("n07", tvwChild, "n0701", "Buildings", 1)
    Set NewNode = TreeView1.Nodes.Add("n07", tvwChild, "n0702", "Models", 1)
    ' Division 17
    Set NewNode = TreeView1.Nodes.Add("n08", tvwChild, "n0801", "Project Maintenance", 1)
    Set NewNode = TreeView1.Nodes.Add("n08", tvwChild, "n0802", "Project Analysis", 1)
    ' City Cost Index Nodes
    If DEBUGON Then Stop
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0901", "Material Price", 1)
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0902", "Equipment Rate", 1)
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0903", "Labor Rate", 1)
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0904", "Mat/Equ Exception", 1)
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0905", "Labor Exception", 1)
'    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0910", "City Detail", 1)   'rlh 02/25/2010
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0906", "Index Detail", 1)
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0907", "Dollar Listing", 1)
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0908", "Index Detail Exception", 1)
    Set NewNode = TreeView1.Nodes.Add("n09", tvwChild, "n0909", "Component Usage", 1)
    ' City Cost Index Administration Nodes
    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1000", "Extend Quarter Dates", 1)
    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1001", "Clone Qtr Mat/Equ Prices", 1)
    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1002", "Publish Qtr Labor Rates", 1)
'    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1003", "Report Qtr Mat/Equ Prices", 1)    'rlh 02/27/2010
    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1004", "Generate MasterFormat Index", 1)
    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1005", "Generate UNIFormat Index", 1)
    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1006", "Generate Residential Index", 1)
'    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1007", "Generate Dollar Listing", 1)   'rlh 02/27/2010
    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1008", "Generate MF Exception", 1)
    Set NewNode = TreeView1.Nodes.Add("n10", tvwChild, "n1009", "Generate Mailing Labels", 1)
    
    ' Information Source
    Set NewNode = TreeView1.Nodes.Add("n11", tvwChild, "n1101", "Information Sources", 1)
    
    ' Admin
    Set NewNode = TreeView1.Nodes.Add("n12", tvwChild, "n1200", "Admin Control Panel", 1)
    Set NewNode = TreeView1.Nodes.Add("n12", tvwChild, "n1201", "User Administration", 1)
    
    ' Report Nodes
    ' 10/4/2005 RTD - CORRECTED NODE KEYS TO 13xx
    ' 10/12/2005 RTD - ADDED NODE FOR REPORTS MENU
    Set NewNode = TreeView1.Nodes.Add("n13", tvwChild, "n1300", "Reports Menu", 1)
    Set NewNode = TreeView1.Nodes.Add("n13", tvwChild, "n1301", "MatPrice Div 1-14", 1)
    Set NewNode = TreeView1.Nodes.Add("n13", tvwChild, "n1302", "MatPrice Div 15-16", 1)

    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub TreeView1_Click()
    ' 10/4/2005 RTD - CORRECTED NODE KEYS TO 13xx
    Select Case TreeView1.SelectedItem.Key
    Case "n1301", "n1302"
        Forms(0).mnuFilePageSetup.Enabled = True
        Forms(0).mnuFilePrintPreview.Enabled = True
        Forms(0).mnuFilePrint.Enabled = True
    Case Else
        Forms(0).mnuFilePageSetup.Enabled = False
        Forms(0).mnuFilePrintPreview.Enabled = False
        Forms(0).mnuFilePrint.Enabled = False
    End Select
End Sub

Public Sub PreviewReport()
    TreeView1_DblClick
End Sub

Private Sub TreeView1_DblClick()
Dim frm As Form
    Screen.MousePointer = vbHourglass
    
    If DEBUGON Then
        Stop 'rlh
    End If
    
    Select Case TreeView1.SelectedItem.Key
    Case "n0101"
        Set frm = New frmMatPriceGrid
        frm.Show
    Case "n0102"
        Set frm = New frmMatUsageGrid
        frm.Show
    Case "n0104"
        Set frm = New frmMaterialGrid
        frm.Show
    Case "n0105"
        Set frm = New frmMatManufacturerGrid
        frm.Show
    Case "n0106"
        Set frm = New frmMatPubRollupGrid
        frm.Show
    Case "n0201"
        Set frm = New frmLaborRateGrid
        frm.Show
    Case "n0202"
        Set frm = New frmTradeGroupGrid
        frm.Show
    Case "n0203"
        If DEBUGON Then
            Stop
        End If
'         Set frm = New frmOutOfDateReport       'rlh 02/26/2010 form not yet built!
'        frm.Show
        CCI_Admin 10                             'rlh 03/05/2010 don't need a form, use CCI ADMIN methodology
    Case "n0204"
    
'         Set frm = New frmExtendTermDate        'rlh 02/26/2010 form not yet built!
'        frm.Show
        CCI_Admin 11                             'rlh 03/05/2010 don't need a form, use CCI ADMIN methodology
    Case "n0301"
        Set frm = New frmEquipRateGrid
        frm.Show
    Case "n0302"
        Set frm = New frmEquipmentGrid
        frm.Show
    Case "n0401"
        Set frm = New frmCrewGrid
        frm.Show
    Case "n0501"
        Set frm = New frmUnitCostGrid
        frm.Show
    Case "n0502"
        Set frm = New frmUCostUsageGrid
        frm.Show
    Case "n0503"
        Set frm = New frmMatUsageGrid
        frm.Show
    Case "n0505"
        Set frm = New frmLongDescriptionGrid
        frm.Show
    Case "n0599"
        Set frm = New frmUCostHistoryGrid
        frm.JumpIn ("0151040101")
        frm.Show
    Case "n0601"
        Set frm = New frmAssemblyGrid
        frm.Show
    Case "n0602"
        Set frm = New frmUCostUsageGrid
        frm.Show
    Case "n0603"
        Set frm = New frmAssemblyBookGrid
        frm.Show
    Case "n0701"
        Set frm = New frmBuildingGrid
        frm.Show
    Case "n0702"
        Set frm = New frmModelGrid
        frm.Show
    Case "n0801"
        Set frm = New frmProjectGrid
        frm.Show
    Case "n0802"
        Set frm = New frmProjectAnalysis
        frm.Show
    Case "n0901"
        If DEBUGON Then Stop
        Set frm = New frmCCIMatPriceGrid
        frm.Show
    Case "n0902"
    If DEBUGON Then Stop
        Set frm = New frmCCIEquipRateGrid
        frm.Show
    Case "n0903"
    If DEBUGON Then Stop
        Set frm = New frmCCILaborRateGrid
        frm.Show
    Case "n0904"
    If DEBUGON Then Stop
        Set frm = New frmCCIMatEquRptGrid
        frm.Show
    Case "n0905"
    If DEBUGON Then Stop
        Set frm = New frmCCILabExcGrid
        frm.Show
    Case "n0906"
    If DEBUGON Then Stop
        Set frm = New frmCCIIndexDetailGrid
        frm.Show
    Case "n0907"
    If DEBUGON Then Stop
        Set frm = New frmCCICSIFmtSumRptGrid
        frm.Show
    Case "n0908"
    If DEBUGON Then Stop
        Set frm = New frmCCIIdxDtlExcGrid
        frm.Show
    Case "n0909"
    If DEBUGON Then Stop
        Set frm = New frmCCICompUsageGrid
        frm.Show
    Case "n0910"
        Set frm = New frmCCIDetailGrid
        frm.Show
    Case "n1000"        'rlh 02/26/2010
        CCI_Admin 0
    Case "n1001"
        CCI_Admin 1
    Case "n1002"
        CCI_Admin 2
    Case "n1003"
        CCI_Admin 3
    Case "n1004"
        CCI_Admin 4
    Case "n1005"
        CCI_Admin 5
    Case "n1006"
        CCI_Admin 6
    Case "n1007"
        CCI_Admin 7
    Case "n1008" 'Generate Masterformat Exception Report
        CCI_Admin 8
    Case "n1009"
        CCI_Admin 9
    Case "n1101"
        Set frm = New frmInfoSourceGrid
        frm.Show
    ' 10/18/2005 RTD - ADDED ADMIN MENU NODE
    Case "n1200"
        If g_blnIsUserAdmin Then
            Set frm = New frmAdminMenu
            frm.Show
        Else
            MsgBox "Sorry, but that function is limited to CCD Administrators.", vbInformation
        End If
    Case "n1201"
        If g_blnIsUserAdmin Then
            Set frm = New frmAdminUsers
            frm.Show
        Else
            MsgBox "Sorry, but that function is limited to CCD Administrators.", vbInformation
        End If
    ' 10/12/2005 RTD - ADDED REPORTS MENU NODE
    Case "n1300"
        Set frm = New frmReportMenu
        frm.Show
    Case "n1301"
        MatPriceDiv1_14PrintPreview
    Case "n1302"
        MatPriceDiv15_16PrintPreview
    End Select
    Screen.MousePointer = vbNormal
End Sub

Public Sub SizeIt()
    Move 0, 0, Me.Width, fMainForm.ScaleHeight
    TreeView1.Move TreeView1.Left, TreeView1.Top, TreeView1.Width, Me.ScaleHeight - 2 * TreeView1.Top
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

