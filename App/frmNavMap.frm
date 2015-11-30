VERSION 5.00
Begin VB.Form frmNavMap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Navigation Map"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line25 
      BorderWidth     =   3
      X1              =   4560
      X2              =   4380
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   660
      X2              =   1200
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   660
      X2              =   1200
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   660
      X2              =   1200
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   2520
      X2              =   4380
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   2520
      X2              =   2700
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   2520
      X2              =   2700
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   4200
      X2              =   4560
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   2700
      X2              =   2880
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Line Line11 
      BorderWidth     =   3
      X1              =   2700
      X2              =   2880
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   3540
      X2              =   3540
      Y1              =   5100
      Y2              =   3960
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      X1              =   5880
      X2              =   6240
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line Line14 
      BorderWidth     =   3
      X1              =   4380
      X2              =   4560
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line15 
      BorderWidth     =   3
      X1              =   7560
      X2              =   7860
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line Line16 
      BorderWidth     =   3
      X1              =   660
      X2              =   3540
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Line Line18 
      BorderWidth     =   3
      X1              =   5220
      X2              =   5220
      Y1              =   3960
      Y2              =   4980
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      X1              =   8580
      X2              =   8580
      Y1              =   3960
      Y2              =   4980
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      X1              =   6900
      X2              =   6900
      Y1              =   4320
      Y2              =   3960
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      X1              =   5220
      X2              =   6240
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      X1              =   7540
      X2              =   8580
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      X1              =   660
      X2              =   4380
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      X1              =   4380
      X2              =   4380
      Y1              =   5580
      Y2              =   3660
   End
   Begin VB.Image imgInfoSource 
      Height          =   6300
      Left            =   180
      Picture         =   "frmNavMap.frx":0000
      Top             =   180
      Width           =   450
   End
   Begin VB.Image imgMaterial 
      Height          =   1200
      Left            =   1260
      Picture         =   "frmNavMap.frx":9732
      Top             =   180
      Width           =   1200
   End
   Begin VB.Image imgEquipment 
      Height          =   1200
      Left            =   1260
      Picture         =   "frmNavMap.frx":E274
      Top             =   1860
      Width           =   1200
   End
   Begin VB.Image imgLabor 
      Height          =   1200
      Left            =   1260
      Picture         =   "frmNavMap.frx":12DB6
      Top             =   3540
      Width           =   1200
   End
   Begin VB.Image imgCrews 
      Height          =   1200
      Left            =   2940
      Picture         =   "frmNavMap.frx":178F8
      Top             =   2700
      Width           =   1200
   End
   Begin VB.Image imgUnitCost 
      Height          =   1200
      Left            =   4620
      Picture         =   "frmNavMap.frx":1C43A
      Top             =   2700
      Width           =   1200
   End
   Begin VB.Image Image7 
      Height          =   1200
      Left            =   6300
      Picture         =   "frmNavMap.frx":20F7C
      Top             =   2700
      Width           =   1200
   End
   Begin VB.Image imgModels 
      Height          =   1200
      Left            =   7920
      Picture         =   "frmNavMap.frx":25ABE
      Top             =   2700
      Width           =   1200
   End
   Begin VB.Image imgProjects 
      Height          =   1200
      Left            =   7920
      Picture         =   "frmNavMap.frx":2A600
      Top             =   1020
      Width           =   1200
   End
   Begin VB.Image Image10 
      Height          =   1200
      Left            =   6300
      Picture         =   "frmNavMap.frx":2F142
      Top             =   4380
      Width           =   1200
   End
   Begin VB.Image Image11 
      Height          =   450
      Left            =   4620
      Picture         =   "frmNavMap.frx":33C84
      Top             =   6060
      Width           =   4500
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   2700
      X2              =   2700
      Y1              =   4140
      Y2              =   3540
   End
   Begin VB.Line Line10 
      BorderWidth     =   3
      X1              =   2700
      X2              =   2700
      Y1              =   2460
      Y2              =   3060
   End
   Begin VB.Line Line17 
      BorderWidth     =   3
      X1              =   4380
      X2              =   4380
      Y1              =   2940
      Y2              =   780
   End
   Begin VB.Image Image12 
      Height          =   450
      Left            =   1200
      Picture         =   "frmNavMap.frx":3A63E
      Top             =   6060
      Width           =   3105
   End
End
Attribute VB_Name = "frmNavMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_blnDblClick As Boolean

Private Sub Form_Activate()
    ShowMinimizedForms
    SizeIt
    m_blnDblClick = False
    OutputView False
End Sub

Public Sub SizeIt()
    Move NAV_TREE_WIDTH, 0, fMainForm.ScaleWidth - NAV_TREE_WIDTH, fMainForm.ScaleHeight
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub Image10_DblClick()
    Dim frm As frmCCIMatPriceGrid
    Set frm = New frmCCIMatPriceGrid
    frm.Show
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub Image11_DblClick()
    Dim frm As New frmReportMenu
    frm.Show
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub Image12_Click()
    ShowMinimizedForms
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub Image12_DblClick()
    Dim frm As frmAdminMenu
    If g_blnIsUserAdmin Then
        Set frm = New frmAdminMenu
        frm.Show
    Else
        MsgBox "Sorry, but this function is limited to CCD Administrators.", vbInformation
    End If
End Sub

Private Sub Image7_DblClick()
    Dim frm As frmAssemblyGrid
    Set frm = New frmAssemblyGrid
    frm.Show
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub imgProjects_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub imgCrews_DblClick()
    Dim frm As frmCrewGrid
    Screen.MousePointer = vbHourglass
    Set frm = New frmCrewGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub imgInfoSource_DblClick()
    Dim frm As frmInfoSourceGrid
    Set frm = New frmInfoSourceGrid
    frm.Show
End Sub

Private Sub imgInfoSource_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub imgLabor_DblClick()
    Screen.MousePointer = vbHourglass
    Dim frm As frmLaborRateGrid
    Set frm = New frmLaborRateGrid
    frm.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub imgLabor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub imgMaterial_DblClick()
    Dim frm As frmMatPriceGrid
    Set frm = New frmMatPriceGrid
    frm.m_blnFirstSearch = True
End Sub

Private Sub imgMaterial_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub imgEquipment_DblClick()
    Dim frm As frmEquipRateGrid
    Set frm = New frmEquipRateGrid
    frm.m_blnFirstSearch = True
End Sub

Private Sub imgEquipment_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

Private Sub imgModels_DblClick()
    Dim frm As frmBuildingGrid
    Set frm = New frmBuildingGrid
    frm.Show
End Sub

Private Sub imgProjects_DblClick()
    Dim frm As frmProjectGrid
    Set frm = New frmProjectGrid
    frm.Show
End Sub

Private Sub imgUnitCost_DblClick()
    Dim frm As frmUnitCostGrid
    Set frm = New frmUnitCostGrid
    frm.Show
End Sub

Private Sub imgUnitCost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMinimizedForms
End Sub

