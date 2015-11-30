VERSION 5.00
Begin VB.Form frmToList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parameters"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmToList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Get Defaults Values"
      Height          =   350
      Left            =   0
      TabIndex        =   11
      Top             =   5400
      Width           =   1590
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   2875
      TabIndex        =   13
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   1950
      TabIndex        =   12
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtUserFees 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   4320
      Width           =   1170
   End
   Begin VB.TextBox txtBldgArchFees 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3960
      Width           =   1170
   End
   Begin VB.TextBox txtOPFactor 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   3600
      Width           =   1170
   End
   Begin VB.TextBox txtBldgDoorDensity 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3240
      Width           =   1170
   End
   Begin VB.TextBox txtBldgPartHgt 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   1170
   End
   Begin VB.TextBox txtBldgPartDensity 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2520
      Width           =   1170
   End
   Begin VB.TextBox txtBldgStoriesHgt 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   1170
   End
   Begin VB.TextBox txtBldgStories 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   1170
   End
   Begin VB.TextBox txtPerimeter 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1170
   End
   Begin VB.TextBox txtSFArea 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   1170
   End
   Begin VB.CheckBox chkIncludeBasement 
      Caption         =   "Include Basement in Costs"
      Height          =   375
      Left            =   630
      TabIndex        =   10
      Top             =   4800
      Width           =   2325
   End
   Begin VB.Label lblUserFees 
      Alignment       =   1  'Right Justify
      Caption         =   "User Fees (%)"
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   4440
      Width           =   2430
   End
   Begin VB.Label lblBldgArchFees 
      Alignment       =   1  'Right Justify
      Caption         =   "Architectural Fees (%)"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   4080
      Width           =   2430
   End
   Begin VB.Label lblOPFactor 
      Alignment       =   1  'Right Justify
      Caption         =   "Contractors Overhead && Profit (%)"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   3720
      Width           =   2430
   End
   Begin VB.Label lblBldgDoorDensity 
      Alignment       =   1  'Right Justify
      Caption         =   "Door Density (Ea.)"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   3360
      Width           =   2430
   End
   Begin VB.Label lblBldgPartHgt 
      Alignment       =   1  'Right Justify
      Caption         =   "Partition Height (L.F.)"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   3000
      Width           =   2430
   End
   Begin VB.Label lblBldgPartDensity 
      Alignment       =   1  'Right Justify
      Caption         =   "Partition Density (L.F./S.F.)"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   2640
      Width           =   2430
   End
   Begin VB.Label lblBldgStoriesHgt 
      Alignment       =   1  'Right Justify
      Caption         =   "Story Height (L.F.)"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   2280
      Width           =   2430
   End
   Begin VB.Label lblBldgStories 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Stories (Ea.)"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   1920
      Width           =   2430
   End
   Begin VB.Label lblPerimeter 
      Alignment       =   1  'Right Justify
      Caption         =   "Perimeter (L.F.)"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   1560
      Width           =   2430
   End
   Begin VB.Label lblSFArea 
      Alignment       =   1  'Right Justify
      Caption         =   "Area (S.F.)"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   1200
      Width           =   2430
   End
   Begin VB.Label Label1 
      Caption         =   $"frmToList.frx":014A
      Height          =   1095
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   3270
   End
End
Attribute VB_Name = "frmToList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_rec As New ADODB.RecordSet
'
'   This routine is always called to load the form.
Public Sub SetRow(rec As ADODB.RecordSet)
    
    Screen.MousePointer = vbHourglass
    Set m_rec = rec
    PopulateScreen
    EnableControls
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateScreen()
    With m_rec
        txtSFArea.Text = .Fields("bldg_area_std").Value
        txtPerimeter.Text = .Fields("bldg_perimeter_std").Value
        txtBldgStories.Text = .Fields("bldg_stories").Value
        txtBldgStoriesHgt.Text = .Fields("bldg_stories_hgt").Value
        txtBldgPartDensity.Text = .Fields("bldg_part_density").Value
        txtBldgPartHgt.Text = .Fields("bldg_part_hgt").Value
        txtBldgDoorDensity.Text = .Fields("bldg_door_density").Value
        txtOPFactor.Text = .Fields("op_factor").Value
        txtBldgArchFees.Text = .Fields("bldg_arch_fees").Value
    End With
End Sub

Private Sub EnableControls()
'
End Sub

Private Sub cmdOK_Click()
    Dim frm As New frmToListRpt
    
   ' frm.RunReport rec
        
End Sub

Private Sub cmdCancel_Click()
'
End Sub

Private Sub cmdDefaults_Click()
'
End Sub


