VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmModel 
   Caption         =   "Model Maintenance"
   ClientHeight    =   7125
   ClientLeft      =   1965
   ClientTop       =   495
   ClientWidth     =   13065
   Icon            =   "frmModel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   13065
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "REFRESH ASM GRIDS"
      Height          =   255
      Left            =   7560
      TabIndex        =   207
      Top             =   6360
      Width           =   2085
   End
   Begin VB.CommandButton cmdDeleteClone 
      Caption         =   "Delete Clone"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10560
      TabIndex        =   199
      Top             =   6360
      Width           =   1000
   End
   Begin VB.Frame fraModelMatrixResi 
      Height          =   2100
      Left            =   60
      TabIndex        =   114
      Top             =   0
      Width           =   10960
      Begin VB.Shape shpSelectedAreaPerimeterResi 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   285
         Left            =   1330
         Top             =   120
         Width           =   875
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8300
         TabIndex        =   198
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7425
         TabIndex        =   197
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6560
         TabIndex        =   196
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5685
         TabIndex        =   195
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4810
         TabIndex        =   194
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3950
         TabIndex        =   193
         Top             =   410
         Width           =   875
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3080
         TabIndex        =   192
         Top             =   410
         Width           =   875
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2210
         TabIndex        =   191
         Top             =   410
         Width           =   875
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   190
         Top             =   410
         Width           =   875
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   10920
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   189
         Top             =   1770
         Width           =   875
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2210
         TabIndex        =   188
         Top             =   1770
         Width           =   875
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3080
         TabIndex        =   187
         Top             =   1770
         Width           =   875
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3950
         TabIndex        =   186
         Top             =   1770
         Width           =   875
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4810
         TabIndex        =   185
         Top             =   1770
         Width           =   870
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5685
         TabIndex        =   184
         Top             =   1770
         Width           =   870
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6560
         TabIndex        =   183
         Top             =   1770
         Width           =   870
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7425
         TabIndex        =   182
         Top             =   1770
         Width           =   870
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8300
         TabIndex        =   181
         Top             =   1770
         Width           =   870
      End
      Begin VB.Label lblTotalOPdescResi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   180
         Top             =   1770
         Width           =   1305
      End
      Begin VB.Label lblInstallOPdescResi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Install O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   179
         Top             =   1500
         Width           =   1305
      End
      Begin VB.Label lblEquipmentOPdescResi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Equipment O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   178
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Label lblLaborOPdescResi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Labor O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   177
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8300
         TabIndex        =   176
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8300
         TabIndex        =   175
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8300
         TabIndex        =   174
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8300
         TabIndex        =   173
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7425
         TabIndex        =   172
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7425
         TabIndex        =   171
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7425
         TabIndex        =   170
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7425
         TabIndex        =   169
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6560
         TabIndex        =   168
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6560
         TabIndex        =   167
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6560
         TabIndex        =   166
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6560
         TabIndex        =   165
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5685
         TabIndex        =   164
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5685
         TabIndex        =   163
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5685
         TabIndex        =   162
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5685
         TabIndex        =   161
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4810
         TabIndex        =   160
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4810
         TabIndex        =   159
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4810
         TabIndex        =   158
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4810
         TabIndex        =   157
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3950
         TabIndex        =   156
         Top             =   690
         Width           =   875
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3950
         TabIndex        =   155
         Top             =   960
         Width           =   875
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3950
         TabIndex        =   154
         Top             =   1230
         Width           =   875
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3950
         TabIndex        =   153
         Top             =   1500
         Width           =   875
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3080
         TabIndex        =   152
         Top             =   690
         Width           =   875
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3080
         TabIndex        =   151
         Top             =   960
         Width           =   875
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3080
         TabIndex        =   150
         Top             =   1230
         Width           =   875
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3080
         TabIndex        =   149
         Top             =   1500
         Width           =   875
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2210
         TabIndex        =   148
         Top             =   690
         Width           =   875
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2210
         TabIndex        =   147
         Top             =   960
         Width           =   875
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2210
         TabIndex        =   146
         Top             =   1230
         Width           =   875
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2210
         TabIndex        =   145
         Top             =   1500
         Width           =   875
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   144
         Top             =   690
         Width           =   875
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   143
         Top             =   960
         Width           =   875
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   142
         Top             =   1230
         Width           =   875
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   141
         Top             =   1500
         Width           =   875
      End
      Begin VB.Label lblMaterialOPdescResi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Material O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   140
         Top             =   690
         Width           =   1305
      End
      Begin VB.Label lblSFAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S.F. Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   139
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblLFPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L.F.Perimeter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   138
         Top             =   405
         Width           =   1305
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8300
         TabIndex        =   137
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7425
         TabIndex        =   136
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6560
         TabIndex        =   135
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5685
         TabIndex        =   134
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4810
         TabIndex        =   133
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3950
         TabIndex        =   132
         Top             =   120
         Width           =   875
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3080
         TabIndex        =   131
         Top             =   120
         Width           =   875
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2210
         TabIndex        =   130
         Top             =   120
         Width           =   875
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   129
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   9160
         TabIndex        =   128
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   9160
         TabIndex        =   127
         Top             =   1770
         Width           =   870
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   9160
         TabIndex        =   126
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   9160
         TabIndex        =   125
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   9160
         TabIndex        =   124
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   9160
         TabIndex        =   123
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   9160
         TabIndex        =   122
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblPerimeterResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   10030
         TabIndex        =   121
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblTotalOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   10030
         TabIndex        =   120
         Top             =   1770
         Width           =   870
      End
      Begin VB.Label lblMaterialOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   10030
         TabIndex        =   119
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lblLaborOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   10030
         TabIndex        =   118
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblEquipmentOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   10030
         TabIndex        =   117
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblInstallOPResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   10030
         TabIndex        =   116
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label lblAreaResi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   10030
         TabIndex        =   115
         Top             =   120
         Width           =   870
      End
   End
   Begin TabDlg.SSTab tabModelDetails 
      Height          =   3840
      Left            =   60
      TabIndex        =   92
      Top             =   2160
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   6773
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "&Assembly Components"
      TabPicture(0)   =   "frmModel.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStdSFArea"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblStdPerimeter"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAssemblyCompRowCount"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblFormulaCode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TDBGridAssembly"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDeleteAssembly"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCloneStandardModel"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&Summary Estimate"
      TabPicture(1)   =   "frmModel.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TDBGridSummaryEstimate"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Model Details"
      TabPicture(2)   =   "frmModel.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblbldg_desc"
      Tab(2).Control(1)=   "lblbldg_category"
      Tab(2).Control(2)=   "lblbldg_id"
      Tab(2).Control(3)=   "lblFrameType"
      Tab(2).Control(4)=   "lblWallType"
      Tab(2).Control(5)=   "lblFormatCode"
      Tab(2).Control(6)=   "lblMdlCountryCode"
      Tab(2).Control(7)=   "lblMdlRegionCode"
      Tab(2).Control(8)=   "Label1"
      Tab(2).Control(9)=   "txtbldg_desc"
      Tab(2).Control(10)=   "cbobldg_category"
      Tab(2).Control(11)=   "txtbldg_id"
      Tab(2).Control(12)=   "fratype_code"
      Tab(2).Control(13)=   "cboFrameType"
      Tab(2).Control(14)=   "cboWallType"
      Tab(2).Control(15)=   "cboFormatCode"
      Tab(2).Control(16)=   "fraOPCode"
      Tab(2).Control(17)=   "cboMdlCountryCode"
      Tab(2).Control(18)=   "cboMdlRegionCode"
      Tab(2).Control(19)=   "txtBldgCostDesc"
      Tab(2).Control(20)=   "txtCostWorksDesc"
      Tab(2).ControlCount=   21
      Begin VB.CommandButton cmdCloneStandardModel 
         Caption         =   "Clone Standard Model"
         Height          =   495
         Left            =   11400
         TabIndex        =   208
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtCostWorksDesc 
         Height          =   315
         Left            =   -73845
         MaxLength       =   75
         TabIndex        =   13
         Top             =   2820
         Width           =   10815
      End
      Begin VB.TextBox txtBldgCostDesc 
         Height          =   525
         Left            =   -74895
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   113
         Top             =   480
         Width           =   5800
      End
      Begin VB.ComboBox cboMdlRegionCode 
         Height          =   315
         ItemData        =   "frmModel.frx":019E
         Left            =   -65610
         List            =   "frmModel.frx":01A5
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   1005
      End
      Begin VB.ComboBox cboMdlCountryCode 
         Height          =   315
         ItemData        =   "frmModel.frx":01AE
         Left            =   -63870
         List            =   "frmModel.frx":01B5
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   615
         Width           =   885
      End
      Begin VB.Frame fraOPCode 
         Caption         =   "OP Code"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   -68070
         TabIndex        =   110
         Top             =   480
         Width           =   1755
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1575
            TabIndex        =   201
            Top             =   240
            Width           =   1575
            Begin VB.OptionButton optOpen 
               Caption         =   "Open"
               Height          =   240
               Left            =   855
               TabIndex        =   203
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton optUnion 
               Caption         =   "Union"
               Height          =   255
               Left            =   0
               TabIndex        =   202
               Top             =   0
               Value           =   -1  'True
               Width           =   795
            End
         End
      End
      Begin VB.ComboBox cboFormatCode 
         Height          =   315
         ItemData        =   "frmModel.frx":01BE
         Left            =   -66720
         List            =   "frmModel.frx":01C0
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cboWallType 
         Height          =   315
         ItemData        =   "frmModel.frx":01C2
         Left            =   -73845
         List            =   "frmModel.frx":01C4
         TabIndex        =   11
         Text            =   "cboWallTypeSelection"
         Top             =   2280
         Width           =   4740
      End
      Begin VB.ComboBox cboFrameType 
         Height          =   315
         ItemData        =   "frmModel.frx":01C6
         Left            =   -66720
         List            =   "frmModel.frx":01C8
         TabIndex        =   12
         Text            =   "cboFrameTypeSelection"
         Top             =   2280
         Width           =   3690
      End
      Begin VB.CommandButton cmdDeleteAssembly 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   10200
         TabIndex        =   16
         Top             =   3240
         Width           =   1000
      End
      Begin VB.Frame fratype_code 
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   -72000
         TabIndex        =   96
         Top             =   1120
         Width           =   2925
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   2655
            TabIndex        =   204
            Top             =   180
            Width           =   2655
            Begin VB.OptionButton opttype_codeR 
               Caption         =   "Residential"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1320
               TabIndex        =   206
               Top             =   0
               Width           =   1245
            End
            Begin VB.OptionButton opttype_codeC 
               Caption         =   "Commercial"
               Enabled         =   0   'False
               Height          =   255
               Left            =   0
               TabIndex        =   205
               Top             =   0
               Value           =   -1  'True
               Width           =   1215
            End
         End
      End
      Begin VB.TextBox txtbldg_id 
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
         Left            =   -73845
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   1320
         Width           =   1320
      End
      Begin VB.ComboBox cbobldg_category 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmModel.frx":01CA
         Left            =   -66705
         List            =   "frmModel.frx":01CC
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   1320
         Width           =   2115
      End
      Begin VB.TextBox txtbldg_desc 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   -73845
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   93
         Top             =   1800
         Width           =   4740
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridAssembly 
         Height          =   2415
         Left            =   105
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   4260
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
         DeadAreaBackColor=   12632256
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
      Begin TrueOleDBGrid80.TDBGrid TDBGridSummaryEstimate 
         Height          =   3075
         Left            =   -74895
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   5424
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
         DataMode        =   2
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CostWorks Description:"
         Height          =   495
         Left            =   -74880
         TabIndex        =   200
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label lblMdlRegionCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Region:"
         Height          =   255
         Left            =   -66285
         TabIndex        =   112
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lblMdlCountryCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Country:"
         Height          =   255
         Left            =   -64605
         TabIndex        =   111
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lblFormatCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Format Code:"
         Height          =   255
         Left            =   -67800
         TabIndex        =   109
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblWallType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wall Type:"
         Height          =   255
         Left            =   -74895
         TabIndex        =   108
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label lblFrameType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frame Type:"
         Height          =   255
         Left            =   -68040
         TabIndex        =   107
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label lblFormulaCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   105
         TabIndex        =   105
         Top             =   3240
         Width           =   9960
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAssemblyCompRowCount 
         Alignment       =   1  'Right Justify
         Caption         =   "0 rows returned"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   6480
         TabIndex        =   104
         Top             =   480
         Width           =   4725
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Std SF Area:"
         Height          =   255
         Left            =   105
         TabIndex        =   103
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Std Perimeter:"
         Height          =   255
         Left            =   2520
         TabIndex        =   102
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lblStdPerimeter 
         Height          =   255
         Left            =   3675
         TabIndex        =   101
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblStdSFArea 
         Height          =   255
         Left            =   1260
         TabIndex        =   100
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblbldg_id 
         Alignment       =   1  'Right Justify
         Caption         =   "Building ID:"
         Height          =   255
         Left            =   -74895
         TabIndex        =   99
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label lblbldg_category 
         Alignment       =   1  'Right Justify
         Caption         =   "Category:"
         Height          =   255
         Left            =   -68040
         TabIndex        =   98
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblbldg_desc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Building:"
         Height          =   255
         Left            =   -74895
         TabIndex        =   97
         Top             =   1800
         Width           =   960
      End
   End
   Begin VB.Frame fraGoTo 
      Caption         =   "Go To"
      Height          =   855
      Left            =   60
      TabIndex        =   106
      Top             =   6120
      Width           =   3330
      Begin VB.CommandButton cmdAssemblyComponentsReport 
         Caption         =   "Component Report"
         Height          =   540
         Left            =   2160
         TabIndex        =   21
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdReports 
         Caption         =   "Summary &Report"
         Height          =   540
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdUnitCost 
         Caption         =   "Unit &Cost"
         Height          =   540
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame fraModelMatrix 
      Height          =   2100
      Left            =   60
      TabIndex        =   30
      Top             =   0
      Width           =   9750
      Begin VB.Shape shpSelectedAreaPerimeter 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   285
         Left            =   1330
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8770
         TabIndex        =   91
         Top             =   410
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7840
         TabIndex        =   90
         Top             =   410
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6910
         TabIndex        =   89
         Top             =   410
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5980
         TabIndex        =   88
         Top             =   410
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5050
         TabIndex        =   87
         Top             =   410
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4120
         TabIndex        =   86
         Top             =   410
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3190
         TabIndex        =   85
         Top             =   410
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2260
         TabIndex        =   84
         Top             =   410
         Width           =   930
      End
      Begin VB.Label lblPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   83
         Top             =   410
         Width           =   930
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   11970
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   82
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2260
         TabIndex        =   81
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3190
         TabIndex        =   80
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4120
         TabIndex        =   79
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5050
         TabIndex        =   78
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5980
         TabIndex        =   77
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6910
         TabIndex        =   76
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7840
         TabIndex        =   75
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8770
         TabIndex        =   74
         Top             =   1770
         Width           =   930
      End
      Begin VB.Label lblTotalOPdesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   73
         Top             =   1770
         Width           =   1305
      End
      Begin VB.Label lblInstallOPdesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Install O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   72
         Top             =   1500
         Width           =   1305
      End
      Begin VB.Label lblEquipmentOPdesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Equipment O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   71
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Label lblLaborOPdesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Labor O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   70
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8770
         TabIndex        =   69
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8770
         TabIndex        =   68
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8770
         TabIndex        =   67
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8770
         TabIndex        =   66
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7840
         TabIndex        =   65
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7840
         TabIndex        =   64
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7840
         TabIndex        =   63
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7840
         TabIndex        =   62
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6910
         TabIndex        =   61
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6910
         TabIndex        =   60
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6915
         TabIndex        =   59
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6910
         TabIndex        =   58
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5980
         TabIndex        =   57
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5980
         TabIndex        =   56
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5980
         TabIndex        =   55
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5980
         TabIndex        =   54
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5050
         TabIndex        =   53
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5050
         TabIndex        =   52
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5050
         TabIndex        =   51
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5050
         TabIndex        =   50
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4120
         TabIndex        =   49
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4120
         TabIndex        =   48
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4120
         TabIndex        =   47
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4120
         TabIndex        =   46
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3190
         TabIndex        =   45
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3190
         TabIndex        =   44
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3190
         TabIndex        =   43
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3190
         TabIndex        =   42
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2260
         TabIndex        =   41
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2260
         TabIndex        =   40
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2260
         TabIndex        =   39
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2260
         TabIndex        =   38
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   37
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblLaborOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   36
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblEquipmentOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   35
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label lblInstallOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   34
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblMaterialOPdesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Material O&&P  "
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   33
         Top             =   690
         Width           =   1305
      End
      Begin VB.Label lblSFArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S.F. Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   32
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblLFPerimeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L.F.Perimeter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   25
         TabIndex        =   31
         Top             =   405
         Width           =   1305
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8770
         TabIndex        =   8
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7840
         TabIndex        =   7
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6910
         TabIndex        =   6
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5980
         TabIndex        =   5
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5050
         TabIndex        =   4
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4120
         TabIndex        =   3
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3190
         TabIndex        =   2
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2260
         TabIndex        =   1
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1330
         TabIndex        =   0
         Top             =   120
         Width           =   930
      End
   End
   Begin VB.TextBox txtbldg_skey 
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "1N"
      Top             =   6300
      Width           =   750
   End
   Begin VB.TextBox txtlast_update_person 
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "S"
      Top             =   6690
      Width           =   1170
   End
   Begin VB.TextBox txtlast_update_date 
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6690
      Width           =   2310
   End
   Begin VB.TextBox txtbldg_model_skey 
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "1N"
      Top             =   6300
      Width           =   630
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   11760
      TabIndex        =   18
      Top             =   6360
      Width           =   1000
   End
   Begin VB.Label lblbldg_skey 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Skey:"
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label lbllast_update_person 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Updated By:"
      Height          =   255
      Left            =   7425
      TabIndex        =   27
      Top             =   6720
      Width           =   1035
   End
   Begin VB.Label lbllast_update_date 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Updated:"
      Height          =   255
      Left            =   3870
      TabIndex        =   26
      Top             =   6720
      Width           =   705
   End
   Begin VB.Label lblbldg_model_skey 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mdl Skey:"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   6360
      Width           =   810
   End
End
Attribute VB_Name = "frmModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cnTemp As New ADODB.Connection

Dim m_rec As ADODB.RecordSet
'
'   Common Assemblies grid recordset.
Dim m_recAssembly As New ADODB.RecordSet
'
Dim m_recModelComponent As New ADODB.RecordSet
'
'   Published Bldg Matrix Cost = AreasPerimeters
Dim m_recModelMatrix As New ADODB.RecordSet
'
'   Tells if we are doing an insert or update.
Dim m_blnInsert As Boolean
'
'   Indicate if clone is in progress
Dim m_blnClone As Boolean
'
'   Tells us we're loading the screen for the 1st time
'   so the cbo clicks won't run.
Dim bIsInitialLoad  As Boolean
'
'   Indicates user has modified the data.
Dim bIsPendingChange As Boolean
'
'   Indicates a field that will affect overall cost rollups
'   has been changed so we must RefreshCosts if they update.
Dim bRefreshCosts As Boolean
'
'   Class to handle Rt Pg grid.
Dim m_objModelComponentGridMap As New CMdlComponent
'
'   Class to handle Assemblies grid.
Private WithEvents m_objModelAssembliesGridMap As CMdlAssembly
Attribute m_objModelAssembliesGridMap.VB_VarHelpID = -1
'
'   Notifies that it wants to see changes.
Dim sEventSubscriberID              As String
'
'   Indicates the Area to bold.
'   Set to 0 if we're not on a model that is bolded.
Dim nAreaInd As Integer
'
'   Indicates where the shpSelectedArea is for modelmaint button click.
'   In the format of 1,1 meaning row 1 col area1.
Dim sshpSelectedArea As String
'
'   Used to indicate how the bldg costs are calculated as shown in the book.
Dim Bldg_Cost_Desc_Container As String
'
'
Private Const FORM_WIDTH As Integer = 13300
Private Const FORM_HEIGHT As Integer = 7600

' These row/col settings are set when we populate the model matrix grid.  We use them when the user
' clicks the clone standard assembly button on the assembly components tab to read in the components for
' the standard model instead of the one they originally selected.  This saves them a lot of time
' when typing in the assembly ids.
Private m_iStandardModelRow As Integer
Private m_iStandardModelCol As Integer

Private Const STANDARD_MODEL_CODE_MIN As String = "1"
Private Const STANDARD_MODEL_CODE_MAX As String = "6"

Private Sub cmdAssemblyComponentsReport_Click()
    PreviewAssemblyComponentsReport
End Sub

' This sub reads the assembly components from the standard model for this building, regardless of what model
' the user originally clicked on.  The assembly grid is populated with the standard model's assembly components.
Private Sub cmdCloneStandardModel_Click()
    
    On Error GoTo Err_Handler
    
    ' Ask the user if it is ok to overwrite the original components.
    If (TDBGridAssembly.ApproxCount > 0) Then
        Dim Result As Integer
        Result = MsgBox("Do you wish to replace the existing assembly components with the ones from " & _
                        "the standard model?  If you select Yes, the current assembly components for this model will be " & _
                        "immediately and permanently deleted.", vbYesNoCancel)
        If (Result <> vbYes) Then
            Exit Sub
        End If
    End If
    
    m_objModelAssembliesGridMap.bCloneAssembliesInProcess = True

    
    Dim bldgModelSKey As Long
    Dim bldgArea As Long
    Dim recTemp As New ADODB.RecordSet
    Dim strSelect As String
    Dim origBldgModelSKey As Long
    Dim strErrMsg As String
    
    ' This is the skey of the model the user clicked originally.
    origBldgModelSKey = m_rec.Fields("bldg_model_skey")
    
    ' We saved the standard model row and col when we built the upper matrix.
    ' We need to get the building model skey for the standard model row for this building.
    strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & m_rec.Fields("bldg_skey").Value & _
                "' AND model_code = '" & m_iStandardModelRow & "'"
            
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Or recTemp.RecordCount = 0 Then
        Exit Sub
    End If
    
    bldgModelSKey = GetBldgModelSKey(recTemp.Fields("bldg_model_skey").Value)
    recTemp.Close
    
    ' This is the area for the standard model column.
    bldgArea = Trim(lblArea(m_iStandardModelCol - 1).Caption)
    
    ' Check to see if the user is trying to clone while on the standard model itself which is
    ' not allowed.
    If (origBldgModelSKey = bldgModelSKey) Then
        MsgBox "You cannot clone from the standard model itself.", vbInformation
        Exit Sub
    End If
    
    
    ' We first have to delete the existing assemblies from the grid and db.  Do this by first
    ' selecting every row, then calling the delete routine.
    Dim Index As Integer
    TDBGridAssembly.MoveFirst
    
    ' Select all rows
    Do While TDBGridAssembly.EOF = False
    
        TDBGridAssembly.SelBookmarks.Add TDBGridAssembly.Bookmark
        TDBGridAssembly.MoveNext
    
    Loop
    
    ' Now delete them all.
    m_objModelAssembliesGridMap.Delete
    
    
    ' Since we are cloning, it's as if the user added the assembly components to the model, so force
    ' a refresh of the costs.
    bRefreshCosts = True
    
    ' Creates the stored proc call, performs the fix-up, and populates the on-screen grid control with the
    ' assemblies from the standard model.
    Dim populateResult As Boolean
    populateResult = PopulateAssemblyComponents(bldgModelSKey, bldgArea, origBldgModelSKey, strErrMsg)
    If (populateResult = True) Then
        Err.Raise ERROR_CLONING_ASSEMBLIES, "frmModel", "Error populating assembly grid: " & _
            strErrMsg
    End If
       
    ' Need to set up the row info structure so that when the user clicks Update, it will
    ' write the grid data to the db (otherwise the update routine won't know of our new rows in the grid).
    ' Note:  we can't do this until after we call the populate routine because that routine will clear out the
    ' data structure every time.
    m_objModelAssembliesGridMap.UpdateRowInfo m_recAssembly
    
    cmdUpdate.Enabled = True
    bIsPendingChange = True
    
    Exit Sub
    
    
Err_Handler:

    Dim errMsg As String
    Dim errNum As Long
    Dim errSrc As String
    
    errMsg = Err.Description
    errNum = Err.Number
    errSrc = Err.Source
    
    Screen.MousePointer = vbNormal
    MsgBox "Error " & errNum & " cloning assemblies:  " & errMsg & ", Source = " & errSrc, vbCritical
    Status ("")
    
End Sub

Private Sub cmdRefresh_Click()
Dim strUpdate       As String
    Dim cmdTemp         As New ADODB.Command
    Dim i               As Integer

    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
'    RefreshCostsCommercial = True
    Status ("Updating Building Cost Information For Model: " & txtbldg_model_skey.Text & " ...")
    With cnTemp
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
        
        'rlh - 06/06/07 CCD 8.2 release
        strUpdate = "exec sp_rollup_bldg_model_new @BLDG_ID_PARM = '"
        strUpdate = strUpdate & Trim(txtbldg_id.Text) & "'"
        
        If txtbldg_model_skey = "" Then
            'exit can't refresh yet.
            Exit Sub
        Else
            'strUpdate = strUpdate & Trim(txtbldg_model_skey.Text) & "',"
        End If
        '---------------------------------------------------
        'DO FOR "STD" SHOP      'rlh 06/17/2008
        '---------------------------------------------------
        strUpdate = strUpdate & ",@op_code_parm = 'STD'"
       
        With cmdTemp
            .CommandTimeout = 50000
            .CommandType = adCmdText
            .CommandText = strUpdate
            
            ' rlh
            If DEBUGON Then
                Debug.Print "frmModel: cmdRefresh (#1): " & strUpdate
            End If
            
            .Execute adExecuteNoRecords
        End With
        '---------------------------------------------------
        'DO FOR "OPN" shop      'rlh 06/17/2008
        '---------------------------------------------------
        If cnTemp.Errors.Count = 0 Then
            strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
            With cmdTemp
                .CommandTimeout = 50000
                .CommandType = adCmdText
                .CommandText = strUpdate
                 ' rlh
                If DEBUGON Then
                    Debug.Print "frmModel: cmdRefresh (errors count=0): " & strUpdate
                End If
                
                .Execute adExecuteNoRecords
            End With
            If cnTemp.Errors.Count <> 0 Then
                Screen.MousePointer = vbNormal
                'MsgBox "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                    & Trim(txtbldg_model_skey.Text) & " " & vbCrLf & cnTemp.Errors(0).Description, vbCritical
                    MsgBox "Errors in the cmdRefresh_Click() routine for Building Model skey: " _
                    & Trim(txtbldg_model_skey.Text) & " " & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            Else
                MsgBox "Assembly Grid Refresh is Complete"
                Screen.MousePointer = vbNormal
            End If
        Else
            Screen.MousePointer = vbNormal
            MsgBox "Errors in the cmdRefresh_Click() routine for Building Model skey: " _
                & Trim(txtbldg_model_skey.Text) & " " & vbCrLf & cnTemp.Errors(0).Description, vbCritical
        End If
    End With
    Exit Sub

errorHandler:
    Screen.MousePointer = vbNormal
'    RefreshCostsCommercial = False
    MsgBox "Errors in the cmdRefresh_Click() routine: " & Err.Description
    Status ("")
End Sub


Private Sub Form_Activate()
    ShowPrintToolbar True
End Sub

Private Sub Form_Deactivate()
    ShowPrintToolbar False
End Sub

Private Sub Form_Initialize()
    Screen.MousePointer = vbHourglass
    Status ("Loading Model Maintenance ...")
    sEventSubscriberID = EventSubscriberAdd(Me)
    
    Bldg_Cost_Desc_Container = "Model costs calculated for a [BLDG_STORIES] story building with [STORIES_HGT]' " & _
            "story height and" & vbCrLf & " [BLDG_AREA] square feet of floor area"
    
    sshpSelectedArea = "0,0"
    Set m_objModelAssembliesGridMap = New CMdlAssembly

    With cnTemp
        .ConnectionTimeout = 0
        .CommandTimeout = 0
        '.Open "UID=" + strUserName + ";PWD=;DATABASE=" + strConnectDatabase + ";SERVER=" + strConnectServer + ";DRIVER={SQL SERVER};DSN='';"
        .Open strConnect
    End With
    
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Move 150, 200
    Me.Height = FORM_HEIGHT
    Me.Width = FORM_WIDTH
    ColorLockedFields Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button              As String
    Dim blnPendingChange    As Boolean
    Dim bln_New             As Boolean
    '
    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        If bIsPendingChange = True Or m_objModelComponentGridMap.IsPendingChange _
            Or m_objModelAssembliesGridMap.IsPendingChange Then
            
            Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Close Model Form")
            If Button = vbYes Then
                '
                '   If there were errors, cancel the close.
                If Not Update Then
                    Cancel = True
                End If
            ElseIf Button = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ShowPrintToolbar False
    EventSubscriberRemove sEventSubscriberID
    m_rec.Close
    m_recAssembly.Close
    m_recModelComponent.Close
    m_recModelMatrix.Close
    Set m_rec = Nothing
    Set m_recAssembly = Nothing
    Set m_recModelComponent = Nothing
    Set m_recModelMatrix = Nothing
    cnTemp.Close
    Set cnTemp = Nothing
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState <> vbMinimized Then
        Me.Height = FORM_HEIGHT
        Me.Width = FORM_WIDTH
    End If
    
End Sub

Public Sub EventNotify(eNotifyType As EEventSubscriberNotifyType, sAffectedRecordIdentifier As String)
   
    On Error Resume Next
    '
    '   If the record that was updated is for our bldg
    '   we need to refresh.
    If eNotifyType = esnBuildingRecordUpdated And _
        Trim(txtbldg_id.Text) = Trim(sAffectedRecordIdentifier) Then
        
        SearchForNewModel False, Trim(txtbldg_model_skey.Text), Trim(txtbldg_id.Text)
    End If
End Sub
'
'   This routine is always called to load the form.
Public Sub SetRow(ByVal rec As ADODB.RecordSet, Optional blnInsert As Boolean = False, _
                Optional sAreaToPopulate As String, Optional sOpCodeToUse As String)
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    bIsInitialLoad = True
    Set m_rec = rec
    m_blnInsert = blnInsert
    '
    '   If we are inserting/cloning.
    If m_blnInsert Then
        '
        '   Do this so OriginalValue will be set to
        '   the values copied into the row.
        m_rec.UpdateBatch
        
        If Trim(m_rec.Fields("bldg_model_skey").Value) <> 0 Then
            m_blnClone = True
        Else
            m_blnClone = False
        End If
    End If
    '
    '   Initialize grids.  Do this here because
    '   if we are on a format_row 7 or 8 then we have
    '   to lock all columns on the assemblies grid.
    With m_objModelComponentGridMap
        .SetGrid TDBGridSummaryEstimate
        .InitGrid
    End With
    If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
    Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
        With m_objModelAssembliesGridMap
            .SetGrid TDBGridAssembly
            .InitGrid False
        End With
    Else
        With m_objModelAssembliesGridMap
            .SetGrid TDBGridAssembly
            .InitGrid ((m_rec.Fields("format_row").Value) = 1)
        End With
    End If

    If sOpCodeToUse = "Open" Then
        optOpen.Value = True
    Else
        optUnion.Value = True
    End If
    '
    '   The following routines are called from here on initial load and
    '   when user selects different wall/frame.
    PopulateScreen True, sAreaToPopulate
    
    If m_blnInsert And m_blnClone = False Then
        tabModelDetails.Tab = 2
    ElseIf sAreaToPopulate <> "" And m_rec.Fields("format_row").Value <> 1 Then
        tabModelDetails.Tab = 1
    Else
        tabModelDetails.Tab = 0
    End If
    EnableControls
    bIsInitialLoad = False
    bIsPendingChange = False
    Status ("")
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateScreen(b1stBldgSearch As Boolean, Optional sAreaToPopulate As String)
    Dim recBuildingDetails  As New ADODB.RecordSet
    Dim strErrMsg As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass

    If b1stBldgSearch Then
        '
        '   Set defaults for 1st run, not called when user selects different wall/frame.
        '   Get the available wall type & frame type for all models for this bldg.
        PopulateFormatCodes
        Set recBuildingDetails = PopulateBuildingInformation
        '
        '   Get the type code for the building since the opttype_code not populated yet.
        PopulateAvailWallTypesFrameTypes recBuildingDetails.Fields("type_code").Value
    End If
    
    PopulateModelDetails recBuildingDetails
    
    If m_blnInsert And Not m_blnClone Then
        '
        '   Don't let them add assemblies or summary estimate
        '   until a model_skey is assigned.
        tabModelDetails.TabEnabled(0) = False
        tabModelDetails.TabEnabled(1) = False
        Me.Caption = "Model Maintenance [Bldg ID:" & Trim(txtbldg_id.Text) & " | New Model]"
    Else
        tabModelDetails.TabEnabled(0) = True
        '
        '   Always populate the matrix before Assembly or SummaryEstimate
        '   because it sets the op_code to the correct value.
        PopulateModelMatrix sAreaToPopulate
        PopulateAssemblyComponents GetBldgModelSKey(m_rec.Fields("bldg_model_skey").Value), GetBuildingArea(), -1, strErrMsg
        '
        '   Don't show SummaryEstimate for format rows.
        If (m_rec.Fields("format_row").Value) <> 1 Then
            tabModelDetails.TabEnabled(1) = True
            PopulateSummaryEstimateComponents
        Else
            If opttype_codeR.Value = True Then
                If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
                Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
                    '
                    '   Show that assemblies are applied to all model_codes 7 & 8 for the Quality Series.
                    TDBGridAssembly.ToolTipText = "Assemblies for Quality Series models are applied to all basement models for buildings within that Quality Series."
                Else
                    '
                    '   Show that assemblies are applied to all model_codes 7 & 8 for the Quality Series.
                    TDBGridAssembly.ToolTipText = "Assemblies for fomat rows match the assemblies for the corresponding format row within the Quality Series building."
                End If
            Else
                '
                '   Show that assemblies are for model_code 1.
                TDBGridAssembly.ToolTipText = "Assemblies for format rows match the assemblies for Model Code 1."
            End If
            tabModelDetails.TabEnabled(1) = False
        End If
        Me.Caption = "Model Maintenance [Bldg ID: " & Trim(txtbldg_id.Text) & " | Model Code: " _
            & Trim(m_rec.Fields("model_code").Value) & " | " & Trim(cboWallType.Text) & "]"
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateAvailWallTypesFrameTypes(sTypeCode As String)
    Dim recModels       As ADODB.RecordSet
    Dim strSelect       As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
        
    cboWallType.Clear
    cboFrameType.Clear
        
    If sTypeCode = "R" Then
        cboFrameType.Visible = False
    Else
        cboFrameType.Visible = True
    End If
    
    If sTypeCode = "C" Then
        strSelect = "SELECT DISTINCT frame_type FROM bldg_model WHERE model_code != '0' AND " _
                        & "model_code != '7' AND model_code != '8' AND frame_type != '' ORDER BY frame_type"
        '
        '   Use DAL to perform select.
        If Not g_objDAL.GetRecordset(vbNullString, strSelect, recModels) Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching.", vbCritical
        Else
           With recModels
                Do Until .EOF
                    cboFrameType.AddItem Trim(.Fields("frame_type").Value)
                    .MoveNext
                Loop
                .Close
            End With
        End If
    End If

    strSelect = "SELECT DISTINCT wall_type FROM bldg_model WHERE model_code != '0' AND " _
                    & "model_code != '7' AND model_code != '8' AND wall_type != '' ORDER BY wall_type"
    '
    '   Use DAL to perform select.
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, recModels) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching to populate available wall types.", vbCritical
    Else
       With recModels
            Do Until .EOF
                cboWallType.AddItem Trim(.Fields("wall_type").Value)
                .MoveNext
            Loop
            .Close
        End With
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateFormatCodes()

    On Error Resume Next
    cboFormatCode.Clear
    With cboFormatCode
        .AddItem "A3"
        .AddItem "A4"
    End With
End Sub

Private Sub PopulateModelDetails(recBuildingDetails As ADODB.RecordSet)
    Dim i As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   If we're not inserting then get the model information from the
    '   recordset passed in to the form.
    If Not m_blnInsert Or m_blnClone Then
        With m_rec
            For i = 0 To cboWallType.listcount - 1
                If Trim(.Fields("wall_type").Value) = Trim(cboWallType.List(i)) Then
                    cboWallType.ListIndex = i
                    Exit For
                End If
            Next i
            '
            '   In case we never found a match in our cbo add it.
            If Trim(cboWallType.Text) = "" Then
                cboWallType.AddItem Trim(.Fields("wall_type").Value)
                cboWallType.ListIndex = cboWallType.listcount - 1
            End If
            
            If opttype_codeC.Value = True Then
                For i = 0 To cboFrameType.listcount - 1
                    If Trim(.Fields("frame_type").Value) = Trim(cboFrameType.List(i)) Then
                        cboFrameType.ListIndex = i
                        Exit For
                    End If
                Next i
                '
                '   In case we never found a match in our cbo add it.
                If Trim(cboFrameType.Text) = "" Then
                    cboFrameType.AddItem Trim(.Fields("frame_type"))
                    cboFrameType.ListIndex = cboFrameType.listcount - 1
                End If
            End If
            
            If Trim(m_rec.Fields("format_row").Value) = 1 Then
                cboFormatCode.AddItem Trim(.Fields("format_code").Value)
                cboFormatCode.ListIndex = cboFormatCode.listcount - 1
            Else
                For i = 0 To cboFormatCode.listcount - 1
                    If Trim(.Fields("format_code").Value) = cboFormatCode.List(i) Then
                        cboFormatCode.ListIndex = i
                        Exit For
                    End If
                Next i
            End If
            
            'ADDED 7/5/2005 RTD CR#1530
            txtCostWorksDesc.Text = .Fields("costworks_desc").Value
        
            txtbldg_model_skey.Text = .Fields("bldg_model_skey").Value
            txtlast_update_date.Text = .Fields("last_update_date").Value
            txtlast_update_person.Text = .Fields("last_update_person").Value
        End With
    End If
    With recBuildingDetails
        If Not IsNull(.Fields("bldg_desc").Value) Then
            txtbldg_desc.Text = Trim(.Fields("bldg_desc").Value)
        Else
            txtbldg_desc.Text = ""
        End If
        
        If Not IsNull(.Fields("bldg_id").Value) Then
            txtbldg_id.Text = .Fields("bldg_id").Value
        Else
            txtbldg_id.Text = ""
        End If
        
        If Not IsNull(.Fields("bldg_skey").Value) Then
            txtbldg_skey.Text = .Fields("bldg_skey").Value
        Else
            txtbldg_skey.Text = ""
        End If
        
        If .Fields("type_code") = "C" Or IsNull(.Fields("type_code")) Then
            opttype_codeC.Value = True
        ElseIf .Fields("type_code") = "R" Then
            opttype_codeR.Value = True
        End If
        '
        '   They cannot change the bldg category once they have saved.
        cbobldg_category.AddItem Trim(.Fields("bldg_category"))
        cbobldg_category.ListIndex = 0
    End With
    Screen.MousePointer = vbNormal
End Sub

Private Function PopulateBuildingInformation() As ADODB.RecordSet
    Dim strSelect       As String
    Dim strBldgCostDesc As String
    Dim m_recBuilding   As New ADODB.RecordSet
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    strSelect = "exec sp_select_building @type_code = '%', @bldg_category = '%', @bldg_id = '" & m_rec.Fields("bldg_id").Value & "', @bldg_desc = '%'"
    '
    '   Use DAL to perform select.
    If g_objDAL.GetRecordset(vbNullString, strSelect, m_recBuilding) Then
        With m_recBuilding
            m_objModelAssembliesGridMap.bldgAreaStd = .Fields("bldg_area_std").Value
            m_objModelAssembliesGridMap.bldgPerimeterStd = .Fields("bldg_perimeter_std").Value
            m_objModelAssembliesGridMap.bldgDoorDensity = .Fields("bldg_door_density").Value
            m_objModelAssembliesGridMap.bldgPartDensity = .Fields("bldg_part_density").Value
            m_objModelAssembliesGridMap.bldgPartHgt = .Fields("bldg_part_hgt").Value
            m_objModelAssembliesGridMap.bldgStories = .Fields("bldg_stories").Value
            m_objModelAssembliesGridMap.bldgStoriesHgt = .Fields("bldg_stories_hgt").Value
            m_objModelAssembliesGridMap.BldgStairFactor = .Fields("bldg_stair_factor").Value
            m_objModelAssembliesGridMap.BldgFHeight = .Fields("bldg_f_height").Value
            m_objModelAssembliesGridMap.BldgFactor = .Fields("bldg_factor").Value
            m_objModelAssembliesGridMap.BldgType = .Fields("bldg_type").Value
            m_objModelAssembliesGridMap.WindowArea = .Fields("window_area").Value
            m_objModelAssembliesGridMap.TypeCode = .Fields("type_code").Value
            
            If optUnion.Value = True Then
                m_objModelAssembliesGridMap.opCode = "STD"
            Else
                m_objModelAssembliesGridMap.opCode = "OPN'"
            End If
            
            lblStdSFArea.Caption = .Fields("bldg_area_std").Value
            lblStdPerimeter.Caption = .Fields("bldg_perimeter_std").Value

            If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
            Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
                
                txtBldgCostDesc.Text = "Basement Models are used for all buildings within the Quality Series.  " & _
                    "Only their Assemblies may be modified."
            Else
                strBldgCostDesc = Replace(Bldg_Cost_Desc_Container, "[BLDG_STORIES]", .Fields("bldg_stories").Value)
                strBldgCostDesc = Replace(strBldgCostDesc, "[STORIES_HGT]", .Fields("bldg_stories_hgt").Value)
                '
                '   Now we have the bldg_stories & stories_hgt which won't change while we're in
                '   here but the area they select can change so reset the container to make it easy to
                '   replace the area each time they select a new one.
                Bldg_Cost_Desc_Container = strBldgCostDesc
                strBldgCostDesc = Replace(strBldgCostDesc, "[BLDG_AREA]", FormatNumber(.Fields("bldg_area_std").Value, 0))
                txtBldgCostDesc.Text = strBldgCostDesc
            End If
            Set PopulateBuildingInformation = m_recBuilding
        End With
    End If
    
    Screen.MousePointer = vbNormal
End Function

Private Sub PopulateModelMatrix(Optional sAreaToPopulate As String)
    Dim strSelect       As String
    Dim recTemp         As New ADODB.RecordSet
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    m_rec.MoveFirst
    '
    '   Make sure it is closed.
    With m_recModelMatrix
        .Close
        '
        '   Set the maximum number to bring back.
        .MaxRecords = MAX_RECORDS
    End With

    If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
    Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
        strSelect = "exec sp_select_building_pub_matrix_cost_basements @bldg_id = '"
        strSelect = strSelect & Trim(m_rec.Fields("bldg_id").Value)
        strSelect = strSelect & "', @bldg_model_skey = '%'"
    Else
        strSelect = "exec sp_select_building_pub_matrix_cost @bldg_id = '"
        If Len(Trim(txtbldg_id.Text)) > 0 Then
            strSelect = strSelect & Trim(txtbldg_id.Text)
        Else
            strSelect = strSelect & "%"
        End If
        
        strSelect = strSelect & "', @bldg_model_skey = '"
        If Len(Trim(m_rec.Fields("bldg_model_skey"))) > 0 Then
            strSelect = strSelect & Trim(m_rec.Fields("bldg_model_skey"))
        Else
            strSelect = strSelect & "%"
        End If
        
        strSelect = strSelect & "', @bldg_desc = '"
        If Len(Trim(txtbldg_desc.Text)) > 0 Then
           '
           '   We never know if we might have apos ' in our
           '   desc so replace for query.
           strSelect = strSelect & Replace(Trim(txtbldg_desc.Text), "'", "''") & "'"
        Else
            strSelect = strSelect & "%'"
        End If
        
        strSelect = strSelect & ", @op_code = '"
        If optUnion.Value = True Then
           strSelect = strSelect & "STD'"
        Else
           strSelect = strSelect & "OPN'"
        End If
        
        strSelect = strSelect & ", @country_code = '"
        If Len(Trim(cboMdlCountryCode.Text)) > 0 Then
            strSelect = strSelect & cboMdlCountryCode.Text & "'"
        Else
            strSelect = strSelect & "USA'"
        End If
        
        strSelect = strSelect & ", @region_code = '"
        If Len(Trim(cboMdlRegionCode.Text)) > 0 Then
            strSelect = strSelect & cboMdlRegionCode.Text & "'"
        Else
            strSelect = strSelect & "NAT'"
        End If
        
        If opttype_codeC.Value = True Then
            strSelect = strSelect & ", @type_code = 'C'"
        Else
            strSelect = strSelect & ", @type_code = 'R'"
        End If
    End If
    With cnTemp
        m_recModelMatrix.CursorLocation = adUseClient
        m_recModelMatrix.Open _
            Source:=strSelect, _
            ActiveConnection:=cnTemp, _
            CursorType:=adOpenStatic, _
            LockType:=adLockBatchOptimistic

        If .Errors.Count <> 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "Errors in the PopulateModelMatrix routine: " _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
        Else
            '
            '   Populate different frames based upon Commercial or Residential
            If opttype_codeC.Value = True Then
                PopulateModelMatrixCommercial sAreaToPopulate
            Else
                PopulateModelMatrixResi sAreaToPopulate
            End If
        End If
    End With
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateModelMatrixCommercial(sAreaToPopulate As String)
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With m_recModelMatrix
        If Not .EOF Then
            .MoveFirst
        
            lblArea(0).Caption = Trim(.Fields("area1").Value)
            lblPerimeter(0).Caption = Trim(.Fields("perimeter1").Value)
            lblArea(1).Caption = Trim(.Fields("area2").Value)
            lblPerimeter(1).Caption = Trim(.Fields("perimeter2").Value)
            lblArea(2).Caption = Trim(.Fields("area3").Value)
            lblPerimeter(2).Caption = Trim(.Fields("perimeter3").Value)
            lblArea(3).Caption = Trim(.Fields("area4").Value)
            lblPerimeter(3).Caption = Trim(.Fields("perimeter4").Value)
            lblArea(4).Caption = Trim(.Fields("area5").Value)
            lblPerimeter(4).Caption = Trim(.Fields("perimeter5").Value)
            lblArea(5).Caption = Trim(.Fields("area6").Value)
            lblPerimeter(5).Caption = Trim(.Fields("perimeter6").Value)
            lblArea(6).Caption = Trim(.Fields("area7").Value)
            lblPerimeter(6).Caption = Trim(.Fields("perimeter7").Value)
            lblArea(7).Caption = Trim(.Fields("area8").Value)
            lblPerimeter(7).Caption = Trim(.Fields("perimeter8").Value)
            lblArea(8).Caption = Trim(.Fields("area9").Value)
            lblPerimeter(8).Caption = Trim(.Fields("perimeter9").Value)

            lblMaterialOP(0).Caption = FormatNumber(.Fields("col1_mat_cost_op").Value, 2)
            lblMaterialOP(0).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col1_mat_cost_op").Value) * Trim(lblArea(0).Caption)), 2)
            
            lblLaborOP(0).Caption = FormatNumber(.Fields("col1_labor_cost_op").Value, 2)
            lblLaborOP(0).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col1_labor_cost_op").Value) * Trim(lblArea(0).Caption)), 2)
            
            lblEquipmentOP(0).Caption = FormatNumber(.Fields("col1_equip_cost_op").Value, 2)
            lblEquipmentOP(0).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col1_equip_cost_op").Value) * Trim(lblArea(0).Caption)), 2)
            
            lblInstallOP(0).Caption = FormatNumber(.Fields("col1_inst_cost_op").Value, 2)
            lblInstallOP(0).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col1_inst_cost_op").Value) * Trim(lblArea(0).Caption)), 2)
            
            lblTotalOP(0).Caption = FormatNumber(.Fields("col1_total_cost_op").Value, 2)
            lblTotalOP(0).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col1_total_cost_op").Value) * Trim(lblArea(0).Caption)), 2)
          
            lblMaterialOP(1).Caption = FormatNumber(.Fields("col2_mat_cost_op").Value, 2)
            lblMaterialOP(1).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col2_mat_cost_op").Value) * Trim(lblArea(1).Caption)), 2)
            
            lblLaborOP(1).Caption = FormatNumber(.Fields("col2_labor_cost_op").Value, 2)
            lblLaborOP(1).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col2_labor_cost_op").Value) * Trim(lblArea(1).Caption)), 2)
            
            lblEquipmentOP(1).Caption = FormatNumber(.Fields("col2_equip_cost_op").Value, 2)
            lblEquipmentOP(1).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col2_equip_cost_op").Value) * Trim(lblArea(1).Caption)), 2)
            
            lblInstallOP(1).Caption = FormatNumber(.Fields("col2_inst_cost_op").Value, 2)
            lblInstallOP(1).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col2_inst_cost_op").Value) * Trim(lblArea(1).Caption)), 2)
            
            lblTotalOP(1).Caption = FormatNumber(.Fields("col2_total_cost_op").Value, 2)
            lblTotalOP(1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col2_total_cost_op").Value) * Trim(lblArea(1).Caption)), 2)

            lblMaterialOP(2).Caption = FormatNumber(.Fields("col3_mat_cost_op").Value, 2)
            lblMaterialOP(2).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col3_mat_cost_op").Value) * Trim(lblArea(2).Caption)), 2)
            
            lblLaborOP(2).Caption = FormatNumber(.Fields("col3_labor_cost_op").Value, 2)
            lblLaborOP(2).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col3_labor_cost_op").Value) * Trim(lblArea(2).Caption)), 2)
            
            lblEquipmentOP(2).Caption = FormatNumber(.Fields("col3_equip_cost_op").Value, 2)
            lblEquipmentOP(2).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col3_equip_cost_op").Value) * Trim(lblArea(2).Caption)), 2)
            
            lblInstallOP(2).Caption = FormatNumber(.Fields("col3_inst_cost_op").Value, 2)
            lblInstallOP(2).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col3_inst_cost_op").Value) * Trim(lblArea(2).Caption)), 2)
            
            lblTotalOP(2).Caption = FormatNumber(.Fields("col3_total_cost_op").Value, 2)
            lblTotalOP(2).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col3_total_cost_op").Value) * Trim(lblArea(2).Caption)), 2)

            lblMaterialOP(3).Caption = FormatNumber(.Fields("col4_mat_cost_op").Value, 2)
            lblMaterialOP(3).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col4_mat_cost_op").Value) * Trim(lblArea(3).Caption)), 2)
            
            lblLaborOP(3).Caption = FormatNumber(.Fields("col4_labor_cost_op").Value, 2)
            lblLaborOP(3).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col4_labor_cost_op").Value) * Trim(lblArea(3).Caption)), 2)
            
            lblEquipmentOP(3).Caption = FormatNumber(.Fields("col4_equip_cost_op").Value, 2)
            lblEquipmentOP(3).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col4_equip_cost_op").Value) * Trim(lblArea(3).Caption)), 2)
            
            lblInstallOP(3).Caption = FormatNumber(.Fields("col4_inst_cost_op").Value, 2)
            lblInstallOP(3).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col4_inst_cost_op").Value) * Trim(lblArea(3).Caption)), 2)
            
            lblTotalOP(3).Caption = FormatNumber(.Fields("col4_total_cost_op").Value, 2)
            lblTotalOP(3).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col4_total_cost_op").Value) * Trim(lblArea(3).Caption)), 2)

            lblMaterialOP(4).Caption = FormatNumber(.Fields("col5_mat_cost_op").Value, 2)
            lblMaterialOP(4).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col5_mat_cost_op").Value) * Trim(lblArea(4).Caption)), 2)
            
            lblLaborOP(4).Caption = FormatNumber(.Fields("col5_labor_cost_op").Value, 2)
            lblLaborOP(4).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col5_labor_cost_op").Value) * Trim(lblArea(4).Caption)), 2)
            
            lblEquipmentOP(4).Caption = FormatNumber(.Fields("col5_equip_cost_op").Value, 2)
            lblEquipmentOP(4).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col5_equip_cost_op").Value) * Trim(lblArea(4).Caption)), 2)
            
            lblInstallOP(4).Caption = FormatNumber(.Fields("col5_inst_cost_op").Value, 2)
            lblInstallOP(4).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col5_inst_cost_op").Value) * Trim(lblArea(4).Caption)), 2)
            
            lblTotalOP(4).Caption = FormatNumber(.Fields("col5_total_cost_op").Value, 2)
            lblTotalOP(4).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col5_total_cost_op").Value) * Trim(lblArea(4).Caption)), 2)

            lblMaterialOP(5).Caption = FormatNumber(.Fields("col6_mat_cost_op").Value, 2)
            lblMaterialOP(5).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col6_mat_cost_op").Value) * Trim(lblArea(5).Caption)), 2)
            
            lblLaborOP(5).Caption = FormatNumber(.Fields("col6_labor_cost_op").Value, 2)
            lblLaborOP(5).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col6_labor_cost_op").Value) * Trim(lblArea(5).Caption)), 2)
            
            lblEquipmentOP(5).Caption = FormatNumber(.Fields("col6_equip_cost_op").Value, 2)
            lblEquipmentOP(5).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col6_equip_cost_op").Value) * Trim(lblArea(5).Caption)), 2)
            
            lblInstallOP(5).Caption = FormatNumber(.Fields("col6_inst_cost_op").Value, 2)
            lblInstallOP(5).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col6_inst_cost_op").Value) * Trim(lblArea(5).Caption)), 2)
            
            lblTotalOP(5).Caption = FormatNumber(.Fields("col6_total_cost_op").Value, 2)
            lblTotalOP(5).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col6_total_cost_op").Value) * Trim(lblArea(5).Caption)), 2)

            lblMaterialOP(6).Caption = FormatNumber(.Fields("col7_mat_cost_op").Value, 2)
            lblMaterialOP(6).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col7_mat_cost_op").Value) * Trim(lblArea(6).Caption)), 2)
            
            lblLaborOP(6).Caption = FormatNumber(.Fields("col7_labor_cost_op").Value, 2)
            lblLaborOP(6).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col7_labor_cost_op").Value) * Trim(lblArea(6).Caption)), 2)
            
            lblEquipmentOP(6).Caption = FormatNumber(.Fields("col7_equip_cost_op").Value, 2)
            lblEquipmentOP(6).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col7_equip_cost_op").Value) * Trim(lblArea(6).Caption)), 2)
            
            lblInstallOP(6).Caption = FormatNumber(.Fields("col7_inst_cost_op").Value, 2)
            lblInstallOP(6).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col7_inst_cost_op").Value) * Trim(lblArea(6).Caption)), 2)
            
            lblTotalOP(6).Caption = FormatNumber(.Fields("col7_total_cost_op").Value, 2)
            lblTotalOP(6).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col7_total_cost_op").Value) * Trim(lblArea(6).Caption)), 2)

            lblMaterialOP(7).Caption = FormatNumber(.Fields("col8_mat_cost_op").Value, 2)
            lblMaterialOP(7).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col8_mat_cost_op").Value) * Trim(lblArea(7).Caption)), 2)
            
            lblLaborOP(7).Caption = FormatNumber(.Fields("col8_labor_cost_op").Value, 2)
            lblLaborOP(7).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col8_labor_cost_op").Value) * Trim(lblArea(7).Caption)), 2)
            
            lblEquipmentOP(7).Caption = FormatNumber(.Fields("col8_equip_cost_op").Value, 2)
            lblEquipmentOP(7).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col8_equip_cost_op").Value) * Trim(lblArea(7).Caption)), 2)
            
            lblInstallOP(7).Caption = FormatNumber(.Fields("col8_inst_cost_op").Value, 2)
            lblInstallOP(7).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col8_inst_cost_op").Value) * Trim(lblArea(7).Caption)), 2)
            
            lblTotalOP(7).Caption = FormatNumber(.Fields("col8_total_cost_op").Value, 2)
            lblTotalOP(7).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col8_total_cost_op").Value) * Trim(lblArea(7).Caption)), 2)

            lblMaterialOP(8).Caption = FormatNumber(.Fields("col9_mat_cost_op").Value, 2)
            lblMaterialOP(8).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col9_mat_cost_op").Value) * Trim(lblArea(8).Caption)), 2)
            
            lblLaborOP(8).Caption = FormatNumber(.Fields("col9_labor_cost_op").Value, 2)
            lblLaborOP(8).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col9_labor_cost_op").Value) * Trim(lblArea(8).Caption)), 2)
            
            lblEquipmentOP(8).Caption = FormatNumber(.Fields("col9_equip_cost_op").Value, 2)
            lblEquipmentOP(8).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col9_equip_cost_op").Value) * Trim(lblArea(8).Caption)), 2)
            
            lblInstallOP(8).Caption = FormatNumber(.Fields("col9_inst_cost_op").Value, 2)
            lblInstallOP(8).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col9_inst_cost_op").Value) * Trim(lblArea(8).Caption)), 2)
            
            lblTotalOP(8).Caption = FormatNumber(.Fields("col9_total_cost_op").Value, 2)
            lblTotalOP(8).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col9_total_cost_op").Value) * Trim(lblArea(8).Caption)), 2)

            If Trim(.Fields("op_code").Value) = "STD" Then
                optUnion.Value = True
            Else
                optOpen.Value = True
            End If
            
            PopulateComboCountryRegion Trim(.Fields("country_code").Value), _
                            Trim(.Fields("region_code").Value)
            '
            '   If the model we're on is the one to bold then get the
            '   area col.
            If .Fields("model_code").Value = .Fields("model_code_to_bold").Value Then
                nAreaInd = .Fields("areaind").Value
                ChangeOpCostBackcolor nAreaInd, False
            Else
                nAreaInd = 0
                ChangeOpCostBackcolor nAreaInd, True
            End If
            
            ' Cache the standard model row/col so we can read the standard model's assembly components
            ' if the user tries to clone it.
            m_iStandardModelRow = .Fields("Model_Code_To_Bold").Value
            m_iStandardModelCol = .Fields("AreaInd").Value
            
            '
            '   Set the location of the shpSelectedAreaPerimeter for highlighting.
            '   sAreaToPopulate is in the format of 1,1 meaning row 1 col area1.
            If Trim(sAreaToPopulate) = "" Then
                '
                '   If the sshpSelectedArea has never been set default to the
                '   1st area.
                If sshpSelectedArea = "" Then
                    lblArea_Click IIf((nAreaInd = 0), 0, (nAreaInd - 1))
                Else
                    lblArea_Click Right$(sshpSelectedArea, 1)
                End If
            Else
                lblArea_Click Right$(sAreaToPopulate, 1)
            End If
        End If
    End With
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateModelMatrixResi(sAreaToPopulate As String)
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With m_recModelMatrix
        If Not .EOF Then
            .MoveFirst
        
            lblAreaResi(0).Caption = Trim(.Fields("area1").Value)
            lblPerimeterResi(0).Caption = Trim(.Fields("perimeter1").Value)
            lblAreaResi(1).Caption = Trim(.Fields("area2").Value)
            lblPerimeterResi(1).Caption = Trim(.Fields("perimeter2").Value)
            lblAreaResi(2).Caption = Trim(.Fields("area3").Value)
            lblPerimeterResi(2).Caption = Trim(.Fields("perimeter3").Value)
            lblAreaResi(3).Caption = Trim(.Fields("area4").Value)
            lblPerimeterResi(3).Caption = Trim(.Fields("perimeter4").Value)
            lblAreaResi(4).Caption = Trim(.Fields("area5").Value)
            lblPerimeterResi(4).Caption = Trim(.Fields("perimeter5").Value)
            lblAreaResi(5).Caption = Trim(.Fields("area6").Value)
            lblPerimeterResi(5).Caption = Trim(.Fields("perimeter6").Value)
            lblAreaResi(6).Caption = Trim(.Fields("area7").Value)
            lblPerimeterResi(6).Caption = Trim(.Fields("perimeter7").Value)
            lblAreaResi(7).Caption = Trim(.Fields("area8").Value)
            lblPerimeterResi(7).Caption = Trim(.Fields("perimeter8").Value)
            '
            '   Only 8 Areas for Wings & Ells
            If Trim(m_rec.Fields("bldg_type").Value) <> "H" And Trim(m_rec.Fields("bldg_type").Value) <> "I" And Trim(m_rec.Fields("bldg_type").Value) <> "J" Then
                lblAreaResi(8).Caption = Trim(.Fields("area9").Value)
                lblPerimeterResi(8).Caption = Trim(.Fields("perimeter9").Value)
                lblAreaResi(9).Caption = Trim(.Fields("area10").Value)
                lblPerimeterResi(9).Caption = Trim(.Fields("perimeter10").Value)
                lblAreaResi(10).Caption = Trim(.Fields("area11").Value)
                lblPerimeterResi(10).Caption = Trim(.Fields("perimeter11").Value)
            End If
            
            lblMaterialOPResi(0).Caption = FormatNumber(.Fields("col1_mat_cost_op").Value, 2)
            lblMaterialOPResi(0).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col1_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblLaborOPResi(0).Caption = FormatNumber(.Fields("col1_labor_cost_op").Value, 2)
            lblLaborOPResi(0).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col1_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblEquipmentOPResi(0).Caption = FormatNumber(.Fields("col1_equip_cost_op").Value, 2)
            lblEquipmentOPResi(0).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col1_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblInstallOPResi(0).Caption = FormatNumber(.Fields("col1_inst_cost_op").Value, 2)
            lblInstallOPResi(0).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col1_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblTotalOPResi(0).Caption = FormatNumber(.Fields("col1_total_cost_op").Value, 2)
            lblTotalOPResi(0).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col1_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
          
            lblMaterialOPResi(1).Caption = FormatNumber(.Fields("col2_mat_cost_op").Value, 2)
            lblMaterialOPResi(1).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col2_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblLaborOPResi(1).Caption = FormatNumber(.Fields("col2_labor_cost_op").Value, 2)
            lblLaborOPResi(1).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col2_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblEquipmentOPResi(1).Caption = FormatNumber(.Fields("col2_equip_cost_op").Value, 2)
            lblEquipmentOPResi(1).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col2_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblInstallOPResi(1).Caption = FormatNumber(.Fields("col2_inst_cost_op").Value, 2)
            lblInstallOPResi(1).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col2_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblTotalOPResi(1).Caption = FormatNumber(.Fields("col2_total_cost_op").Value, 2)
            lblTotalOPResi(1).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col2_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)

            lblMaterialOPResi(2).Caption = FormatNumber(.Fields("col3_mat_cost_op").Value, 2)
            lblMaterialOPResi(2).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col3_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblLaborOPResi(2).Caption = FormatNumber(.Fields("col3_labor_cost_op").Value, 2)
            lblLaborOPResi(2).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col3_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblEquipmentOPResi(2).Caption = FormatNumber(.Fields("col3_equip_cost_op").Value, 2)
            lblEquipmentOPResi(2).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col3_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblInstallOPResi(2).Caption = FormatNumber(.Fields("col3_inst_cost_op").Value, 2)
            lblInstallOPResi(2).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col3_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblTotalOPResi(2).Caption = FormatNumber(.Fields("col3_total_cost_op").Value, 2)
            lblTotalOPResi(2).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col3_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)

            lblMaterialOPResi(3).Caption = FormatNumber(.Fields("col4_mat_cost_op").Value, 2)
            lblMaterialOPResi(3).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col4_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblLaborOPResi(3).Caption = FormatNumber(.Fields("col4_labor_cost_op").Value, 2)
            lblLaborOPResi(3).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col4_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblEquipmentOPResi(3).Caption = FormatNumber(.Fields("col4_equip_cost_op").Value, 2)
            lblEquipmentOPResi(3).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col4_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblInstallOPResi(3).Caption = FormatNumber(.Fields("col4_inst_cost_op").Value, 2)
            lblInstallOPResi(3).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col4_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblTotalOPResi(3).Caption = FormatNumber(.Fields("col4_total_cost_op").Value, 2)
            lblTotalOPResi(3).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col4_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)

            lblMaterialOPResi(4).Caption = FormatNumber(.Fields("col5_mat_cost_op").Value, 2)
            lblMaterialOPResi(4).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col5_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblLaborOPResi(4).Caption = FormatNumber(.Fields("col5_labor_cost_op").Value, 2)
            lblLaborOPResi(4).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col5_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblEquipmentOPResi(4).Caption = FormatNumber(.Fields("col5_equip_cost_op").Value, 2)
            lblEquipmentOPResi(4).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col5_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblInstallOPResi(4).Caption = FormatNumber(.Fields("col5_inst_cost_op").Value, 2)
            lblInstallOPResi(4).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col5_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblTotalOPResi(4).Caption = FormatNumber(.Fields("col5_total_cost_op").Value, 2)
            lblTotalOPResi(4).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col5_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)

            lblMaterialOPResi(5).Caption = FormatNumber(.Fields("col6_mat_cost_op").Value, 2)
            lblMaterialOPResi(5).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col6_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblLaborOPResi(5).Caption = FormatNumber(.Fields("col6_labor_cost_op").Value, 2)
            lblLaborOPResi(5).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col6_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblEquipmentOPResi(5).Caption = FormatNumber(.Fields("col6_equip_cost_op").Value, 2)
            lblEquipmentOPResi(5).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col6_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblInstallOPResi(5).Caption = FormatNumber(.Fields("col6_inst_cost_op").Value, 2)
            lblInstallOPResi(5).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col6_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblTotalOPResi(5).Caption = FormatNumber(.Fields("col6_total_cost_op").Value, 2)
            lblTotalOPResi(5).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col6_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)

            lblMaterialOPResi(6).Caption = FormatNumber(.Fields("col7_mat_cost_op").Value, 2)
            lblMaterialOPResi(6).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col7_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblLaborOPResi(6).Caption = FormatNumber(.Fields("col7_labor_cost_op").Value, 2)
            lblLaborOPResi(6).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col7_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblEquipmentOPResi(6).Caption = FormatNumber(.Fields("col7_equip_cost_op").Value, 2)
            lblEquipmentOPResi(6).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col7_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblInstallOPResi(6).Caption = FormatNumber(.Fields("col7_inst_cost_op").Value, 2)
            lblInstallOPResi(6).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col7_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblTotalOPResi(6).Caption = FormatNumber(.Fields("col7_total_cost_op").Value, 2)
            lblTotalOPResi(6).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col7_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)

            lblMaterialOPResi(7).Caption = FormatNumber(.Fields("col8_mat_cost_op").Value, 2)
            lblMaterialOPResi(7).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col8_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblLaborOPResi(7).Caption = FormatNumber(.Fields("col8_labor_cost_op").Value, 2)
            lblLaborOPResi(7).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col8_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblEquipmentOPResi(7).Caption = FormatNumber(.Fields("col8_equip_cost_op").Value, 2)
            lblEquipmentOPResi(7).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col8_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblInstallOPResi(7).Caption = FormatNumber(.Fields("col8_inst_cost_op").Value, 2)
            lblInstallOPResi(7).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col8_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            
            lblTotalOPResi(7).Caption = FormatNumber(.Fields("col8_total_cost_op").Value, 2)
            lblTotalOPResi(7).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col8_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            '
            '   Only 8 Areas for Wings & Ells
            If Trim(m_rec.Fields("bldg_type").Value) <> "H" And Trim(m_rec.Fields("bldg_type").Value) <> "I" And Trim(m_rec.Fields("bldg_type").Value) <> "J" Then

                lblMaterialOPResi(8).Caption = FormatNumber(.Fields("col9_mat_cost_op").Value, 2)
                lblMaterialOPResi(8).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col9_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblLaborOPResi(8).Caption = FormatNumber(.Fields("col9_labor_cost_op").Value, 2)
                lblLaborOPResi(8).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col9_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblEquipmentOPResi(8).Caption = FormatNumber(.Fields("col9_equip_cost_op").Value, 2)
                lblEquipmentOPResi(8).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col9_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblInstallOPResi(8).Caption = FormatNumber(.Fields("col9_inst_cost_op").Value, 2)
                lblInstallOPResi(8).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col9_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblTotalOPResi(8).Caption = FormatNumber(.Fields("col9_total_cost_op").Value, 2)
                lblTotalOPResi(8).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col9_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblMaterialOPResi(9).Caption = FormatNumber(.Fields("col10_mat_cost_op").Value, 2)
                lblMaterialOPResi(9).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col10_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblLaborOPResi(9).Caption = FormatNumber(.Fields("col10_labor_cost_op").Value, 2)
                lblLaborOPResi(9).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col10_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblEquipmentOPResi(9).Caption = FormatNumber(.Fields("col10_equip_cost_op").Value, 2)
                lblEquipmentOPResi(9).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col10_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblInstallOPResi(9).Caption = FormatNumber(.Fields("col10_inst_cost_op").Value, 2)
                lblInstallOPResi(9).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col10_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblTotalOPResi(9).Caption = FormatNumber(.Fields("col10_total_cost_op").Value, 2)
                lblTotalOPResi(9).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col10_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                            
                lblMaterialOPResi(10).Caption = FormatNumber(.Fields("col11_mat_cost_op").Value, 2)
                lblMaterialOPResi(10).ToolTipText = "Total Material Cost: " & FormatNumber((Trim(.Fields("col11_mat_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblLaborOPResi(10).Caption = FormatNumber(.Fields("col11_labor_cost_op").Value, 2)
                lblLaborOPResi(10).ToolTipText = "Total Labor Cost: " & FormatNumber((Trim(.Fields("col11_labor_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblEquipmentOPResi(10).Caption = FormatNumber(.Fields("col11_equip_cost_op").Value, 2)
                lblEquipmentOPResi(10).ToolTipText = "Total Equipment Cost: " & FormatNumber((Trim(.Fields("col11_equip_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblInstallOPResi(10).Caption = FormatNumber(.Fields("col11_inst_cost_op").Value, 2)
                lblInstallOPResi(10).ToolTipText = "Total Installation Cost: " & FormatNumber((Trim(.Fields("col11_inst_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
                
                lblTotalOPResi(10).Caption = FormatNumber(.Fields("col11_total_cost_op").Value, 2)
                lblTotalOPResi(10).ToolTipText = "Total Cost: " & FormatNumber((Trim(.Fields("col11_total_cost_op").Value) * Trim(lblAreaResi(0).Caption)), 2)
            End If
            
            If Trim(.Fields("op_code").Value) = "STD" Then
                optUnion.Value = True
            Else
                optOpen.Value = True
            End If
            
            PopulateComboCountryRegion Trim(.Fields("country_code").Value), _
                            Trim(.Fields("region_code").Value)
            '
            '   If the model we're on is the one to bold then get the
            '   area col.
            If .Fields("model_code").Value = .Fields("model_code_to_bold").Value Then
                nAreaInd = .Fields("areaind").Value
                ChangeOpCostBackcolorResi nAreaInd, False
            Else
                nAreaInd = 0
                ChangeOpCostBackcolorResi nAreaInd, True
            End If
            
            ' Cache the standard model row/col so we can read the standard model's assembly components
            ' if the user tries to clone it.
            m_iStandardModelRow = .Fields("Model_Code_To_Bold").Value
            m_iStandardModelCol = .Fields("AreaInd").Value
            
            '
            '   Set the location of the shpSelectedAreaPerimeter for highlighting.
            '   sAreaToPopulate is in the format of 1,1 meaning row 1 col area1.
            If Trim(sAreaToPopulate) = "" Then
                '
                '   If the sshpSelectedArea has never been set default to the
                '   1st area.
                If sshpSelectedArea = "" Then
                    lblAreaResi_Click IIf((nAreaInd = 0), 0, (nAreaInd - 1))
                Else
                    lblAreaResi_Click Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))
                End If
            Else
                lblAreaResi_Click Right$(sAreaToPopulate, Len(sAreaToPopulate) - InStr(1, sAreaToPopulate, ","))
            End If
        End If
    End With
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateComboCountryRegion(sCountryCodeValueToSelect As String, sRegionCodeValueToSelect As String)
    Dim i As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    For i = 0 To cboMdlCountryCode.listcount - 1
        If cboMdlCountryCode.List(i) = sCountryCodeValueToSelect Then
            cboMdlCountryCode.ListIndex = i
            Exit For
        End If
    Next i
    
    For i = 0 To cboMdlRegionCode.listcount - 1
        If cboMdlRegionCode.List(i) = sRegionCodeValueToSelect Then
            cboMdlRegionCode.ListIndex = i
            Exit For
        End If
    Next i
    Screen.MousePointer = vbNormal
End Sub

Private Sub ChangeOpCostBackcolor(nAreaInd As Integer, bOnlySetToDefaultColor As Boolean)
    Dim ctrl As Control

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    For Each ctrl In Me.Controls
        '
        '   If backcolor is teal = &HFFFF00
        If ctrl.BackColor = &HFFFF00 Then
            If TypeOf ctrl Is TextBox Or TypeOf ctrl Is Label Then
                ctrl.BackColor = &H80000005 'white
            End If
        End If
    Next ctrl
    If Not bOnlySetToDefaultColor Then
        lblArea(nAreaInd - 1).BackColor = &HFFFF00
        lblPerimeter(nAreaInd - 1).BackColor = &HFFFF00
        lblMaterialOP(nAreaInd - 1).BackColor = &HFFFF00
        lblLaborOP(nAreaInd - 1).BackColor = &HFFFF00
        lblEquipmentOP(nAreaInd - 1).BackColor = &HFFFF00
        lblInstallOP(nAreaInd - 1).BackColor = &HFFFF00
        lblTotalOP(nAreaInd - 1).BackColor = &HFFFF00
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub ChangeOpCostBackcolorResi(nAreaInd As Integer, bOnlySetToDefaultColor As Boolean)
    Dim ctrl As Control

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    For Each ctrl In Me.Controls
        '
        '   If backcolor is teal = &HFFFF00
        If ctrl.BackColor = &HFFFF00 Then
            If TypeOf ctrl Is TextBox Or TypeOf ctrl Is Label Then
                ctrl.BackColor = &H80000005 'white
            End If
        End If
    Next ctrl
    '
    '   Only bother setting the backcolor to teal if this model
    '   has the area to bold.
    If Not bOnlySetToDefaultColor Then
        lblAreaResi(nAreaInd - 1).BackColor = &HFFFF00
        lblPerimeterResi(nAreaInd - 1).BackColor = &HFFFF00
        lblMaterialOPResi(nAreaInd - 1).BackColor = &HFFFF00
        lblLaborOPResi(nAreaInd - 1).BackColor = &HFFFF00
        lblEquipmentOPResi(nAreaInd - 1).BackColor = &HFFFF00
        lblInstallOPResi(nAreaInd - 1).BackColor = &HFFFF00
        lblTotalOPResi(nAreaInd - 1).BackColor = &HFFFF00
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateSummaryEstimateComponents()
    Dim strSelect               As String

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    m_rec.MoveFirst
    '
    '   Make sure it is closed.
    With m_recModelComponent
        .Close
        '
        '   Set the maximum number to bring back.
        .MaxRecords = MAX_RECORDS
    End With
    '
    '   Was having a problem with the grid refreshing if
    '   the user had changed data and not save and switched to
    '   a different area it would not refresh.  So have to close
    '   and reopen.
    With TDBGridSummaryEstimate
        .Close
        .ReOpen
    End With
    
    'UPDATED 06/29/2005 RTD - VERSION 7.4.0
    'CORRECTS PROBLEM WHERE MODEL MAINTENANCE IS VERY SLOW TO OPEN
    If Len(Trim(m_rec.Fields("bldg_model_skey").Value)) > 0 And InStr(m_rec.Fields("bldg_model_skey").Value, "*") = 0 Then
        'NEW STORED PROCEDURE FOR EXACT BLDG_MODEL_SKEY - FAST
        strSelect = "exec usp_select_building_pub_component_cost_ex @bldg_model_skey = '"
        strSelect = strSelect & Trim(m_rec.Fields("bldg_model_skey").Value)
    Else
        'OLD STORED PROCEDURE FOR BLDG_MODEL_SKEY WITH WILDCARD - SLOW
        strSelect = "exec sp_select_building_pub_component_cost @bldg_model_skey = '"
        If Trim(m_rec.Fields("bldg_model_skey").Value) = "" Then
            strSelect = strSelect & "%"
        Else
            strSelect = strSelect & SQLChangeWildcard(Trim(m_rec.Fields("bldg_model_skey").Value))
        End If
    End If
    
     strSelect = strSelect & "', @bldg_area = '"
    If opttype_codeC.Value = True Then
        If Len(Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption)) > 0 Then
             strSelect = strSelect & Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption)
        Else
            'abort
            Exit Sub
        End If
    Else
        If Len(Trim(lblAreaResi(Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))).Caption)) > 0 Then
             strSelect = strSelect & Trim(lblAreaResi(Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))).Caption)
        Else
            'abort
            Exit Sub
        End If
    End If
     
    strSelect = strSelect & "', @op_code = '"
    If optUnion.Value = True Then
        strSelect = strSelect & "STD'"
    Else
        strSelect = strSelect & "OPN'"
    End If
    
    strSelect = strSelect & ", @country_code = '"
    If Len(Trim(cboMdlCountryCode.Text)) > 0 Then
        strSelect = strSelect & cboMdlCountryCode.Text & "'"
    Else
        strSelect = strSelect & "USA'"
    End If
    
    strSelect = strSelect & ", @region_code = '"
    If Len(Trim(cboMdlRegionCode.Text)) > 0 Then
        strSelect = strSelect & cboMdlRegionCode.Text & "'"
    Else
        strSelect = strSelect & "NAT'"
    End If
    '
    '   Use DAL to perform select.
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, m_recModelComponent) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while retreiving Summary Estimate Components."
        DoEvents
    Else
        '
        '   Pass recordset to handler class.
        m_objModelComponentGridMap.RecordSet = m_recModelComponent
        '
        '   Need to make sure that the user cannot set
        '   max_records = 0
        With m_recModelComponent
            If .RecordCount > 0 Then
                '
                ' If the upper bound was hit, inform user.
                If .RecordCount = MAX_RECORDS And .State = adStateOpen Then
                    MsgBox "The search returned the maximum number of records allowed. More records may be available."
                End If
            End If
         End With
         DoEvents
         '
         '   Reset the grid contents
         With TDBGridSummaryEstimate
             .Bookmark = Null
             .ReBind
             .ApproxCount = m_recModelComponent.RecordCount
         End With
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub PreviewAssemblyComponentsReport()
    Dim fPreviewWindow As New frmReportPreview
    Dim strSelect As String
    Dim recTemp As ADODB.RecordSet
    Dim sBldgDesc As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status "Generating Assembly Components report..."
    '   If we're on a format row 7 or 8 for Commercial we have to display the assemblies for model_code 1
    '   since 7 & 8 do not have rows in assembly_usage.  Per Tom 4/26/02.
    If m_rec.Fields("format_row").Value = "1" And opttype_codeC.Value = True Then
        strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & m_rec.Fields("bldg_skey").Value & "' AND model_code = '1'"
        If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Or recTemp.RecordCount = 0 Then
            Exit Sub
        Else
            'rlh - updated to fix Missing Component Report reported by BB
            strSelect = "exec usp_rpt_model_assembly_units @bldg_model_skey = '" & recTemp.Fields("bldg_model_skey").Value
            recTemp.Close
        End If
    ElseIf m_rec.Fields("format_row").Value = "1" And opttype_codeR.Value = True Then
        'rlh - updated to fix Missing Component Report reported by BB
        strSelect = "exec usp_rpt_model_assembly_units @bldg_model_skey = '"
        If Len(Trim(m_rec.Fields("bldg_model_skey").Value)) > 0 Then
            strSelect = strSelect & Trim(m_rec.Fields("bldg_model_skey").Value)
        Else
            Exit Sub
        End If
    Else
        'rlh - updated to fix Missing Component Report reported by BB
        strSelect = "exec usp_rpt_model_assembly_units @bldg_model_skey = '"
        If Len(Trim(m_rec.Fields("bldg_model_skey").Value)) > 0 Then
            strSelect = strSelect & Trim(m_rec.Fields("bldg_model_skey").Value)
        Else
            Exit Sub
        End If
    End If
    
    strSelect = strSelect & "', @bldg_area = '"
    If opttype_codeC.Value = True Then
        If Len(Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption)) > 0 Then
             strSelect = strSelect & Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption)
        Else
            'abort
            Exit Sub
        End If
    Else
        If Len(Trim(lblAreaResi(Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))).Caption)) > 0 Then
             strSelect = strSelect & Trim(lblAreaResi(Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))).Caption)
        Else
            'abort
            Exit Sub
        End If
    End If
    strSelect = strSelect & "', @op_code = '"
    If optUnion.Value = True Then
        strSelect = strSelect & "STD'"
    Else
        strSelect = strSelect & "OPN'"
    End If
    
    strSelect = strSelect & ", @country_code = '"
    If Len(Trim(cboMdlCountryCode.Text)) > 0 Then
        strSelect = strSelect & cboMdlCountryCode.Text & "'"
    Else
        strSelect = strSelect & "USA'"
    End If
     
    strSelect = strSelect & ", @region_code = '"
    If Len(Trim(cboMdlRegionCode.Text)) > 0 Then
        strSelect = strSelect & cboMdlRegionCode.Text & "'"
    Else
        strSelect = strSelect & "NAT'"
    End If
    
    sBldgDesc = Trim(txtbldg_desc.Text) & " [Bldg ID: " & Trim(txtbldg_id.Text) & " | Model Code: " & Trim(m_rec.Fields("model_code").Value) & " | " & Trim(cboWallType.Text) & "]"
    
    fPreviewWindow.ReportName = "Assembly Components"
    fPreviewWindow.ReportFile = "rptSummaryEstimate.xml"
    fPreviewWindow.RecordSource = strSelect
    fPreviewWindow.OpenEvent = "Building_Description = """ & sBldgDesc & """"
    fPreviewWindow.RenderReport
    fPreviewWindow.Show
    Screen.MousePointer = vbDefault
    Status ""

End Sub

' Returns the building model skey from the result set, or -1 if it is empty.
Private Function GetBldgModelSKey(strBldgModelSKey As String) As Long

    If (Len(Trim(strBldgModelSKey)) > 0) Then
        GetBldgModelSKey = Trim(strBldgModelSKey)
    Else
        GetBldgModelSKey = -1
    End If
    
End Function

' Returns the building area (square footage) based on the column value of the selected area variable, or -1 if it is empty.
Private Function GetBuildingArea() As Long
    
    ' Commercial
    If opttype_codeC.Value = True Then
        If Len(Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption)) > 0 Then
            GetBuildingArea = Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption)
        Else
            GetBuildingArea = -1
        End If
    Else
        ' Residential.  Note that for residential, we may have up to 10 columns which is why the Right$() call is different here.
        If Len(Trim(lblAreaResi(Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))).Caption)) > 0 Then
            GetBuildingArea = Trim(lblAreaResi(Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))).Caption)
        Else
            GetBuildingArea = -1
        End If
    End If
End Function

' This routine will populate the Assembly Components grid on the form.  It is called by the
' regular form load code, and also when the user clones the standard model.
' This routine will clear out both the grid and the recordset.
Private Function PopulateAssemblyComponents(lBldgModelSKey As Long, lBldgArea As Long, lOrigBldgModelSKey As Long, strErrMsg As String) As Boolean
    Dim strSelect       As String
    Dim recTemp         As New ADODB.RecordSet
    
    PopulateAssemblyComponents = False
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   Make sure it is closed.
    With m_recAssembly
        .Close
        '
        '   Set the maximum number to bring back.
        .MaxRecords = MAX_RECORDS
    End With
    '
    '   Was having a problem with the grid refreshing if
    '   the user had changed data and not save and switched to
    '   a different area it would not refresh.  So have to close
    '   and reopen.
    With TDBGridAssembly
        .Close
        .ReOpen
    End With
    '
    '   If we're on a Quality Series building we have to get the assemblies directly
    '   from assembly_usage not published_bldg_model_assembly_cost.
    If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
    Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
        strSelect = "exec sp_select_model_assembly_usage_basements @parent_skey = '"
        If lBldgModelSKey >= 0 Then
            strSelect = strSelect & lBldgModelSKey
        Else
            Exit Function
        End If
        strSelect = strSelect & "', @op_code = '"
        If optUnion.Value = True Then
            strSelect = strSelect & "STD'"
        Else
            strSelect = strSelect & "OPN'"
        End If
        
        strSelect = strSelect & ", @country_code = '"
        If Len(Trim(cboMdlCountryCode.Text)) > 0 Then
            strSelect = strSelect & cboMdlCountryCode.Text & "'"
        Else
            strSelect = strSelect & "USA'"
        End If
         
        strSelect = strSelect & ", @region_code = '"
        If Len(Trim(cboMdlRegionCode.Text)) > 0 Then
            strSelect = strSelect & cboMdlRegionCode.Text & "'"
        Else
            strSelect = strSelect & "NAT'"
        End If
    Else
        '
        '   If we're on a format row 7 or 8 for Commercial we have to display the assemblies for model_code 1
        '   since 7 & 8 do not have rows in assembly_usage.  Per Tom 4/26/02.
        If m_rec.Fields("format_row").Value = "1" And opttype_codeC.Value = True Then
            strSelect = "SELECT bldg_model_skey FROM bldg_model WHERE bldg_skey = '" & m_rec.Fields("bldg_skey").Value & "' AND model_code = '1'"
            
            If Not g_objDAL.GetRecordset(vbNullString, strSelect, recTemp) Or recTemp.RecordCount = 0 Then
                Exit Function
            Else
                strSelect = "exec sp_select_published_bldg_model_assembly_cost @bldg_model_skey = '" & recTemp.Fields("bldg_model_skey").Value
                recTemp.Close
            End If
        ElseIf m_rec.Fields("format_row").Value = "1" And opttype_codeR.Value = True Then
            strSelect = "exec sp_select_published_bldg_model_assembly_cost_basements @bldg_model_skey = '"
            If lBldgModelSKey >= 0 Then
                strSelect = strSelect & lBldgModelSKey
            Else
                Exit Function
            End If
        Else
            strSelect = "exec sp_select_published_bldg_model_assembly_cost @bldg_model_skey = '"
            If lBldgModelSKey >= 0 Then
                strSelect = strSelect & lBldgModelSKey
            Else
                Exit Function
            End If
        End If
        
        ' Note:  commercial vs. resi was handled already in GetBuildingArea().
        strSelect = strSelect & "', @bldg_area = '"
        If lBldgArea >= 0 Then
             strSelect = strSelect & lBldgArea
        Else
            'abort
            Exit Function
        End If
                    
        strSelect = strSelect & "', @op_code = '"
        If optUnion.Value = True Then
            strSelect = strSelect & "STD'"
        Else
            strSelect = strSelect & "OPN'"
        End If
        
        strSelect = strSelect & ", @country_code = '"
        If Len(Trim(cboMdlCountryCode.Text)) > 0 Then
            strSelect = strSelect & cboMdlCountryCode.Text & "'"
        Else
            strSelect = strSelect & "USA'"
        End If
         
        strSelect = strSelect & ", @region_code = '"
        If Len(Trim(cboMdlRegionCode.Text)) > 0 Then
            strSelect = strSelect & cboMdlRegionCode.Text & "'"
        Else
            strSelect = strSelect & "NAT'"
        End If
    End If
    With cnTemp
        m_recAssembly.CursorLocation = adUseClient
        m_recAssembly.Open _
            Source:=strSelect, _
            ActiveConnection:=cnTemp, _
            CursorType:=adOpenStatic, _
            LockType:=adLockBatchOptimistic

        If .Errors.Count <> 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "Errors in the PopulateAssemblyComponents routine: " _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
        Else
            If (m_objModelAssembliesGridMap.bCloneAssembliesInProcess = True) Then
                ' If we are called from a clone action, we must overlay the building skey, last update
                ' person, date and id.
                Dim fixupResult As Boolean
                fixupResult = FixupCloneInfo(lOrigBldgModelSKey, strErrMsg)
                If (fixupResult = True) Then
                    PopulateAssemblyComponents = True
                    Exit Function
                End If
            End If
            '
            '   Pass recordset to handler class.
            m_objModelAssembliesGridMap.RecordSet = m_recAssembly
            
            '
            '   Need to make sure that the user cannot set
            '   max_records = 0
            With m_recAssembly
                .MoveFirst
                
                If .RecordCount > 0 Then
                    lblAssemblyCompRowCount.Caption = .RecordCount & " rows returned."
                    '
                    ' If the upper bound was hit, inform user.
                    If .RecordCount = MAX_RECORDS And .State = adStateOpen Then
                        MsgBox "The search returned the maximum number of records allowed. More records may be available."
                    End If
                Else
                    lblAssemblyCompRowCount.Caption = "0 rows returned."
                End If
             End With
             DoEvents
             '
             '   Reset the grid contents
             With TDBGridAssembly
                If m_recAssembly.RecordCount = 0 Then
                    .Bookmark = Null
                Else
                    .Bookmark = 1
                End If
                .ReBind
                .ApproxCount = m_recAssembly.RecordCount
                m_objModelAssembliesGridMap.SetupAssembliesFormulas False
             End With
        End If
    End With
    Screen.MousePointer = vbNormal
End Function

' When we are called from a clone action, we need to overlay several fields in the record set because they
' are currently pointing to the model we cloned from.  We need to change the values to the model we are
' cloning to.
Private Function FixupCloneInfo(lOrigBldgModelSKey As Long, strErrMsg As String) As Boolean

    FixupCloneInfo = False
    
    ' This is necessary since the Reset_Orig_Values() routine always throws an error.
    ' If it is ever fixed though, we want it to be called below like in all of the other places we overlay values.
    On Error Resume Next
    
    If (m_recAssembly.EOF) Then
        strErrMsg = "No assembly components found for the standard model.  Please verify that the standard model " & _
                 "contains at least one assembly component"
        FixupCloneInfo = True
        Exit Function
    End If
    
    m_recAssembly.MoveFirst
    
    Do While Not m_recAssembly.EOF
    
        m_recAssembly.Fields("bldg_model_skey").Value = lOrigBldgModelSKey
        m_recAssembly.Fields("last_update_person").Value = strUserName
        m_recAssembly.Fields("last_update_date").Value = Now
        ' Need to set the id to 0 so the stored proc will treat this as an insert rather than an update.
        m_recAssembly.Fields("last_update_id").Value = 0
        
        ' Now make sure the original value for each column in the recordset is set to what we just
        ' overlayed in the value field.
        Reset_Orig_Values m_recAssembly
    
        m_recAssembly.MoveNext
    
    Loop
    
    ' Reset to the first record in case the caller expects that.
    m_recAssembly.MoveFirst


End Function
Private Sub SetShpTopLocation()

    Dim strErrMsg As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   Get the row we're in.
    shpSelectedAreaPerimeter.Top = 120
    '
    '   Get the column we're in.
    Select Case Right$(sshpSelectedArea, 1)
        Case 0
            shpSelectedAreaPerimeter.Left = 1330
        Case 1
            shpSelectedAreaPerimeter.Left = 2260
        Case 2
            shpSelectedAreaPerimeter.Left = 3190
        Case 3
            shpSelectedAreaPerimeter.Left = 4120
        Case 4
            shpSelectedAreaPerimeter.Left = 5050
        Case 5
            shpSelectedAreaPerimeter.Left = 5980
        Case 6
            shpSelectedAreaPerimeter.Left = 6910
        Case 7
            shpSelectedAreaPerimeter.Left = 7840
        Case 8
            shpSelectedAreaPerimeter.Left = 8770
    End Select
    If Not bIsInitialLoad Then
        PopulateSummaryEstimateComponents
        PopulateAssemblyComponents GetBldgModelSKey(m_rec.Fields("bldg_model_skey").Value), GetBuildingArea(), -1, strErrMsg
        txtBldgCostDesc.Text = Replace(Bldg_Cost_Desc_Container, "[BLDG_AREA]", Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption))
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub SetShpTopLocationResi()
    Dim strErrMsg As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    '
    '   Get the row we're in.
    shpSelectedAreaPerimeterResi.Top = 120
    '
    '   Get the column we're in.
    Select Case Right$(sshpSelectedArea, Len(sshpSelectedArea) - InStr(1, sshpSelectedArea, ","))
        Case 0
            shpSelectedAreaPerimeterResi.Left = 1330
        Case 1
            shpSelectedAreaPerimeterResi.Left = 2210
        Case 2
            shpSelectedAreaPerimeterResi.Left = 3080
        Case 3
            shpSelectedAreaPerimeterResi.Left = 3950
        Case 4
            shpSelectedAreaPerimeterResi.Left = 4810
        Case 5
            shpSelectedAreaPerimeterResi.Left = 5685
        Case 6
            shpSelectedAreaPerimeterResi.Left = 6560
        Case 7
            shpSelectedAreaPerimeterResi.Left = 7425
        Case 8
            shpSelectedAreaPerimeterResi.Left = 8300
        Case 9
            shpSelectedAreaPerimeterResi.Left = 9160
        Case 10
            shpSelectedAreaPerimeterResi.Left = 10030
    End Select
    If Not bIsInitialLoad Then
        PopulateSummaryEstimateComponents
        PopulateAssemblyComponents GetBldgModelSKey(m_rec.Fields("bldg_model_skey").Value), GetBuildingArea(), -1, strErrMsg
        txtBldgCostDesc.Text = Replace(Bldg_Cost_Desc_Container, "[BLDG_AREA]", Trim(lblAreaResi(Right$(sshpSelectedArea, 1)).Caption))
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateFormulaValues(sFormulaCode As String)
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Select Case Trim(UCase(sFormulaCode))
        Case "G"
            lblFormulaCode.Caption = "G = (assembly_cost * Factor) / bldg_stories"
        Case "L"
            lblFormulaCode.Caption = "L = (assembly_cost * bldg_perimeter * Factor) / bldg_area"
        Case "EF"
            lblFormulaCode.Caption = "EF = ((bldg_stories - 1) * assembly_cost * factor) / bldg_stories"
        Case "F"
            lblFormulaCode.Caption = "F = assembly_cost * Factor"
        Case "W"
            lblFormulaCode.Caption = "W = (assembly_cost * bldg_perimeter * bldg_stories * bldg_stories_hgt * factor) / bldg_area"
        Case "WW"
            lblFormulaCode.Caption = "WW = (assembly_cost * bldg_perimeter * bldg_stories * bldg_stories_hgt * Factor) / (bldg_area * window_area)"
        Case "EA"
            lblFormulaCode.Caption = Trim(UCase(sFormulaCode)) & " = (assembly_cost * factor) / bldg_area"
        Case "E", "DE"
            lblFormulaCode.Caption = Trim(UCase(sFormulaCode)) & " = (assembly_cost * factor) / bldg_area_std"
        Case "P"
            lblFormulaCode.Caption = "P = (assembly_cost * bldg_part_hgt * factor) / bldg_part_density"
        Case "DI"
            lblFormulaCode.Caption = "DI = (assembly_cost * factor) / bldg_door_density"
        Case "DW"
            lblFormulaCode.Caption = "DW = ((assembly_cost * 144) * factor * bldg_stories_hgt * bldg_perimeter * bldg_stories) / bldg_area"
        Case "S"
            lblFormulaCode.Caption = "S = (assembly_cost * 2 * bldg_part_hgt * factor) / bldg_part_density"
    End Select
    Screen.MousePointer = vbNormal
End Sub

Private Sub PopulateFormulaValuesResi(sFormulaCode As String)
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Select Case Trim(UCase(sFormulaCode))
        Case "G"
            lblFormulaCode.Caption = "G = (assembly_cost * ground_area * Factor) / bldg_area"
        Case "L"
            lblFormulaCode.Caption = "L = (assembly_cost * perim_hold * Factor) / bldg_area"
        Case "EF"
            lblFormulaCode.Caption = "EF = (assembly_cost * (bldg_area - ground_area)) / bldg_area"
        Case "F"
            lblFormulaCode.Caption = "F = assembly_cost * Factor"
        Case "W"
            lblFormulaCode.Caption = "W = (assembly_cost * perim_hold * bldg_f_height * factor) / bldg_area"
        Case "WW"
            lblFormulaCode.Caption = "WW = (assembly_cost * perim_hold * bldg_f_height * Factor) / (bldg_area * window_area)"
        Case "EA", "BS"
            lblFormulaCode.Caption = Trim(UCase(sFormulaCode)) & " = (assembly_cost * factor) / bldg_area"
        Case "E"
            lblFormulaCode.Caption = "E = (assembly_cost * factor) / 1200"
        Case "DE"
            lblFormulaCode.Caption = "DE = (assembly_cost * factor * bldg_door_exterior) / bldg_area"
        Case "P"
            lblFormulaCode.Caption = "P = (assembly_cost * factor * bldg_part_hgt) / bldg_partition"
        Case "DI"
            lblFormulaCode.Caption = "DI = (assembly_cost * factor) / bldg_door_interior"
        Case "DW"
            lblFormulaCode.Caption = "DW = ((assembly_cost /144) * Factor * bldg_stories_hgt * perim_hold * bldg_stories) / bldg_area"
        Case "S"
            lblFormulaCode.Caption = "S = (assembly_cost * 2 * bldg_part_hgt * factor) / bldg_partition"
        Case "ST" 'AND bldg_type not in ('H','I','J') Enforced by calling application.
            lblFormulaCode.Caption = "ST = (assembly_cost * bldg_stair_factor) / bldg_area"
    End Select
    Screen.MousePointer = vbNormal
End Sub

Public Sub SearchForNewModel(b1stBldgSearch As Boolean, sBldgModelSkey As String, sBldgID As String)
    Dim strSelect               As String
    Dim Button                  As String
    
    On Error Resume Next
    If bIsPendingChange = True Then
        Button = MsgBox("Do you want to save your changes to " & Me.Caption & "?", vbYesNoCancel, "Search For New Model")
        If Button = vbYes Then
            '
            '   If there were errors, cancel the search
            If Not Update Then
                Exit Sub
            End If
        ElseIf Button = vbCancel Then
            '
            ' Cancel the search
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    '
    '   If we're searching for a new bldg we aren't
    '   inserting or cloning anymore.
    m_blnInsert = False
    m_blnClone = False
    bRefreshCosts = False
    '
    '   Make sure it is closed.
    With m_rec
        .Close
        '
        '   Set the maximum number to bring back.
        .MaxRecords = MAX_RECORDS
    End With
    '
    '   7/2/02 Fill in all the fields available to get the model
    '   note for inserts that since we will not have the model_skey we have to move
    '   to the end of the recordset so we get the most recently inserted model.
    strSelect = "exec sp_select_model @type_code = '" & IIf(opttype_codeC.Value = True, "C", "R") & "',"
    strSelect = strSelect & "@bldg_category = '%" & Trim(cbobldg_category.Text) & "%',"
    strSelect = strSelect & "@bldg_id = '"
    If Len(Trim(sBldgID)) > 0 Then
        strSelect = strSelect & SQLChangeWildcard(sBldgID) & "',"
    Else
        strSelect = strSelect & "%',"
    End If
    
    strSelect = strSelect & "@bldg_desc = '%" & Trim(txtbldg_desc.Text) & "%',"
    strSelect = strSelect & "@frame_type = '%" & Trim(cboFrameType.Text) & "%',"
    strSelect = strSelect & "@wall_type = '%" & Trim(cboWallType.Text) & "%',"
    strSelect = strSelect & "@bldg_model_skey = '%" & Trim(sBldgModelSkey) & "'"
    '
    '   Use DAL to perform select.
    ' rlh
    If DEBUGON Then
        Debug.Print "frmModel: SearchForNewModel: " & strSelect
    End If
    
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, m_rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for model.", vbCritical
    Else
        If m_rec.RecordCount > 1 Then
            m_rec.MoveLast
        End If
        
        bIsInitialLoad = True
        PopulateScreen b1stBldgSearch
        EnableControls
        bIsInitialLoad = False
        bIsPendingChange = False
    End If
    Status ""

End Sub
'
'   Called from frmMain when the user clicks on the
'   toolbar buttons for sorting.
Public Sub Sort(intDir As Integer)
    If tabModelDetails.Tab = 0 Then
        m_objModelAssembliesGridMap.Sort intDir
    Else
        m_objModelComponentGridMap.Sort intDir
    End If
End Sub

Private Sub EnableControls()
    Dim i As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If bIsInitialLoad Then
        cmdUpdate.Enabled = False
        If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
        Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then

            cmdDeleteClone.Enabled = False
            cmdReports.Enabled = False
            fraModelMatrixResi.Visible = True
            fraModelMatrix.Visible = False
            cmdAssemblyComponentsReport.Enabled = False
            '
            '   No Frame Type to select.
            cboFrameType.Visible = False
            lblFrameType.Visible = False
           
            cboFormatCode.Enabled = False
            cboFormatCode.BackColor = LTGREY
            '
            '   Can't change op_codes for Quality Series it's all the same.
            fraOPCode.Enabled = False
            fraOPCode.BackColor = LTGREY
            optUnion.Enabled = False
            optUnion.BackColor = LTGREY
            optOpen.Enabled = False
            optOpen.BackColor = LTGREY
            cboMdlRegionCode.Enabled = False
            cboMdlRegionCode.BackColor = LTGREY
            cboMdlCountryCode.Enabled = False
            cboMdlCountryCode.BackColor = LTGREY
        Else
            If opttype_codeC.Value = True Then
                fraModelMatrix.Visible = True
                fraModelMatrixResi.Visible = False
                '
                '   They cannot run a report for model_code 7 & 8 for Commercial
                '   assemblies are same as model_code 1 anyway.
                If m_rec.Fields("format_row").Value = 1 Then
                    cmdReports.Enabled = False
                End If
            Else
                fraModelMatrixResi.Visible = True
                fraModelMatrix.Visible = False
                '
                '   No Frame Type to select.
                cboFrameType.Visible = False
                lblFrameType.Visible = False

                If Trim(m_rec.Fields("bldg_type").Value) = "H" Or _
                    Trim(m_rec.Fields("bldg_type").Value) = "I" Or _
                    Trim(m_rec.Fields("bldg_type").Value) = "J" Then
                
                    fraModelMatrixResi.Width = 8325
                    For i = 8 To 10
                        lblAreaResi(i).Visible = False
                        lblPerimeterResi(i).Visible = False
                        lblMaterialOPResi(i).Visible = False
                        lblLaborOPResi(i).Visible = False
                        lblEquipmentOPResi(i).Visible = False
                        lblInstallOPResi(i).Visible = False
                        lblTotalOPResi(i).Visible = False
                    Next i
                End If
            End If
            
            If m_blnInsert And m_blnClone = False Then
                cmdDeleteClone.Enabled = False
            Else
                cmdDeleteClone.Enabled = m_blnClone
            End If
        End If
    End If
    If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
    Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
    Else
        '
        '   If we're on a format_row...
        If ((m_rec.Fields("format_row").Value) = 1) Then
            cboFormatCode.Enabled = False
            '
            '   disabled gray color
            cboFormatCode.BackColor = LTGREY
            cmdDeleteAssembly.Enabled = False
        End If
    End If
    '
    '   No current record - disable buttons
    If TDBGridAssembly.Bookmark >= 1 Then
        cmdDeleteAssembly.Enabled = True
    Else
        cmdDeleteAssembly.Enabled = False
    End If
    
    ' We only allow cloning for commercial, and model codes 1-6.
    If (opttype_codeC.Value = True And _
        Trim(m_rec.Fields("model_code").Value) >= STANDARD_MODEL_CODE_MIN And _
        Trim(m_rec.Fields("model_code").Value) <= STANDARD_MODEL_CODE_MAX) Then
        
        cmdCloneStandardModel.Enabled = True
    Else
        cmdCloneStandardModel.Enabled = False
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Function ValidateAssemblies() As String

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With TDBGridAssembly
        If m_objModelAssembliesGridMap.bInsertInProcess Then
            ValidateAssemblies = "Your last assembly insert is still in progress.  " & vbCrLf & "Please move to another row allowing the insert to save within the grid, then click the Update button."
            '
            '   If we failed on this assembly validation then
            '   re-enable the cmdupdate since the user doesn't
            '   have to actually change text just move to another
            '   row and then click update.
            If ValidateAssemblies <> "" Then
                cmdUpdate.Enabled = True
            End If
        Else
            .MoveFirst
        
            Do Until IsNull(.Bookmark)
                If Trim(.Columns("Assembly ID").Value) = "" Then
                    ValidateAssemblies = "Please provide an Assembly ID."
                    Exit Do
                ElseIf Trim(.Columns("Algorithm").Value) = "" Then
                    ValidateAssemblies = "Please provide an Algorithm for Assembly ID: " & Trim(.Columns("Assembly ID").Value)
                    Exit Do
                ElseIf IsNumeric(Trim(.Columns("Algorithm").Value)) = True Then
                    ValidateAssemblies = "Please provide a non-numeric Algorithm for Assembly ID: " & Trim(.Columns("Assembly ID").Value)
                    Exit Do
                ElseIf Trim(.Columns("Formula Factor").Value) = "" Then
                    ValidateAssemblies = "Please provide an Formula Factor for Assembly ID: " & Trim(.Columns("Assembly ID").Value)
                    Exit Do
                ElseIf IsNumeric(Trim(.Columns("Formula Factor").Value)) = False Then
                    ValidateAssemblies = "Please provide an numeric Formula Factor for Assembly ID: " & Trim(.Columns("Assembly ID").Value)
                    Exit Do
                ElseIf Val(Trim(.Columns("Formula Factor").Value)) <= 0 Then
                    ValidateAssemblies = "Please provide a Formula Factor that is greater than zero for Assembly ID: " & Trim(.Columns("Assembly ID").Value)
                    Exit Do
                Else
                    If opttype_codeC.Value = True Then
                        Select Case UCase(Trim(.Columns("Algorithm").Value))
                            Case "G", "L", "EF", "F", "W", "WW", "EA", "E", "DE", "P", "DI", "DW", "S"
                            Case Else
                                ValidateAssemblies = "Please provide a valid Algorithm for Assembly ID: " & Trim(.Columns("Assembly ID").Value)
                                Exit Do
                        End Select
                    Else
                        Select Case UCase(Trim(.Columns("Algorithm").Value))
                            Case "G", "L", "EF", "F", "W", "WW", "EA", "E", "DE", "BS", "P", "DI", "DW", "S", "ST"
                            Case Else
                                ValidateAssemblies = "Please provide a valid Algorithm for Assembly ID: " & Trim(.Columns("Assembly ID").Value)
                                Exit Do
                        End Select
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    Screen.MousePointer = vbNormal
End Function

Private Function ValidateSummaryEstimate() As String

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    With TDBGridSummaryEstimate
        .MoveFirst

        Do Until IsNull(.Bookmark)
            If Len(Trim(.Columns("Unit").Value)) > 15 Then
                ValidateSummaryEstimate = "Please provide a Unit that is less that 15 characters for System Component: " & Trim(.Columns("System Component").Value)
            ElseIf Len(Trim(.Columns("Specifications").Value)) > 300 Then
                ValidateSummaryEstimate = "Please provide a Specification that is less that 300 characters for System Component: " & Trim(.Columns("System Component").Value)
            End If
            .MoveNext
        Loop
    End With
    Screen.MousePointer = vbNormal
End Function

Private Function RefreshCostsCommercial() As Boolean
    Dim strUpdate       As String
    Dim cmdTemp         As New ADODB.Command
    Dim i               As Integer

    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    RefreshCostsCommercial = True
    Status ("Updating Building Cost Information For Model: " & txtbldg_model_skey.Text & " ...")
    With cnTemp
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
        'rlh - 06/06/07 CCD 8.2 release
        strUpdate = "exec sp_update_bldg_model @bldg_model_skey = '"
        'strUpdate = "exec sp_update_bldg_model_rlh @bldg_model_skey = '"
        If txtbldg_model_skey = "" Then
            'exit can't refresh yet.
            Exit Function
        Else
            strUpdate = strUpdate & Trim(txtbldg_model_skey.Text) & "',"
        End If
        strUpdate = strUpdate & "@op_code = 'STD',"
        '
        'allow to update & change order of models?
        strUpdate = strUpdate & "@country_code = '" & cboMdlCountryCode.Text & "',"
        strUpdate = strUpdate & "@region_code = '" & cboMdlRegionCode.Text & "'"
        With cmdTemp
            .CommandTimeout = 50000
            .CommandType = adCmdText
            .CommandText = strUpdate
            
            ' rlh
            If DEBUGON Then
                Debug.Print "frmModel: RefreshCostsCommercial (#1): " & strUpdate
            End If
            
            .Execute adExecuteNoRecords
        End With
        
        If cnTemp.Errors.Count = 0 Then
            strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
            With cmdTemp
                .CommandTimeout = 50000
                .CommandType = adCmdText
                .CommandText = strUpdate
                 ' rlh
                If DEBUGON Then
                    Debug.Print "frmModel: RefreshCostsCommercial (errors count=0): " & strUpdate
                End If
                
                .Execute adExecuteNoRecords
            End With
            If cnTemp.Errors.Count <> 0 Then
                Screen.MousePointer = vbNormal
                MsgBox "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                    & Trim(txtbldg_model_skey.Text) & " " & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            End If
        Else
            Screen.MousePointer = vbNormal
            MsgBox "Errors in the RefreshCostsCommercial routine for Building Model skey: " _
                & Trim(txtbldg_model_skey.Text) & " " & vbCrLf & cnTemp.Errors(0).Description, vbCritical
        End If
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    RefreshCostsCommercial = False
    MsgBox "Errors in the RefreshCostsCommercial routine: " & Err.Description
    Status ("")
End Function

Private Function RefreshCostsResidential() As Boolean
    Dim strUpdate       As String
    Dim cmdTemp         As New ADODB.Command
    Dim i               As Integer

    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    RefreshCostsResidential = True
    Status ("Updating Building Cost Information For Model: " & txtbldg_model_skey.Text & " ...")
    With cnTemp
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
        '
        '   Updates cost for STD & OPN op_codes.
        strUpdate = "exec sp_update_bldg_model_resi @bldg_model_skey = '"
        If txtbldg_model_skey = "" Then
            'exit can't refresh yet.
            Exit Function
        Else
            strUpdate = strUpdate & Trim(txtbldg_model_skey.Text) & "',"
        End If
        strUpdate = strUpdate & "@op_code = 'STD',"
        '
        'allow to update & change order of models?
        strUpdate = strUpdate & "@country_code = '" & cboMdlCountryCode.Text & "',"
        strUpdate = strUpdate & "@region_code = '" & cboMdlRegionCode.Text & "'"
        With cmdTemp
            .CommandTimeout = 50000
            .CommandType = adCmdText
            .CommandText = strUpdate
            .Execute adExecuteNoRecords
        End With
        DoEvents
        If cnTemp.Errors.Count = 0 Then
            strUpdate = Replace(strUpdate, "@op_code = 'STD'", "@op_code = 'OPN'", 1)
            With cmdTemp
                .CommandTimeout = 50000
                .CommandType = adCmdText
                .CommandText = strUpdate
                .Execute adExecuteNoRecords
            End With

            DoEvents
            If cnTemp.Errors.Count <> 0 Then
                Screen.MousePointer = vbNormal
                MsgBox "Errors in the RefreshCosts routine for Building Model skey: " _
                    & Trim(txtbldg_model_skey.Text) & " " & vbCrLf & cnTemp.Errors(0).Description, vbCritical
                RefreshCostsResidential = False
            End If
        Else
            Screen.MousePointer = vbNormal
            MsgBox "Errors in the RefreshCosts routine for Building Model skey: " _
                & Trim(txtbldg_model_skey.Text) & " " & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            RefreshCostsResidential = False
        End If
    End With
    Exit Function
    
errorHandler:
    Screen.MousePointer = vbNormal
    RefreshCostsResidential = False
    MsgBox "Errors in the RefreshCostsResidential routine: " & Err.Description
    Status ("")
End Function

Private Sub cmdDeleteAssembly_Click()
    Dim varButton
    
    ' rlh
    If DEBUGON Then
        Debug.Print "DELETE button clicked..............................."
    End If
    
    On Error Resume Next
    With TDBGridAssembly
        If .SelBookmarks.Count = 1 Then
            varButton = MsgBox(CStr(.SelBookmarks.Count) + " records will be deleted.  Are you sure you want to delete this row permanently?", vbYesNo + vbCritical)
            If varButton = vbYes Then
                .Delete
            End If
        '
        '   If multiple records are selected.
        ElseIf .SelBookmarks.Count > 1 Then
            varButton = MsgBox(CStr(.SelBookmarks.Count) + " records will be deleted.  Are you sure you want to delete these rows permanently?", vbYesNo + vbCritical)
            If varButton = vbYes Then
                m_objModelAssembliesGridMap.Delete
            End If
        End If
    End With

End Sub

Private Sub cmdDeleteClone_Click()
    Dim Button          As String
    Dim strUpdate       As String
    Dim strError        As String
    
    On Error Resume Next
    Button = MsgBox("Are you sure you want to delete this cloned building model?", vbYesNo + vbCritical)
    If Button = vbYes Then
      Screen.MousePointer = vbHourglass

      strUpdate = "exec sp_delete_bldg_model @bldg_model_skey = '"
      strUpdate = strUpdate & Trim(txtbldg_model_skey.Text) & "'"
      
      If Not g_objDAL.ExecQuery(vbNullString, strUpdate, strError) Then
          MsgBox "Error deleting building model clone. " & vbCrLf & strError & ".", vbCritical
      Else
        '
        '   Always refresh forms that are listening for changes in case part of the update succeeded.
        '   ie -the bldg was updated and the grid has an old last_update_id.
        EventSubscriberNotify esnModelRecordUpdated, m_rec.Fields("bldg_id").Value
        Me.Hide
        Unload Me
      End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdReports_Click()
    If opttype_codeC.Value = True Then
        RunReportCommercial
    Else
        RunReportResidential
    End If
End Sub

Private Sub RunReportCommercial()
    Dim strSelect               As String
    'Dim frm                     As New frmSummaryEstimateRpt

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    strSelect = "exec sp_rpt_summary_estimate_commercial @bldg_model_skey = '"
    strSelect = strSelect & Trim(txtbldg_model_skey.Text)
    '
    '   Indicates where the shpSelectedArea is for modelmaint button click.
    '   In the format of 1,1 meaning row 1 col area1.
    strSelect = strSelect & "', @bldg_area = '"
    strSelect = strSelect & lblArea(Right$(sshpSelectedArea, 1)).Caption
         
    strSelect = strSelect & "', @op_code = '"
    If optUnion.Value = True Then
        strSelect = strSelect & "STD"
    Else
        strSelect = strSelect & "OPN"
    End If
    
    strSelect = strSelect & "', @country_code = '"
    If Len(Trim(cboMdlCountryCode.Text)) = 0 Then
        strSelect = strSelect & "USA"
    Else
        strSelect = strSelect & SQLChangeWildcard(cboMdlCountryCode.Text)
    End If
    
    strSelect = strSelect & "', @region_code = '"
    If Len(Trim(cboMdlRegionCode.Text)) = 0 Then
        strSelect = strSelect & "NAT'"
    Else
        strSelect = strSelect & SQLChangeWildcard(cboMdlRegionCode.Text) & "'"
    End If
    
    Screen.MousePointer = vbNormal
    'frm.RunReportCommercial strSELECT
    CommercialEstimatePreview strSelect
    
End Sub

Private Sub RunReportResidential()
    Dim strSelect               As String
    'Dim frm                     As New frmSummaryEstimateRpt

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    strSelect = "exec sp_rpt_summary_estimate_residential @bldg_model_skey = '"
    strSelect = strSelect & Trim(txtbldg_model_skey.Text)
    '
    '   Indicates where the shpSelectedArea is for modelmaint button click.
    '   In the format of 1,1 meaning row 1 col area1.
    strSelect = strSelect & "', @bldg_area = '"
    strSelect = strSelect & lblAreaResi(Right$(sshpSelectedArea, 1)).Caption
         
    strSelect = strSelect & "', @op_code = '"
    If optUnion.Value = True Then
        strSelect = strSelect & "STD"
    Else
        strSelect = strSelect & "OPN"
    End If
    
    strSelect = strSelect & "', @country_code = '"
    If Len(Trim(cboMdlCountryCode.Text)) = 0 Then
        strSelect = strSelect & "USA"
    Else
        strSelect = strSelect & SQLChangeWildcard(cboMdlCountryCode.Text)
    End If
    
    strSelect = strSelect & "', @region_code = '"
    If Len(Trim(cboMdlRegionCode.Text)) = 0 Then
        strSelect = strSelect & "NAT'"
    Else
        strSelect = strSelect & SQLChangeWildcard(cboMdlRegionCode.Text) & "'"
    End If
    
    Screen.MousePointer = vbNormal
    'frm.RunReportResidential strSELECT
    ResidentialEstimatePreview strSelect

End Sub

Private Sub cmdUpdate_Click()
    '
    '   Set cmdUpdate to not enabled so they must change
    '   data to update again or if they fail screen validation
    '   they must change before updating.
    '
    '   NOTE that if they fail the common add validation insert in progress
    '   then ValidateAssemblies will re-enable the cmdupdate since the user doesn't
    '   have to actually change text just move to another
    '   row and then click update.
    cmdUpdate.Enabled = False
    ' rlh
    If DEBUGON Then
        Debug.Print "                             "
        Debug.Print "UPDATE button clicked---------------------------: "
    End If
    
    Update
End Sub
' This routine will do the following steps:
' For commercial:
' 1.  Validate the fields on the screen.
' 2.  Gets a new form id from the db (like an identity column).  Note that sometimes these ids are re-used
'     because they will be deleted during cleanup.  But if multiple users are active, their id's are deleted, but
'     the next available id might be higher than their original one.
' 3.  Inserts the assemblies into a temp table (since there will be a variable (and sometimes large) amount
'     of assemblies, which would make it problematic to pass a large list into the stored proc.
' 4.  Update the summary estimate.
' 5.  Update the costs for the assemblies, and write the assemblies to the permanent assembly_usage table.
' 6.  Refresh the costs if requested.
' 7.  Cleanup the temp tables.
' For residential:


Private Function Update() As Boolean
    Dim nmodel_form     As Integer
    Dim bAllAreas       As Boolean
    Dim varButton
    
    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    
    If opttype_codeC.Value = True Then
        If ValidateCommercialScreen Then
            '
            '   Now branch between inserting a new building or updating an existing building.
            If m_blnInsert = True And m_blnClone = False Then
                Status ("Updating Model Details ...")
                If InsertCommercial Then
                    Update = True
                End If
            Else
                Status ("Getting Temporary Model Form ID ...")
                '
                '   Get the form id that uniquely identifies us so that if we
                '   added assemblies or summary estimate we can get them from the
                '   assembly_usage_holding_table and published_bldg_component_cost_holding_table
                '   This is a required parameter for the update sp's for Commercial & Residential.
                nmodel_form = GetFormID
                If nmodel_form = 0 Then
                    '
                    '   There was an error so we can't update.
                    Exit Function
                End If
                '
                '   If we're on row 7 or 8 then they can't update Assemblies or Summary Estimate.
                If m_rec.Fields("format_row").Value <> 1 Then
                    If m_objModelComponentGridMap.IsPendingChange Then
                        varButton = MsgBox("Would you like to apply Summary Estimate changes to all Areas for this model?", vbYesNo, "Summary Estimate Updates")
                        If varButton = vbYes Then
                            bAllAreas = True
                        End If
                    End If

                    Status ("Updating Temporary Assemblies Table ...")
                    DoEvents
                    '
                    '   Insert the Assemblies in the temp table that the
                    '   model update sp uses to commit finals from.
                    If UpdateAssemblies(nmodel_form) Then
                        Status ("Updating Temporary Summary Estimate Table ...")
                        DoEvents
                        '
                        '   Insert the Summary Estimate in the temp table that the
                        '   model update sp uses to commit finals from.
                        If UpdateSummaryEstimate(bAllAreas, nmodel_form) Then
                            Status ("Updating Model Details ...")
                            DoEvents
                            If UpdateCommercial(nmodel_form) Then
                                If bRefreshCosts Then
                                    If RefreshCostsCommercial Then
                                        Update = True
                                    End If
                                Else
                                    Update = True
                                End If
                            End If
                        End If
                    End If
                Else
                    Status ("Updating Model Details ...")
                    DoEvents
                    If UpdateCommercial(nmodel_form) Then
                        Update = True
                    End If
                End If
            End If
            Status ("Cleaning Temporary Tables ...")
            '
            '   Regardless if we updated ok or not, cleanup the assembly_usage_holding_table
            '   published_bldg_component_cost_holding_table and tmp_form_id tables for
            '   assemblies and summary estimate.
            CleanupTmpTables nmodel_form

            If Update Then
                Status ("Model Details Updated Successfully...")
                MsgBox "Model Updated Successfully.", vbInformation
                Status ("Refreshing Model Maintenance Screen...")
                bIsPendingChange = False
                m_objModelAssembliesGridMap.bCloneAssembliesInProcess = False
                SearchForNewModel True, txtbldg_model_skey.Text, txtbldg_id.Text
            End If
        End If
    Else
        If ValidateResidentialScreen Then
            '
            '   Now branch between inserting a new building or updating an existing building.
            If m_blnInsert = True And m_blnClone = False Then
                Status ("Updating Model Details ...")
                If InsertResidential Then
                    Update = True
                End If
            Else
                Status ("Getting Temporary Model Form ID ...")
                DoEvents
                '
                '   Get the form id that uniquely identifies us so that if we
                '   added assemblies or summary estimate we can get them from the
                '   assembly_usage_holding_table and published_bldg_component_cost_holding_table
                '   This is a required parameter for the update sp's for Commercial & Residential.
                nmodel_form = GetFormID
                If nmodel_form = 0 Then
                    '
                    '   There was an error so we can't update.
                    Exit Function
                End If
                '
                '   If we're on row 7 or 8 then they can't update Assemblies or Summary Estimate.
                '   Unless their on a Quality Series building model. bldg_id 100,200,300,400.
                If m_rec.Fields("format_row").Value <> 1 Then
                     
                    If m_objModelComponentGridMap.IsPendingChange Then
                        varButton = MsgBox("Would you like to apply Summary Estimate changes to all Areas for this model?", vbYesNo, "Summary Estimate Updates")
                        If varButton = vbYes Then
                            bAllAreas = True
                        End If
                    End If
    
                    Status ("Updating Temporary Assemblies Table ...")
                    DoEvents
                    '
                    '   Insert the Assemblies in the temp table that the
                    '   model update sp uses to commit finals from.
                    If UpdateAssemblies(nmodel_form) Then
                        Status ("Updating Temporary Summary Estimate Table ...")
                        '
                        '   Insert the Summary Estimate in the temp table that the
                        '   model update sp uses to commit finals from.
                        If UpdateSummaryEstimate(bAllAreas, nmodel_form) Then
                            Status ("Updating Model Details ...")
                            If UpdateResidential(nmodel_form) Then
                                If bRefreshCosts Then
                                    If RefreshCostsResidential Then
                                        Update = True
                                    End If
                                Else
                                    Update = True
                                End If
                            End If
                        End If
                    End If
                '
                '   They can't add Summary Estimate for Quality Series buildings.
                ElseIf m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
                Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
                    Status ("Updating Temporary Assemblies Table ...")
                    '
                    '   Insert the Assemblies in the temp table that the
                    '   model update sp uses to commit finals from.
                    If UpdateAssemblies(nmodel_form) Then
                        Status ("Updating Model Details ...")
                        
                        If UpdateResidential(nmodel_form) Then
                            If RefreshCostsResidential Then
                                Update = True
                            End If
                        End If
                    End If
                Else
                    Status ("Updating Model Details ...")
                    
                    If UpdateResidential(nmodel_form) Then
                        Update = True
                    End If
                End If
            End If
            Status ("Cleaning Temporary Tables ...")
            '
            '   Regardless if we updated ok or not, cleanup the assembly_usage_holding_table
            '   published_bldg_component_cost_holding_table and tmp_form_id tables for
            '   assemblies and summary estimate.
            CleanupTmpTables nmodel_form
            
            If Update Then
                Status ("Model Details Updated Successfully...")
                MsgBox "Model Updated Successfully.", vbInformation
                Status ("Refreshing Model Maintenance Screen...")
                bIsPendingChange = False
                SearchForNewModel True, txtbldg_model_skey.Text, txtbldg_id.Text
            End If
        End If
    End If
    '
    '   Always refresh forms that are listening for changes in case part of the update succeeded.
    '   ie -the model was updated and the grid or bldg has an old last_update_id.
    EventSubscriberNotify esnModelRecordUpdated, txtbldg_id.Text
    Status ("")
    Screen.MousePointer = vbNormal
    Exit Function
    
errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the Update routine: " & Err.Description, vbCritical
    Status ("")

End Function
' Gets the next available "form id" from the database.  This is a key that uniquely identifies our session
' within the temp tables.
Private Function GetFormID() As Integer
    Dim recTemp         As New ADODB.RecordSet
    Dim sErrorDesc      As String
    Dim bOK             As Boolean
    Dim org_form_id     As Integer
    Dim strUpdate       As String
    
    On Error Resume Next
    Do Until bOK
        '
        '   Get bldg_form number to use in tmp table to indentify common adds as ours.
        If g_objDAL.GetRecordset(vbNullString, "SELECT MAX(form_id) AS form_id FROM form_id", recTemp) Then
            If sErrorDesc = "" Then
                org_form_id = recTemp.Fields("form_id").Value
                recTemp.Close
            
                strUpdate = "INSERT INTO form_id(form_id, form_type) VALUES('" & org_form_id + 1 & "', 'M')"
                If g_objDAL.ExecQuery(vbNullString, strUpdate, sErrorDesc) Then
                    If sErrorDesc = "" Then
                        GetFormID = org_form_id + 1
                        bOK = True
                    '
                    '   If the error is due to primary key constraint then just get the next
                    '   form_id and try again.  Otherwise we have a real error so exit.
                    ElseIf sErrorDesc = "[Microsoft][ODBC SQL Server Driver][SQL Server]Violation of PRIMARY KEY constraint 'PK_tmp_form_id'. Cannot insert duplicate key in object 'tmp_form_id'." Then
                        sErrorDesc = ""
                    Else
                        Screen.MousePointer = vbNormal
                        MsgBox "Error setting Form ID in temporary table 'form_id' " _
                            & vbCrLf & "Error: " & sErrorDesc, vbCritical
                        GetFormID = 0
                        Exit Do
                    End If
                Else
                    Screen.MousePointer = vbNormal
                    MsgBox "Error setting Form ID in temporary table 'form_id' " _
                        & vbCrLf & "Error: " & sErrorDesc, vbCritical
                    GetFormID = 0
                    Exit Do
                End If
            Else
                Screen.MousePointer = vbNormal
                MsgBox "Error selecting Form ID in temporary table 'form_id' " _
                    & vbCrLf & "Error: " & sErrorDesc, vbCritical
                GetFormID = 0
                Exit Do
            End If
        Else
            Screen.MousePointer = vbNormal
            MsgBox "Error selecting Form ID in temporary table 'form_id' " _
                & vbCrLf & "Error: " & sErrorDesc, vbCritical
            GetFormID = 0
            Exit Do
        End If
    Loop
End Function

Private Function UpdateAssemblies(nmodel_form As Integer) As Boolean
     
    On Error GoTo errorHandler:
    '
    '   Now update the Assemblies grid
    With TDBGridAssembly
        .MoveFirst
        .Update
        UpdateAssemblies = m_objModelAssembliesGridMap.Update("SF", Trim(txtbldg_model_skey.Text), nmodel_form)
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the UpdateAssemblies routine: " & Err.Description
    Status ("")
End Function

Private Function UpdateSummaryEstimate(bAllAreas As Boolean, nmodel_form As Integer) As Boolean
    
    On Error GoTo errorHandler:
    '
    '   Now update the Summary Estimate grid
    With TDBGridSummaryEstimate
        .MoveFirst
        .Update
        UpdateSummaryEstimate = m_objModelComponentGridMap.Update(bAllAreas, nmodel_form)
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the UpdateSummaryEstimate routine: " & Err.Description
    Status ("")
End Function

Private Function UpdateCommercial(nmodel_form As Integer) As Boolean
    
    Dim cmdTemp         As New ADODB.Command
    Dim strError        As String
    Dim strUpdate       As String
    Dim bInTrans        As Boolean
    
    bInTrans = False
    
    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    strUpdate = "exec sp_update_commercial_model @bldg_model_skey = '" & Trim(txtbldg_model_skey.Text) & "',"
    strUpdate = strUpdate & "@frame_type = '" & SQLFixString(Trim(cboFrameType.Text)) & "',"
    strUpdate = strUpdate & "@wall_type = '" & SQLFixString(Trim(cboWallType.Text)) & "',"
    strUpdate = strUpdate & "@costworks_desc = '" & SQLFixString(Trim(txtCostWorksDesc.Text)) & "',"
    strUpdate = strUpdate & "@format_code = '" & Trim(cboFormatCode.Text) & "',"
    strUpdate = strUpdate & "@model_form = " & nmodel_form & ","

    strUpdate = strUpdate & " @last_update_id = '" & Trim(m_rec.Fields("last_update_id").Value) & "',"
    strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"

    With cnTemp
        .BeginTrans
        bInTrans = True
        
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            
            ' rlh
            If DEBUGON Then
                Debug.Print "frmModel: UpdateCommercial:Update: " & strUpdate
            End If
            
            .Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the UpdateCommercial routine. " _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            
            .RollbackTrans
            bInTrans = False
        Else
            .CommitTrans
            bInTrans = False
            UpdateCommercial = True
        End If
    End With
    Exit Function

errorHandler:
    If (bInTrans = True) Then
        cnTemp.RollbackTrans
    End If
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the UpdateCommercial routine: " & Err.Description, vbCritical
    Status ("")
End Function

Private Function UpdateResidential(nmodel_form As Integer) As Boolean
    
    Dim cmdTemp         As New ADODB.Command
    Dim strError        As String
    Dim strUpdate       As String
    
    On Error GoTo errorHandler:
    Screen.MousePointer = vbHourglass
    strUpdate = "exec sp_update_residential_model @bldg_model_skey = '" & Trim(txtbldg_model_skey.Text) & "',"
    '
    '   Frame type is not populated for Residential it is contained in the wall type.
    strUpdate = strUpdate & "@frame_type = '',"
    strUpdate = strUpdate & "@wall_type = '" & SQLFixString(Trim(cboWallType.Text)) & "',"
    strUpdate = strUpdate & "@costworks_desc = '" & SQLFixString(Trim(txtCostWorksDesc.Text)) & "',"
    strUpdate = strUpdate & "@format_code = '" & Trim(cboFormatCode.Text) & "',"
    strUpdate = strUpdate & "@model_form = " & nmodel_form & ","

    strUpdate = strUpdate & " @last_update_id = '" & Trim(m_rec.Fields("last_update_id").Value) & "',"
    strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"

    With cnTemp
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            .Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the UpdateResidential routine. " _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            UpdateResidential = True
        End If
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the UpdateResidential routine: " & Err.Description, vbCritical
    Status ("")
End Function
'
'   Used to enforce business rules and data integrity.
Private Function ValidateCommercialScreen() As Boolean
    Dim strMessage  As String
    
    On Error Resume Next
    ValidateCommercialScreen = True
    If m_blnInsert And m_blnClone = False Then
        If Trim(cboWallType.Text) = "" Then
            strMessage = "Please provide a wall type."
            cboWallType.SetFocus
    
        ElseIf Len(Trim(cboWallType.Text)) > 30 Then
            strMessage = "Maximum length for wall type is 30 characters."
            cboWallType.SetFocus
    
        ElseIf Trim(cboFrameType.Text) = "" Then
            strMessage = "Please provide a frame type."
            cboFrameType.SetFocus

        ElseIf Len(Trim(cboFrameType.Text)) > 30 Then
            strMessage = "Maximum length for frame type is 30 characters."
            cboFrameType.SetFocus
        
        ElseIf Trim(cboFormatCode.Text) = "" Then
            strMessage = "Please select a format code."
            cboFormatCode.SetFocus
        End If
    Else
        If Trim(cboWallType.Text) = "" Then
            strMessage = "Please provide a wall type."
            cboWallType.SetFocus
        
        ElseIf Len(Trim(cboWallType.Text)) > 30 Then
            strMessage = "Maximum length for wall type is 30 characters."
            cboWallType.SetFocus

        ElseIf Trim(cboFrameType.Text) = "" Then
            strMessage = "Please provide a frame type."
            cboFrameType.SetFocus
            
        ElseIf Len(Trim(cboFrameType.Text)) > 30 Then
            strMessage = "Maximum length for frame type is 30 characters."
            cboFrameType.SetFocus
        End If
        If strMessage = "" Then
            strMessage = ValidateAssemblies
        End If
        If strMessage = "" Then
            strMessage = ValidateSummaryEstimate
        End If
    End If
    
    If strMessage <> "" Then
        Screen.MousePointer = vbNormal
        ValidateCommercialScreen = False
        MsgBox strMessage, vbCritical
    End If
End Function

Private Sub CleanupTmpTables(nmodel_form As Integer)
    Dim strUpdate   As String
    Dim sErrorDesc  As String
    
    On Error Resume Next
    strUpdate = "DELETE FROM assembly_usage_holding_table WHERE model_form = '" & nmodel_form & "'"
    If Not g_objDAL.ExecQuery(vbNullString, strUpdate, sErrorDesc) Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Assemblies temporary table 'assembly_usage_holding_table' for form_id: " & nmodel_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    ElseIf sErrorDesc <> "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Assemblies temporary table 'assembly_usage_holding_table' for form_id: " & nmodel_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    End If
    
    strUpdate = "DELETE FROM published_bldg_component_cost_holding_table WHERE model_form = '" & nmodel_form & "'"
    If Not g_objDAL.ExecQuery(vbNullString, strUpdate, sErrorDesc) Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Summary Estimate temporary table 'published_bldg_component_cost_holding_table' for form_id: " & nmodel_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    ElseIf sErrorDesc <> "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Summary Estimate temporary table 'published_bldg_component_cost_holding_table' for form_id: " & nmodel_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    End If
    
    strUpdate = "DELETE FROM form_id WHERE form_id = '" & nmodel_form & "'"
    If Not g_objDAL.ExecQuery(vbNullString, strUpdate, sErrorDesc) Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Form ID temporary table 'form_id' for form_id: " & nmodel_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    ElseIf sErrorDesc <> "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Error cleaning up Form ID temporary table 'form_id' for form_id: " & nmodel_form _
            & vbCrLf & "Error: " & sErrorDesc, vbCritical
        Exit Sub
    End If
End Sub
'
'   Used to enforce business rules and data integrity.
Private Function ValidateResidentialScreen() As Boolean
    Dim strMessage  As String
    
    On Error Resume Next
    ValidateResidentialScreen = True
        
    If m_blnInsert And m_blnClone = False Then
        If Trim(cboWallType.Text) = "" Then
            strMessage = "Please provide a wall type."
            cboWallType.SetFocus
    
        ElseIf Len(Trim(cboWallType.Text)) > 30 Then
            strMessage = "Maximum length for wall type is 30 characters."
            cboWallType.SetFocus
           
        ElseIf Trim(cboFormatCode.Text) = "" Then
            strMessage = "Please select a format code."
            cboFormatCode.SetFocus
        End If
    Else
        If Trim(cboWallType.Text) = "" Then
            strMessage = "Please provide a wall type."
            cboWallType.SetFocus
        
        ElseIf Len(Trim(cboWallType.Text)) > 30 Then
            strMessage = "Maximum length for wall type is 30 characters."
            cboWallType.SetFocus
        End If
        
        If strMessage = "" Then
            strMessage = ValidateAssemblies
        End If
        If strMessage = "" Then
            strMessage = ValidateSummaryEstimate
        End If
    End If
    
    If strMessage <> "" Then
        Screen.MousePointer = vbNormal
        ValidateResidentialScreen = False
        MsgBox strMessage, vbCritical
    End If
End Function

Private Function InsertCommercial() As Boolean
    Dim cmdTemp         As New ADODB.Command
    Dim strError        As String
    Dim strUpdate       As String
    
    On Error GoTo errorHandler:
    
    strUpdate = "exec sp_insert_commercial_model @bldg_skey= '" & Trim(txtbldg_skey.Text) & "',"
    strUpdate = strUpdate & "@frame_type = '" & SQLFixString(Trim(cboFrameType.Text)) & "',"
    strUpdate = strUpdate & "@wall_type = '" & SQLFixString(Trim(cboWallType.Text)) & "',"
    strUpdate = strUpdate & "@format_code = '" & Trim(cboFormatCode.Text) & "',"
    strUpdate = strUpdate & "@costworks_desc = '" & SQLFixString(Trim(txtCostWorksDesc.Text)) & "',"
    strUpdate = strUpdate & "@last_update_person = '" & strUserName & "'"
    
    With cnTemp
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            .Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the InsertCommercial routine. " _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            InsertCommercial = True
        End If
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the InsertCommercial routine: " & Err.Description, vbCritical
    Status ("")
End Function

Private Function InsertResidential() As Boolean
    Dim cmdTemp         As New ADODB.Command
    Dim strError        As String
    Dim strUpdate       As String
    
    On Error GoTo errorHandler:
    
    strUpdate = "exec sp_insert_residential_model @bldg_skey= '" & Trim(txtbldg_skey.Text) & "',"
    strUpdate = strUpdate & "@frame_type = '" & SQLFixString(Trim(cboFrameType.Text)) & "',"
    strUpdate = strUpdate & "@wall_type = '" & SQLFixString(Trim(cboWallType.Text)) & "',"
    strUpdate = strUpdate & "@format_code = '" & SQLFixString(Trim(cboFormatCode.Text)) & "',"
    strUpdate = strUpdate & "@costworks_desc = '" & SQLFixString(Trim(txtCostWorksDesc.Text)) & "',"
    strUpdate = strUpdate & " @last_update_person = '" & strUserName & "'"

    With cnTemp
        .BeginTrans
        Set cmdTemp = New ADODB.Command
        Set cmdTemp.ActiveConnection = cnTemp
    
        With cmdTemp
            .CommandTimeout = 0
            .CommandType = adCmdText
            .CommandText = strUpdate
            .Execute 'adExecuteNoRecords
        End With
    
        If .Errors.Count <> 0 Then
            MsgBox "Errors in the InsertResidential routine. " _
                & vbCrLf & cnTemp.Errors(0).Description, vbCritical
            
            .RollbackTrans
        Else
            .CommitTrans
            InsertResidential = True
        End If
    End With
    Exit Function

errorHandler:
    Screen.MousePointer = vbNormal
    MsgBox "Errors in the InsertResidential routine: " & Err.Description, vbCritical
    Status ("")
End Function

Private Sub cmdUnitCost_Click()
    Dim ID As String
    Dim frm As frmUCostUsageGrid
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Status ("Loading Unit Cost Usage Grid ...")
    '
    '   Open single record view with data from row selected.
    With TDBGridAssembly
        If Not IsNull(.Bookmark) Then
            Set frm = New frmUCostUsageGrid
    
            If Len(Trim(.Columns("Assembly ID").Value)) = 14 Then
                ID = Left$(.Columns("Assembly ID").Value, 5) & Right$(Left$(.Columns("Assembly ID").Value, 9), 3) & Right$(.Columns("Assembly ID").Value, 4)
            ElseIf Len(Trim(.Columns("Assembly ID").Value)) = 12 And InStr(1, Trim(.Columns("Assembly ID").Value), " ") <> 0 Then
                ID = Left$(.Columns("Assembly ID").Value, 3) & Right$(Left$(.Columns("Assembly ID").Value, 7), 3) & Right$(.Columns("Assembly ID").Value, 4)
            Else
                ID = .Columns("Assembly ID").Value
            End If
            frm.JumpIn2 Trim(ID)
        End If
    End With
    Screen.MousePointer = vbNormal
    Status ""
End Sub

Private Sub cboMdlCountryCode_Click()
    Screen.MousePointer = vbHourglass
    If Not bIsInitialLoad Then
        PopulateModelMatrix
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub cboMdlRegionCode_Click()
    Screen.MousePointer = vbHourglass
    If Not bIsInitialLoad Then
        PopulateModelMatrix
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub lblArea_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    '
    '   If they click on a label that is not populated
    '   don't do anything, unless this is the 1st time we're
    '   loading meaning the ChangeOpCostBackcolor routine is calling us.
    If lblArea(Index).Caption <> "" Then
        '
        '   If we're not already on this area go to it and repopulate everything.
        If sshpSelectedArea <> "0," & Index Then
            '
            '   We are in 1st row & the value of index is the column.
            sshpSelectedArea = "0," & Index
            SetShpTopLocation
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub lblArea_DblClick(Index As Integer)
    tabModelDetails.Tab = 1
End Sub

Private Sub lblAreaResi_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    '
    '   We don't have areas for Quality Series buildings so don't do anything.
    If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
    Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
    Else
        '
        '   If they click on a label that is not populated
        '   don't do anything, unless this is the 1st time we're
        '   loading meaning the ChangeOpCostBackcolor routine is calling us.
        If lblAreaResi(Index).Caption <> "" Then
            '
            '   If we're not already on this area go to it and repopulate everything.
            If sshpSelectedArea <> "0," & Index Then
                '
                '   We are in 1st row & the value of index is the column.
                sshpSelectedArea = "0," & Index
                SetShpTopLocationResi
            End If
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub lblAreaResi_DblClick(Index As Integer)
    tabModelDetails.Tab = 1
End Sub


Private Sub optOpen_Click()
    Dim strErrMsg As String
    
    Screen.MousePointer = vbHourglass
    If Not bIsInitialLoad Then
        PopulateModelMatrix
        PopulateAssemblyComponents GetBldgModelSKey(m_rec.Fields("bldg_model_skey").Value), GetBuildingArea(), -1, strErrMsg
        PopulateSummaryEstimateComponents
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub optUnion_Click()
    Dim strErrMsg As String
    Screen.MousePointer = vbHourglass
    If Not bIsInitialLoad Then
        PopulateModelMatrix
        PopulateAssemblyComponents GetBldgModelSKey(m_rec.Fields("bldg_model_skey").Value), GetBuildingArea(), -1, strErrMsg
        PopulateSummaryEstimateComponents
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub cboFormatCode_Click()
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        EnableControls
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub cboFrameType_Change()
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        EnableControls
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub cboFrameType_Click()
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        EnableControls
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub cboFrameType_LostFocus()
    'MODIFIED 6/20/2005 RTD
    'cboFrameType.Text = RemoveCharacters(cboFrameType.Text, "\?/|[]{}*&"":;?,~`_'-=+@!#$%^().<>")
    cboFrameType.Text = RemoveCharacters(cboFrameType.Text, "\?|[]{}*&:;?~`_=+@!#$%^()")

End Sub

Private Sub cboWallType_Change()
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        EnableControls
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
    
End Sub

Private Sub cboWallType_Click()
    
    On Error Resume Next
    If Not bIsInitialLoad Then
        EnableControls
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub cboWallType_LostFocus()
    'MODIFIED 6/20/2005 RTD
    'cboWallType.Text = RemoveCharacters(cboWallType.Text, "\?/|[]{}*&"":;?,~`_'-=+@!#$%^().")
    cboWallType.Text = RemoveCharacters(cboWallType.Text, "\?|[]{}*&:;?~`_=+@!#$%^()")
End Sub
'
'   Event raised from grid assembly class indicating an assembly
'   was permanently deleted and to refresh the costs.
Private Sub m_objModelAssembliesGridMap_RefreshCostsAssemblyDeleted()
    Screen.MousePointer = vbHourglass
    If opttype_codeC.Value = True Then
        RefreshCostsCommercial
    Else
        RefreshCostsResidential
    End If
    Status ("Refreshing Model Matrix ...")
    PopulateModelMatrix
    '
    '   Always refresh forms that are listening for changes.
    '   ie -the model was updated and the grid or bldg has an old last_update_id.
    EventSubscriberNotify esnModelRecordUpdated, txtbldg_id.Text
    Status ("")
    Screen.MousePointer = vbNormal
End Sub

'************************
Private Sub TDBGridAssembly_GotFocus()
    TDBGridAssembly.TabStop = True
End Sub

Private Sub TDBGridAssembly_LostFocus()
    TDBGridAssembly.TabStop = False
End Sub

Private Sub TDBGridAssembly_DblClick()
    cmdUnitCost_Click
End Sub

Private Sub TDBGridAssembly_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGridAssembly.DataChanged = False
End Sub

Private Sub TDBGridAssembly_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With TDBGridAssembly
        If Button = vbRightButton And IsNumeric(.Bookmark) Then
            If Len(m_objModelAssembliesGridMap.GetError(.Bookmark)) > 0 Then
                MsgBox m_objModelAssembliesGridMap.GetError(.Bookmark)
            End If
        End If
    End With
End Sub

Private Sub TDBGridAssembly_KeyUp(KeyCode As Integer, Shift As Integer)
    EnableControls
End Sub

Private Sub TDBGridAssembly_AfterDelete()
    EnableControls
End Sub

Private Sub TDBGridAssembly_AfterInsert()
    bIsPendingChange = True
    cmdUpdate.Enabled = True
    EnableControls
End Sub

Private Sub TDBGridAssembly_AfterColUpdate(ByVal ColIndex As Integer)
    On Error Resume Next
    
    With TDBGridAssembly
        '
        '   If we're on a Quality Series building don't compute the cost per sf
        '   since the bldg_detail variables are just placeholders and the actuals are
        '   for the model that uses the assemblies from this bldg's model's assemblies.
        If m_rec.Fields("bldg_id").Value = "100" Or m_rec.Fields("bldg_id").Value = "200" _
        Or m_rec.Fields("bldg_id").Value = "300" Or m_rec.Fields("bldg_id").Value = "400" Then
        Else
            If .Columns(ColIndex).Caption = "Algorithm" Then
                If .Columns("Formula Factor").Value <> "" Then
                    'm_objModelAssembliesGridMap.ComputeValues Trim(lblStdSFArea.Caption), Trim(lblStdPerimeter.Caption)
                    If opttype_codeC.Value = True Then
                        m_objModelAssembliesGridMap.ComputeValues Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption), Trim(lblPerimeter(Right$(sshpSelectedArea, 1)).Caption)
                    Else
                        m_objModelAssembliesGridMap.ComputeValuesResi Trim(lblAreaResi(Right$(sshpSelectedArea, 1)).Caption), Trim(lblPerimeterResi(Right$(sshpSelectedArea, 1)).Caption)
                    End If
                End If
                bRefreshCosts = True
                
            ElseIf .Columns(ColIndex).Caption = "Formula Factor" Then
                If .Columns("Algorithm").Value <> "" Then
                    'm_objModelAssembliesGridMap.ComputeValues Trim(lblStdSFArea.Caption), Trim(lblStdPerimeter.Caption)
                    If opttype_codeC.Value = True Then
                        m_objModelAssembliesGridMap.ComputeValues Trim(lblArea(Right$(sshpSelectedArea, 1)).Caption), Trim(lblPerimeter(Right$(sshpSelectedArea, 1)).Caption)
                    Else
                        m_objModelAssembliesGridMap.ComputeValuesResi Trim(lblAreaResi(Right$(sshpSelectedArea, 1)).Caption), Trim(lblPerimeterResi(Right$(sshpSelectedArea, 1)).Caption)
                    End If
                End If
                bRefreshCosts = True
                
            ElseIf .Columns(ColIndex).Caption = "Assembly ID" Or .Columns(ColIndex).Caption = "Alt Assembly ID" Then
                bRefreshCosts = True
            End If
        End If
        bIsPendingChange = True
        cmdUpdate.Enabled = True
        EnableControls
    End With
End Sub

Private Sub TDBGridAssembly_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    With TDBGridAssembly
        .Columns("Algorithm").Value = UCase(Trim(.Columns("Algorithm").Value))
    End With
End Sub

Private Sub TDBGridAssembly_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If Not bIsInitialLoad Then
        With TDBGridAssembly
            If Not IsNull(.Bookmark) Then
                If opttype_codeC.Value = True Then
                    PopulateFormulaValues .Columns("Algorithm").Value
                Else
                    PopulateFormulaValuesResi .Columns("Algorithm").Value
                End If
            End If
        End With
    End If
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGridAssembly_UnboundAddData(ByVal RowBuf As TrueOleDBGrid70.RowBuffer, NewRowBookmark As Variant)
Private Sub TDBGridAssembly_UnboundAddData(ByVal RowBuf As TrueOleDBGrid80.RowBuffer, NewRowBookmark As Variant)
    lblAssemblyCompRowCount.Caption = TDBGridAssembly.ApproxCount + 1 & " rows."
    bRefreshCosts = True
End Sub

Private Sub TDBGridAssembly_UnboundDeleteRow(Bookmark As Variant)
    lblAssemblyCompRowCount.Caption = TDBGridAssembly.ApproxCount - 1 & " rows."
    bRefreshCosts = True
End Sub

Private Sub TDBGridSummaryEstimate_GotFocus()
    TDBGridSummaryEstimate.TabStop = True
End Sub

Private Sub TDBGridSummaryEstimate_LostFocus()
    TDBGridSummaryEstimate.TabStop = False
End Sub

Private Sub TDBGridSummaryEstimate_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGridSummaryEstimate.DataChanged = False
End Sub

Private Sub TDBGridSummaryEstimate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With TDBGridSummaryEstimate
        If Button = vbRightButton And IsNumeric(.Bookmark) Then
            If Len(m_objModelComponentGridMap.GetError(.Bookmark)) > 0 Then
                MsgBox m_objModelComponentGridMap.GetError(.Bookmark)
            End If
        End If
    End With
    'SetButtons USEBOOKMARK
End Sub

Private Sub TDBGridSummaryEstimate_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    With TDBGridSummaryEstimate
        If .Columns(ColIndex).Caption = "Unit" Then
            If Len(Trim(.Columns("Unit").Value)) > 15 Then
                MsgBox "Please provide a Unit that is less that 15 characters for System Component: " & Trim(.Columns("System Component").Value) & vbCrLf _
                    & "Only the 1st 15 characters will be kept.", vbInformation
                .Columns("Unit").Value = Left$(Trim(.Columns("Unit").Value), 15)
            End If
        ElseIf .Columns(ColIndex).Caption = "Specifications" Then
            If Len(Trim(.Columns("Specifications").Value)) > 300 Then
                MsgBox "Please provide a Specification that is less that 300 characters for System Component: " & Trim(.Columns("System Component").Value) & vbCrLf _
                    & "Only the 1st 300 characters will be kept.", vbInformation
                .Columns("Specifications").Value = Left$(Trim(.Columns("Specifications").Value), 300)
            End If
        End If
    End With
End Sub

Private Sub TDBGridSummaryEstimate_AfterColUpdate(ByVal ColIndex As Integer)
    bIsPendingChange = True
    cmdUpdate.Enabled = True
End Sub

Public Sub PrintReport()
    cmdReports_Click
End Sub

Public Sub PreviewReport()
    cmdReports_Click
End Sub

Public Sub ShowPrintToolbar(ByVal bShowIcons As Boolean)

    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
    fMainForm.mnuFilePageSetup.Enabled = bShowIcons
    fMainForm.mnuFilePrint.Enabled = bShowIcons
    fMainForm.mnuFilePrintPreview.Enabled = bShowIcons

End Sub

Private Sub txtCostWorksDesc_Change()
    On Error Resume Next
    If Not bIsInitialLoad Then
        EnableControls
        bIsPendingChange = True
        cmdUpdate.Enabled = True
    End If
    
End Sub
