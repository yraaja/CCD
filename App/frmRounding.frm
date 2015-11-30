VERSION 5.00
Begin VB.Form frmRounding 
   Caption         =   "Rounding"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMessage 
      Enabled         =   0   'False
      Height          =   735
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   114
      Top             =   6120
      Width           =   7000
   End
   Begin VB.CommandButton cmdCopyValues 
      Caption         =   "Copy Values"
      Height          =   495
      Left            =   8400
      TabIndex        =   113
      Top             =   6120
      Visible         =   0   'False
      Width           =   1630
   End
   Begin VB.TextBox res_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7560
      TabIndex        =   95
      Top             =   5610
      Width           =   795
   End
   Begin VB.TextBox res_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   94
      Top             =   5610
      Width           =   795
   End
   Begin VB.TextBox res_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   93
      Top             =   5610
      Width           =   795
   End
   Begin VB.TextBox res_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   92
      Top             =   5610
      Width           =   795
   End
   Begin VB.TextBox res_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   91
      Text            =   " "
      Top             =   5610
      Width           =   795
   End
   Begin VB.TextBox res_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   90
      Text            =   " "
      Top             =   5610
      Width           =   795
   End
   Begin VB.TextBox res_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   89
      Top             =   5610
      Width           =   795
   End
   Begin VB.TextBox res_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   88
      Top             =   5610
      Width           =   795
   End
   Begin VB.TextBox metric_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   87
      Top             =   5280
      Width           =   795
   End
   Begin VB.TextBox metric_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   86
      Top             =   5280
      Width           =   795
   End
   Begin VB.TextBox metric_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   85
      Top             =   5280
      Width           =   795
   End
   Begin VB.TextBox rr_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   84
      Top             =   4560
      Width           =   795
   End
   Begin VB.TextBox rr_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   83
      Top             =   4560
      Width           =   795
   End
   Begin VB.TextBox rr_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   82
      Top             =   4560
      Width           =   795
   End
   Begin VB.TextBox opn_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   81
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox opn_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   80
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox opn_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   79
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox std_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   78
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox std_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   77
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox std_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   76
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox metric_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7560
      TabIndex        =   75
      Top             =   5280
      Width           =   795
   End
   Begin VB.TextBox rr_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7560
      TabIndex        =   74
      Top             =   4560
      Width           =   795
   End
   Begin VB.TextBox opn_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7560
      TabIndex        =   73
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox std_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7560
      TabIndex        =   72
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox metric_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   71
      Top             =   5280
      Width           =   795
   End
   Begin VB.TextBox metric_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   70
      Top             =   5280
      Width           =   795
   End
   Begin VB.TextBox metric_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   69
      Top             =   5280
      Width           =   795
   End
   Begin VB.TextBox metric_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   68
      Top             =   5280
      Width           =   795
   End
   Begin VB.TextBox rr_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   67
      Top             =   4560
      Width           =   795
   End
   Begin VB.TextBox rr_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   66
      Top             =   4560
      Width           =   795
   End
   Begin VB.TextBox rr_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   65
      Top             =   4560
      Width           =   795
   End
   Begin VB.TextBox rr_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1305
      TabIndex        =   64
      Top             =   4560
      Width           =   795
   End
   Begin VB.TextBox opn_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   63
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox opn_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   62
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox opn_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   61
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox opn_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   60
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox std_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   59
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox std_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   58
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox std_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   57
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox std_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   56
      Top             =   4200
      Width           =   795
   End
   Begin VB.CommandButton cmdApplyRounding 
      Caption         =   "Apply Rounding"
      Height          =   495
      Left            =   8400
      TabIndex        =   55
      Top             =   2880
      Visible         =   0   'False
      Width           =   1630
   End
   Begin VB.TextBox std_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   39
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox std_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   38
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox std_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   37
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox std_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   36
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox opn_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   35
      Top             =   1800
      Width           =   795
   End
   Begin VB.TextBox opn_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   34
      Top             =   1800
      Width           =   795
   End
   Begin VB.TextBox opn_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   33
      Top             =   1800
      Width           =   795
   End
   Begin VB.TextBox opn_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   32
      Top             =   1800
      Width           =   795
   End
   Begin VB.TextBox rr_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1305
      TabIndex        =   31
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox rr_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   30
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox rr_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   29
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox rr_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   28
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox metric_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   27
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox metric_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   26
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox metric_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   25
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox metric_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   24
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox std_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   7560
      TabIndex        =   23
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox opn_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   7560
      TabIndex        =   22
      Top             =   1800
      Width           =   795
   End
   Begin VB.TextBox rr_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   7560
      TabIndex        =   21
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox metric_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   7560
      TabIndex        =   20
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox std_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   19
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox std_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   18
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox std_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   17
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox opn_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   16
      Top             =   1800
      Width           =   795
   End
   Begin VB.TextBox opn_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   15
      Top             =   1800
      Width           =   795
   End
   Begin VB.TextBox opn_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   14
      Top             =   1800
      Width           =   795
   End
   Begin VB.TextBox rr_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   13
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox rr_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   12
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox rr_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   11
      Top             =   1440
      Width           =   795
   End
   Begin VB.TextBox metric_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   10
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox metric_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   9
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox metric_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   8
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox res_mat_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   2490
      Width           =   795
   End
   Begin VB.TextBox res_labor_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   2490
      Width           =   795
   End
   Begin VB.TextBox res_equip_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Text            =   " "
      Top             =   2490
      Width           =   795
   End
   Begin VB.TextBox res_total_cost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   4
      Text            =   " "
      Top             =   2490
      Width           =   795
   End
   Begin VB.TextBox res_mat_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   2490
      Width           =   795
   End
   Begin VB.TextBox res_labor_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   2
      Top             =   2490
      Width           =   795
   End
   Begin VB.TextBox res_equip_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   1
      Top             =   2490
      Width           =   795
   End
   Begin VB.TextBox res_total_cost_op 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   7560
      TabIndex        =   0
      Top             =   2490
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Error Message"
      Height          =   615
      Left            =   480
      TabIndex        =   115
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Rounding Applied"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   112
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   111
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "Resi"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   110
      Top             =   5610
      Width           =   735
   End
   Begin VB.Label Label76 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
      Height          =   255
      Index           =   1
      Left            =   7680
      TabIndex        =   109
      Top             =   3840
      Width           =   555
   End
   Begin VB.Label Label75 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipment"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   108
      Top             =   3840
      Width           =   795
   End
   Begin VB.Label Label74 
      Alignment       =   1  'Right Justify
      Caption         =   "Labor"
      Height          =   255
      Index           =   1
      Left            =   5985
      TabIndex        =   107
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label73 
      Alignment       =   1  'Right Justify
      Caption         =   "Material"
      Height          =   255
      Index           =   1
      Left            =   5025
      TabIndex        =   106
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label72 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   105
      Top             =   3840
      Width           =   555
   End
   Begin VB.Label Label71 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipment"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   104
      Top             =   3840
      Width           =   795
   End
   Begin VB.Label Label70 
      Alignment       =   1  'Right Justify
      Caption         =   "Labor"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   103
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label69 
      Alignment       =   1  'Right Justify
      Caption         =   "Material"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   102
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      Caption         =   "Overhead && Profit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5625
      TabIndex        =   101
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label67 
      Caption         =   "Bare Costs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2385
      TabIndex        =   100
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label66 
      Caption         =   "Metric"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   99
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label65 
      Caption         =   "Open"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   98
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label64 
      Caption         =   "R&&R"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   97
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label63 
      Caption         =   "Standard"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   96
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label63 
      Caption         =   "Standard"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   54
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label64 
      Caption         =   "R&&R"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   53
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label65 
      Caption         =   "Open"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   52
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label66 
      Caption         =   "Metric"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   51
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label67 
      Caption         =   "Bare Costs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2385
      TabIndex        =   50
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      Caption         =   "Overhead && Profit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5625
      TabIndex        =   49
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label69 
      Alignment       =   1  'Right Justify
      Caption         =   "Material"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   48
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label70 
      Alignment       =   1  'Right Justify
      Caption         =   "Labor"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   47
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label71 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipment"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   46
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label72 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   45
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label73 
      Alignment       =   1  'Right Justify
      Caption         =   "Material"
      Height          =   255
      Index           =   0
      Left            =   5025
      TabIndex        =   44
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label74 
      Alignment       =   1  'Right Justify
      Caption         =   "Labor"
      Height          =   255
      Index           =   0
      Left            =   5985
      TabIndex        =   43
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label75 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipment"
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   42
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label76 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
      Height          =   255
      Index           =   0
      Left            =   7680
      TabIndex        =   41
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label17 
      Caption         =   "Resi"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   40
      Top             =   2490
      Width           =   735
   End
End
Attribute VB_Name = "frmRounding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public frmCallingForm As Form


'Private Function ApplyRoundingToRow1()
'
'    'using Tag Values "R" for rounding, "T" for retotalling (this is for the totals only)
'
'    'std_mat_cost = 0 and std_labor_cost = 0 and std_equip_cost = 0 and std_total_cost != 0, then round std_total_cost
'    If (((std_mat_cost(0).Text = "" Or Val(std_mat_cost(0).Text) = 0) And (std_labor_cost(0).Text = "" Or Val(std_labor_cost(0).Text) = 0) And (std_equip_cost(0).Text = "" Or Val(std_equip_cost(0).Text) = 0)) And (std_total_cost(0).Text <> "" And Val(std_total_cost(0).Text) <> 0)) Then
'        'call sp_rounding_routing 'U',std_total_cost.Text
'        std_total_cost(0).Tag = "R"
'    'everything is 0
'    ElseIf ((std_mat_cost(0).Text = "" Or Val(std_mat_cost(0).Text) = 0) And (std_labor_cost(0).Text = "" Or Val(std_labor_cost(0).Text) = 0) And (std_equip_cost(0).Text = "" Or Val(std_equip_cost(0).Text) = 0)) Then
'        std_total_cost(0).Tag = ""
'        std_total_cost(1).Text = std_total_cost(0).Text
'    Else
'
'        'retotal after rounding
'
'        std_total_cost(0).Tag = "T"
'
'        If (std_mat_cost(0).Text <> "" And Val(std_mat_cost(0).Text) <> 0) Then
'            std_mat_cost(0).Tag = "R"
'        End If
'        If (std_labor_cost(0).Text <> "" And Val(std_labor_cost(0).Text) <> 0) Then
'            std_labor_cost(0).Tag = "R"
'        End If
'        If (std_equip_cost(0).Text <> "" And Val(std_equip_cost(0).Text) <> 0) Then
'            std_equip_cost(0).Tag = "R"
'        End If
'    End If
'
'
'


'
'
'    'get the rounded values by executing RoundingFromStoredProc (only for those that have tag = "R"
'    RoundingFromStoredProc
'    'then retotal those that require retotalling
'
'
''    'res_mat_cost_op = 0 and res_labor_cost_op = 0 and res_equip_cost_op = 0 and res_total_cost_op != 0, then round res_total_cost_op
''    If (((res_mat_cost_op(0).Text = "" Or Val(res_mat_cost_op(0).Text) = 0) And (res_labor_cost_op(0).Text = "" Or Val(res_labor_cost_op(0).Text) = 0) And (res_equip_cost_op(0).Text = "" Or Val(res_equip_cost_op(0).Text) = 0)) And (res_total_cost_op(0).Text <> "" And Val(res_total_cost_op(0).Text) <> 0)) Then
''        'call sp_rounding_routing 'U',res_total_cost_op.Text
''        retRounding = Rounding(res_total_cost_op(0).Text)
''        If (retRounding <> "-1") Then
''            res_total_cost_op(1).Text = Format(retRounding, ReplaceCharactersForFormat(res_total_cost_op(0).Text))
''        Else
''            errMessage = errMessage + "There was a problem rounding Overhead & Profit - Resi Total Cost" + vbCrLf
''        End If
''    End If
'
'    txtMessage.Text = errMessage
'
'    RoundingFromStoredProc
'

'End Function

'
'Private Sub ApplyRetotaling(ByRef total As TextBox, ByRef mat2 As TextBox, ByRef labor2 As TextBox, ByRef equip2 As TextBox, ByRef total2 As TextBox)
'    Dim retotal As Double
'    If (total.Tag = "T") Then
'        If (mat2.Text <> "" And Val(mat2.Text) <> 0) Then
'            retotal = retotal + mat2
'        End If
'        If (labor2.Text <> "" And Val(labor2.Text) <> 0) Then
'            retotal = retotal + labor2
'        End If
'        If (equip2.Text <> "" And Val(equip2.Text) <> 0) Then
'            retotal = retotal + equip2
'        End If
'
'        total2.Text = Format(retotal, ReplaceCharactersForFormat(total.Text))
'
'    End If
'End Sub


'Private Sub ApplyRoundingRule(ByRef mat As TextBox, ByRef labor As TextBox, ByRef equip As TextBox, ByRef total As TextBox, ByRef total2 As TextBox)
'
'
'    'using Tag Values "R" for rounding, "T" for retotalling (this is for the totals only)
'
'    'std_mat_cost = 0 and std_labor_cost = 0 and std_equip_cost = 0 and std_total_cost != 0, then round std_total_cost
'    If (((mat.Text = "" Or Val(mat.Text) = 0) And (labor.Text = "" Or Val(labor.Text) = 0) And (equip.Text = "" Or Val(equip.Text) = 0)) And (total.Text <> "" And Val(total.Text) <> 0)) Then
'        'call sp_rounding_routing 'U',std_total_cost.Text
'        total.Tag = "R"
'    'everything is 0
'    ElseIf ((mat.Text = "" Or Val(mat.Text) = 0) And (labor.Text = "" Or Val(labor.Text) = 0) And (equip.Text = "" Or Val(equip.Text) = 0)) Then
'        total.Tag = ""
'        total2.Text = total.Text
'    Else
'
'        'retotal after rounding
'
'        total.Tag = "T"
'
'        If (mat.Text <> "" And Val(mat.Text) <> 0) Then
'            mat.Tag = "R"
'        End If
'        If (labor.Text <> "" And Val(labor.Text) <> 0) Then
'            labor.Tag = "R"
'        End If
'        If (equip.Text <> "" And Val(equip.Text) <> 0) Then
'            equip.Tag = "R"
'        End If
'    End If
'
'
'
'End Sub
'



'Function RoundingFromStoredProc(ByVal std_mat_cost as Double,) As String
'
'
'
'    'pass info into stored proc and update hier tree
'
'    'On Error GoTo ErrHandler
'
'    Dim Cmd As New ADODB.Command
'
'    Cmd.CommandText = "sp_round_cost_values"
'    Cmd.CommandType = CommandTypeEnum.adCmdStoredProc
'
'    If Trim(std_mat_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("std_mat_cost", adCurrency, adParamInputOutput, , std_mat_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(std_labor_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("std_labor_cost", adCurrency, adParamInputOutput, , std_labor_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(std_equip_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("std_equip_cost", adCurrency, adParamInputOutput, , std_equip_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(std_total_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("std_total_cost", adCurrency, adParamInputOutput, , std_total_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(rr_mat_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("rr_mat_cost", adCurrency, adParamInputOutput, , rr_mat_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(rr_labor_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("rr_labor_cost", adCurrency, adParamInputOutput, , rr_labor_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(rr_equip_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("rr_equip_cost", adCurrency, adParamInputOutput, , rr_equip_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(rr_total_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("rr_total_cost", adCurrency, adParamInputOutput, , rr_total_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(opn_mat_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("opn_mat_cost", adCurrency, adParamInputOutput, , opn_mat_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(opn_labor_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("opn_labor_cost", adCurrency, adParamInputOutput, , opn_labor_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(opn_equip_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("opn_equip_cost", adCurrency, adParamInputOutput, , opn_equip_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(opn_total_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("opn_total_cost", adCurrency, adParamInputOutput, , opn_total_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(metric_mat_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("metric_mat_cost", adCurrency, adParamInputOutput, , metric_mat_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(metric_labor_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("metric_labor_cost", adCurrency, adParamInputOutput, , metric_labor_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(metric_equip_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("metric_equip_cost", adCurrency, adParamInputOutput, , metric_equip_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(metric_total_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("metric_total_cost", adCurrency, adParamInputOutput, , metric_total_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(res_mat_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("res_mat_cost", adCurrency, adParamInputOutput, , res_mat_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(res_labor_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("res_labor_cost", adCurrency, adParamInputOutput, , res_labor_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(res_equip_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("res_equip_cost", adCurrency, adParamInputOutput, , res_equip_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(res_total_cost(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("res_total_cost", adCurrency, adParamInputOutput, , res_total_cost(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(std_mat_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("std_mat_cost_op", adCurrency, adParamInputOutput, , std_mat_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(std_labor_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("std_labor_cost_op", adCurrency, adParamInputOutput, , std_labor_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(std_equip_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("std_equip_cost_op", adCurrency, adParamInputOutput, , std_equip_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(std_total_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("std_total_cost_op", adCurrency, adParamInputOutput, , std_total_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(rr_mat_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("rr_mat_cost_op", adCurrency, adParamInputOutput, , rr_mat_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(rr_labor_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("rr_labor_cost_op", adCurrency, adParamInputOutput, , rr_labor_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(rr_equip_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("rr_equip_cost_op", adCurrency, adParamInputOutput, , rr_equip_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(rr_total_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("rr_total_cost_op", adCurrency, adParamInputOutput, , rr_total_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(opn_mat_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("opn_mat_cost_op", adCurrency, adParamInputOutput, , opn_mat_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(opn_labor_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("opn_labor_cost_op", adCurrency, adParamInputOutput, , opn_labor_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(opn_equip_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("opn_equip_cost_op", adCurrency, adParamInputOutput, , opn_equip_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(opn_total_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("opn_total_cost_op", adCurrency, adParamInputOutput, , opn_total_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(metric_mat_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("metric_mat_cost_op", adCurrency, adParamInputOutput, , metric_mat_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(metric_labor_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("metric_labor_cost_op", adCurrency, adParamInputOutput, , metric_labor_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(metric_equip_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("metric_equip_cost_op", adCurrency, adParamInputOutput, , metric_equip_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(metric_total_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("metric_total_cost_op", adCurrency, adParamInputOutput, , metric_total_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(res_mat_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("res_mat_cost_op", adCurrency, adParamInputOutput, , res_mat_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(res_labor_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("res_labor_cost_op", adCurrency, adParamInputOutput, , res_labor_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(res_equip_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("res_equip_cost_op", adCurrency, adParamInputOutput, , res_equip_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'    If Trim(res_total_cost_op(0).Tag) = "R" Then
'    Set Param = Cmd.CreateParameter("res_total_cost_op", adCurrency, adParamInputOutput, , res_total_cost_op(0).Text)
'    Cmd.Parameters.Append Param
'    End If
'
'
'    ' Assuming a connection has been established and a recordset has
'    ' been created
'    Set Cmd.ActiveConnection = g_cnShared
'    Cmd.Execute
'
'
'    If Trim(std_mat_cost(0).Tag) = "R" Then
'    std_mat_cost(1).Text = Format(Cmd("std_mat_cost").Value, ReplaceCharactersForFormat(std_mat_cost(0).Text))
'    Else
'    std_mat_cost(1).Text = std_mat_cost(0).Text
'    End If
'    If Trim(std_labor_cost(0).Tag) = "R" Then
'    std_labor_cost(1).Text = Format(Cmd("std_labor_cost").Value, ReplaceCharactersForFormat(std_labor_cost(0).Text))
'    Else
'    std_labor_cost(1).Text = std_labor_cost(0).Text
'    End If
'    If Trim(std_equip_cost(0).Tag) = "R" Then
'    std_equip_cost(1).Text = Format(Cmd("std_equip_cost").Value, ReplaceCharactersForFormat(std_equip_cost(0).Text))
'    Else
'    std_equip_cost(1).Text = std_equip_cost(0).Text
'    End If
'    If Trim(std_total_cost(0).Tag) = "R" Then
'    std_total_cost(1).Text = Format(Cmd("std_total_cost").Value, ReplaceCharactersForFormat(std_total_cost(0).Text))
'    Else
'    std_total_cost(1).Text = std_total_cost(0).Text
'    End If
'    If Trim(rr_mat_cost(0).Tag) = "R" Then
'    rr_mat_cost(1).Text = Format(Cmd("rr_mat_cost").Value, ReplaceCharactersForFormat(rr_mat_cost(0).Text))
'    Else
'    rr_mat_cost(1).Text = rr_mat_cost(0).Text
'    End If
'    If Trim(rr_labor_cost(0).Tag) = "R" Then
'    rr_labor_cost(1).Text = Format(Cmd("rr_labor_cost").Value, ReplaceCharactersForFormat(rr_labor_cost(0).Text))
'    Else
'    rr_labor_cost(1).Text = rr_labor_cost(0).Text
'    End If
'    If Trim(rr_equip_cost(0).Tag) = "R" Then
'    rr_equip_cost(1).Text = Format(Cmd("rr_equip_cost").Value, ReplaceCharactersForFormat(rr_equip_cost(0).Text))
'    Else
'    rr_equip_cost(1).Text = rr_equip_cost(0).Text
'    End If
'    If Trim(rr_total_cost(0).Tag) = "R" Then
'    rr_total_cost(1).Text = Format(Cmd("rr_total_cost").Value, ReplaceCharactersForFormat(rr_total_cost(0).Text))
'    Else
'    rr_total_cost(1).Text = rr_total_cost(0).Text
'    End If
'    If Trim(opn_mat_cost(0).Tag) = "R" Then
'    opn_mat_cost(1).Text = Format(Cmd("opn_mat_cost").Value, ReplaceCharactersForFormat(opn_mat_cost(0).Text))
'    Else
'    opn_mat_cost(1).Text = opn_mat_cost(0).Text
'    End If
'    If Trim(opn_labor_cost(0).Tag) = "R" Then
'    opn_labor_cost(1).Text = Format(Cmd("opn_labor_cost").Value, ReplaceCharactersForFormat(opn_labor_cost(0).Text))
'    Else
'    opn_labor_cost(1).Text = opn_labor_cost(0).Text
'    End If
'    If Trim(opn_equip_cost(0).Tag) = "R" Then
'    opn_equip_cost(1).Text = Format(Cmd("opn_equip_cost").Value, ReplaceCharactersForFormat(opn_equip_cost(0).Text))
'    Else
'    opn_equip_cost(1).Text = opn_equip_cost(0).Text
'    End If
'    If Trim(opn_total_cost(0).Tag) = "R" Then
'    opn_total_cost(1).Text = Format(Cmd("opn_total_cost").Value, ReplaceCharactersForFormat(opn_total_cost(0).Text))
'    Else
'    opn_total_cost(1).Text = opn_total_cost(0).Text
'    End If
'    If Trim(metric_mat_cost(0).Tag) = "R" Then
'    metric_mat_cost(1).Text = Format(Cmd("metric_mat_cost").Value, ReplaceCharactersForFormat(metric_mat_cost(0).Text))
'    Else
'    metric_mat_cost(1).Text = metric_mat_cost(0).Text
'    End If
'    If Trim(metric_labor_cost(0).Tag) = "R" Then
'    metric_labor_cost(1).Text = Format(Cmd("metric_labor_cost").Value, ReplaceCharactersForFormat(metric_labor_cost(0).Text))
'    Else
'    metric_labor_cost(1).Text = metric_labor_cost(0).Text
'    End If
'    If Trim(metric_equip_cost(0).Tag) = "R" Then
'    metric_equip_cost(1).Text = Format(Cmd("metric_equip_cost").Value, ReplaceCharactersForFormat(metric_equip_cost(0).Text))
'    Else
'    metric_equip_cost(1).Text = metric_equip_cost(0).Text
'    End If
'    If Trim(metric_total_cost(0).Tag) = "R" Then
'    metric_total_cost(1).Text = Format(Cmd("metric_total_cost").Value, ReplaceCharactersForFormat(metric_total_cost(0).Text))
'    Else
'    metric_total_cost(1).Text = metric_total_cost(0).Text
'    End If
'    If Trim(res_mat_cost(0).Tag) = "R" Then
'    res_mat_cost(1).Text = Format(Cmd("res_mat_cost").Value, ReplaceCharactersForFormat(res_mat_cost(0).Text))
'    Else
'    res_mat_cost(1).Text = res_mat_cost(0).Text
'    End If
'    If Trim(res_labor_cost(0).Tag) = "R" Then
'    res_labor_cost(1).Text = Format(Cmd("res_labor_cost").Value, ReplaceCharactersForFormat(res_labor_cost(0).Text))
'    Else
'    res_labor_cost(1).Text = res_labor_cost(0).Text
'    End If
'    If Trim(res_equip_cost(0).Tag) = "R" Then
'    res_equip_cost(1).Text = Format(Cmd("res_equip_cost").Value, ReplaceCharactersForFormat(res_equip_cost(0).Text))
'    Else
'    res_equip_cost(1).Text = res_equip_cost(0).Text
'    End If
'    If Trim(res_total_cost(0).Tag) = "R" Then
'    res_total_cost(1).Text = Format(Cmd("res_total_cost").Value, ReplaceCharactersForFormat(res_total_cost(0).Text))
'    Else
'    res_total_cost(1).Text = res_total_cost(0).Text
'    End If
'    If Trim(std_mat_cost_op(0).Tag) = "R" Then
'    std_mat_cost_op(1).Text = Format(Cmd("std_mat_cost_op").Value, ReplaceCharactersForFormat(std_mat_cost_op(0).Text))
'    Else
'    std_mat_cost_op(1).Text = std_mat_cost_op(0).Text
'    End If
'    If Trim(std_labor_cost_op(0).Tag) = "R" Then
'    std_labor_cost_op(1).Text = Format(Cmd("std_labor_cost_op").Value, ReplaceCharactersForFormat(std_labor_cost_op(0).Text))
'    Else
'    std_labor_cost_op(1).Text = std_labor_cost_op(0).Text
'    End If
'    If Trim(std_equip_cost_op(0).Tag) = "R" Then
'    std_equip_cost_op(1).Text = Format(Cmd("std_equip_cost_op").Value, ReplaceCharactersForFormat(std_equip_cost_op(0).Text))
'    Else
'    std_equip_cost_op(1).Text = std_equip_cost_op(0).Text
'    End If
'    If Trim(std_total_cost_op(0).Tag) = "R" Then
'    std_total_cost_op(1).Text = Format(Cmd("std_total_cost_op").Value, ReplaceCharactersForFormat(std_total_cost_op(0).Text))
'    Else
'    std_total_cost_op(1).Text = std_total_cost_op(0).Text
'    End If
'    If Trim(rr_mat_cost_op(0).Tag) = "R" Then
'    rr_mat_cost_op(1).Text = Format(Cmd("rr_mat_cost_op").Value, ReplaceCharactersForFormat(rr_mat_cost_op(0).Text))
'    Else
'    rr_mat_cost_op(1).Text = rr_mat_cost_op(0).Text
'    End If
'    If Trim(rr_labor_cost_op(0).Tag) = "R" Then
'    rr_labor_cost_op(1).Text = Format(Cmd("rr_labor_cost_op").Value, ReplaceCharactersForFormat(rr_labor_cost_op(0).Text))
'    Else
'    rr_labor_cost_op(1).Text = rr_labor_cost_op(0).Text
'    End If
'    If Trim(rr_equip_cost_op(0).Tag) = "R" Then
'    rr_equip_cost_op(1).Text = Format(Cmd("rr_equip_cost_op").Value, ReplaceCharactersForFormat(rr_equip_cost_op(0).Text))
'    Else
'    rr_equip_cost_op(1).Text = rr_equip_cost_op(0).Text
'    End If
'    If Trim(rr_total_cost_op(0).Tag) = "R" Then
'    rr_total_cost_op(1).Text = Format(Cmd("rr_total_cost_op").Value, ReplaceCharactersForFormat(rr_total_cost_op(0).Text))
'    Else
'    rr_total_cost_op(1).Text = rr_total_cost_op(0).Text
'    End If
'    If Trim(opn_mat_cost_op(0).Tag) = "R" Then
'    opn_mat_cost_op(1).Text = Format(Cmd("opn_mat_cost_op").Value, ReplaceCharactersForFormat(opn_mat_cost_op(0).Text))
'    Else
'    opn_mat_cost_op(1).Text = opn_mat_cost_op(0).Text
'    End If
'    If Trim(opn_labor_cost_op(0).Tag) = "R" Then
'    opn_labor_cost_op(1).Text = Format(Cmd("opn_labor_cost_op").Value, ReplaceCharactersForFormat(opn_labor_cost_op(0).Text))
'    Else
'    opn_labor_cost_op(1).Text = opn_labor_cost_op(0).Text
'    End If
'    If Trim(opn_equip_cost_op(0).Tag) = "R" Then
'    opn_equip_cost_op(1).Text = Format(Cmd("opn_equip_cost_op").Value, ReplaceCharactersForFormat(opn_equip_cost_op(0).Text))
'    Else
'    opn_equip_cost_op(1).Text = opn_equip_cost_op(0).Text
'    End If
'    If Trim(opn_total_cost_op(0).Tag) = "R" Then
'    opn_total_cost_op(1).Text = Format(Cmd("opn_total_cost_op").Value, ReplaceCharactersForFormat(opn_total_cost_op(0).Text))
'    Else
'    opn_total_cost_op(1).Text = opn_total_cost_op(0).Text
'    End If
'    If Trim(metric_mat_cost_op(0).Tag) = "R" Then
'    metric_mat_cost_op(1).Text = Format(Cmd("metric_mat_cost_op").Value, ReplaceCharactersForFormat(metric_mat_cost_op(0).Text))
'    Else
'    metric_mat_cost_op(1).Text = metric_mat_cost_op(0).Text
'    End If
'    If Trim(metric_labor_cost_op(0).Tag) = "R" Then
'    metric_labor_cost_op(1).Text = Format(Cmd("metric_labor_cost_op").Value, ReplaceCharactersForFormat(metric_labor_cost_op(0).Text))
'    Else
'    metric_labor_cost_op(1).Text = metric_labor_cost_op(0).Text
'    End If
'    If Trim(metric_equip_cost_op(0).Tag) = "R" Then
'    metric_equip_cost_op(1).Text = Format(Cmd("metric_equip_cost_op").Value, ReplaceCharactersForFormat(metric_equip_cost_op(0).Text))
'    Else
'    metric_equip_cost_op(1).Text = metric_equip_cost_op(0).Text
'    End If
'    If Trim(metric_total_cost_op(0).Tag) = "R" Then
'    metric_total_cost_op(1).Text = Format(Cmd("metric_total_cost_op").Value, ReplaceCharactersForFormat(metric_total_cost_op(0).Text))
'    Else
'    metric_total_cost_op(1).Text = metric_total_cost_op(0).Text
'    End If
'    If Trim(res_mat_cost_op(0).Tag) = "R" Then
'    res_mat_cost_op(1).Text = Format(Cmd("res_mat_cost_op").Value, ReplaceCharactersForFormat(res_mat_cost_op(0).Text))
'    Else
'    res_mat_cost_op(1).Text = res_mat_cost_op(0).Text
'    End If
'    If Trim(res_labor_cost_op(0).Tag) = "R" Then
'    res_labor_cost_op(1).Text = Format(Cmd("res_labor_cost_op").Value, ReplaceCharactersForFormat(res_labor_cost_op(0).Text))
'    Else
'    res_labor_cost_op(1).Text = res_labor_cost_op(0).Text
'    End If
'    If Trim(res_equip_cost_op(0).Tag) = "R" Then
'    res_equip_cost_op(1).Text = Format(Cmd("res_equip_cost_op").Value, ReplaceCharactersForFormat(res_equip_cost_op(0).Text))
'    Else
'    res_equip_cost_op(1).Text = res_equip_cost_op(0).Text
'    End If
'    If Trim(res_total_cost_op(0).Tag) = "R" Then
'    res_total_cost_op(1).Text = Format(Cmd("res_total_cost_op").Value, ReplaceCharactersForFormat(res_total_cost_op(0).Text))
'    Else
'    res_total_cost_op(1).Text = res_total_cost_op(0).Text
'    End If
'
'
'
'RoundingFromStoredProc = ""
'
'    Exit Function
'
'ErrHandler:
'    RoundingFromStoredProc = Err.Description
'
'End Function

Private Sub cmdApplyRounding_Click()

    Dim errMessage As String
    Dim retRounding As String
    errMessage = ""

    
    Dim retStr As String
    
    Dim dbl_std_mat_cost As Double
    Dim dbl_std_labor_cost As Double
    Dim dbl_std_equip_cost As Double
    Dim dbl_std_total_cost As Double
    Dim dbl_rr_mat_cost As Double
    Dim dbl_rr_labor_cost As Double
    Dim dbl_rr_equip_cost As Double
    Dim dbl_rr_total_cost As Double
    Dim dbl_opn_mat_cost As Double
    Dim dbl_opn_labor_cost As Double
    Dim dbl_opn_equip_cost As Double
    Dim dbl_opn_total_cost As Double
    Dim dbl_metric_mat_cost As Double
    Dim dbl_metric_labor_cost As Double
    Dim dbl_metric_equip_cost As Double
    Dim dbl_metric_total_cost As Double
    Dim dbl_res_mat_cost As Double
    Dim dbl_res_labor_cost As Double
    Dim dbl_res_equip_cost As Double
    Dim dbl_res_total_cost As Double
    Dim dbl_std_mat_cost_op As Double
    Dim dbl_std_labor_cost_op As Double
    Dim dbl_std_equip_cost_op As Double
    Dim dbl_std_total_cost_op As Double
    Dim dbl_rr_mat_cost_op As Double
    Dim dbl_rr_labor_cost_op As Double
    Dim dbl_rr_equip_cost_op As Double
    Dim dbl_rr_total_cost_op As Double
    Dim dbl_opn_mat_cost_op As Double
    Dim dbl_opn_labor_cost_op As Double
    Dim dbl_opn_equip_cost_op As Double
    Dim dbl_opn_total_cost_op As Double
    Dim dbl_metric_mat_cost_op As Double
    Dim dbl_metric_labor_cost_op As Double
    Dim dbl_metric_equip_cost_op As Double
    Dim dbl_metric_total_cost_op As Double
    Dim dbl_res_mat_cost_op As Double
    Dim dbl_res_labor_cost_op As Double
    Dim dbl_res_equip_cost_op As Double
    Dim dbl_res_total_cost_op As Double
        
    dbl_std_mat_cost = Val(std_mat_cost(0).Text)
    dbl_std_labor_cost = Val(std_labor_cost(0).Text)
    dbl_std_equip_cost = Val(std_equip_cost(0).Text)
    dbl_std_total_cost = Val(std_total_cost(0).Text)
    dbl_rr_mat_cost = Val(rr_mat_cost(0).Text)
    dbl_rr_labor_cost = Val(rr_labor_cost(0).Text)
    dbl_rr_equip_cost = Val(rr_equip_cost(0).Text)
    dbl_rr_total_cost = Val(rr_total_cost(0).Text)
    dbl_opn_mat_cost = Val(opn_mat_cost(0).Text)
    dbl_opn_labor_cost = Val(opn_labor_cost(0).Text)
    dbl_opn_equip_cost = Val(opn_equip_cost(0).Text)
    dbl_opn_total_cost = Val(opn_total_cost(0).Text)
    dbl_metric_mat_cost = Val(metric_mat_cost(0).Text)
    dbl_metric_labor_cost = Val(metric_labor_cost(0).Text)
    dbl_metric_equip_cost = Val(metric_equip_cost(0).Text)
    dbl_metric_total_cost = Val(metric_total_cost(0).Text)
    dbl_res_mat_cost = Val(res_mat_cost(0).Text)
    dbl_res_labor_cost = Val(res_labor_cost(0).Text)
    dbl_res_equip_cost = Val(res_equip_cost(0).Text)
    dbl_res_total_cost = Val(res_total_cost(0).Text)
    dbl_std_mat_cost_op = Val(std_mat_cost_op(0).Text)
    dbl_std_labor_cost_op = Val(std_labor_cost_op(0).Text)
    dbl_std_equip_cost_op = Val(std_equip_cost_op(0).Text)
    dbl_std_total_cost_op = Val(std_total_cost_op(0).Text)
    dbl_rr_mat_cost_op = Val(rr_mat_cost_op(0).Text)
    dbl_rr_labor_cost_op = Val(rr_labor_cost_op(0).Text)
    dbl_rr_equip_cost_op = Val(rr_equip_cost_op(0).Text)
    dbl_rr_total_cost_op = Val(rr_total_cost_op(0).Text)
    dbl_opn_mat_cost_op = Val(opn_mat_cost_op(0).Text)
    dbl_opn_labor_cost_op = Val(opn_labor_cost_op(0).Text)
    dbl_opn_equip_cost_op = Val(opn_equip_cost_op(0).Text)
    dbl_opn_total_cost_op = Val(opn_total_cost_op(0).Text)
    dbl_metric_mat_cost_op = Val(metric_mat_cost_op(0).Text)
    dbl_metric_labor_cost_op = Val(metric_labor_cost_op(0).Text)
    dbl_metric_equip_cost_op = Val(metric_equip_cost_op(0).Text)
    dbl_metric_total_cost_op = Val(metric_total_cost_op(0).Text)
    dbl_res_mat_cost_op = Val(res_mat_cost_op(0).Text)
    dbl_res_labor_cost_op = Val(res_labor_cost_op(0).Text)
    dbl_res_equip_cost_op = Val(res_equip_cost_op(0).Text)
    dbl_res_total_cost_op = Val(res_total_cost_op(0).Text)
        
 
    
        
    
    retStr = CostRoundingFromStoredProc(dbl_std_mat_cost, dbl_std_labor_cost, dbl_std_equip_cost, dbl_std_total_cost, _
        dbl_rr_mat_cost, dbl_rr_labor_cost, dbl_rr_equip_cost, dbl_rr_total_cost, _
        dbl_opn_mat_cost, dbl_opn_labor_cost, dbl_opn_equip_cost, dbl_opn_total_cost, _
        dbl_metric_mat_cost, dbl_metric_labor_cost, dbl_metric_equip_cost, dbl_metric_total_cost, _
        dbl_res_mat_cost, dbl_res_labor_cost, dbl_res_equip_cost, dbl_res_total_cost, _
        dbl_std_mat_cost_op, dbl_std_labor_cost_op, dbl_std_equip_cost_op, dbl_std_total_cost_op, _
        dbl_rr_mat_cost_op, dbl_rr_labor_cost_op, dbl_rr_equip_cost_op, dbl_rr_total_cost_op, _
        dbl_opn_mat_cost_op, dbl_opn_labor_cost_op, dbl_opn_equip_cost_op, dbl_opn_total_cost_op, _
        dbl_metric_mat_cost_op, dbl_metric_labor_cost_op, dbl_metric_equip_cost_op, dbl_metric_total_cost_op, _
        dbl_res_mat_cost_op, dbl_res_labor_cost_op, dbl_res_equip_cost_op, dbl_res_total_cost_op)
    
    
    If retStr = "" Then
        
        
        If Trim(std_mat_cost(0).Text) <> "" And Val(std_mat_cost(0).Text) <> 0 Then
        std_mat_cost(1).Text = Format(dbl_std_mat_cost, ReplaceCharactersForFormat(std_mat_cost(0).Text))
        Else
        std_mat_cost(1).Text = std_mat_cost(0).Text
        End If
        If Trim(std_labor_cost(0).Text) <> "" And Val(std_labor_cost(0).Text) <> 0 Then
        std_labor_cost(1).Text = Format(dbl_std_labor_cost, ReplaceCharactersForFormat(std_labor_cost(0).Text))
        Else
        std_labor_cost(1).Text = std_labor_cost(0).Text
        End If
        If Trim(std_equip_cost(0).Text) <> "" And Val(std_equip_cost(0).Text) <> 0 Then
        std_equip_cost(1).Text = Format(dbl_std_equip_cost, ReplaceCharactersForFormat(std_equip_cost(0).Text))
        Else
        std_equip_cost(1).Text = std_equip_cost(0).Text
        End If
        If (Trim(std_total_cost(0).Text) <> "" And Val(std_total_cost(0).Text) <> 0) Or (dbl_std_mat_cost <> 0 Or dbl_std_labor_cost <> 0 Or dbl_std_equip_cost <> 0) Then
            'If Trim(std_total_cost(0).Text <> "") Then
            If (dbl_std_total_cost <> 0) Then
'                std_total_cost(1).Text = Format(dbl_std_total_cost, ReplaceCharactersForFormat(std_total_cost(0).Text))
'            Else
                std_total_cost(1).Text = Format(dbl_std_total_cost, ReplaceCharactersForFormat(CStr(dbl_std_total_cost)))
            End If
        Else
        std_total_cost(1).Text = std_total_cost(0).Text
        End If
        If Trim(rr_mat_cost(0).Text) <> "" And Val(rr_mat_cost(0).Text) <> 0 Then
        rr_mat_cost(1).Text = Format(dbl_rr_mat_cost, ReplaceCharactersForFormat(rr_mat_cost(0).Text))
        Else
        rr_mat_cost(1).Text = rr_mat_cost(0).Text
        End If
        If Trim(rr_labor_cost(0).Text) <> "" And Val(rr_labor_cost(0).Text) <> 0 Then
        rr_labor_cost(1).Text = Format(dbl_rr_labor_cost, ReplaceCharactersForFormat(rr_labor_cost(0).Text))
        Else
        rr_labor_cost(1).Text = rr_labor_cost(0).Text
        End If
        If Trim(rr_equip_cost(0).Text) <> "" And Val(rr_equip_cost(0).Text) <> 0 Then
        rr_equip_cost(1).Text = Format(dbl_rr_equip_cost, ReplaceCharactersForFormat(rr_equip_cost(0).Text))
        Else
        rr_equip_cost(1).Text = rr_equip_cost(0).Text
        End If
        If (Trim(rr_total_cost(0).Text) <> "" And Val(rr_total_cost(0).Text) <> 0) Or (dbl_rr_mat_cost <> 0 Or dbl_rr_labor_cost <> 0 Or dbl_rr_equip_cost <> 0) Then
            'If Trim(rr_total_cost(0).Text) <> "" Then
            If (dbl_rr_total_cost <> 0) Then
'                rr_total_cost(1).Text = Format(dbl_rr_total_cost, ReplaceCharactersForFormat(rr_total_cost(0).Text))
'            Else
                rr_total_cost(1).Text = Format(dbl_rr_total_cost, ReplaceCharactersForFormat(CStr(dbl_rr_total_cost)))
            End If
        Else
        rr_total_cost(1).Text = rr_total_cost(0).Text
        End If
        If Trim(opn_mat_cost(0).Text) <> "" And Val(opn_mat_cost(0).Text) <> 0 Then
        opn_mat_cost(1).Text = Format(dbl_opn_mat_cost, ReplaceCharactersForFormat(opn_mat_cost(0).Text))
        Else
        opn_mat_cost(1).Text = opn_mat_cost(0).Text
        End If
        If Trim(opn_labor_cost(0).Text) <> "" And Val(opn_labor_cost(0).Text) <> 0 Then
        opn_labor_cost(1).Text = Format(dbl_opn_labor_cost, ReplaceCharactersForFormat(opn_labor_cost(0).Text))
        Else
        opn_labor_cost(1).Text = opn_labor_cost(0).Text
        End If
        If Trim(opn_equip_cost(0).Text) <> "" And Val(opn_equip_cost(0).Text) <> 0 Then
        opn_equip_cost(1).Text = Format(dbl_opn_equip_cost, ReplaceCharactersForFormat(opn_equip_cost(0).Text))
        Else
        opn_equip_cost(1).Text = opn_equip_cost(0).Text
        End If
        If (Trim(opn_total_cost(0).Text) <> "" And Val(opn_total_cost(0).Text) <> 0) Or (dbl_opn_mat_cost <> 0 Or dbl_opn_labor_cost <> 0 Or dbl_opn_equip_cost <> 0) Then
            'If Trim(opn_total_cost(0).Text) <> "" Then
            If dbl_opn_total_cost <> 0 Then
'                opn_total_cost(1).Text = Format(dbl_opn_total_cost, ReplaceCharactersForFormat(opn_total_cost(0).Text))
'            Else
                opn_total_cost(1).Text = Format(dbl_opn_total_cost, ReplaceCharactersForFormat(CStr(dbl_opn_total_cost)))
            End If
        Else
        opn_total_cost(1).Text = opn_total_cost(0).Text
        End If
        If Trim(metric_mat_cost(0).Text) <> "" And Val(metric_mat_cost(0).Text) <> 0 Then
        metric_mat_cost(1).Text = Format(dbl_metric_mat_cost, ReplaceCharactersForFormat(metric_mat_cost(0).Text))
        Else
        metric_mat_cost(1).Text = metric_mat_cost(0).Text
        End If
        If Trim(metric_labor_cost(0).Text) <> "" And Val(metric_labor_cost(0).Text) <> 0 Then
        metric_labor_cost(1).Text = Format(dbl_metric_labor_cost, ReplaceCharactersForFormat(metric_labor_cost(0).Text))
        Else
        metric_labor_cost(1).Text = metric_labor_cost(0).Text
        End If
        If Trim(metric_equip_cost(0).Text) <> "" And Val(metric_equip_cost(0).Text) <> 0 Then
        metric_equip_cost(1).Text = Format(dbl_metric_equip_cost, ReplaceCharactersForFormat(metric_equip_cost(0).Text))
        Else
        metric_equip_cost(1).Text = metric_equip_cost(0).Text
        End If
        If (Trim(metric_total_cost(0).Text) <> "" And Val(metric_total_cost(0).Text) <> 0) Or (dbl_metric_mat_cost <> 0 Or dbl_metric_labor_cost <> 0 Or dbl_metric_equip_cost <> 0) Then
            'If Trim(metric_total_cost(0).Text) <> "" Then
            If dbl_metric_total_cost <> 0 Then
'                metric_total_cost(1).Text = Format(dbl_metric_total_cost, ReplaceCharactersForFormat(metric_total_cost(0).Text))
'            Else
                metric_total_cost(1).Text = Format(dbl_metric_total_cost, ReplaceCharactersForFormat(CStr(dbl_metric_total_cost)))
            End If
        Else
        metric_total_cost(1).Text = metric_total_cost(0).Text
        End If
        If Trim(res_mat_cost(0).Text) <> "" And Val(res_mat_cost(0).Text) <> 0 Then
        res_mat_cost(1).Text = Format(dbl_res_mat_cost, ReplaceCharactersForFormat(res_mat_cost(0).Text))
        Else
        res_mat_cost(1).Text = res_mat_cost(0).Text
        End If
        If Trim(res_labor_cost(0).Text) <> "" And Val(res_labor_cost(0).Text) <> 0 Then
        res_labor_cost(1).Text = Format(dbl_res_labor_cost, ReplaceCharactersForFormat(res_labor_cost(0).Text))
        Else
        res_labor_cost(1).Text = res_labor_cost(0).Text
        End If
        If Trim(res_equip_cost(0).Text) <> "" And Val(res_equip_cost(0).Text) <> 0 Then
        res_equip_cost(1).Text = Format(dbl_res_equip_cost, ReplaceCharactersForFormat(res_equip_cost(0).Text))
        Else
        res_equip_cost(1).Text = res_equip_cost(0).Text
        End If
        If (Trim(res_total_cost(0).Text) <> "" And Val(res_total_cost(0).Text) <> 0) Or (dbl_res_mat_cost <> 0 Or dbl_res_labor_cost <> 0 Or dbl_res_equip_cost <> 0) Then
            'If (Trim(res_total_cost(0).Text) <> "") Then
            If dbl_res_total_cost <> 0 Then
'                res_total_cost(1).Text = Format(dbl_res_total_cost, ReplaceCharactersForFormat(res_total_cost(0).Text))
'            Else
                res_total_cost(1).Text = Format(dbl_res_total_cost, ReplaceCharactersForFormat(CStr(dbl_res_total_cost)))
            End If
        Else
        res_total_cost(1).Text = res_total_cost(0).Text
        End If
        
        
        If Trim(std_mat_cost_op(0).Text) <> "" And Val(std_mat_cost_op(0).Text) <> 0 Then
        std_mat_cost_op(1).Text = Format(dbl_std_mat_cost_op, ReplaceCharactersForFormat(std_mat_cost_op(0).Text))
        Else
        std_mat_cost_op(1).Text = std_mat_cost_op(0).Text
        End If
        If Trim(std_labor_cost_op(0).Text) <> "" And Val(std_labor_cost_op(0).Text) <> 0 Then
        std_labor_cost_op(1).Text = Format(dbl_std_labor_cost_op, ReplaceCharactersForFormat(std_labor_cost_op(0).Text))
        Else
        std_labor_cost_op(1).Text = std_labor_cost_op(0).Text
        End If
        If Trim(std_equip_cost_op(0).Text) <> "" And Val(std_equip_cost_op(0).Text) <> 0 Then
        std_equip_cost_op(1).Text = Format(dbl_std_equip_cost_op, ReplaceCharactersForFormat(std_equip_cost_op(0).Text))
        Else
        std_equip_cost_op(1).Text = std_equip_cost_op(0).Text
        End If
        If (Trim(std_total_cost_op(0).Text) <> "" And Val(std_total_cost_op(0).Text) <> 0) Or (dbl_std_mat_cost_op <> 0 Or dbl_std_labor_cost_op <> 0 Or dbl_std_equip_cost_op <> 0) Then
            'If Trim(std_total_cost_op(0).Text <> "") Then
            If dbl_std_total_cost_op <> 0 Then
'                std_total_cost_op(1).Text = Format(dbl_std_total_cost_op, ReplaceCharactersForFormat(std_total_cost_op(0).Text))
'            Else
                std_total_cost_op(1).Text = Format(dbl_std_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_std_total_cost_op)))
            End If
        Else
        std_total_cost_op(1).Text = std_total_cost_op(0).Text
        End If
        If Trim(rr_mat_cost_op(0).Text) <> "" And Val(rr_mat_cost_op(0).Text) <> 0 Then
        rr_mat_cost_op(1).Text = Format(dbl_rr_mat_cost_op, ReplaceCharactersForFormat(rr_mat_cost_op(0).Text))
        Else
        rr_mat_cost_op(1).Text = rr_mat_cost_op(0).Text
        End If
        If Trim(rr_labor_cost_op(0).Text) <> "" And Val(rr_labor_cost_op(0).Text) <> 0 Then
        rr_labor_cost_op(1).Text = Format(dbl_rr_labor_cost_op, ReplaceCharactersForFormat(rr_labor_cost_op(0).Text))
        Else
        rr_labor_cost_op(1).Text = rr_labor_cost_op(0).Text
        End If
        If Trim(rr_equip_cost_op(0).Text) <> "" And Val(rr_equip_cost_op(0).Text) <> 0 Then
        rr_equip_cost_op(1).Text = Format(dbl_rr_equip_cost_op, ReplaceCharactersForFormat(rr_equip_cost_op(0).Text))
        Else
        rr_equip_cost_op(1).Text = rr_equip_cost_op(0).Text
        End If
        If (Trim(rr_total_cost_op(0).Text) <> "" And Val(rr_total_cost_op(0).Text) <> 0) Or (dbl_rr_mat_cost_op <> 0 Or dbl_rr_labor_cost_op <> 0 Or dbl_rr_equip_cost_op <> 0) Then
            'If Trim(rr_total_cost_op(0).Text) <> "" Then
            If dbl_rr_total_cost_op <> 0 Then
'                rr_total_cost_op(1).Text = Format(dbl_rr_total_cost_op, ReplaceCharactersForFormat(rr_total_cost_op(0).Text))
'            Else
                rr_total_cost_op(1).Text = Format(dbl_rr_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_rr_total_cost_op)))
            End If
        Else
        rr_total_cost_op(1).Text = rr_total_cost_op(0).Text
        End If
        If Trim(opn_mat_cost_op(0).Text) <> "" And Val(opn_mat_cost_op(0).Text) <> 0 Then
        opn_mat_cost_op(1).Text = Format(dbl_opn_mat_cost_op, ReplaceCharactersForFormat(opn_mat_cost_op(0).Text))
        Else
        opn_mat_cost_op(1).Text = opn_mat_cost_op(0).Text
        End If
        If Trim(opn_labor_cost_op(0).Text) <> "" And Val(opn_labor_cost_op(0).Text) <> 0 Then
        opn_labor_cost_op(1).Text = Format(dbl_opn_labor_cost_op, ReplaceCharactersForFormat(opn_labor_cost_op(0).Text))
        Else
        opn_labor_cost_op(1).Text = opn_labor_cost_op(0).Text
        End If
        If Trim(opn_equip_cost_op(0).Text) <> "" And Val(opn_equip_cost_op(0).Text) <> 0 Then
        opn_equip_cost_op(1).Text = Format(dbl_opn_equip_cost_op, ReplaceCharactersForFormat(opn_equip_cost_op(0).Text))
        Else
        opn_equip_cost_op(1).Text = opn_equip_cost_op(0).Text
        End If
        If (Trim(opn_total_cost_op(0).Text) <> "" And Val(opn_total_cost_op(0).Text) <> 0) Or (dbl_opn_mat_cost_op <> 0 Or dbl_opn_labor_cost_op <> 0 Or dbl_opn_equip_cost_op <> 0) Then
            'If Trim(opn_total_cost_op(0).Text) <> "" Then
            If dbl_opn_total_cost_op <> 0 Then
'            opn_total_cost_op(1).Text = Format(dbl_opn_total_cost_op, ReplaceCharactersForFormat(opn_total_cost_op(0).Text))
'            Else
            opn_total_cost_op(1).Text = Format(dbl_opn_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_opn_total_cost_op)))
            End If
        Else
        opn_total_cost_op(1).Text = opn_total_cost_op(0).Text
        End If
        If Trim(metric_mat_cost_op(0).Text) <> "" And Val(metric_mat_cost_op(0).Text) <> 0 Then
        metric_mat_cost_op(1).Text = Format(dbl_metric_mat_cost_op, ReplaceCharactersForFormat(metric_mat_cost_op(0).Text))
        Else
        metric_mat_cost_op(1).Text = metric_mat_cost_op(0).Text
        End If
        If Trim(metric_labor_cost_op(0).Text) <> "" And Val(metric_labor_cost_op(0).Text) <> 0 Then
        metric_labor_cost_op(1).Text = Format(dbl_metric_labor_cost_op, ReplaceCharactersForFormat(metric_labor_cost_op(0).Text))
        Else
        metric_labor_cost_op(1).Text = metric_labor_cost_op(0).Text
        End If
        If Trim(metric_equip_cost_op(0).Text) <> "" And Val(metric_equip_cost_op(0).Text) <> 0 Then
        metric_equip_cost_op(1).Text = Format(dbl_metric_equip_cost_op, ReplaceCharactersForFormat(metric_equip_cost_op(0).Text))
        Else
        metric_equip_cost_op(1).Text = metric_equip_cost_op(0).Text
        End If
        If (Trim(metric_total_cost_op(0).Text) <> "" And Val(metric_total_cost_op(0).Text) <> 0) Or (dbl_metric_mat_cost_op <> 0 Or dbl_metric_labor_cost_op <> 0 Or dbl_metric_equip_cost_op <> 0) Then
            'If Trim(metric_total_cost_op(0).Text) <> "" Then
            If dbl_metric_total_cost_op <> 0 Then
'                metric_total_cost_op(1).Text = Format(dbl_metric_total_cost_op, ReplaceCharactersForFormat(metric_total_cost_op(0).Text))
'            Else
                metric_total_cost_op(1).Text = Format(dbl_metric_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_metric_total_cost_op)))
            End If
        Else
        metric_total_cost_op(1).Text = metric_total_cost_op(0).Text
        End If
        If Trim(res_mat_cost_op(0).Text) <> "" And Val(res_mat_cost_op(0).Text) <> 0 Then
        res_mat_cost_op(1).Text = Format(dbl_res_mat_cost_op, ReplaceCharactersForFormat(res_mat_cost_op(0).Text))
        Else
        res_mat_cost_op(1).Text = res_mat_cost_op(0).Text
        End If
        If Trim(res_labor_cost_op(0).Text) <> "" And Val(res_labor_cost_op(0).Text) <> 0 Then
        res_labor_cost_op(1).Text = Format(dbl_res_labor_cost_op, ReplaceCharactersForFormat(res_labor_cost_op(0).Text))
        Else
        res_labor_cost_op(1).Text = res_labor_cost_op(0).Text
        End If
        If Trim(res_equip_cost_op(0).Text) <> "" And Val(res_equip_cost_op(0).Text) <> 0 Then
        res_equip_cost_op(1).Text = Format(dbl_res_equip_cost_op, ReplaceCharactersForFormat(res_equip_cost_op(0).Text))
        Else
        res_equip_cost_op(1).Text = res_equip_cost_op(0).Text
        End If
        If (Trim(res_total_cost_op(0).Text) <> "" And Val(res_total_cost_op(0).Text) <> 0) Or (dbl_res_mat_cost_op <> 0 Or dbl_res_labor_cost_op <> 0 Or dbl_res_equip_cost_op <> 0) Then
            'If (Trim(res_total_cost_op(0).Text) <> "") Then
            If dbl_res_total_cost_op <> 0 Then
'                res_total_cost_op(1).Text = Format(dbl_res_total_cost_op, ReplaceCharactersForFormat(res_total_cost_op(0).Text))
'            Else
                res_total_cost_op(1).Text = Format(dbl_res_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_res_total_cost_op)))
            End If
        Else
        res_total_cost_op(1).Text = res_total_cost_op(0).Text
        End If
        
           
        Dim strOpGreaterThanBare As String
        strOpGreaterThanBare = ""
        If (dbl_std_mat_cost_op < dbl_std_mat_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Standard Material" + vbCrLf
        End If
        If (dbl_std_labor_cost_op < dbl_std_labor_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Standard Labor" + vbCrLf
        End If
        If (dbl_std_equip_cost_op < dbl_std_equip_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Standard Equipment" + vbCrLf
        End If
        If (dbl_std_total_cost_op < dbl_std_total_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Standard Total" + vbCrLf
        End If
         If (dbl_rr_mat_cost_op < dbl_rr_mat_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for R&R Material" + vbCrLf
        End If
        If (dbl_rr_labor_cost_op < dbl_rr_labor_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for R&R Labor" + vbCrLf
        End If
        If (dbl_rr_equip_cost_op < dbl_rr_equip_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for R&R Equipment" + vbCrLf
        End If
        If (dbl_rr_total_cost_op < dbl_rr_total_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for R&R Total" + vbCrLf
        End If
        If (dbl_opn_mat_cost_op < dbl_opn_mat_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Open Material" + vbCrLf
        End If
        If (dbl_opn_labor_cost_op < dbl_opn_labor_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Open Labor" + vbCrLf
        End If
        If (dbl_opn_equip_cost_op < dbl_opn_equip_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Open Equipment" + vbCrLf
        End If
        If (dbl_opn_total_cost_op < dbl_opn_total_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Open Total" + vbCrLf
        End If
        
        If (dbl_metric_mat_cost_op < dbl_metric_mat_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Metric Material" + vbCrLf
        End If
        If (dbl_metric_labor_cost_op < dbl_metric_labor_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Metric Labor" + vbCrLf
        End If
        If (dbl_metric_equip_cost_op < dbl_metric_equip_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Metric Equipment" + vbCrLf
        End If
        If (dbl_metric_total_cost_op < dbl_metric_total_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Metric Total" + vbCrLf
        End If
        
        If (dbl_res_mat_cost_op < dbl_res_mat_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Resi Material" + vbCrLf
        End If
        If (dbl_res_labor_cost_op < dbl_res_labor_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Resi Labor" + vbCrLf
        End If
        If (dbl_res_equip_cost_op < dbl_res_equip_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Resi Equipment" + vbCrLf
        End If
        If (dbl_res_total_cost_op < dbl_res_total_cost) Then
            strOpGreaterThanBare = strOpGreaterThanBare + "O&P must be >= Bare for Resi Total" + vbCrLf
        End If
        
         If (strOpGreaterThanBare <> "") Then
            MsgBox (strOpGreaterThanBare + vbCrLf + vbCrLf + "If you are ok with case, then please make sure you choose 'Yes' when prompted on the Unit Cost Maintenance screen.")
        End If
        
        ' everything ok and exiting sub
        
    Else
        'error while rounding in db
        MsgBox "There was a database error while trying to apply the rounding rules - " + vbCrLf + retStr
        
    End If



End Sub


'Private Sub cmdApplyRounding_Click()
'
'    Dim errMessage As String
'    Dim retRounding As String
'    errMessage = ""
'
'    ApplyRoundingRule std_mat_cost(0), std_labor_cost(0), std_equip_cost(0), std_total_cost(0), std_total_cost(1)
'    ApplyRoundingRule rr_mat_cost(0), rr_labor_cost(0), rr_equip_cost(0), rr_total_cost(0), rr_total_cost(1)
'    ApplyRoundingRule opn_mat_cost(0), opn_labor_cost(0), opn_equip_cost(0), opn_total_cost(0), opn_total_cost(1)
'    ApplyRoundingRule metric_mat_cost(0), metric_labor_cost(0), metric_equip_cost(0), metric_total_cost(0), metric_total_cost(1)
'    ApplyRoundingRule res_mat_cost(0), res_labor_cost(0), res_equip_cost(0), res_total_cost(0), res_total_cost(1)
'
'
'    ApplyRoundingRule std_mat_cost_op(0), std_labor_cost_op(0), std_equip_cost_op(0), std_total_cost_op(0), std_total_cost_op(1)
'    ApplyRoundingRule rr_mat_cost_op(0), rr_labor_cost_op(0), rr_equip_cost_op(0), rr_total_cost_op(0), rr_total_cost_op(1)
'    ApplyRoundingRule opn_mat_cost_op(0), opn_labor_cost_op(0), opn_equip_cost_op(0), opn_total_cost_op(0), opn_total_cost_op(1)
'    ApplyRoundingRule metric_mat_cost_op(0), metric_labor_cost_op(0), metric_equip_cost_op(0), metric_total_cost_op(0), metric_total_cost_op(1)
'    ApplyRoundingRule res_mat_cost_op(0), res_labor_cost_op(0), res_equip_cost_op(0), res_total_cost_op(0), res_total_cost_op(1)
'
'    Dim retRoundingFromStoredProc As String
'    retRoundingFromStoredProc = RoundingFromStoredProc()
'
'    If retRoundingFromStoredProc <> "" Then
'      errMessage = errMessage + "There was a problem rounding Overhead & Profit - Resi Total Cost" + vbCrLf
'    End If
'    If errMessage <> "" Then
'        txtMessage.Text = errMessage
'    End If
'
'
'    ApplyRetotaling std_total_cost(0), std_mat_cost(1), std_labor_cost(1), std_equip_cost(1), std_total_cost(1)
'    ApplyRetotaling rr_total_cost(0), rr_mat_cost(1), rr_labor_cost(1), rr_equip_cost(1), rr_total_cost(1)
'    ApplyRetotaling opn_total_cost(0), opn_mat_cost(1), opn_labor_cost(1), opn_equip_cost(1), opn_total_cost(1)
'    ApplyRetotaling metric_total_cost(0), metric_mat_cost(1), metric_labor_cost(1), metric_equip_cost(1), metric_total_cost(1)
'    ApplyRetotaling res_total_cost(0), res_mat_cost(1), res_labor_cost(1), res_equip_cost(1), res_total_cost(1)
'
'    ApplyRetotaling std_total_cost_op(0), std_mat_cost_op(1), std_labor_cost_op(1), std_equip_cost_op(1), std_total_cost_op(1)
'    ApplyRetotaling rr_total_cost_op(0), rr_mat_cost_op(1), rr_labor_cost_op(1), rr_equip_cost_op(1), rr_total_cost_op(1)
'    ApplyRetotaling opn_total_cost_op(0), opn_mat_cost_op(1), opn_labor_cost_op(1), opn_equip_cost_op(1), opn_total_cost_op(1)
'    ApplyRetotaling metric_total_cost_op(0), metric_mat_cost_op(1), metric_labor_cost_op(1), metric_equip_cost_op(1), metric_total_cost_op(1)
'    ApplyRetotaling res_total_cost_op(0), res_mat_cost_op(1), res_labor_cost_op(1), res_equip_cost_op(1), res_total_cost_op(1)
'
'
'End Sub

Private Sub cmdCopyValues_Click()
 frmCallingForm.std_total_cost.Text = std_total_cost(1).Text
 

    
      frmCallingForm.std_mat_cost.Text = std_mat_cost(1).Text
      frmCallingForm.std_labor_cost.Text = std_labor_cost(1).Text
      frmCallingForm.std_equip_cost.Text = std_equip_cost(1).Text
      frmCallingForm.std_total_cost.Text = std_total_cost(1).Text


      frmCallingForm.rr_mat_cost.Text = rr_mat_cost(1).Text
      frmCallingForm.rr_labor_cost.Text = rr_labor_cost(1).Text
      frmCallingForm.rr_equip_cost.Text = rr_equip_cost(1).Text
      frmCallingForm.rr_total_cost.Text = rr_total_cost(1).Text

 
      frmCallingForm.opn_mat_cost.Text = opn_mat_cost(1).Text
      frmCallingForm.opn_labor_cost.Text = opn_labor_cost(1).Text
      frmCallingForm.opn_equip_cost.Text = opn_equip_cost(1).Text
      frmCallingForm.opn_total_cost.Text = opn_total_cost(1).Text


      frmCallingForm.metric_mat_cost.Text = metric_mat_cost(1).Text
      frmCallingForm.metric_labor_cost.Text = metric_labor_cost(1).Text
      frmCallingForm.metric_equip_cost.Text = metric_equip_cost(1).Text
      frmCallingForm.metric_total_cost.Text = metric_total_cost(1).Text


      frmCallingForm.res_mat_cost.Text = res_mat_cost(1).Text
      frmCallingForm.res_labor_cost.Text = res_labor_cost(1).Text
      frmCallingForm.res_equip_cost.Text = res_equip_cost(1).Text
      frmCallingForm.res_total_cost.Text = res_total_cost(1).Text


      frmCallingForm.std_mat_cost_op.Text = std_mat_cost_op(1).Text
      frmCallingForm.std_labor_cost_op.Text = std_labor_cost_op(1).Text
      frmCallingForm.std_equip_cost_op.Text = std_equip_cost_op(1).Text
      frmCallingForm.std_total_cost_op.Text = std_total_cost_op(1).Text

      frmCallingForm.rr_mat_cost_op.Text = rr_mat_cost_op(1).Text
      frmCallingForm.rr_labor_cost_op.Text = rr_labor_cost_op(1).Text
      frmCallingForm.rr_equip_cost_op.Text = rr_equip_cost_op(1).Text
      frmCallingForm.rr_total_cost_op.Text = rr_total_cost_op(1).Text

      frmCallingForm.opn_mat_cost_op.Text = opn_mat_cost_op(1).Text
      frmCallingForm.opn_labor_cost_op.Text = opn_labor_cost_op(1).Text
      frmCallingForm.opn_equip_cost_op.Text = opn_equip_cost_op(1).Text
      frmCallingForm.opn_total_cost_op.Text = opn_total_cost_op(1).Text

      frmCallingForm.metric_mat_cost_op.Text = metric_mat_cost_op(1).Text
      frmCallingForm.metric_labor_cost_op.Text = metric_labor_cost_op(1).Text
      frmCallingForm.metric_equip_cost_op.Text = metric_equip_cost_op(1).Text
      frmCallingForm.metric_total_cost_op.Text = metric_total_cost_op(1).Text


      frmCallingForm.res_mat_cost_op.Text = res_mat_cost_op(1).Text
      frmCallingForm.res_labor_cost_op.Text = res_labor_cost_op(1).Text
      frmCallingForm.res_equip_cost_op.Text = res_equip_cost_op(1).Text
      frmCallingForm.res_total_cost_op.Text = res_total_cost_op(1).Text
 
 
 
End Sub

Private Sub Form_Activate()
    cmdApplyRounding_Click
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Move (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)

    'apply rounding for the values
    
End Sub
