VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEquipRate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipment Rate"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frmEquipRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   10215
   Begin VB.TextBox equiprate_last_update_id_x 
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
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   120
      Tag             =   "N"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraEquipRateException 
      Caption         =   "Equipment Rate Exception"
      Height          =   2775
      Left            =   120
      TabIndex        =   55
      Top             =   7800
      Width           =   9975
      Begin VB.ComboBox region_code 
         Height          =   315
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "3S"
         Top             =   240
         Width           =   915
      End
      Begin VB.ComboBox country_code 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "3S"
         Top             =   240
         Width           =   915
      End
      Begin VB.Frame Frame1 
         Caption         =   "Metric Costs"
         Height          =   615
         Left            =   120
         TabIndex        =   106
         Top             =   1260
         Width           =   9735
         Begin VB.TextBox metric_rent_per_day_x 
            Height          =   315
            Left            =   960
            TabIndex        =   28
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox metric_rent_per_week_x 
            Height          =   315
            Left            =   2940
            TabIndex        =   29
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox metric_rent_per_month_x 
            Height          =   315
            Left            =   4920
            TabIndex        =   30
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox metric_operating_cost_hrly_x 
            Height          =   315
            Left            =   6960
            TabIndex        =   31
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox metric_crew_equip_cost_x 
            Height          =   315
            Left            =   8820
            TabIndex        =   32
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            Caption         =   "Daily Rent:"
            Height          =   255
            Left            =   60
            TabIndex        =   111
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            Caption         =   "Weekly Rent:"
            Height          =   255
            Left            =   1800
            TabIndex        =   110
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            Caption         =   "Monthly Rent:"
            Height          =   255
            Left            =   3780
            TabIndex        =   109
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            Caption         =   "Hrly Operation:"
            Height          =   255
            Left            =   5760
            TabIndex        =   108
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            Caption         =   "Crew Equip:"
            Height          =   255
            Left            =   7800
            TabIndex        =   107
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.TextBox start_date_x 
         BackColor       =   &H8000000F&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   103
         Tag             =   "D"
         Top             =   1980
         Width           =   1035
      End
      Begin VB.TextBox term_date_x 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   1980
         Width           =   1035
      End
      Begin VB.CheckBox pct_ind 
         Caption         =   "Percent"
         Height          =   255
         Left            =   4740
         TabIndex        =   22
         Tag             =   "3"
         Top             =   300
         Width           =   1215
      End
      Begin VB.Frame Frame8 
         Caption         =   "Costs"
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   9735
         Begin VB.TextBox crew_equip_cost_x 
            Height          =   315
            Left            =   8820
            TabIndex        =   27
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox operating_cost_hrly_x 
            Height          =   315
            Left            =   6960
            TabIndex        =   26
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox rent_per_month_x 
            Height          =   315
            Left            =   4920
            TabIndex        =   25
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox rent_per_week_x 
            Height          =   315
            Left            =   2940
            TabIndex        =   24
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox rent_per_day_x 
            Height          =   315
            Left            =   960
            TabIndex        =   23
            Tag             =   "3N"
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            Caption         =   "Crew Equip:"
            Height          =   255
            Left            =   7800
            TabIndex        =   61
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            Caption         =   "Hrly Operation:"
            Height          =   255
            Left            =   5760
            TabIndex        =   60
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            Caption         =   "Monthly Rent:"
            Height          =   255
            Left            =   3780
            TabIndex        =   59
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            Caption         =   "Weekly Rent:"
            Height          =   255
            Left            =   1800
            TabIndex        =   58
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            Caption         =   "Daily Rent:"
            Height          =   255
            Left            =   60
            TabIndex        =   57
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         Caption         =   "Region:"
         Height          =   255
         Left            =   2100
         TabIndex        =   113
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         Caption         =   "Country:"
         Height          =   255
         Left            =   240
         TabIndex        =   112
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   180
         TabIndex        =   105
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         Caption         =   "End Date:"
         Height          =   255
         Left            =   2400
         TabIndex        =   104
         Top             =   2040
         Width           =   795
      End
   End
   Begin VB.Frame fraEquipRate 
      Caption         =   "Equipment Rate"
      Height          =   2775
      Left            =   120
      TabIndex        =   45
      Top             =   3780
      Width           =   8355
      Begin VB.Frame Frame3 
         Caption         =   "Contact"
         Height          =   615
         Left            =   120
         TabIndex        =   114
         Top             =   600
         Width           =   7875
         Begin VB.TextBox contact_id 
            Height          =   315
            Left            =   1140
            TabIndex        =   17
            Tag             =   "2S"
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox company_name 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   116
            Top             =   180
            Width           =   2055
         End
         Begin VB.TextBox contact_name 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   5940
            Locked          =   -1  'True
            TabIndex        =   115
            Top             =   180
            Width           =   1815
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            Caption         =   "Company:"
            Height          =   255
            Left            =   2400
            TabIndex        =   118
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Name:"
            Height          =   255
            Left            =   5340
            TabIndex        =   117
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CheckBox factor_ind 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5340
         TabIndex        =   15
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox estimated_ind 
         Caption         =   "Estimated"
         Height          =   315
         Left            =   6360
         TabIndex        =   16
         Tag             =   "2"
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox comment 
         Height          =   495
         Left            =   1260
         MultiLine       =   -1  'True
         TabIndex        =   19
         Tag             =   "2S"
         Top             =   2160
         Width           =   6975
      End
      Begin VB.TextBox term_date 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1740
         Width           =   1035
      End
      Begin VB.TextBox start_date 
         BackColor       =   &H8000000F&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   46
         Tag             =   "D"
         Top             =   1740
         Width           =   1035
      End
      Begin VB.TextBox info_source_ref 
         Height          =   315
         Left            =   1260
         TabIndex        =   18
         Tag             =   "2S"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox operating_cost_hrly 
         Height          =   315
         Left            =   3480
         TabIndex        =   14
         Tag             =   "2N"
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox rent_per_week 
         Height          =   315
         Left            =   1260
         TabIndex        =   13
         Tag             =   "2N"
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label27 
         Caption         =   "Factor"
         Height          =   195
         Left            =   5595
         TabIndex        =   54
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment:"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   2220
         Width           =   915
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "End Date:"
         Height          =   255
         Left            =   2460
         TabIndex        =   52
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   60
         TabIndex        =   51
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Info Src Ref:"
         Height          =   255
         Left            =   180
         TabIndex        =   50
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Hrly Operation:"
         Height          =   255
         Left            =   2280
         TabIndex        =   49
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Weekly Rent:"
         Height          =   255
         Left            =   180
         TabIndex        =   48
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.TextBox alt_equip_id 
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
      Left            =   4080
      TabIndex        =   1
      Tag             =   "1S"
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox equip_id 
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
      Left            =   1200
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   60
      Width           =   1215
   End
   Begin VB.ComboBox type_code 
      Height          =   315
      ItemData        =   "frmEquipRate.frx":0442
      Left            =   6960
      List            =   "frmEquipRate.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "1S"
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox equiprate_last_update_id 
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
      Left            =   6180
      Locked          =   -1  'True
      TabIndex        =   41
      Tag             =   "N"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   3000
      TabIndex        =   33
      Top             =   7140
      Width           =   1150
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4440
      TabIndex        =   34
      Top             =   7140
      Width           =   1150
   End
   Begin VB.TextBox last_update_person 
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
      Left            =   4500
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox last_update_date 
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
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox equip_skey 
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
      Left            =   6180
      Locked          =   -1  'True
      TabIndex        =   36
      Tag             =   "N"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox equip_last_update_id 
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
      Left            =   6180
      Locked          =   -1  'True
      TabIndex        =   35
      Tag             =   "N"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3075
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   5424
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Equipment"
      TabPicture(0)   =   "frmEquipRate.frx":0462
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label30"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label31"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label32"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label33"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "metric_unit"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "unit"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "index_desc"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "index_code"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "model_name"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "traces_ind"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Descriptions"
      TabPicture(1)   =   "frmEquipRate.frx":047E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label41"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label42"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label43"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label44"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "tech_desc"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "book_desc"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "crew_equip_desc"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "crew_equip_desc_plural"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Metric Descriptions"
      TabPicture(2)   =   "frmEquipRate.frx":049A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label45"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label46"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label47"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label48"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "metric_crew_equip_desc_plural"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "metric_crew_equip_desc"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "metric_book_desc"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "metric_tech_desc"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.CheckBox traces_ind 
         Caption         =   "TRACES"
         Height          =   255
         Left            =   6120
         TabIndex        =   9
         Tag             =   "1"
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox model_name 
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Tag             =   "1S"
         Top             =   540
         Width           =   2535
      End
      Begin VB.TextBox index_code 
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
         Left            =   4620
         TabIndex        =   8
         Tag             =   "1S"
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox index_desc 
         Height          =   315
         Left            =   1260
         TabIndex        =   7
         Tag             =   "1S"
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox unit 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "1S"
         Top             =   540
         Width           =   1215
      End
      Begin VB.ComboBox metric_unit 
         Height          =   315
         Left            =   7020
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "1S"
         Top             =   540
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Formatting"
         Height          =   675
         Left            =   240
         TabIndex        =   70
         Top             =   1560
         Width           =   7875
         Begin VB.TextBox indent_code 
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
            Left            =   1320
            TabIndex        =   10
            Tag             =   "1N"
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox format_code 
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
            Left            =   6060
            TabIndex        =   12
            Tag             =   "1S"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox format_characters 
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
            Left            =   3660
            TabIndex        =   11
            Tag             =   "1N"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Format Chars:"
            Height          =   255
            Left            =   2580
            TabIndex        =   73
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Format Code:"
            Height          =   255
            Left            =   4980
            TabIndex        =   72
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Indent Code:"
            Height          =   255
            Left            =   360
            TabIndex        =   71
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.TextBox metric_tech_desc 
         Height          =   495
         Left            =   -73740
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   69
         Tag             =   "1S"
         Top             =   540
         Width           =   6855
      End
      Begin VB.TextBox metric_book_desc 
         Height          =   495
         Left            =   -73740
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   68
         Tag             =   "1S"
         Top             =   1140
         Width           =   6855
      End
      Begin VB.TextBox metric_crew_equip_desc 
         Height          =   315
         Left            =   -73740
         MaxLength       =   35
         TabIndex        =   67
         Tag             =   "1S"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox metric_crew_equip_desc_plural 
         Height          =   315
         Left            =   -69480
         MaxLength       =   35
         TabIndex        =   66
         Tag             =   "1S"
         Top             =   1740
         Width           =   2595
      End
      Begin VB.TextBox crew_equip_desc_plural 
         Height          =   315
         Left            =   -69480
         MaxLength       =   35
         TabIndex        =   65
         Tag             =   "1S"
         Top             =   1740
         Width           =   2595
      End
      Begin VB.TextBox crew_equip_desc 
         Height          =   315
         Left            =   -73740
         MaxLength       =   35
         TabIndex        =   64
         Tag             =   "1S"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox book_desc 
         Height          =   495
         Left            =   -73740
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   63
         Tag             =   "1S"
         Top             =   1140
         Width           =   6855
      End
      Begin VB.TextBox tech_desc 
         Height          =   495
         Left            =   -73740
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   62
         Tag             =   "1S"
         Top             =   540
         Width           =   6855
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip Plural:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   101
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   100
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   99
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Tech Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   98
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip Plural:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   97
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   96
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   95
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Tech Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   94
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Indent Code:"
         Height          =   255
         Left            =   -74580
         TabIndex        =   93
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "Table Ref Col:"
         Height          =   255
         Left            =   -69960
         TabIndex        =   92
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "Format Code:"
         Height          =   255
         Left            =   -69960
         TabIndex        =   91
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "Format Chars:"
         Height          =   255
         Left            =   -72360
         TabIndex        =   90
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Graphic Ref ID:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   89
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Table Ref ID:"
         Height          =   255
         Left            =   -72300
         TabIndex        =   88
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Chng Notice Cd:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   87
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Model Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Index Code:"
         Height          =   255
         Left            =   3600
         TabIndex        =   85
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "Index Desc:"
         Height          =   255
         Left            =   300
         TabIndex        =   84
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit:"
         Height          =   255
         Left            =   4080
         TabIndex        =   83
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Metric Unit:"
         Height          =   255
         Left            =   6060
         TabIndex        =   82
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Tech Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   81
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   80
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   79
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip Plural:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   78
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip Plural:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   77
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Crew Equip:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   76
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   75
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Tech Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   74
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8520
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "Alt Equip ID:"
      Height          =   255
      Left            =   3000
      TabIndex        =   44
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Equip ID:"
      Height          =   255
      Left            =   180
      TabIndex        =   43
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Type Code:"
      Height          =   255
      Left            =   5760
      TabIndex        =   42
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3120
      TabIndex        =   40
      Top             =   6780
      Width           =   1275
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   180
      TabIndex        =   39
      Top             =   6780
      Width           =   1095
   End
End
Attribute VB_Name = "frmEquipRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As ADODB.RecordSet
Dim m_blnRecFlag As Boolean ' True if a populated RecordSet was passed, then we show data
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
Dim strLast_equip_id As String ' Holds last equip_id so we know if it changed

' Fills all fields with data
Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If Not m_rec.Fields("equip_skey") = 0 Then
        m_blnRecFlag = True
    End If
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String

    Dim varButton
    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
    If varButton = vbNo Then
        Exit Sub
    End If

    If m_rec.Fields("type_code").Value = "M" Then
        strUpdate = "exec sp_delete_equipment_rate "
        strUpdate = strUpdate + "@equip_skey=" + str(Me.Controls("equip_skey")) + ","
        strUpdate = strUpdate + " @contact_id='" + Me.Controls("contact_id") + "',"
        strUpdate = strUpdate + " @start_date='" + Me.Controls("start_date") + "',"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    ElseIf m_rec.Fields("type_code").Value = "E" Then
        strUpdate = "exec sp_delete_equipment_rate_x "
        strUpdate = strUpdate + "@equip_skey=" + str(Me.Controls("equip_skey")) + ","
        strUpdate = strUpdate + " @country_code='" + Me.Controls("country_code") + "',"
        strUpdate = strUpdate + " @region_code='" + Me.Controls("region_code") + "',"
        strUpdate = strUpdate + " @start_date='" + Me.Controls("start_date") + "',"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    End If
    
    blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
    Else
        MsgBox "Delete successful."
        m_rec.Delete
        m_blnDeleted = True
        Unload Me
    End If
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim strTempMat As String
    Dim strUpdate As String
    Dim strMaterial As String
    Dim strPrice As String
    Dim strMatWhere As String
    Dim strPriceWhere As String
    Dim strError As String
    Dim blnRet As Boolean
    Dim fld As ADODB.Field
    Dim ctr As Control
    Dim blnUpdateEquip As Boolean
    Dim blnUpdateEquipRate As Boolean
    Dim blnUpdateEquipRateEx As Boolean
    Dim strSelect As String
    Dim rec As New ADODB.RecordSet
    
    m_blnWereErrors = False
    
    ' if we are updating
    If m_blnInsert = False Then
        Dim recClone As ADODB.RecordSet
        Set recClone = m_rec.Clone
        recClone.AddNew
        UpdateRecordsetFromForm Me, recClone ' m_rec
        For Each fld In m_rec.Fields
            ' If the value changed
            'If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor (fld.Value = "")) Then
            If Not fld.Value = recClone.Fields(fld.Name).Value Or ((IsNull(fld.Value) Or fld.Value = "") Xor (recClone.Fields(fld.Name).Value = "")) Then
                Set ctr = Nothing
                Set ctr = Me.Controls(fld.Name)
                If Not ctr Is Nothing Then
                    ' See what table the field is from
                    ' Mark the table we should update
                    If Left(Me.Controls(fld.Name).Tag, 1) = 1 Then
                        blnUpdateEquip = True
                    ElseIf Left(Me.Controls(fld.Name).Tag, 1) = 2 Then
                        blnUpdateEquipRate = True
                    ElseIf Left(Me.Controls(fld.Name).Tag, 1) = 3 Then
                        blnUpdateEquipRateEx = True
                    End If
                End If
            End If
        Next
        ' Undo the changes made by the UpdateRecordsetFromForm call above
        recClone.CancelUpdate
        recClone.Close
        Set recClone = Nothing
        ' What gets updated depends on type_code
        If type_code.Text = "M" Then
            If blnUpdateEquipRate And blnUpdateEquip Then
                blnRet = False
                strUpdate = "exec sp_update_equipment_and_rate @equip_skey=" + equip_skey.Text + ", @equiprate_last_update_id=" + equiprate_last_update_id.Text + ", @equip_last_update_id=" + equip_last_update_id.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 1
                BuildStoredProcSQL Me, strUpdate, 2
                strUpdate = strUpdate + " @old_contact_id='" + m_rec.Fields("contact_id").OriginalValue + "', "
                strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
                strUpdate = strUpdate + " @factor_ind=" + str(factor_ind.Value) + ", "
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                    equiprate_last_update_id.Text = m_rec.Fields("equiprate_last_update_id").Value
                    m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                    equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
                    strSelect = "select start_date from material_price where mat_skey=" + str(m_rec.Fields("equip_skey")) + " and last_update_id=" + str(m_rec.Fields("equiprate_last_update_id"))
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
               End If
            ElseIf blnUpdateEquipRate Then
                blnRet = False
                strUpdate = "exec sp_update_equipment_rate @equip_skey=" + equip_skey.Text + ", @equiprate_last_update_id=" + equiprate_last_update_id.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 2
                strUpdate = strUpdate + " @old_contact_id='" + m_rec.Fields("contact_id").OriginalValue + "', "
                strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
                strUpdate = strUpdate + " @factor_ind=" + str(factor_ind.Value) + ", "
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                     m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                    equiprate_last_update_id.Text = m_rec.Fields("equiprate_last_update_id").Value
                    strSelect = "select start_date from equipment_rate where equip_skey=" + str(m_rec.Fields("equip_skey")) + " and last_update_id=" + str(m_rec.Fields("equiprate_last_update_id"))
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
               End If
            ElseIf blnUpdateEquip Then
                blnRet = False
                strUpdate = "exec sp_update_equipment @equip_skey=" + equip_skey.Text + ", @equip_last_update_id=" + equip_last_update_id.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 1
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                    equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
                End If
            Else
                MsgBox "You must modify a field before updating."
                Exit Sub
            End If
        ElseIf type_code.Text = "E" Then
            If blnUpdateEquipRateEx And blnUpdateEquip Then
                blnRet = False
                strUpdate = "exec sp_update_equipment_and_rate_x @equip_skey=" + equip_skey.Text + ", @equiprate_last_update_id=" + equiprate_last_update_id.Text + ", @equip_last_update_id=" + equip_last_update_id.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 1
                BuildStoredProcSQL Me, strUpdate, 2
                strUpdate = strUpdate + " @old_region_code='" + m_rec.Fields("region_code").OriginalValue + "', "
                strUpdate = strUpdate + " @old_country_code='" + m_rec.Fields("country_code").OriginalValue + "', "
                strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                    equiprate_last_update_id_x.Text = m_rec.Fields("equiprate_last_update_id_x").Value
                    m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                    equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
                    strSelect = "select start_date from published_equipment_rate_excep where equip_skey=" + str(m_rec.Fields("equip_skey")) + " and last_update_id=" + str(m_rec.Fields("equiprate_last_update_id_x"))
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
               End If
            ElseIf blnUpdateEquipRateEx Then
                blnRet = False
                strUpdate = "exec sp_update_equipment_rate_x @equip_skey=" + equip_skey.Text + ", @equiprate_last_update_id_x=" + equiprate_last_update_id_x.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 3
                strUpdate = strUpdate + " @old_region_code='" + m_rec.Fields("region_code").OriginalValue + "', "
                strUpdate = strUpdate + " @old_country_code='" + m_rec.Fields("country_code").OriginalValue + "', "
                strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                     m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                    equiprate_last_update_id_x.Text = m_rec.Fields("equiprate_last_update_id_x").Value
                    strSelect = "select start_date from published_equipment_rate_excep where equip_skey=" + str(m_rec.Fields("equip_skey")) + " and last_update_id=" + str(m_rec.Fields("equiprate_last_update_id_x"))
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
               End If
            ElseIf blnUpdateEquip Then
                blnRet = False
                strUpdate = "exec sp_update_equipment @equip_skey=" + equip_skey.Text + ", @equip_last_update_id=" + equip_last_update_id.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 1
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                    equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
                End If
            Else
                MsgBox "You must modify a field before updating."
                Exit Sub
            End If
        ElseIf type_code.Text = "H" Then
            If blnUpdateEquip Then
                blnRet = False
                strUpdate = "exec sp_update_equipment @equip_skey=" + equip_skey.Text + ", @equip_last_update_id=" + equip_last_update_id.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 1
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                    equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
                End If
            Else
                MsgBox "You must modify a field before updating."
                Exit Sub
            End If
        End If
    ' If we are inserting (or cloning)
    Else
        ' If equip_skey is set, then just insert rate and maybe update equipment
        If Len(equip_skey.Text) > 0 Then
            ' What gets inserted depends on type_code
            If type_code.Text = "M" Then
                strUpdate = "exec sp_insert_equipment_rate @equip_skey=" + equip_skey.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 1
                BuildStoredProcSQL Me, strUpdate, 2
                strUpdate = strUpdate + " @equip_last_update_id=" + equip_last_update_id.Text + ", "
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                    equiprate_last_update_id.Text = m_rec.Fields("equiprate_last_update_id").Value
                    strSelect = "select start_date from equipment_rate where equip_skey=" + str(m_rec.Fields("equip_skey")) + " and last_update_id=" + str(m_rec.Fields("equiprate_last_update_id") - 1)
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                End If
            ElseIf type_code.Text = "E" Then
                strUpdate = "exec sp_insert_equipment_rate_x @equip_skey=" + equip_skey.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 1
                BuildStoredProcSQL Me, strUpdate, 3
                strUpdate = strUpdate + " @equip_last_update_id=" + equip_last_update_id.Text + ", "
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                    equiprate_last_update_id_x.Text = m_rec.Fields("equiprate_last_update_id_x").Value
                    strSelect = "select start_date from published_equipment_rate_excep where equip_skey=" + str(m_rec.Fields("equip_skey")) + " and last_update_id=" + str(m_rec.Fields("equiprate_last_update_id_x") - 1)
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                End If
            End If
            ' Now check if we need to update equipment
'            Dim recClone As ADODB.RecordSet
'            Set recClone = m_rec.Clone
'            UpdateRecordsetFromForm Me, recClone
'            For Each fld In recClone.Fields
'                ' If the value changed
'                If Not fld.OriginalValue = fld.Value Or (IsNull(fld.OriginalValue) Xor (IsNull(fld.Value) Or (fld.Value = ""))) Then
'                    Set ctr = Nothing
'                    Set ctr = Me.Controls(fld.Name)
'                    If Not ctr Is Nothing Then
'                        ' See what table the field is from
'                         ' If it is from Equipment
'                        If Left(Me.Controls(fld.Name).Tag, 1) = 1 Then
'                            blnRet = False
'                            strUpdate = "exec sp_update_equipment @equip_skey=" + equip_skey.Text + ", @equip_last_update_id=" + equip_last_update_id.Text + ", "
'                            BuildStoredProcSQL Me, strUpdate, 1
'                            strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
'                            blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
'                            If blnRet = False Then
'                                MsgBox strError
'                                m_blnWereErrors = True
'                            Else
'                               ' Put latest data into source recordset
'                               UpdateRecordsetFromForm Me, m_rec
'                               m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("euqip_last_update_id").Value + 1
'                               equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
'                              UpdateFormFromRecordset Me, m_rec
'                           End If
'                            Exit For
'                        End If
'                    End If
'                End If
'            Next
'            recClone.CancelUpdate
'            recClone.Close
            If m_blnWereErrors = False Then
                MsgBox "Update successful."
            End If
        ' Insert equipment and rate
        Else
            ' Insert with rate
            If type_code.Text = "M" Then
                blnRet = False
                strUpdate = "exec sp_insert_equipment_and_rate "
                BuildStoredProcSQL Me, strUpdate, 1
                BuildStoredProcSQL Me, strUpdate, 2
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                    equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
                    m_rec.Fields("equiprate_last_update_id").Value = m_rec.Fields("equiprate_last_update_id").Value + 1
                    equiprate_last_update_id.Text = m_rec.Fields("equiprate_last_update_id").Value
                    strSelect = "select start_date from equipment_rate where equip_skey=" + str(m_rec.Fields("equip_skey")) + " and last_update_id=" + str(m_rec.Fields("equiprate_last_update_id"))
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
                End If
            ElseIf type_code.Text = "E" Then
                blnRet = False
                strUpdate = "exec sp_insert_equipment_and_rate_x "
                BuildStoredProcSQL Me, strUpdate, 1
                BuildStoredProcSQL Me, strUpdate, 3
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                    equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
                    m_rec.Fields("equiprate_last_update_id_x").Value = m_rec.Fields("equiprate_last_update_id_x").Value + 1
                    equiprate_last_update_id_x.Text = m_rec.Fields("equiprate_last_update_id_x").Value
                    strSelect = "select start_date from published_equipment_rate_excep where equip_skey=" + str(m_rec.Fields("equip_skey")) + " and last_update_id=" + str(m_rec.Fields("equiprate_last_update_id_x"))
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
                End If
            Else
                blnRet = False
                strUpdate = "exec sp_insert_equipment "
                BuildStoredProcSQL Me, strUpdate, 1
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("equip_last_update_id").Value = m_rec.Fields("equip_last_update_id").Value + 1
                    equip_last_update_id.Text = m_rec.Fields("equip_last_update_id").Value
                    UpdateFormFromRecordset Me, m_rec
                    MsgBox "Update successful."
                End If
            End If
        End If
    End If
End Sub


Private Sub contact_id_Validate(Cancel As Boolean)
    Dim rec As ADODB.RecordSet
    g_objDAL.GetRecordset vbNullString, "select company_name, first_name, last_name from information_source where contact_id = '" + contact_id.Text + "'", rec
    If rec.RecordCount = 0 Then
        MsgBox "You must enter a valid Contact ID."
        rec.Close
        Cancel = True
    Else
        company_name.Text = rec.Fields("company_name").Value
        contact_name.Text = rec.Fields("first_name")
        If Len(contact_name.Text) > 0 Then contact_name.Text = contact_name.Text + " "
        contact_name.Text = contact_name.Text + rec.Fields("last_name")
        rec.Close
    End If
End Sub

Private Sub Form_Activate()
    OutputView False

End Sub

Private Sub Form_Initialize()
    m_blnInsert = False
    m_blnDeleted = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Move START_LEFT, START_TOP, 10305, 8115
    
    g_objDAL.GetRecordset CONNECT, "select unit from unit_of_measure order by unit", rec
    While Not rec.EOF
        unit.AddItem (rec.Fields("unit").Value)
        metric_unit.AddItem (rec.Fields("unit").Value)
        rec.MoveNext
    Wend
    rec.Close
    g_objDAL.GetRecordset CONNECT, "select country_code from country order by country_code", rec
    While Not rec.EOF
        country_code.AddItem (rec.Fields("country_code").Value)
        rec.MoveNext
    Wend
    rec.Close
    g_objDAL.GetRecordset CONNECT, "select region_code from region order by region_code", rec
    While Not rec.EOF
        region_code.AddItem (rec.Fields("region_code").Value)
        rec.MoveNext
    Wend
    rec.Close
    
    ' If we are showing data
'    If m_blnRecFlag = True Then
        If Not m_rec.State = adStateClosed Then
            UpdateFormFromRecordset Me, m_rec
        End If
        strLast_equip_id = m_rec.Fields("equip_id").Value
        g_objDAL.GetRecordset vbNullString, "select company_name, first_name, last_name from information_source where contact_id = '" + contact_id.Text + "'", rec

        ' Build company and contact name
        company_name.Text = rec.Fields("company_name")
        contact_name.Text = rec.Fields("first_name")
        If Len(contact_name.Text) > 0 Then contact_name.Text = contact_name.Text + " "
        contact_name.Text = contact_name.Text + rec.Fields("last_name")
        rec.Close
'    Else
'        ' Set all controls to blanks
'        ' Loop through all controls on form
'        For Each ctr In Me.Controls
'            ' Check type of control
'            If TypeOf ctr Is TextBox Then
'                ctr = ""
'            ElseIf TypeOf ctr Is CheckBox Then
'                ctr = 0
'            End If
'        Next ctr
'    End If
    ' If we are NOT inserting
    If m_blnInsert = False Then
        ' Lock fields that can't be changed
        equip_id.Locked = True
        equip_id.BackColor = LTGREY
        
        Me.Caption = Me.Caption + " [" + m_rec.Fields("equip_id").Value + "]"
    Else
        ' If we are inserting and not showing data
        ' Set some defaults
        If Not m_blnRecFlag Then
'            active_status_ind.Value = 1
            Me.Caption = Me.Caption + " [New]"
        Else
            Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("equip_id").Value + "]"
        End If
    End If
    SSTab1.Tab = 0
    ' Make the form show the right fields
    type_code_LostFocus
End Sub

Private Sub equip_id_LostFocus()
    On Error Resume Next
    If equip_id.Locked = False And Not strLast_equip_id = equip_id.Text Then
        Dim strSelect As String
        Dim blnReturn As Boolean
        Dim ctr As Control
        Dim rec As New ADODB.RecordSet
        
        strLast_equip_id = equip_id.Text
                
        ' Check to see if the mat_id entered exists already
        strSelect = "Select *, last_update_id as equip_last_update_id from Equipment where equip_id='" + equip_id.Text + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        ' If it does, copy that data into fields
        If rec.RecordCount > 0 Then
            Dim fld As ADODB.Field
            For Each fld In rec.Fields
                m_rec.Fields(fld.Name).Value = fld.Value
            Next
            For Each ctr In Me.Controls
                If Left(ctr.Tag, 1) = "1" Then
                    ' Check type of control
                    If TypeOf ctr Is TextBox Then
                        ctr.Text = rec.Fields(ctr.Name)
                    ElseIf TypeOf ctr Is CheckBox Then
                        ' Convert from True/False to 1/0
                        If rec.Fields(ctr.Name) Then
                            ctr.Value = 1
                        Else
                            ctr.Value = 0
                        End If
                    ElseIf TypeOf ctr Is ComboBox Then
                        ctr.Text = rec.Fields(ctr.Name)
                    End If
                End If
            Next
            equip_skey.Text = rec.Fields("equip_skey")
            equip_last_update_id.Text = rec.Fields("equip_last_update_id")
        Else
            ' Only blank out fields if we are not inserting
            If m_blnInsert = False Then
                For Each ctr In Me.Controls
                    If Left(ctr.Tag, 1) = "1" And Not ctr.Name = "equip_id" Then
                        ' Check type of control
                        If TypeOf ctr Is TextBox Then
                            ctr.Text = ""
                        ElseIf TypeOf ctr Is CheckBox Then
                            ' Convert from True/False to 1/0
                            ctr.Value = 0
                        ElseIf TypeOf ctr Is ComboBox Then
                            ctr.ListIndex = -1
                        End If
                    End If
                Next
            End If
            equip_skey.Text = ""
        End If
'        m_rec.Close
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button
    Dim blnPendingChange As Boolean
    
    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        blnPendingChange = IsControlChanged(Me, m_rec)
    
        If blnPendingChange = True Then
            Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
            If Button = vbYes Then
                cmdUpdate_Click
                ' If there were errors, cancel the close
                If m_blnWereErrors Then
                    Cancel = True
                End If
            ElseIf Button = vbCancel Then
                Cancel = True
                Exit Sub
            ElseIf m_blnInsert = True Then
                m_rec.Delete
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
ResizeForm Me
End Sub

Private Sub type_code_LostFocus()
    On Error Resume Next
    If type_code.Text = "E" Then
        fraEquipRate.Visible = False
        fraEquipRateException.Move 120, 3780
        fraEquipRateException.Visible = True
        last_update_person.Text = m_rec.Fields("equiprate_last_update_person_x").Value
        last_update_date.Text = m_rec.Fields("equiprate_last_update_date_x").Value
    ElseIf type_code.Text = "M" Then
        fraEquipRateException.Visible = False
        fraEquipRate.Move 120, 3780
        fraEquipRate.Visible = True
        last_update_person.Text = m_rec.Fields("equiprate_last_update_person").Value
        last_update_date.Text = m_rec.Fields("equiprate_last_update_date").Value
    Else
        fraEquipRate.Visible = False
        fraEquipRateException.Visible = False
        last_update_person.Text = m_rec.Fields("equip_last_update_person").Value
        last_update_date.Text = m_rec.Fields("equip_last_update_date").Value
    End If
End Sub

Private Sub rent_per_week_Validate(Cancel As Boolean)
    CheckValueForNumber rent_per_week.Text, Cancel
End Sub

Private Sub operating_cost_hrly_Validate(Cancel As Boolean)
    CheckValueForNumber operating_cost_hrly.Text, Cancel
End Sub

Private Sub rent_per_day_x_Validate(Cancel As Boolean)
    CheckValueForNumber rent_per_day_x.Text, Cancel
End Sub

Private Sub rent_per_week_x_Validate(Cancel As Boolean)
    CheckValueForNumber rent_per_week_x.Text, Cancel
End Sub

Private Sub rent_per_month_x_Validate(Cancel As Boolean)
    CheckValueForNumber rent_per_month_x.Text, Cancel
End Sub

Private Sub operating_cost_hrly_x_Validate(Cancel As Boolean)
    CheckValueForNumber operating_cost_hrly_x.Text, Cancel
End Sub

Private Sub crew_equip_cost_x_Validate(Cancel As Boolean)
    CheckValueForNumber crew_equip_cost_x.Text, Cancel
End Sub

Private Sub metric_rent_per_day_x_Validate(Cancel As Boolean)
    CheckValueForNumber metric_rent_per_day_x.Text, Cancel
End Sub

Private Sub metric_rent_per_week_x_Validate(Cancel As Boolean)
    CheckValueForNumber metric_rent_per_week_x.Text, Cancel
End Sub

Private Sub metric_rent_per_month_x_Validate(Cancel As Boolean)
    CheckValueForNumber metric_rent_per_month_x.Text, Cancel
End Sub

Private Sub metric_operating_cost_hrly_x_Validate(Cancel As Boolean)
    CheckValueForNumber metric_operating_cost_hrly_x.Text, Cancel
End Sub

Private Sub metric_crew_equip_cost_x_Validate(Cancel As Boolean)
    CheckValueForNumber metric_crew_equip_cost_x.Text, Cancel
End Sub

