VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmUnitCost 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Unit Cost Maintenance"
   ClientHeight    =   7770
   ClientLeft      =   2925
   ClientTop       =   1800
   ClientWidth     =   11655
   Icon            =   "frmUnitCost.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11655
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4355
      Left            =   0
      ScaleHeight     =   4350
      ScaleWidth      =   11895
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   120
      Width           =   11895
      Begin VB.TextBox ext_unit_cost_id 
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
         Left            =   4240
         TabIndex        =   3
         Tag             =   "1S"
         Top             =   30
         Width           =   1455
      End
      Begin VB.TextBox alt_unit_cost_id 
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
         Left            =   7440
         TabIndex        =   5
         Tag             =   "1S"
         Top             =   30
         Width           =   1455
      End
      Begin VB.TextBox unit_cost_id 
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
         Left            =   1120
         TabIndex        =   1
         Tag             =   "1S"
         Top             =   30
         Width           =   1455
      End
      Begin VB.ComboBox type_code 
         Height          =   315
         ItemData        =   "frmUnitCost.frx":0442
         Left            =   10320
         List            =   "frmUnitCost.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "1S"
         Top             =   30
         Width           =   975
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4080
         Left            =   120
         TabIndex        =   8
         Tag             =   " "
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7197
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Unit Cost"
         TabPicture(0)   =   "frmUnitCost.frx":0446
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label33"
         Tab(0).Control(1)=   "Label30"
         Tab(0).Control(2)=   "Label28"
         Tab(0).Control(3)=   "Label7"
         Tab(0).Control(4)=   "Label8"
         Tab(0).Control(5)=   "Label3"
         Tab(0).Control(6)=   "metric_unit"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "crew_qty"
         Tab(0).Control(8)=   "unit"
         Tab(0).Control(9)=   "crew_id"
         Tab(0).Control(10)=   "daily_output"
         Tab(0).Control(11)=   "metric_daily_output"
         Tab(0).Control(12)=   "fraFormatting"
         Tab(0).Control(13)=   "Frame1"
         Tab(0).ControlCount=   14
         TabCaption(1)   =   "Descriptions"
         TabPicture(1)   =   "frmUnitCost.frx":0462
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label20"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label10"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label46"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label45"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label42"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label41"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label61"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label12"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label13"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label14"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label62"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label9"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label11"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label15"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "comment"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "traces_book_desc"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "metric_assembly_book_desc"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "metric_book_desc"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "metric_tech_desc"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "book_desc"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "tech_desc"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "assembly_book_desc"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "txtLongDescriptionI"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "txtLongDescriptionM"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "cmdLongDesc"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).ControlCount=   25
         TabCaption(2)   =   "Costs"
         TabPicture(2)   =   "frmUnitCost.frx":047E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdShowRounding"
         Tab(2).Control(1)=   "inhouse_total_cost_op"
         Tab(2).Control(2)=   "inhouse_equip_cost_op"
         Tab(2).Control(3)=   "inhouse_mat_cost_op"
         Tab(2).Control(4)=   "inhouse_labor_cost_op"
         Tab(2).Control(5)=   "res_labor_hour"
         Tab(2).Control(6)=   "res_total_cost_op"
         Tab(2).Control(7)=   "res_equip_cost_op"
         Tab(2).Control(8)=   "res_labor_cost_op"
         Tab(2).Control(9)=   "res_mat_cost_op"
         Tab(2).Control(10)=   "res_total_cost"
         Tab(2).Control(11)=   "res_equip_cost"
         Tab(2).Control(12)=   "res_labor_cost"
         Tab(2).Control(13)=   "res_mat_cost"
         Tab(2).Control(14)=   "pct_ind"
         Tab(2).Control(15)=   "metric_labor_hour"
         Tab(2).Control(16)=   "rr_labor_hour"
         Tab(2).Control(17)=   "opn_labor_hour"
         Tab(2).Control(18)=   "std_labor_hour"
         Tab(2).Control(19)=   "metric_equip_cost_op"
         Tab(2).Control(20)=   "metric_labor_cost_op"
         Tab(2).Control(21)=   "metric_mat_cost_op"
         Tab(2).Control(22)=   "rr_equip_cost_op"
         Tab(2).Control(23)=   "rr_labor_cost_op"
         Tab(2).Control(24)=   "rr_mat_cost_op"
         Tab(2).Control(25)=   "opn_equip_cost_op"
         Tab(2).Control(26)=   "opn_labor_cost_op"
         Tab(2).Control(27)=   "opn_mat_cost_op"
         Tab(2).Control(28)=   "std_equip_cost_op"
         Tab(2).Control(29)=   "std_labor_cost_op"
         Tab(2).Control(30)=   "std_mat_cost_op"
         Tab(2).Control(31)=   "metric_total_cost_op"
         Tab(2).Control(32)=   "rr_total_cost_op"
         Tab(2).Control(33)=   "opn_total_cost_op"
         Tab(2).Control(34)=   "std_total_cost_op"
         Tab(2).Control(35)=   "metric_total_cost"
         Tab(2).Control(36)=   "metric_equip_cost"
         Tab(2).Control(37)=   "metric_labor_cost"
         Tab(2).Control(38)=   "metric_mat_cost"
         Tab(2).Control(39)=   "rr_total_cost"
         Tab(2).Control(40)=   "rr_equip_cost"
         Tab(2).Control(41)=   "rr_labor_cost"
         Tab(2).Control(42)=   "rr_mat_cost"
         Tab(2).Control(43)=   "opn_total_cost"
         Tab(2).Control(44)=   "opn_equip_cost"
         Tab(2).Control(45)=   "opn_labor_cost"
         Tab(2).Control(46)=   "opn_mat_cost"
         Tab(2).Control(47)=   "std_total_cost"
         Tab(2).Control(48)=   "std_equip_cost"
         Tab(2).Control(49)=   "std_labor_cost"
         Tab(2).Control(50)=   "std_mat_cost"
         Tab(2).Control(51)=   "Label19"
         Tab(2).Control(52)=   "Label18"
         Tab(2).Control(53)=   "Label17"
         Tab(2).Control(54)=   "Line3"
         Tab(2).Control(55)=   "Label77"
         Tab(2).Control(56)=   "Line2"
         Tab(2).Control(57)=   "Line1"
         Tab(2).Control(58)=   "Label76"
         Tab(2).Control(59)=   "Label75"
         Tab(2).Control(60)=   "Label74"
         Tab(2).Control(61)=   "Label73"
         Tab(2).Control(62)=   "Label72"
         Tab(2).Control(63)=   "Label71"
         Tab(2).Control(64)=   "Label70"
         Tab(2).Control(65)=   "Label69"
         Tab(2).Control(66)=   "Label68"
         Tab(2).Control(67)=   "Label67"
         Tab(2).Control(68)=   "Label66"
         Tab(2).Control(69)=   "Label65"
         Tab(2).Control(70)=   "Label64"
         Tab(2).Control(71)=   "Label63"
         Tab(2).ControlCount=   72
         Begin VB.CommandButton cmdShowRounding 
            Caption         =   "Show Rounding"
            Height          =   495
            Left            =   -65280
            TabIndex        =   149
            Top             =   3240
            Width           =   1335
         End
         Begin VB.TextBox inhouse_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -67200
            TabIndex        =   147
            Tag             =   "3G"
            Top             =   3480
            Width           =   795
         End
         Begin VB.TextBox inhouse_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68025
            TabIndex        =   146
            Tag             =   "3G"
            Top             =   3480
            Width           =   795
         End
         Begin VB.TextBox inhouse_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -69705
            TabIndex        =   145
            Tag             =   "3G"
            Top             =   3480
            Width           =   795
         End
         Begin VB.TextBox inhouse_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68880
            TabIndex        =   144
            Tag             =   "3G"
            Text            =   " "
            Top             =   3480
            Width           =   795
         End
         Begin VB.TextBox res_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   -65985
            TabIndex        =   142
            Tag             =   "3G"
            Top             =   2760
            Width           =   795
         End
         Begin VB.TextBox res_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -67185
            TabIndex        =   141
            Tag             =   "3G"
            Top             =   2760
            Width           =   795
         End
         Begin VB.TextBox res_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68025
            TabIndex        =   140
            Tag             =   "3G"
            Top             =   2760
            Width           =   795
         End
         Begin VB.TextBox res_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68865
            TabIndex        =   139
            Tag             =   "3G"
            Top             =   2760
            Width           =   795
         End
         Begin VB.TextBox res_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -69705
            TabIndex        =   138
            Tag             =   "3G"
            Top             =   2760
            Width           =   795
         End
         Begin VB.TextBox res_total_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -70905
            TabIndex        =   137
            Tag             =   "3G"
            Text            =   " "
            Top             =   2760
            Width           =   795
         End
         Begin VB.TextBox res_equip_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -71745
            TabIndex        =   136
            Tag             =   "3G"
            Text            =   " "
            Top             =   2760
            Width           =   795
         End
         Begin VB.TextBox res_labor_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72585
            TabIndex        =   135
            Tag             =   "3G"
            Top             =   2760
            Width           =   795
         End
         Begin VB.TextBox res_mat_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -73425
            TabIndex        =   134
            Tag             =   "3G"
            Top             =   2760
            Width           =   795
         End
         Begin VB.Frame Frame1 
            Caption         =   "Book Index"
            Height          =   855
            Left            =   -74400
            TabIndex        =   121
            Top             =   1440
            Width           =   9375
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
               Left            =   6720
               TabIndex        =   123
               Tag             =   "1S"
               Top             =   300
               Width           =   555
            End
            Begin VB.TextBox index_desc 
               Height          =   315
               Left            =   1440
               TabIndex        =   122
               Tag             =   "1S"
               Top             =   300
               Width           =   3135
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               Caption         =   "Index Code:"
               Height          =   255
               Left            =   5700
               TabIndex        =   125
               Top             =   360
               Width           =   915
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               Caption         =   "Index Desc:"
               Height          =   255
               Left            =   480
               TabIndex        =   124
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdLongDesc 
            Height          =   375
            Left            =   10920
            Picture         =   "frmUnitCost.frx":049A
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Edit Long Descriptions"
            Top             =   2560
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.TextBox txtLongDescriptionM 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   43
            Tag             =   "ignore"
            Top             =   2780
            Width           =   9555
         End
         Begin VB.TextBox txtLongDescriptionI 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   41
            Tag             =   "ignore"
            Top             =   2460
            Width           =   9555
         End
         Begin VB.Frame fraFormatting 
            Caption         =   "Book Formatting"
            Height          =   1215
            Left            =   -74400
            TabIndex        =   21
            Top             =   2400
            Width           =   9375
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
               Left            =   4200
               TabIndex        =   25
               Tag             =   "1S"
               Top             =   240
               Width           =   735
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
               Left            =   6720
               TabIndex        =   27
               Tag             =   "1S"
               Top             =   240
               Width           =   735
            End
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
               Left            =   1560
               TabIndex        =   23
               Tag             =   "1S"
               Top             =   240
               Width           =   435
            End
            Begin VB.Label lblPreview3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2520
               TabIndex        =   129
               Top             =   720
               Width           =   90
            End
            Begin VB.Label lblPreview2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2280
               TabIndex        =   128
               Top             =   720
               Width           =   90
            End
            Begin VB.Label lblPreview1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   127
               Top             =   720
               Width           =   615
            End
            Begin VB.Line Line4 
               X1              =   2190
               X2              =   2190
               Y1              =   720
               Y2              =   1005
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               BorderStyle     =   3  'Dot
               Height          =   285
               Left            =   1560
               Top             =   720
               Width           =   7455
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Preview:"
               Height          =   255
               Left            =   480
               TabIndex        =   126
               Top             =   720
               Width           =   915
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Indent Code:"
               Height          =   255
               Left            =   480
               TabIndex        =   22
               Top             =   300
               Width           =   915
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Format Code:"
               Height          =   255
               Left            =   5580
               TabIndex        =   26
               Top             =   300
               Width           =   1035
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Format Chars:"
               Height          =   255
               Left            =   3000
               TabIndex        =   24
               Top             =   300
               Width           =   1035
            End
         End
         Begin VB.CheckBox pct_ind 
            Caption         =   "&Percent"
            Height          =   255
            Left            =   -73440
            TabIndex        =   93
            Tag             =   "S"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox metric_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   -65985
            Locked          =   -1  'True
            TabIndex        =   92
            Tag             =   "3G"
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox rr_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   -65985
            Locked          =   -1  'True
            TabIndex        =   74
            Tag             =   "3G"
            Top             =   1710
            Width           =   795
         End
         Begin VB.TextBox opn_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   -65985
            Locked          =   -1  'True
            TabIndex        =   83
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox std_labor_hour 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   -65985
            Locked          =   -1  'True
            TabIndex        =   65
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox metric_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68025
            TabIndex        =   90
            Tag             =   "3G"
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox metric_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68865
            TabIndex        =   89
            Tag             =   "3G"
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox metric_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -69705
            TabIndex        =   88
            Tag             =   "3G"
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox rr_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68025
            TabIndex        =   72
            Tag             =   "3G"
            Top             =   1710
            Width           =   795
         End
         Begin VB.TextBox rr_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68865
            TabIndex        =   71
            Tag             =   "3G"
            Top             =   1680
            Width           =   795
         End
         Begin VB.TextBox rr_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -69705
            TabIndex        =   70
            Tag             =   "3G"
            Top             =   1710
            Width           =   795
         End
         Begin VB.TextBox opn_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68025
            TabIndex        =   81
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox opn_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68865
            TabIndex        =   80
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox opn_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -69705
            TabIndex        =   79
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox std_equip_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68025
            TabIndex        =   63
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox std_labor_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68865
            TabIndex        =   62
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox std_mat_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -69705
            TabIndex        =   61
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox metric_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -67185
            TabIndex        =   91
            Tag             =   "3G"
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox rr_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -67185
            TabIndex        =   73
            Tag             =   "3G"
            Top             =   1710
            Width           =   795
         End
         Begin VB.TextBox opn_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -67185
            TabIndex        =   82
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox std_total_cost_op 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -67185
            TabIndex        =   64
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox metric_total_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -70905
            TabIndex        =   87
            Tag             =   "3G"
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox metric_equip_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -71745
            TabIndex        =   86
            Tag             =   "3G"
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox metric_labor_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72585
            TabIndex        =   85
            Tag             =   "3G"
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox metric_mat_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -73425
            TabIndex        =   84
            Tag             =   "3G"
            Top             =   2430
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
            Height          =   285
            Left            =   -70905
            TabIndex        =   69
            Tag             =   "3G"
            Top             =   1710
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
            Height          =   285
            Left            =   -71745
            TabIndex        =   68
            Tag             =   "3G"
            Top             =   1710
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
            Height          =   285
            Left            =   -72585
            TabIndex        =   67
            Tag             =   "3G"
            Top             =   1710
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
            Height          =   285
            Left            =   -73440
            TabIndex        =   66
            Tag             =   "3G"
            Top             =   1710
            Width           =   795
         End
         Begin VB.TextBox opn_total_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -70905
            TabIndex        =   78
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox opn_equip_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -71745
            TabIndex        =   77
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox opn_labor_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72585
            TabIndex        =   76
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox opn_mat_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -73425
            TabIndex        =   75
            Tag             =   "3G"
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox std_total_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -70905
            TabIndex        =   60
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox std_equip_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -71745
            TabIndex        =   59
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox std_labor_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72585
            TabIndex        =   58
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox std_mat_cost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -73425
            TabIndex        =   57
            Tag             =   "3G"
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox metric_daily_output 
            Height          =   315
            Left            =   -66120
            TabIndex        =   20
            Tag             =   "1S"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox daily_output 
            Height          =   315
            Left            =   -69240
            TabIndex        =   18
            Tag             =   "1S"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox crew_id 
            Height          =   315
            Left            =   -73560
            TabIndex        =   10
            Tag             =   "1S"
            Top             =   660
            Width           =   1095
         End
         Begin VB.ComboBox unit 
            Height          =   315
            Left            =   -69240
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Tag             =   "1S"
            Top             =   660
            Width           =   1215
         End
         Begin VB.TextBox crew_qty 
            Height          =   315
            Left            =   -71340
            TabIndex        =   12
            Tag             =   "1S"
            Top             =   660
            Width           =   795
         End
         Begin VB.TextBox assembly_book_desc 
            Height          =   285
            Left            =   1320
            MaxLength       =   75
            TabIndex        =   37
            Tag             =   "1S"
            Top             =   1785
            Width           =   9555
         End
         Begin VB.TextBox tech_desc 
            Height          =   285
            Left            =   1320
            MaxLength       =   75
            TabIndex        =   29
            Tag             =   "1S"
            Top             =   420
            Width           =   9555
         End
         Begin VB.TextBox book_desc 
            Height          =   285
            Left            =   1320
            MaxLength       =   75
            TabIndex        =   33
            Tag             =   "1S"
            Top             =   1110
            Width           =   9555
         End
         Begin VB.TextBox metric_tech_desc 
            Height          =   285
            Left            =   1320
            MaxLength       =   75
            TabIndex        =   31
            Tag             =   "1S"
            Top             =   745
            Width           =   9555
         End
         Begin VB.TextBox metric_book_desc 
            Height          =   285
            Left            =   1320
            MaxLength       =   75
            TabIndex        =   35
            Tag             =   "1S"
            Top             =   1420
            Width           =   9555
         End
         Begin VB.TextBox metric_assembly_book_desc 
            Height          =   285
            Left            =   1320
            MaxLength       =   75
            TabIndex        =   39
            Tag             =   "1S"
            Top             =   2110
            Width           =   9555
         End
         Begin VB.TextBox traces_book_desc 
            Height          =   285
            Left            =   1320
            MaxLength       =   75
            TabIndex        =   45
            Top             =   3095
            Visible         =   0   'False
            Width           =   9555
         End
         Begin VB.ComboBox metric_unit 
            Height          =   315
            ItemData        =   "frmUnitCost.frx":07DC
            Left            =   -66120
            List            =   "frmUnitCost.frx":07DE
            Style           =   2  'Dropdown List
            TabIndex        =   16
            TabStop         =   0   'False
            Tag             =   "1S"
            Top             =   660
            Width           =   1215
         End
         Begin VB.TextBox comment 
            Height          =   555
            Left            =   1320
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   47
            Tag             =   "1S"
            Top             =   3180
            Width           =   9555
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "In-House Overhead && Profit r/o"
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
            Left            =   -69720
            TabIndex        =   148
            Top             =   3120
            Width           =   3375
         End
         Begin VB.Label Label18 
            Caption         =   "FMR"
            Height          =   255
            Left            =   -74280
            TabIndex        =   143
            Top             =   3480
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "Resi"
            Height          =   255
            Left            =   -74270
            TabIndex        =   133
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Imp:"
            Height          =   255
            Left            =   960
            TabIndex        =   40
            Top             =   2460
            Width           =   315
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric:"
            Height          =   255
            Left            =   720
            TabIndex        =   42
            Top             =   2745
            Width           =   555
         End
         Begin VB.Label Label9 
            Caption         =   "Long Desc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   119
            Top             =   2480
            Width           =   495
         End
         Begin VB.Line Line3 
            X1              =   -73425
            X2              =   -65040
            Y1              =   1230
            Y2              =   1230
         End
         Begin VB.Label Label77 
            Alignment       =   2  'Center
            Caption         =   "Labor Hours"
            Height          =   495
            Left            =   -65985
            TabIndex        =   56
            Top             =   810
            Width           =   795
         End
         Begin VB.Line Line2 
            X1              =   -66200
            X2              =   -66200
            Y1              =   1440
            Y2              =   2580
         End
         Begin VB.Line Line1 
            X1              =   -69900
            X2              =   -69900
            Y1              =   1440
            Y2              =   2580
         End
         Begin VB.Label Label76 
            Alignment       =   1  'Right Justify
            Caption         =   "Total"
            Height          =   255
            Left            =   -67065
            TabIndex        =   55
            Top             =   990
            Width           =   555
         End
         Begin VB.Label Label75 
            Alignment       =   1  'Right Justify
            Caption         =   "Equipment"
            Height          =   255
            Left            =   -68025
            TabIndex        =   54
            Top             =   990
            Width           =   795
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            Caption         =   "Labor"
            Height          =   255
            Left            =   -68760
            TabIndex        =   53
            Top             =   990
            Width           =   495
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            Caption         =   "Material"
            Height          =   255
            Left            =   -69720
            TabIndex        =   52
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            Caption         =   "Total"
            Height          =   255
            Left            =   -70785
            TabIndex        =   51
            Top             =   990
            Width           =   555
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            Caption         =   "Equipment"
            Height          =   255
            Left            =   -71745
            TabIndex        =   50
            Top             =   990
            Width           =   795
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            Caption         =   "Labor"
            Height          =   255
            Left            =   -72465
            TabIndex        =   49
            Top             =   990
            Width           =   495
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Caption         =   "Material"
            Height          =   255
            Left            =   -73425
            TabIndex        =   48
            Top             =   990
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
            Left            =   -69120
            TabIndex        =   118
            Top             =   630
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
            Left            =   -72360
            TabIndex        =   117
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label Label66 
            Caption         =   "Metric"
            Height          =   255
            Left            =   -74265
            TabIndex        =   116
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "Open"
            Height          =   255
            Left            =   -74265
            TabIndex        =   115
            Top             =   2070
            Width           =   975
         End
         Begin VB.Label Label64 
            Caption         =   "R&&R"
            Height          =   255
            Left            =   -74265
            TabIndex        =   114
            Top             =   1710
            Width           =   375
         End
         Begin VB.Label Label63 
            Caption         =   "Standard"
            Height          =   255
            Left            =   -74265
            TabIndex        =   113
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric:"
            Height          =   255
            Left            =   720
            TabIndex        =   38
            Top             =   2115
            Width           =   555
         End
         Begin VB.Label Label14 
            Caption         =   "Asbly Book"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   112
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Book"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Tech"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric Daily Output:"
            Height          =   255
            Left            =   -67680
            TabIndex        =   19
            Top             =   1140
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Daily Output:"
            Height          =   255
            Left            =   -70440
            TabIndex        =   17
            Top             =   1140
            Width           =   1035
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Crew ID:"
            Height          =   255
            Left            =   -74400
            TabIndex        =   9
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric Unit:"
            Height          =   255
            Left            =   -67080
            TabIndex        =   15
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Unit:"
            Height          =   255
            Left            =   -69840
            TabIndex        =   13
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Crew Qty:"
            Height          =   255
            Left            =   -72480
            TabIndex        =   11
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            Caption         =   "Imp:"
            Height          =   255
            Left            =   960
            TabIndex        =   36
            Top             =   1785
            Width           =   315
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "Imp:"
            Height          =   255
            Left            =   960
            TabIndex        =   28
            Top             =   435
            Width           =   315
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "Imp:"
            Height          =   255
            Left            =   960
            TabIndex        =   32
            Top             =   1110
            Width           =   315
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric:"
            Height          =   255
            Left            =   720
            TabIndex        =   30
            Top             =   750
            Width           =   555
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric:"
            Height          =   255
            Left            =   720
            TabIndex        =   34
            Top             =   1425
            Width           =   555
         End
         Begin VB.Label Label10 
            Caption         =   "Traces Bk:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   3095
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "Comment:"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   3240
            Width           =   915
         End
      End
      Begin VB.Label lbl_ext_unit_cost_id 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit Cost ID 04:"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   90
         Width           =   1275
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Alt Unit Cost ID:"
         Height          =   255
         Left            =   6120
         TabIndex        =   4
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label lbl_unit_cost_id 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit Cost ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   90
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Type Code:"
         Height          =   255
         Left            =   9120
         TabIndex        =   6
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   11895
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   4560
      Width           =   11895
      Begin VB.TextBox txtMasterFormat 
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
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   130
         TabStop         =   0   'False
         Tag             =   "ignore"
         Top             =   2700
         Width           =   600
      End
      Begin VB.TextBox unit_cost_skey 
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
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Tag             =   "1N"
         Top             =   2700
         Width           =   1215
      End
      Begin VB.TextBox last_update_date 
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
         Height          =   315
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1695
      End
      Begin VB.TextBox last_update_person 
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
         Height          =   315
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Tag             =   "S"
         Top             =   2700
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   10320
         TabIndex        =   104
         Top             =   2640
         Visible         =   0   'False
         Width           =   1150
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   9000
         TabIndex        =   103
         Top             =   2640
         Width           =   1150
      End
      Begin VB.Frame fraUnitCost 
         Caption         =   "Unit Cost"
         Height          =   2535
         Left            =   120
         TabIndex        =   94
         Top             =   0
         Width           =   11415
         Begin VB.CommandButton cmdMatUsageDelete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   120
            TabIndex        =   96
            Top             =   2040
            Width           =   1150
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid 
            Height          =   1575
            Left            =   120
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   300
            Width           =   8655
            _ExtentX        =   15266
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
            AllowAddNew     =   -1  'True
            DataMode        =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
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
      End
      Begin VB.Label lblMasterFormat 
         Alignment       =   1  'Right Justify
         Caption         =   "MF:"
         Height          =   255
         Left            =   6960
         TabIndex        =   131
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblSkey 
         Alignment       =   1  'Right Justify
         Caption         =   "Skey:"
         Height          =   255
         Left            =   5040
         TabIndex        =   101
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblUpdated 
         Alignment       =   1  'Right Justify
         Caption         =   "Updated:"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblUpdatedBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Updated By:"
         Height          =   255
         Left            =   2685
         TabIndex        =   99
         Top             =   2760
         Width           =   915
      End
   End
   Begin VB.TextBox crew_type_code 
      Height          =   285
      Left            =   360
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox cstw_last_update_id 
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
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   107
      TabStop         =   0   'False
      Tag             =   "0N"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox ucd_last_update_id 
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
      Left            =   6660
      Locked          =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      Tag             =   "0N"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H000000FF&
      Caption         =   "Resi"
      Height          =   255
      Left            =   840
      TabIndex        =   132
      Top             =   3360
      Width           =   615
   End
End
Attribute VB_Name = "frmUnitCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <modulename> frmUnitCost</modulename>
' <functionname>General (Main) </functionname>
'
' <summary>
'This form is the "child" form of frmUnitCostGrid.frm the "Unit Cost Grid" form.
'It will be displayed upon double-clicking a line on the datagrid or clicking the "Unit Cost" button shown toward the bottom left of the "Unit Cost Grid" form.
'
' Displays (3) tabs as follows:
'1.  Unit Cost
'o   Crew Id
'o unit
'o   Daily Output
'o   Format Code
'o   Indent Code
'o   etc.
'2.  Descriptions
'o Tech
'o Book
'o   Long
'3.  Costs
'"Bare Costs" (raw costs) and "Overhead & Profits" over each of the following
'o op - codes:
'o STD
'o OPN
'o RR
'o metric
'
'
'HELPER Class: CUCostMatMap.Cls
'
' </summary>
' <seealso>N/A</seealso>
' <datastruct>m_rec</datastruct>
' <storedprocedurename>usp_update_unit_cost_driver_ext_rlh
'</storedprocedurename>
' <storedprocedurename> sp_delete_unit_cost</storedprocedurename>
' <storedprocedurename>sp_copy_output_usage</storedprocedurename>
' <storedprocedurename> usp_select_attribute_value_ext</storedprocedurename>
' <storedprocedurename> sp_update_object_attribute_value2
'</storedprocedurename>
' <storedprocedurename>sp_update_object_description</storedprocedurename>
'
'
' <returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
' <code>
'************************************************************************
'IMPORTANT
'************************************************************************
'    'This block of code was arbitrarily lifted from the VB code to show the settings of
'    costs for both "bare costs" and "Overhead and Profit" costs across (these op-codes):
'o   STD  (standard union shop)
'o   OPN (non union)
'o RR(Repair & Remodeling)
'o   METRIC  (metric version of STD)
'o RESI(residential)
'o   FMR   (Facility maintenance & repair or "in-house")
'
'   Each row has the following columns to be set:
'    (Bare Costs)
'o MATERIAL
'o Labor
'o EQUIPMENT
'o   Total Cost
'
'(Overhead & Profit)
'o MATERIAL(op)
'o Labor(op)
'o EQUIPMENT(op)
'o   Total Cost (op)
'
'    Labor Hours
'   '-------------------------------------------------------------------------------------------
'   'typical block of code:
'
'    ' STD
'    std_mat_cost_o = std_mat_cost.Text      '(Bare Cost) Material
'    std_labor_cost_o = std_labor_cost.Text      '(Bare Cost) Labor
'    std_labor_hour_o = std_labor_hour.Text      'Labor Hours
'    std_mat_cost_op_o = std_mat_cost_op.Text    '(Overhead & Profit) Material
'    std_labor_cost_op_o = std_labor_cost_op.Text    '(Overhead & profit) Labor
'    std_equip_cost_o = std_equip_cost.Text      '(Bare Cost) Equipment
'    std_equip_cost_op_o = std_equip_cost_op.Text    '(Overhead & profit) Equipment
'    std_total_cost_o = std_total_cost.Text      '(Bare Cost) Total Cost
'    std_total_cost_op_o = std_total_cost_op.Text    '(Overhead & profit) Total Cost
'
'   'OPN
'    opn_labor_hour_o = opn_labor_hour.Text
'    opn_mat_cost_o = opn_mat_cost.Text
'    opn_mat_cost_op_o = opn_mat_cost_op.Text
'    opn_labor_cost_o = opn_labor_cost.Text
'    opn_labor_cost_op_o = opn_labor_cost_op.Text
'    opn_equip_cost_o = opn_equip_cost.Text
'    opn_equip_cost_op_o = opn_equip_cost_op.Text
'    opn_total_cost_o = opn_total_cost.Text
'    opn_total_cost_op_o = opn_total_cost_op.Text
'
'   'RR
'    rr_labor_hour_o = rr_labor_hour.Text
'    rr_mat_cost_o = rr_mat_cost.Text
'    rr_mat_cost_op_o = rr_mat_cost_op.Text
'    rr_labor_cost_o = rr_labor_cost.Text
'    rr_labor_cost_op_o = rr_labor_cost_op.Text
'    rr_equip_cost_o = rr_equip_cost.Text
'    rr_equip_cost_op_o = rr_equip_cost_op.Text
'    rr_total_cost_o = rr_total_cost.Text
'    rr_total_cost_op_o = rr_total_cost_op.Text
'
'   'METRIC
'    metric_labor_hour_o = metric_labor_hour.Text
'    metric_mat_cost_o = metric_mat_cost.Text
'    metric_mat_cost_op_o = metric_mat_cost_op.Text
'    metric_labor_cost_o = metric_labor_cost.Text
'    metric_labor_cost_op_o = metric_labor_cost_op.Text
'    metric_equip_cost_o = metric_equip_cost.Text
'    metric_equip_cost_op_o = metric_equip_cost_op.Text
'    metric_total_cost_o = metric_total_cost.Text
'    metric_total_cost_op_o = metric_total_cost_op.Text
'
'   'RESI
'    res_labor_hour_o = res_labor_hour.Text
'    res_mat_cost_o = res_mat_cost.Text
'    res_mat_cost_op_o = res_mat_cost_op.Text
'    res_labor_cost_o = res_labor_cost.Text
'    res_labor_cost_op_o = res_labor_cost_op.Text
'    res_equip_cost_o = res_equip_cost.Text
'    res_equip_cost_op_o = res_equip_cost_op.Text
'    res_total_cost_o = res_total_cost.Text
'    res_total_cost_op_o = res_total_cost_op.Text
'    'FMR (In-House)
'    inhouse_mat_cost_op_o = inhouse_mat_cost_op.Text
'    inhouse_equip_cost_op_o = inhouse_equip_cost_op.Text
'    inhouse_labor_cost_op_o = inhouse_labor_cost_op.Text
'    inhouse_total_cost_op_o = inhouse_total_cost_op.Text
'     </code>
'    <code>
'exec usp_update_unit_cost_driver_ext_rlh  @ext_unit_cost_id='681020300409', @alt_unit_cost_id='031104023003', @unit_cost_id='030110203006', @type_code='M', @index_code='', @index_desc='', @format_characters=0, @format_code='F1', @indent_code=1,
'@metric_daily_output='23.22500', @daily_output='350.00000', @crew_id='C3', @unit='SFCA', @crew_qty='1', @assembly_book_desc='Forms (place & strip), beams, plywood, 1 use', @tech_desc='Forms (place & strip), beams, plywood, 1 use',
'@book_desc='Forms (place & strip), beams, plywood, 1 use', @metric_tech_desc='Forms (place & strip), beams, plywood, 1 use', @metric_book_desc='Forms (place & strip), beams, plywood, 1 use', @metric_assembly_book_desc='Forms (place & strip), beams, plywood, 1 use',
'@metric_unit='m2CA    ', @comment='', @unit_cost_skey=109032, @inhouse_total_cost_op='0.00', @inhouse_equip_cost_op='0.00', @inhouse_mat_cost_op='0.00', @inhouse_labor_cost_op='0.00', @res_labor_hour='0.25600', @res_total_cost_op='11.25', @res_equip_cost_op='0.43',
'@res_labor_cost_op='10.80', @res_mat_cost_op='', @res_total_cost='6.25', @res_equip_cost='0.05', @res_labor_cost='6.20', @res_mat_cost='', @metric_labor_hour='2.75565',
'@rr_labor_hour='0.25600', @opn_labor_hour='0.03200', @std_labor_hour='0.25600', @metric_equip_cost_op='4.63', @metric_labor_cost_op='168.00', @metric_mat_cost_op='', @rr_equip_cost_op='0.43', @rr_labor_cost_op='16.30', @rr_mat_cost_op='', @opn_equip_cost_op='0.43', @opn_labor_cost_op='',
'@opn_mat_cost_op='', @std_equip_cost_op='0.43', @std_labor_cost_op='15.60', @std_mat_cost_op='', @metric_total_cost_op='173.00', @rr_total_cost_op='16.75', @opn_total_cost_op='0.43', @std_total_cost_op='16.05', @metric_total_cost='108.22', @metric_equip_cost='4.22', @metric_labor_cost='104.00',
'@metric_mat_cost='', @rr_total_cost='10.09', @rr_equip_cost='0.39', @rr_labor_cost='9.70', @rr_mat_cost='', @opn_total_cost='0.39', @opn_equip_cost='0.39', @opn_labor_cost='', @opn_mat_cost='', @std_total_cost='10.09', @std_equip_cost='0.39', @std_labor_cost='9.70', @std_mat_cost=''
'
' @percent_flag='',  @last_update_person='Hancockrl', @bypass_ucd_ind = 0, @ucd_last_update_id=9,  @cstw_last_update_id=8, @update_material_usage_ind=1, @cost_change_ind=1
'</code>
'<code>
'exec sp_copy_output_usage @type = 'U', @FromSkey = '27810', @ToSkey='107709',  @last_update_date='7/31/2006 2:40:04 PM',  @last_update_person='Hancockrl',  @last_update_id='1'
'</code>
'<code>
'exec sp_update_object_attribute_value2 107709, 'U', 1, NULL, 'wire rope thimble', '', 'A', 'I', 0
'</code>
'<code>
'exec usp_select_attribute_value_ext @min_object_id = '050105100020%', @max_object_id = '', @skey_type = 'U', @meas_sys_cd = 'A', @obj_desc_filter = '', @master_format = 2004
'</code>
'<code>
'    exec sp_update_object_description 107709, 'U', 1
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'o frmUnitCostGrid.frm
'o CUCostMatMap.Cls
'o CGridMap.Cls
'o   CCDdal.CRSMDataAccess (
'Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>



Dim m_rec As ADODB.RecordSet
Dim m_rec2 As New ADODB.RecordSet   'Material Usage grid
Dim m_blnRecFlag As Boolean ' True if a populated RecordSet was passed, then we show data
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update


Dim m_blnNew As Boolean    ' Tells if we are doing a NEW unit cost line  'rlh 07/01/2008
Dim m_blnClone As Boolean  'Indicate if clone is in progress

Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
Dim strLast_unit_cost_id As String ' Holds last unit cost_id so we know if it changed
Dim m_objGridMap As New CUCostMatMap ' Class to handle grid
Dim m_recUsage As ADODB.RecordSet
Dim m_lngOriginalSkey As Long
Dim strOriginalCostID As String  ' Stores the original Unit Cost ID when cloning
Dim m_intMasterFormat As Long
Dim VALIDATE_CANCEL As Boolean      'rlh 05/16/2008

Public frmCallingForm As Form
'*** APEX Migration Utility Code Change ***
'Public tdbCols As TrueOleDBGrid60.Columns
'*** APEX Migration Utility Code Change ***
'Public tdbCols As TrueOleDBGrid70.Columns
Public tdbCols As TrueOleDBGrid80.Columns
'*** APEX Migration Utility Code Change ***
'Public myTDBGrid As TrueOleDBGrid60.TDBGrid
'*** APEX Migration Utility Code Change ***
'Public myTDBGrid As TrueOleDBGrid70.TDBGrid
Public myTDBGrid As TrueOleDBGrid80.TDBGrid
Dim tdbOldCols As Variant

Dim pct_ind_o As Integer
Dim type_code_o As String
Dim m_type_code As String
Dim crew_qty_o As String
Dim crew_id_o As String
Dim daily_output_o  As String
Dim metric_daily_output_o     As String
Dim unit_o As String
Dim metric_unit_o As String
Dim std_mat_cost_o As String
Dim std_labor_hour_o As String
Dim std_mat_cost_op_o As String
Dim std_labor_cost_o As String
Dim std_labor_cost_op_o As String
Dim std_equip_cost_o As String
Dim std_equip_cost_op_o As String
Dim std_total_cost_o As String
Dim std_total_cost_op_o As String
Dim opn_labor_hour_o As String
Dim opn_mat_cost_o As String
Dim opn_mat_cost_op_o As String
Dim opn_labor_cost_o As String
Dim opn_labor_cost_op_o As String
Dim opn_equip_cost_o As String
Dim opn_equip_cost_op_o As String
Dim opn_total_cost_o As String
Dim opn_total_cost_op_o As String
Dim rr_labor_hour_o As String
Dim rr_mat_cost_o As String
Dim rr_mat_cost_op_o As String
Dim rr_labor_cost_o As String
Dim rr_labor_cost_op_o As String
Dim rr_equip_cost_o As String
Dim rr_equip_cost_op_o As String
Dim rr_total_cost_o As String
Dim rr_total_cost_op_o As String
Dim metric_labor_hour_o As String
Dim metric_mat_cost_o As String
Dim metric_mat_cost_op_o As String
Dim metric_labor_cost_o As String
Dim metric_labor_cost_op_o As String
Dim metric_equip_cost_o As String
Dim metric_equip_cost_op_o As String
Dim metric_total_cost_o As String
Dim metric_total_cost_op_o As String
'RESI - rlh 05/06/2010
Dim res_labor_hour_o As String
Dim res_mat_cost_o As String
Dim res_mat_cost_op_o As String
Dim res_labor_cost_o As String
Dim res_labor_cost_op_o As String
Dim res_equip_cost_o As String
Dim res_equip_cost_op_o As String
Dim res_total_cost_o As String
Dim res_total_cost_op_o As String
'FMR
Dim inhouse_mat_cost_op_o As String
Dim inhouse_labor_cost_op_o As String
Dim inhouse_equip_cost_op_o As String
Dim inhouse_total_cost_op_o As String


'To track changes in individual total columns
Dim form_loading As Boolean
Dim std_total_cost_changed As Boolean
Dim rr_total_cost_changed As Boolean
Dim opn_total_cost_changed As Boolean
Dim metric_total_cost_changed As Boolean
Dim res_total_cost_changed As Boolean
Dim std_total_cost_op_changed As Boolean
Dim rr_total_cost_op_changed As Boolean
Dim opn_total_cost_op_changed As Boolean
Dim metric_total_cost_op_changed As Boolean
Dim res_total_cost_op_changed As Boolean

Dim std_NON_total_cost_changed As Boolean
Dim rr_NON_total_cost_changed As Boolean
Dim opn_NON_total_cost_changed As Boolean
Dim metric_NON_total_cost_changed As Boolean
Dim res_NON_total_cost_changed As Boolean
Dim std_NON_total_cost_op_changed As Boolean
Dim rr_NON_total_cost_op_changed As Boolean
Dim opn_NON_total_cost_op_changed As Boolean
Dim metric_NON_total_cost_op_changed As Boolean
Dim res_NON_total_cost_op_changed As Boolean



Public Property Get MasterFormat() As Long
    MasterFormat = m_intMasterFormat
End Property
Public Property Let MasterFormat(NewValue As Long)
    
    m_intMasterFormat = NewValue
    Me.txtMasterFormat.Text = m_intMasterFormat
    If m_intMasterFormat = EXT_MASTERFORMAT_VERSION Then
        lbl_ext_unit_cost_id.Caption = "Unit Cost ID " & Right(UCD_MASTERFORMAT_VERSION, 2) & ":"
        ext_unit_cost_id.Text = Format(Compress_String(ext_unit_cost_id.Text), FORMAT_UNIT_COST_SRV)
        unit_cost_id.Text = Format(Compress_String(unit_cost_id.Text), FORMAT_UNIT_COST_04_SRV)
    Else
        lbl_ext_unit_cost_id.Caption = "Unit Cost ID " & Right(EXT_MASTERFORMAT_VERSION, 2) & ":"
        ext_unit_cost_id.Text = Format(Compress_String(ext_unit_cost_id.Text), FORMAT_UNIT_COST_04_SRV)
        unit_cost_id.Text = Format(Compress_String(unit_cost_id.Text), FORMAT_UNIT_COST_SRV)
    End If
    
End Property

Private Sub RebindTDBGridNow()
    Dim oldRow As Variant
    oldRow = myTDBGrid.Bookmark
    myTDBGrid.ReBind
    myTDBGrid.Bookmark = oldRow
End Sub

Private Sub format_costs()
    std_equip_cost = Format(std_equip_cost, "##,##0.00")
    std_equip_cost_op = Format(std_equip_cost_op, "##,##0.00")
    std_labor_cost = Format(std_labor_cost, "##,##0.00")
    std_labor_cost_op = Format(std_labor_cost_op, "##,##0.00")
    std_mat_cost = Format(std_mat_cost, "##,##0.00")
    std_mat_cost_op = Format(std_mat_cost_op, "##,##0.00")
    std_total_cost = Format(std_total_cost, "##,##0.00")
    std_total_cost_op = Format(std_total_cost_op, "##,##0.00")
    opn_equip_cost = Format(opn_equip_cost, "##,##0.00")
    opn_equip_cost_op = Format(opn_equip_cost_op, "##,##0.00")
    opn_labor_cost = Format(opn_labor_cost, "##,##0.00")
    opn_labor_cost_op = Format(opn_labor_cost_op, "##,##0.00")
    opn_mat_cost = Format(opn_mat_cost, "##,##0.00")
    opn_mat_cost_op = Format(opn_mat_cost_op, "##,##0.00")
    opn_total_cost = Format(opn_total_cost, "##,##0.00")
    opn_total_cost_op = Format(opn_total_cost_op, "##,##0.00")
    rr_equip_cost = Format(rr_equip_cost, "##,##0.00")
    rr_equip_cost_op = Format(rr_equip_cost_op, "##,##0.00")
    rr_labor_cost = Format(rr_labor_cost, "##,##0.00")
    rr_labor_cost_op = Format(rr_labor_cost_op, "##,##0.00")
    rr_mat_cost = Format(rr_mat_cost, "##,##0.00")
    rr_mat_cost_op = Format(rr_mat_cost_op, "##,##0.00")
    rr_total_cost = Format(rr_total_cost, "##,##0.00")
    rr_total_cost_op = Format(rr_total_cost_op, "##,##0.00")
    metric_equip_cost = Format(metric_equip_cost, "##,##0.00")
    metric_equip_cost_op = Format(metric_equip_cost_op, "##,##0.00")
    metric_labor_cost = Format(metric_labor_cost, "##,##0.00")
    metric_labor_cost_op = Format(metric_labor_cost_op, "##,##0.00")
    metric_mat_cost = Format(metric_mat_cost, "##,##0.00")
    metric_mat_cost_op = Format(metric_mat_cost_op, "##,##0.00")
    metric_total_cost = Format(metric_total_cost, "##,##0.00")
    metric_total_cost_op = Format(metric_total_cost_op, "##,##0.00")
    'RESI  - rlh 05/06/2010
    res_equip_cost = Format(res_equip_cost, "##,##0.00")
    res_equip_cost_op = Format(res_equip_cost_op, "##,##0.00")
    res_labor_cost = Format(res_labor_cost, "##,##0.00")
    res_labor_cost_op = Format(res_labor_cost_op, "##,##0.00")
    res_mat_cost = Format(res_mat_cost, "##,##0.00")
    res_mat_cost_op = Format(res_mat_cost_op, "##,##0.00")
    res_total_cost = Format(res_total_cost, "##,##0.00")
    res_total_cost_op = Format(res_total_cost_op, "##,##0.00")
    'FMR (in-house)
    inhouse_mat_cost_op = Format(inhouse_mat_cost_op, "##,##0.00")
    inhouse_labor_cost_op = Format(inhouse_labor_cost_op, "##,##0.00")
    inhouse_equip_cost_op = Format(inhouse_equip_cost_op, "##,##0.00")
    inhouse_total_cost_op = Format(inhouse_total_cost_op, "##,##0.00")
    
End Sub

Private Sub m_rec_unformatfields()
    m_rec.Fields("std_equip_cost") = Format(std_equip_cost, "####0.00")
    m_rec.Fields("std_equip_cost_op") = Format(std_equip_cost_op, "####0.00")
    m_rec.Fields("std_labor_cost") = Format(std_labor_cost, "####0.00")
    m_rec.Fields("std_labor_cost_op") = Format(std_labor_cost_op, "####0.00")
    m_rec.Fields("std_mat_cost") = Format(std_mat_cost, "####0.00")
    m_rec.Fields("std_mat_cost_op") = Format(std_mat_cost_op, "####0.00")
    m_rec.Fields("std_total_cost") = Format(std_total_cost, "####0.00")
    m_rec.Fields("std_total_cost_op") = Format(std_total_cost_op, "####0.00")
    m_rec.Fields("opn_equip_cost") = Format(opn_equip_cost, "####0.00")
    m_rec.Fields("opn_equip_cost_op") = Format(opn_equip_cost_op, "####0.00")
    m_rec.Fields("opn_labor_cost") = Format(opn_labor_cost, "####0.00")
    m_rec.Fields("opn_labor_cost_op") = Format(opn_labor_cost_op, "####0.00")
    m_rec.Fields("opn_mat_cost") = Format(opn_mat_cost, "####0.00")
    m_rec.Fields("opn_mat_cost_op") = Format(opn_mat_cost_op, "####0.00")
    m_rec.Fields("opn_total_cost") = Format(opn_total_cost, "####0.00")
    m_rec.Fields("opn_total_cost_op") = Format(opn_total_cost_op, "####0.00")
    m_rec.Fields("rr_equip_cost") = Format(rr_equip_cost, "####0.00")
    m_rec.Fields("rr_equip_cost_op") = Format(rr_equip_cost_op, "####0.00")
    m_rec.Fields("rr_labor_cost") = Format(rr_labor_cost, "####0.00")
    m_rec.Fields("rr_labor_cost_op") = Format(rr_labor_cost_op, "####0.00")
    m_rec.Fields("rr_mat_cost") = Format(rr_mat_cost, "####0.00")
    m_rec.Fields("rr_mat_cost_op") = Format(rr_mat_cost_op, "####0.00")
    m_rec.Fields("rr_total_cost") = Format(rr_total_cost, "####0.00")
    m_rec.Fields("rr_total_cost_op") = Format(rr_total_cost_op, "####0.00")
    m_rec.Fields("metric_equip_cost") = Format(metric_equip_cost, "####0.00")
    m_rec.Fields("metric_equip_cost_op") = Format(metric_equip_cost_op, "##,##0.00")
    m_rec.Fields("metric_labor_cost") = Format(metric_labor_cost, "####0.00")
    m_rec.Fields("metric_labor_cost_op") = Format(metric_labor_cost_op, "####0.00")
    m_rec.Fields("metric_mat_cost") = Format(metric_mat_cost, "####0.00")
    m_rec.Fields("metric_mat_cost_op") = Format(metric_mat_cost_op, "####0.00")
    m_rec.Fields("metric_total_cost") = Format(metric_total_cost, "####0.00")
    m_rec.Fields("metric_total_cost_op") = Format(metric_total_cost_op, "####0.00")
    'RESI
    m_rec.Fields("res_equip_cost") = Format(res_equip_cost, "####0.00")
    m_rec.Fields("res_equip_cost_op") = Format(res_equip_cost_op, "####0.00")
    m_rec.Fields("res_labor_cost") = Format(res_labor_cost, "####0.00")
    m_rec.Fields("res_labor_cost_op") = Format(res_labor_cost_op, "####0.00")
    m_rec.Fields("res_mat_cost") = Format(res_mat_cost, "####0.00")
    m_rec.Fields("res_mat_cost_op") = Format(res_mat_cost_op, "####0.00")
    m_rec.Fields("res_total_cost") = Format(res_total_cost, "####0.00")
    m_rec.Fields("res_total_cost_op") = Format(res_total_cost_op, "####0.00")
    'FMR (in-house)
    m_rec.Fields("inhouse_mat_cost_op") = Format(inhouse_mat_cost_op, "####0.00")
    m_rec.Fields("inhouse_equip_cost_op") = Format(inhouse_equip_cost_op, "####0.00")
    m_rec.Fields("inhouse_labor_cost_op") = Format(inhouse_labor_cost_op, "####0.00")
    m_rec.Fields("inhouse_total_cost_op") = Format(inhouse_total_cost_op, "####0.00")
    
End Sub

Private Sub SaveOrigCost()
'Save Original Values
    If m_blnClone = True Or m_blnInsert = True Then
        type_code_o = ""
    Else
        type_code_o = type_code.Text
    End If
    crew_qty_o = crew_qty.Text
    set_pct_ind
    pct_ind_o = pct_ind
    crew_id_o = crew_id.Text
    daily_output_o = daily_output.Text
    metric_daily_output_o = metric_daily_output.Text
    unit_o = unit.Text
    metric_unit_o = metric_unit.Text
    std_mat_cost_o = std_mat_cost.Text
    std_labor_cost_o = std_labor_cost.Text
    std_labor_hour_o = std_labor_hour.Text
    std_mat_cost_op_o = std_mat_cost_op.Text
    std_labor_cost_o = std_labor_cost.Text
    std_labor_cost_op_o = std_labor_cost_op.Text
    std_equip_cost_o = std_equip_cost.Text
    std_equip_cost_op_o = std_equip_cost_op.Text
    std_total_cost_o = std_total_cost.Text
    std_total_cost_op_o = std_total_cost_op.Text
    opn_labor_hour_o = opn_labor_hour.Text
    opn_mat_cost_o = opn_mat_cost.Text
    opn_mat_cost_op_o = opn_mat_cost_op.Text
    opn_labor_cost_o = opn_labor_cost.Text
    opn_labor_cost_op_o = opn_labor_cost_op.Text
    opn_equip_cost_o = opn_equip_cost.Text
    opn_equip_cost_op_o = opn_equip_cost_op.Text
    opn_total_cost_o = opn_total_cost.Text
    opn_total_cost_op_o = opn_total_cost_op.Text
    rr_labor_hour_o = rr_labor_hour.Text
    rr_mat_cost_o = rr_mat_cost.Text
    rr_mat_cost_op_o = rr_mat_cost_op.Text
    rr_labor_cost_o = rr_labor_cost.Text
    rr_labor_cost_op_o = rr_labor_cost_op.Text
    rr_equip_cost_o = rr_equip_cost.Text
    rr_equip_cost_op_o = rr_equip_cost_op.Text
    rr_total_cost_o = rr_total_cost.Text
    rr_total_cost_op_o = rr_total_cost_op.Text
    metric_labor_hour_o = metric_labor_hour.Text
    metric_mat_cost_o = metric_mat_cost.Text
    metric_mat_cost_op_o = metric_mat_cost_op.Text
    metric_labor_cost_o = metric_labor_cost.Text
    metric_labor_cost_op_o = metric_labor_cost_op.Text
    metric_equip_cost_o = metric_equip_cost.Text
    metric_equip_cost_op_o = metric_equip_cost_op.Text
    metric_total_cost_o = metric_total_cost.Text
    metric_total_cost_op_o = metric_total_cost_op.Text
    'RESI
    res_labor_hour_o = res_labor_hour.Text
    res_mat_cost_o = res_mat_cost.Text
    res_mat_cost_op_o = res_mat_cost_op.Text
    res_labor_cost_o = res_labor_cost.Text
    res_labor_cost_op_o = res_labor_cost_op.Text
    res_equip_cost_o = res_equip_cost.Text
    res_equip_cost_op_o = res_equip_cost_op.Text
    res_total_cost_o = res_total_cost.Text
    res_total_cost_op_o = res_total_cost_op.Text
    'FMR (In-House)
    inhouse_mat_cost_op_o = inhouse_mat_cost_op.Text
    inhouse_equip_cost_op_o = inhouse_equip_cost_op.Text
    inhouse_labor_cost_op_o = inhouse_labor_cost_op.Text
    inhouse_total_cost_op_o = inhouse_total_cost_op.Text
    
    

End Sub

Private Function SetCostChange() As Integer
'Set the unit cost change flag to determine if a new cost or exception recrd is to be generated
    
    SetCostChange = 0
    If pct_ind_o <> pct_ind Then SetCostChange = 1
    If type_code.Text <> type_code_o Then SetCostChange = 1
    If crew_qty.Text <> crew_qty_o Then SetCostChange = 1
    If crew_id.Text <> crew_id_o Then SetCostChange = 1
    If daily_output.Text <> daily_output_o Then SetCostChange = 1
    If metric_daily_output.Text <> metric_daily_output_o Then SetCostChange = 1
    If unit.Text <> unit_o Then SetCostChange = 1
    If metric_unit.Text <> metric_unit_o Then SetCostChange = 1
    If std_mat_cost.Text <> std_mat_cost_o Then SetCostChange = 1
    If std_labor_cost.Text <> std_labor_cost_o Then SetCostChange = 1
    If std_labor_hour.Text <> std_labor_hour_o Then SetCostChange = 1
    If std_mat_cost_op.Text <> std_mat_cost_op_o Then SetCostChange = 1
    If std_labor_cost.Text <> std_labor_cost_o Then SetCostChange = 1
    If std_labor_cost_op.Text <> std_labor_cost_op_o Then SetCostChange = 1
    If std_equip_cost.Text <> std_equip_cost_o Then SetCostChange = 1
    If std_equip_cost_op.Text <> std_equip_cost_op_o Then SetCostChange = 1
    If std_total_cost.Text <> std_total_cost_o Then SetCostChange = 1
    If std_total_cost_op.Text <> std_total_cost_op_o Then SetCostChange = 1
    If opn_labor_hour.Text <> opn_labor_hour_o Then SetCostChange = 1
    If opn_mat_cost.Text <> opn_mat_cost_o Then SetCostChange = 1
    If opn_mat_cost_op.Text <> opn_mat_cost_op_o Then SetCostChange = 1
    If opn_labor_cost.Text <> opn_labor_cost_o Then SetCostChange = 1
    If opn_labor_cost_op.Text <> opn_labor_cost_op_o Then SetCostChange = 1
    If opn_equip_cost.Text <> opn_equip_cost_o Then SetCostChange = 1
    If opn_equip_cost_op.Text <> opn_equip_cost_op_o Then SetCostChange = 1
    If opn_total_cost.Text <> opn_total_cost_o Then SetCostChange = 1
    If opn_total_cost_op.Text <> opn_total_cost_op_o Then SetCostChange = 1
    If rr_labor_hour.Text <> rr_labor_hour_o Then SetCostChange = 1
    If rr_mat_cost.Text <> rr_mat_cost_o Then SetCostChange = 1
    If rr_mat_cost_op.Text <> rr_mat_cost_op_o Then SetCostChange = 1
    If rr_labor_cost.Text <> rr_labor_cost_o Then SetCostChange = 1
    If rr_labor_cost_op.Text <> rr_labor_cost_op_o Then SetCostChange = 1
    If rr_equip_cost.Text <> rr_equip_cost_o Then SetCostChange = 1
    If rr_equip_cost_op.Text <> rr_equip_cost_op_o Then SetCostChange = 1
    If rr_total_cost.Text <> rr_total_cost_o Then SetCostChange = 1
    If rr_total_cost_op.Text <> rr_total_cost_op_o Then SetCostChange = 1
    If metric_labor_hour.Text <> metric_labor_hour_o Then SetCostChange = 1
    If metric_mat_cost.Text <> metric_mat_cost_o Then SetCostChange = 1
    If metric_mat_cost_op.Text <> metric_mat_cost_op_o Then SetCostChange = 1
    If metric_labor_cost.Text <> metric_labor_cost_o Then SetCostChange = 1
    If metric_labor_cost_op.Text <> metric_labor_cost_op_o Then SetCostChange = 1
    If metric_equip_cost.Text <> metric_equip_cost_o Then SetCostChange = 1
    If metric_equip_cost_op.Text <> metric_equip_cost_op_o Then SetCostChange = 1
    If metric_total_cost.Text <> metric_total_cost_o Then SetCostChange = 1
    If metric_total_cost_op.Text <> metric_total_cost_op_o Then SetCostChange = 1
    'RESI - rlh 05/06/2010
    If res_labor_hour.Text <> res_labor_hour_o Then SetCostChange = 1
    If res_mat_cost.Text <> res_mat_cost_o Then SetCostChange = 1
    If res_mat_cost_op.Text <> res_mat_cost_op_o Then SetCostChange = 1
    If res_labor_cost.Text <> res_labor_cost_o Then SetCostChange = 1
    If res_labor_cost_op.Text <> res_labor_cost_op_o Then SetCostChange = 1
    If res_equip_cost.Text <> res_equip_cost_o Then SetCostChange = 1
    If res_equip_cost_op.Text <> res_equip_cost_op_o Then SetCostChange = 1
    If res_total_cost.Text <> res_total_cost_o Then SetCostChange = 1
    If res_total_cost_op.Text <> res_total_cost_op_o Then SetCostChange = 1
    'FMR (In-house) - rlh 05/14/2010
    If inhouse_mat_cost_op.Text <> inhouse_mat_cost_op_o Then SetCostChange = 1
    If inhouse_equip_cost_op.Text <> inhouse_equip_cost_op_o Then SetCostChange = 1
    If inhouse_labor_cost_op.Text <> inhouse_labor_cost_op_o Then SetCostChange = 1
    If inhouse_total_cost_op.Text <> inhouse_total_cost_op_o Then SetCostChange = 1
    
End Function

Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
' Fill all fields with data
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If Not m_rec.Fields("unit_cost_skey") = 0 Then
        m_blnRecFlag = True
    End If
    
    'unit_cost_id.Text = Format(unit_cost_id.Text, FORMAT_UNIT_COST_SRV)
    'ext_unit_cost_id.Text = Format(ext_unit_cost_id.Text, FORMAT_UNIT_COST_04_SRV)

End Sub

Private Function validate_crew_qty() As Boolean
    validate_crew_qty = True
    If Trim(crew_qty) <> "" Then
        If IsNumeric(crew_qty) = False Then        'qty required
            MsgBox "Please enter a numeric quantity."
            validate_crew_qty = False
        ElseIf crew_qty <= 0 Then
            MsgBox "Please enter a valid quantity."
            validate_crew_qty = False
        ElseIf crew_type_code = "L" And Val(crew_qty.Text) <> CLng(crew_qty.Text) Then
            MsgBox "Please enter a whole number for this type of crew."
            validate_crew_qty = False
        End If
        If validate_crew_qty = False Then
            crew_qty.SetFocus
        End If
    End If
End Function

Private Sub alt_unit_cost_id_Validate(Cancel As Boolean)
    Dim bln_New As Boolean
    If alt_unit_cost_id <> "" Then
        If m_blnInsert Or m_blnClone Then
            bln_New = True
        End If
        If Invalid_ID_Format(Compress_String(alt_unit_cost_id), "alt_unit_cost_id", m_rec, bln_New) = True Then
            Cancel = True
        End If
    End If
End Sub

Private Sub assembly_book_desc_Change()
    Dim intLength As Integer
    Dim intPosition As Integer
    Dim txtSaveassembly_book_desc As String
    Dim txtNewassembly_book_desc As String
    
    If Len(assembly_book_desc) > 0 Then
        intPosition = assembly_book_desc.SelStart
        If intPosition > 0 Then
            If Asc(Mid(assembly_book_desc, intPosition, 1)) >= 0 And Asc(Mid(assembly_book_desc, intPosition, 1)) <= 31 Then
                intLength = Len(assembly_book_desc)
                txtSaveassembly_book_desc = assembly_book_desc.Text
                MsgBox "Non-printable characters are not allowed in the assembly_book_description."
                txtNewassembly_book_desc = Left(txtSaveassembly_book_desc, intPosition - 2)
                If intPosition < intLength Then
                    txtNewassembly_book_desc = txtNewassembly_book_desc + Right(txtSaveassembly_book_desc, intLength - intPosition)
                End If
                assembly_book_desc.Text = txtNewassembly_book_desc
                assembly_book_desc.SelStart = intPosition - 2
            End If
        End If
    End If

End Sub

Private Sub book_desc_Change()
    Dim intLength As Integer
    Dim intPosition As Integer
    Dim txtSavebook_desc As String
    Dim txtNewbook_desc As String
    
    If Len(book_desc) > 0 Then
        intPosition = book_desc.SelStart
        If intPosition > 0 Then
            If Asc(Mid(book_desc, intPosition, 1)) >= 0 And Asc(Mid(book_desc, intPosition, 1)) <= 31 Then
                intLength = Len(book_desc)
                txtSavebook_desc = book_desc.Text
                MsgBox "Non-printable characters are not allowed in the book_description."
                txtNewbook_desc = Left(txtSavebook_desc, intPosition - 2)
                If intPosition < intLength Then
                    txtNewbook_desc = txtNewbook_desc + Right(txtSavebook_desc, intLength - intPosition)
                End If
                book_desc.Text = txtNewbook_desc
                book_desc.SelStart = intPosition - 2
            End If
        End If
    End If
    UpdateBookPreviewLine
    
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String

    Dim varButton
    varButton = MsgBox("Are you sure you want to delete? The CSI Line will be removed. Press the Material Usage delete button to remove a material usage.", vbYesNo + vbCritical)
    If varButton = vbNo Then
        Exit Sub
    End If

    strUpdate = "exec sp_delete_unit_cost "
    strUpdate = strUpdate + "@unit_cost_skey=" + str(Me.Controls("unit_cost_skey")) + ","
    strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    
    blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
    Else
        MsgBox "Delete successful."
        m_rec.Delete
        RebindTDBGridNow
        m_blnDeleted = True
        Unload Me
    End If
End Sub

Private Sub Store_Grid_Old_Values()
    Dim i As Integer
    ReDim tdbOldCols(tdbCols.Count - 1)
    For i = 0 To tdbCols.Count - 1
        tdbOldCols(i) = tdbCols.Item(i).Value
    Next
End Sub

Private Sub RestoreGridValues()
    ' this restores the grid back to its positioin if the user did not choose to save
    Dim i As Integer
    On Error Resume Next
    If m_blnInsert = False Then
        For i = 1 To tdbCols.Count - 1
            If tdbCols.Item(i).Value <> tdbOldCols(i) Then
                tdbCols.Item(i).Value = tdbOldCols(i)
            End If
        Next i
        'myTDBGrid.RefetchRow
        myTDBGrid.DataChanged = False
        myTDBGrid.RefreshRow
        DoEvents
    End If
End Sub

Private Sub cmdLongDesc_Click()
    Dim frm As frmLongDescriptionGrid
    
    Set frm = New frmLongDescriptionGrid
    'ADDED 9/7/2005 RTD - SEND MASTERFORMAT OF UNIT_COST_ID TO LONG DESCRIPTION GRID
    frm.MasterFormat = MasterFormat
    frm.JumpIn Compress_String(Me.unit_cost_id.Text)
    frm.Show
    
End Sub

Private Sub cmdMatUsageDelete_Click()
    
    On Error Resume Next
'    Dim varButton
'    varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
'    If varButton = vbYes Then
        If TDBGrid.AddNewMode > 0 Then
            TDBGrid.ReBind
        Else
            TDBGrid.Delete
        End If
'    End If
End Sub

Private Sub cmdShowRounding_Click()

    Dim frm As frmRounding
    Set frm = New frmRounding
    Set frm.frmCallingForm = Me
    
    
      frm.std_mat_cost(0).Text = std_mat_cost.Text
      frm.std_labor_cost(0).Text = std_labor_cost.Text
      frm.std_equip_cost(0).Text = std_equip_cost.Text
      frm.std_total_cost(0).Text = std_total_cost.Text


      frm.rr_mat_cost(0).Text = rr_mat_cost.Text
      frm.rr_labor_cost(0).Text = rr_labor_cost.Text
      frm.rr_equip_cost(0).Text = rr_equip_cost.Text
      frm.rr_total_cost(0).Text = rr_total_cost.Text

 
      frm.opn_mat_cost(0).Text = opn_mat_cost.Text
      frm.opn_labor_cost(0).Text = opn_labor_cost.Text
      frm.opn_equip_cost(0).Text = opn_equip_cost.Text
      frm.opn_total_cost(0).Text = opn_total_cost.Text


      frm.metric_mat_cost(0).Text = metric_mat_cost.Text
      frm.metric_labor_cost(0).Text = metric_labor_cost.Text
      frm.metric_equip_cost(0).Text = metric_equip_cost.Text
      frm.metric_total_cost(0).Text = metric_total_cost.Text


      frm.res_mat_cost(0).Text = res_mat_cost.Text
      frm.res_labor_cost(0).Text = res_labor_cost.Text
      frm.res_equip_cost(0).Text = res_equip_cost.Text
      frm.res_total_cost(0).Text = res_total_cost.Text


      frm.std_mat_cost_op(0).Text = std_mat_cost_op.Text
      frm.std_labor_cost_op(0).Text = std_labor_cost_op.Text
      frm.std_equip_cost_op(0).Text = std_equip_cost_op.Text
      frm.std_total_cost_op(0).Text = std_total_cost_op.Text

      frm.rr_mat_cost_op(0).Text = rr_mat_cost_op.Text
      frm.rr_labor_cost_op(0).Text = rr_labor_cost_op.Text
      frm.rr_equip_cost_op(0).Text = rr_equip_cost_op.Text
      frm.rr_total_cost_op(0).Text = rr_total_cost_op.Text

      frm.opn_mat_cost_op(0).Text = opn_mat_cost_op.Text
      frm.opn_labor_cost_op(0).Text = opn_labor_cost_op.Text
      frm.opn_equip_cost_op(0).Text = opn_equip_cost_op.Text
      frm.opn_total_cost_op(0).Text = opn_total_cost_op.Text

      frm.metric_mat_cost_op(0).Text = metric_mat_cost_op.Text
      frm.metric_labor_cost_op(0).Text = metric_labor_cost_op.Text
      frm.metric_equip_cost_op(0).Text = metric_equip_cost_op.Text
      frm.metric_total_cost_op(0).Text = metric_total_cost_op.Text


      frm.res_mat_cost_op(0).Text = res_mat_cost_op.Text
      frm.res_labor_cost_op(0).Text = res_labor_cost_op.Text
      frm.res_equip_cost_op(0).Text = res_equip_cost_op.Text
      frm.res_total_cost_op(0).Text = res_total_cost_op.Text
    
    
    
    frm.Show vbModal
    
    

End Sub


Private Function ApplyRoundingRules() As Integer

    
    Dim retStr As String
    retStr = ""
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
        
        
    dbl_std_mat_cost = Val(std_mat_cost.Text)
    dbl_std_labor_cost = Val(std_labor_cost.Text)
    dbl_std_equip_cost = Val(std_equip_cost.Text)
    dbl_std_total_cost = Val(std_total_cost.Text)
    dbl_rr_mat_cost = Val(rr_mat_cost.Text)
    dbl_rr_labor_cost = Val(rr_labor_cost.Text)
    dbl_rr_equip_cost = Val(rr_equip_cost.Text)
    dbl_rr_total_cost = Val(rr_total_cost.Text)
    dbl_opn_mat_cost = Val(opn_mat_cost.Text)
    dbl_opn_labor_cost = Val(opn_labor_cost.Text)
    dbl_opn_equip_cost = Val(opn_equip_cost.Text)
    dbl_opn_total_cost = Val(opn_total_cost.Text)
    dbl_metric_mat_cost = Val(metric_mat_cost.Text)
    dbl_metric_labor_cost = Val(metric_labor_cost.Text)
    dbl_metric_equip_cost = Val(metric_equip_cost.Text)
    dbl_metric_total_cost = Val(metric_total_cost.Text)
    dbl_res_mat_cost = Val(res_mat_cost.Text)
    dbl_res_labor_cost = Val(res_labor_cost.Text)
    dbl_res_equip_cost = Val(res_equip_cost.Text)
    dbl_res_total_cost = Val(res_total_cost.Text)
    dbl_std_mat_cost_op = Val(std_mat_cost_op.Text)
    dbl_std_labor_cost_op = Val(std_labor_cost_op.Text)
    dbl_std_equip_cost_op = Val(std_equip_cost_op.Text)
    dbl_std_total_cost_op = Val(std_total_cost_op.Text)
    dbl_rr_mat_cost_op = Val(rr_mat_cost_op.Text)
    dbl_rr_labor_cost_op = Val(rr_labor_cost_op.Text)
    dbl_rr_equip_cost_op = Val(rr_equip_cost_op.Text)
    dbl_rr_total_cost_op = Val(rr_total_cost_op.Text)
    dbl_opn_mat_cost_op = Val(opn_mat_cost_op.Text)
    dbl_opn_labor_cost_op = Val(opn_labor_cost_op.Text)
    dbl_opn_equip_cost_op = Val(opn_equip_cost_op.Text)
    dbl_opn_total_cost_op = Val(opn_total_cost_op.Text)
    dbl_metric_mat_cost_op = Val(metric_mat_cost_op.Text)
    dbl_metric_labor_cost_op = Val(metric_labor_cost_op.Text)
    dbl_metric_equip_cost_op = Val(metric_equip_cost_op.Text)
    dbl_metric_total_cost_op = Val(metric_total_cost_op.Text)
    dbl_res_mat_cost_op = Val(res_mat_cost_op.Text)
    dbl_res_labor_cost_op = Val(res_labor_cost_op.Text)
    dbl_res_equip_cost_op = Val(res_equip_cost_op.Text)
    dbl_res_total_cost_op = Val(res_total_cost_op.Text)
        

    
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
        
        
        
        If Trim(std_mat_cost.Text) <> "" And Val(std_mat_cost.Text) <> 0 Then
        std_mat_cost.Text = Format(dbl_std_mat_cost, ReplaceCharactersForFormat(std_mat_cost.Text))
        Else
        std_mat_cost.Text = std_mat_cost.Text
        End If
        If Trim(std_labor_cost.Text) <> "" And Val(std_labor_cost.Text) <> 0 Then
        std_labor_cost.Text = Format(dbl_std_labor_cost, ReplaceCharactersForFormat(std_labor_cost.Text))
        Else
        std_labor_cost.Text = std_labor_cost.Text
        End If
        If Trim(std_equip_cost.Text) <> "" And Val(std_equip_cost.Text) <> 0 Then
        std_equip_cost.Text = Format(dbl_std_equip_cost, ReplaceCharactersForFormat(std_equip_cost.Text))
        Else
        std_equip_cost.Text = std_equip_cost.Text
        End If
        If (Trim(std_total_cost.Text) <> "" And Val(std_total_cost.Text) <> 0) Or (dbl_std_mat_cost <> 0 Or dbl_std_labor_cost <> 0 Or dbl_std_equip_cost <> 0) Then
            'If Trim(std_total_cost.Text <> "") Then
            If dbl_std_total_cost <> 0 Then
'                std_total_cost.Text = Format(dbl_std_total_cost, ReplaceCharactersForFormat(std_total_cost.Text))
'            Else
                std_total_cost.Text = Format(dbl_std_total_cost, ReplaceCharactersForFormat(CStr(dbl_std_total_cost)))
            End If
        Else
        std_total_cost.Text = std_total_cost.Text
        End If
        If Trim(rr_mat_cost.Text) <> "" And Val(rr_mat_cost.Text) <> 0 Then
        rr_mat_cost.Text = Format(dbl_rr_mat_cost, ReplaceCharactersForFormat(rr_mat_cost.Text))
        Else
        rr_mat_cost.Text = rr_mat_cost.Text
        End If
        If Trim(rr_labor_cost.Text) <> "" And Val(rr_labor_cost.Text) <> 0 Then
        rr_labor_cost.Text = Format(dbl_rr_labor_cost, ReplaceCharactersForFormat(rr_labor_cost.Text))
        Else
        rr_labor_cost.Text = rr_labor_cost.Text
        End If
        If Trim(rr_equip_cost.Text) <> "" And Val(rr_equip_cost.Text) <> 0 Then
        rr_equip_cost.Text = Format(dbl_rr_equip_cost, ReplaceCharactersForFormat(rr_equip_cost.Text))
        Else
        rr_equip_cost.Text = rr_equip_cost.Text
        End If
        If (Trim(rr_total_cost.Text) <> "" And Val(rr_total_cost.Text) <> 0) Or (dbl_rr_mat_cost <> 0 Or dbl_rr_labor_cost <> 0 Or dbl_rr_equip_cost <> 0) Then

            'If Trim(rr_total_cost.Text) <> "" Then
            If dbl_rr_total_cost <> 0 Then
'                rr_total_cost.Text = Format(dbl_rr_total_cost, ReplaceCharactersForFormat(rr_total_cost.Text))
'            Else
                rr_total_cost.Text = Format(dbl_rr_total_cost, ReplaceCharactersForFormat(CStr(dbl_rr_total_cost)))
            End If
        Else
        rr_total_cost.Text = rr_total_cost.Text
        End If
        If Trim(opn_mat_cost.Text) <> "" And Val(opn_mat_cost.Text) <> 0 Then
        opn_mat_cost.Text = Format(dbl_opn_mat_cost, ReplaceCharactersForFormat(opn_mat_cost.Text))
        Else
        opn_mat_cost.Text = opn_mat_cost.Text
        End If
        If Trim(opn_labor_cost.Text) <> "" And Val(opn_labor_cost.Text) <> 0 Then
        opn_labor_cost.Text = Format(dbl_opn_labor_cost, ReplaceCharactersForFormat(opn_labor_cost.Text))
        Else
        opn_labor_cost.Text = opn_labor_cost.Text
        End If
        If Trim(opn_equip_cost.Text) <> "" And Val(opn_equip_cost.Text) <> 0 Then
        opn_equip_cost.Text = Format(dbl_opn_equip_cost, ReplaceCharactersForFormat(opn_equip_cost.Text))
        Else
        opn_equip_cost.Text = opn_equip_cost.Text
        End If
        If (Trim(opn_total_cost.Text) <> "" And Val(opn_total_cost.Text) <> 0) Or (dbl_opn_mat_cost <> 0 Or dbl_opn_labor_cost <> 0 Or dbl_opn_equip_cost <> 0) Then
            'If Trim(opn_total_cost.Text) <> "" Then
            If dbl_opn_total_cost <> 0 Then
'            opn_total_cost.Text = Format(dbl_opn_total_cost, ReplaceCharactersForFormat(opn_total_cost.Text))
'            Else
            opn_total_cost.Text = Format(dbl_opn_total_cost, ReplaceCharactersForFormat(CStr(dbl_opn_total_cost)))
            End If
        Else
        opn_total_cost.Text = opn_total_cost.Text
        End If
        If Trim(metric_mat_cost.Text) <> "" And Val(metric_mat_cost.Text) <> 0 Then
        metric_mat_cost.Text = Format(dbl_metric_mat_cost, ReplaceCharactersForFormat(metric_mat_cost.Text))
        Else
        metric_mat_cost.Text = metric_mat_cost.Text
        End If
        If Trim(metric_labor_cost.Text) <> "" And Val(metric_labor_cost.Text) <> 0 Then
        metric_labor_cost.Text = Format(dbl_metric_labor_cost, ReplaceCharactersForFormat(metric_labor_cost.Text))
        Else
        metric_labor_cost.Text = metric_labor_cost.Text
        End If
        If Trim(metric_equip_cost.Text) <> "" And Val(metric_equip_cost.Text) <> 0 Then
        metric_equip_cost.Text = Format(dbl_metric_equip_cost, ReplaceCharactersForFormat(metric_equip_cost.Text))
        Else
        metric_equip_cost.Text = metric_equip_cost.Text
        End If
        If (Trim(metric_total_cost.Text) <> "" And Val(metric_total_cost.Text) <> 0) Or (dbl_metric_mat_cost <> 0 Or dbl_metric_labor_cost <> 0 Or dbl_metric_equip_cost <> 0) Then
               'If Trim(metric_total_cost.Text) <> "" Then
               If dbl_metric_total_cost <> 0 Then
'                metric_total_cost.Text = Format(dbl_metric_total_cost, ReplaceCharactersForFormat(metric_total_cost.Text))
'            Else
                metric_total_cost.Text = Format(dbl_metric_total_cost, ReplaceCharactersForFormat(CStr(dbl_metric_total_cost)))
            End If
        Else
        metric_total_cost.Text = metric_total_cost.Text
        End If
        If Trim(res_mat_cost.Text) <> "" And Val(res_mat_cost.Text) <> 0 Then
        res_mat_cost.Text = Format(dbl_res_mat_cost, ReplaceCharactersForFormat(res_mat_cost.Text))
        Else
        res_mat_cost.Text = res_mat_cost.Text
        End If
        If Trim(res_labor_cost.Text) <> "" And Val(res_labor_cost.Text) <> 0 Then
        res_labor_cost.Text = Format(dbl_res_labor_cost, ReplaceCharactersForFormat(res_labor_cost.Text))
        Else
        res_labor_cost.Text = res_labor_cost.Text
        End If
        If Trim(res_equip_cost.Text) <> "" And Val(res_equip_cost.Text) <> 0 Then
        res_equip_cost.Text = Format(dbl_res_equip_cost, ReplaceCharactersForFormat(res_equip_cost.Text))
        Else
        res_equip_cost.Text = res_equip_cost.Text
        End If
        If (Trim(res_total_cost.Text) <> "" And Val(res_total_cost.Text) <> 0) Or (dbl_res_mat_cost <> 0 Or dbl_res_labor_cost <> 0 Or dbl_res_equip_cost <> 0) Then
            'If (Trim(res_total_cost.Text) <> "") Then
            If dbl_res_total_cost <> 0 Then
'                res_total_cost.Text = Format(dbl_res_total_cost, ReplaceCharactersForFormat(res_total_cost.Text))
'            Else
                res_total_cost.Text = Format(dbl_res_total_cost, ReplaceCharactersForFormat(CStr(dbl_res_total_cost)))
            End If
        Else
        res_total_cost.Text = res_total_cost.Text
        End If
        
        
        If Trim(std_mat_cost_op.Text) <> "" And Val(std_mat_cost_op.Text) <> 0 Then
        std_mat_cost_op.Text = Format(dbl_std_mat_cost_op, ReplaceCharactersForFormat(std_mat_cost_op.Text))
        Else
        std_mat_cost_op.Text = std_mat_cost_op.Text
        End If
        If Trim(std_labor_cost_op.Text) <> "" And Val(std_labor_cost_op.Text) <> 0 Then
        std_labor_cost_op.Text = Format(dbl_std_labor_cost_op, ReplaceCharactersForFormat(std_labor_cost_op.Text))
        Else
        std_labor_cost_op.Text = std_labor_cost_op.Text
        End If
        If Trim(std_equip_cost_op.Text) <> "" And Val(std_equip_cost_op.Text) <> 0 Then
        std_equip_cost_op.Text = Format(dbl_std_equip_cost_op, ReplaceCharactersForFormat(std_equip_cost_op.Text))
        Else
        std_equip_cost_op.Text = std_equip_cost_op.Text
        End If
        If (Trim(std_total_cost_op.Text) <> "" And Val(std_total_cost_op.Text) <> 0) Or (dbl_std_mat_cost_op <> 0 Or dbl_std_labor_cost_op <> 0 Or dbl_std_equip_cost_op <> 0) Then
            'If Trim(std_total_cost_op.Text <> "") Then
            If dbl_std_total_cost_op <> 0 Then
'                std_total_cost_op.Text = Format(dbl_std_total_cost_op, ReplaceCharactersForFormat(std_total_cost_op.Text))
'            Else
                std_total_cost_op.Text = Format(dbl_std_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_std_total_cost_op)))
            End If
        Else
        std_total_cost_op.Text = std_total_cost_op.Text
        End If
        If Trim(rr_mat_cost_op.Text) <> "" And Val(rr_mat_cost_op.Text) <> 0 Then
        rr_mat_cost_op.Text = Format(dbl_rr_mat_cost_op, ReplaceCharactersForFormat(rr_mat_cost_op.Text))
        Else
        rr_mat_cost_op.Text = rr_mat_cost_op.Text
        End If
        If Trim(rr_labor_cost_op.Text) <> "" And Val(rr_labor_cost_op.Text) <> 0 Then
        rr_labor_cost_op.Text = Format(dbl_rr_labor_cost_op, ReplaceCharactersForFormat(rr_labor_cost_op.Text))
        Else
        rr_labor_cost_op.Text = rr_labor_cost_op.Text
        End If
        If Trim(rr_equip_cost_op.Text) <> "" And Val(rr_equip_cost_op.Text) <> 0 Then
        rr_equip_cost_op.Text = Format(dbl_rr_equip_cost_op, ReplaceCharactersForFormat(rr_equip_cost_op.Text))
        Else
        rr_equip_cost_op.Text = rr_equip_cost_op.Text
        End If
        If (Trim(rr_total_cost_op.Text) <> "" And Val(rr_total_cost_op.Text) <> 0) Or (dbl_rr_mat_cost_op <> 0 Or dbl_rr_labor_cost_op <> 0 Or dbl_rr_equip_cost_op <> 0) Then
            'If Trim(rr_total_cost_op.Text) <> "" Then
            If dbl_rr_total_cost_op <> 0 Then
'                rr_total_cost_op.Text = Format(dbl_rr_total_cost_op, ReplaceCharactersForFormat(rr_total_cost_op.Text))
'            Else
                rr_total_cost_op.Text = Format(dbl_rr_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_rr_total_cost_op)))
            End If
        Else
        rr_total_cost_op.Text = rr_total_cost_op.Text
        End If
        If Trim(opn_mat_cost_op.Text) <> "" And Val(opn_mat_cost_op.Text) <> 0 Then
        opn_mat_cost_op.Text = Format(dbl_opn_mat_cost_op, ReplaceCharactersForFormat(opn_mat_cost_op.Text))
        Else
        opn_mat_cost_op.Text = opn_mat_cost_op.Text
        End If
        If Trim(opn_labor_cost_op.Text) <> "" And Val(opn_labor_cost_op.Text) <> 0 Then
        opn_labor_cost_op.Text = Format(dbl_opn_labor_cost_op, ReplaceCharactersForFormat(opn_labor_cost_op.Text))
        Else
        opn_labor_cost_op.Text = opn_labor_cost_op.Text
        End If
        If Trim(opn_equip_cost_op.Text) <> "" And Val(opn_equip_cost_op.Text) <> 0 Then
        opn_equip_cost_op.Text = Format(dbl_opn_equip_cost_op, ReplaceCharactersForFormat(opn_equip_cost_op.Text))
        Else
        opn_equip_cost_op.Text = opn_equip_cost_op.Text
        End If
        If (Trim(opn_total_cost_op.Text) <> "" And Val(opn_total_cost_op.Text) <> 0) Or (dbl_opn_mat_cost_op <> 0 Or dbl_opn_labor_cost_op <> 0 Or dbl_opn_equip_cost_op <> 0) Then
            'If Trim(opn_total_cost_op.Text) <> "" Then
            If dbl_opn_total_cost_op <> 0 Then
'            opn_total_cost_op.Text = Format(dbl_opn_total_cost_op, ReplaceCharactersForFormat(opn_total_cost_op.Text))
'            Else
            opn_total_cost_op.Text = Format(dbl_opn_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_opn_total_cost_op)))
            End If
        Else
        opn_total_cost_op.Text = opn_total_cost_op.Text
        End If
        If Trim(metric_mat_cost_op.Text) <> "" And Val(metric_mat_cost_op.Text) <> 0 Then
        metric_mat_cost_op.Text = Format(dbl_metric_mat_cost_op, ReplaceCharactersForFormat(metric_mat_cost_op.Text))
        Else
        metric_mat_cost_op.Text = metric_mat_cost_op.Text
        End If
        If Trim(metric_labor_cost_op.Text) <> "" And Val(metric_labor_cost_op.Text) <> 0 Then
        metric_labor_cost_op.Text = Format(dbl_metric_labor_cost_op, ReplaceCharactersForFormat(metric_labor_cost_op.Text))
        Else
        metric_labor_cost_op.Text = metric_labor_cost_op.Text
        End If
        If Trim(metric_equip_cost_op.Text) <> "" And Val(metric_equip_cost_op.Text) <> 0 Then
        metric_equip_cost_op.Text = Format(dbl_metric_equip_cost_op, ReplaceCharactersForFormat(metric_equip_cost_op.Text))
        Else
        metric_equip_cost_op.Text = metric_equip_cost_op.Text
        End If
        If (Trim(metric_total_cost_op.Text) <> "" And Val(metric_total_cost_op.Text) <> 0) Or (dbl_metric_mat_cost_op <> 0 Or dbl_metric_labor_cost_op <> 0 Or dbl_metric_equip_cost_op <> 0) Then
            'If Trim(metric_total_cost_op.Text) <> "" Then
            If dbl_metric_total_cost_op <> 0 Then
'                metric_total_cost_op.Text = Format(dbl_metric_total_cost_op, ReplaceCharactersForFormat(metric_total_cost_op.Text))
'            Else
                metric_total_cost_op.Text = Format(dbl_metric_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_metric_total_cost_op)))
            End If
        Else
        metric_total_cost_op.Text = metric_total_cost_op.Text
        End If
        If Trim(res_mat_cost_op.Text) <> "" And Val(res_mat_cost_op.Text) <> 0 Then
        res_mat_cost_op.Text = Format(dbl_res_mat_cost_op, ReplaceCharactersForFormat(res_mat_cost_op.Text))
        Else
        res_mat_cost_op.Text = res_mat_cost_op.Text
        End If
        If Trim(res_labor_cost_op.Text) <> "" And Val(res_labor_cost_op.Text) <> 0 Then
        res_labor_cost_op.Text = Format(dbl_res_labor_cost_op, ReplaceCharactersForFormat(res_labor_cost_op.Text))
        Else
        res_labor_cost_op.Text = res_labor_cost_op.Text
        End If
        If Trim(res_equip_cost_op.Text) <> "" And Val(res_equip_cost_op.Text) <> 0 Then
        res_equip_cost_op.Text = Format(dbl_res_equip_cost_op, ReplaceCharactersForFormat(res_equip_cost_op.Text))
        Else
        res_equip_cost_op.Text = res_equip_cost_op.Text
        End If
        If (Trim(res_total_cost_op.Text) <> "" And Val(res_total_cost_op.Text) <> 0) Or (dbl_res_mat_cost_op <> 0 Or dbl_res_labor_cost_op <> 0 Or dbl_res_equip_cost_op <> 0) Then
            'If (Trim(res_total_cost_op.Text) <> "") Then
            If dbl_res_total_cost_op <> 0 Then
'                res_total_cost_op.Text = Format(dbl_res_total_cost_op, ReplaceCharactersForFormat(res_total_cost_op.Text))
'            Else
                res_total_cost_op.Text = Format(dbl_res_total_cost_op, ReplaceCharactersForFormat(CStr(dbl_res_total_cost_op)))
            End If
        Else
        res_total_cost_op.Text = res_total_cost_op.Text
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
        Screen.MousePointer = vbNormal
            'ApplyRoundingRules = strOpGreaterThanBare
            Dim retValOPGreater As Integer
            
            retValOPGreater = MsgBox(strOpGreaterThanBare + vbCrLf + vbCrLf + "Do you want to continue Updating this to the Database?", vbYesNo)
            If retValOPGreater = vbNo Then
                
                ApplyRoundingRules = -2 ' "DO NOT SAVE"
                Exit Function
            End If
            Screen.MousePointer = vbHourglass
        End If

    Else
        
        MsgBox "There was a database error while trying to apply the rounding rules - " + vbCrLf + retStr
        ApplyRoundingRules = -1 'error while rounding
    End If

End Function

Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim blnRet As Boolean
    Dim blnUpdateUnitCost As Boolean
    Dim ctr As Control
    Dim fld As ADODB.Field
    Dim rec As New ADODB.RecordSet
    Dim strError As String
    Dim strPercent_flag As String
    Dim strSelect As String
    Dim strUpdate As String
    Dim strSaveUpdate As String
    Dim intStart As Integer
    Dim varSaveBookmark As Variant
    Dim i As Integer
    Dim bln_Continue As Boolean
      
    'rlh 05/16/2008
    'Added check here to catch invalid MF04 ext unit cost id (2nd box from left) if user goes directly
    'to UPDATE button after respecifying unit_cost_id leftmost box
    If m_blnInsert Or m_blnClone Then           'rlh 05/16/2008
        ext_unit_cost_id_Validate (False)       'rlh 05/16/2008
        If VALIDATE_CANCEL = True Then Exit Sub
    End If                                      'rlh 05/16/2008
    
    Screen.MousePointer = vbHourglass
    TDBGrid.Update
    m_blnWereErrors = m_objGridMap.GridUpdateErrors
    If m_blnWereErrors = True Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'rlh 04/17/2008
    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
        m_blnWereErrors = False
        If m_blnInsert Or m_blnClone Then
            Dim bln_New As Boolean
            bln_New = True
        End If
        'rlh 05/13/08 - Checking on alt unit cost id no longer matters...
'''        If alt_unit_cost_id <> "" Then
'''            m_blnWereErrors = Invalid_ID_Format(Compress_String(alt_unit_cost_id), "alt_unit_cost_id", m_rec, bln_New)
'''        End If
        If Not m_blnWereErrors Then
            m_blnWereErrors = Invalid_ID_Format(Compress_String(unit_cost_id), "unit_cost_id", m_rec, bln_New)
        End If
    End If
    'RLH 04/17/2008  END
    
    
    m_blnWereErrors = False
    If pct_ind = 0 Then
        strPercent_flag = ""
    Else
        strPercent_flag = "Y"
    End If
    
    If Trim(unit_cost_id.Text) = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Unit Cost ID may not be blank.", vbExclamation
        m_blnWereErrors = True
        Exit Sub
    End If
    
    If Trim(type_code.Text) = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "The Type Code may not be blank.", vbExclamation
        m_blnWereErrors = True
        Exit Sub
    End If
    
    If Trim(unit.Text) = "" Then
        If type_code.Text = "M" Then    '3-27-01 EP: CR # 551 Check for the value of type_code IF "M" UOM not blank
            Screen.MousePointer = vbNormal
            MsgBox "The Unit of Measure may not be blank.", vbExclamation
            m_blnWereErrors = True
            Exit Sub
        End If
    End If
    
    If IsNull(TDBGrid.Bookmark) Then
        bln_Continue = True
    ElseIf TDBGrid.AddNewMode = dbgAddNewCurrent Then   'Cursor in new row, no add pending
        bln_Continue = True
    ElseIf Trim(TDBGrid.Columns("mat_id").Text) = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "The Material ID may not be blank.", vbExclamation
    Else
        bln_Continue = True
    End If
    
    
    'code added by Mohan on Jan 24: added rounding code here
    If type_code.Text = "E" And pct_ind.Value = 0 Then
        
        'if only totals were changed
        Dim retCheckOnlyTotalsChanged As String
        Dim retYesNo As Integer
        retCheckOnlyTotalsChanged = CheckOnlyTotalsChanged()
        If retCheckOnlyTotalsChanged <> "" Then
            retYesNo = MsgBox("Only Totals were changed. Would you like to continue?" + vbCrLf + vbCrLf + retCheckOnlyTotalsChanged, vbYesNo)
            If retYesNo = vbNo Then
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
        End If
        
        Dim retApplyRoundingRules As Integer
        retApplyRoundingRules = 0
        
        
        retApplyRoundingRules = ApplyRoundingRules()
        If (retApplyRoundingRules = -1 Or retApplyRoundingRules = -2) Then
            'retApplyRoundingRules = -1 'error while trying to apply the rounding rules - " + vbCrLf + retApplyRoundingRules
            'retApplyRoundingRules = -2 'Then 'DO NOT SAVE because there was a case where  OP<Bare and the user didn't want to save it
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        
        
    End If
    
    If bln_Continue = True Then
        Screen.MousePointer = vbHourglass
        unit_cost_id = Compress_String(unit_cost_id)
        '----------CR # 960 ------------------------------------
        'Added alt_id Compress_String Function
        alt_unit_cost_id = Compress_String(alt_unit_cost_id)
        '-------------------------------------------------------
        ext_unit_cost_id = Compress_String(ext_unit_cost_id)
        Dim recClone As ADODB.RecordSet
        Set recClone = m_rec.Clone
        recClone.AddNew
        UpdateRecordsetFromForm Me, recClone
        
        If pct_ind = 1 Then
            m_rec.Fields("percent_flag").Value = "Y"
        Else
            m_rec.Fields("percent_flag").Value = ""
        End If
    
        For Each fld In m_rec.Fields
            ' If the value changed
            If Not fld.Value = recClone.Fields(fld.Name).Value Or ((IsNull(fld.Value) Or fld.Value = "") Xor (recClone.Fields(fld.Name).Value = "")) Then
                Set ctr = Nothing
                Set ctr = Me.Controls(fld.Name)
                If Not ctr Is Nothing Then
                    ' See what table the field is from
                    If Left(Me.Controls(fld.Name).Tag, 1) = 1 Then
                        blnUpdateUnitCost = True
                    ElseIf Left(Me.Controls(fld.Name).Tag, 1) = 3 Then
                        blnUpdateUnitCost = True
                    End If
                End If
            End If
        Next
        
        'Undo the changes made by the UpdateRecordsetFromForm call above
        recClone.CancelUpdate
        recClone.Close
        Set recClone = Nothing
        'Set the Cost Change flag based on any Unit Cost field changes
        Dim excludeList As New Collection
        
        If blnUpdateUnitCost = True Or m_objGridMap.IsPendingChange Or (m_rec2.RecordCount > 0 And m_blnClone = True) Then
            If MasterFormat = EXT_MASTERFORMAT_VERSION Then
                strUpdate = "exec usp_update_unit_cost_driver_ext_rlh "        'PRODUCTION
                'strUpdate = "exec usp_update_unit_cost_driver_ext_rlh2 "   'rlh - 05/24/207 - STAGE/TEST
                'strUpdate = "exec usp_update_unit_cost_driver_ext_rlh "     'rlh 04/17/2008  -  STAGE/TEST
            Else
                strUpdate = "exec sp_update_unit_cost_driver_res "
                
                ' We need to exclude these controls from the stored proc.
                excludeList.Add "ext_unit_cost_id", "ext_unit_cost_id"
                excludeList.Add "inhouse_total_cost_op", "inhouse_total_cost_op"
                excludeList.Add "inhouse_equip_cost_op", "inhouse_equip_cost_op"
                excludeList.Add "inhouse_mat_cost_op", "inhouse_mat_cost_op"
                excludeList.Add "inhouse_labor_cost_op", "inhouse_labor_cost_op"
            End If
            
            BuildStoredProcSQL Me, strUpdate, 1, m_rec, excludeList
            BuildStoredProcSQL Me, strUpdate, 3, m_rec, excludeList
            
            'strUpdate = strUpdate + " @crew_skey='" + "', "
            strUpdate = strUpdate + " @percent_flag='" + strPercent_flag + "', "
            strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
            strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
            strUpdate = strUpdate + ", @bypass_ucd_ind = 0"
            strSaveUpdate = strUpdate
            If ucd_last_update_id.Text = "" Then ucd_last_update_id.Text = 0
            If cstw_last_update_id.Text = "" Then cstw_last_update_id.Text = 0
            strUpdate = strUpdate + ", @ucd_last_update_id=" + ucd_last_update_id.Text + ", "
            strUpdate = strUpdate + " @cstw_last_update_id=" + cstw_last_update_id.Text
            strUpdate = strUpdate + ", @update_material_usage_ind=0"
            strUpdate = strUpdate + ", @cost_change_ind=" + CStr(SetCostChange())
    
            m_blnWereErrors = False
            If m_blnClone = True Or m_blnInsert = True Or type_code <> "M" Then
                'Need to get skey if adding or update <> M
                ExecUpdate strUpdate
                If m_blnWereErrors = False Then
                    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
                        strSelect = "select unit_cost_skey from unit_cost_detail_ext where unit_cost_id = '" & unit_cost_id.Text & "'"
                    Else
                        strSelect = "select unit_cost_skey from unit_cost_detail where unit_cost_id = '" & unit_cost_id.Text & "'"
                    End If
                    'Stop 'rlh temporary
                    g_objDAL.GetRecordset vbNullString, strSelect, rec
                    If (rec.EOF And rec.BOF) Then
                        Screen.MousePointer = vbNormal
                        MsgBox "Record not added."
                        m_blnWereErrors = True
                        Exit Sub
                    Else
                        If m_blnClone = True Then
                            ' need to decrease for clone - last update id should not be incremented
                            ucd_last_update_id.Text = CInt(ucd_last_update_id.Text) - 1
                            cstw_last_update_id.Text = CInt(cstw_last_update_id.Text) - 1
                            m_rec.Fields("ucd_last_update_id").Value = ucd_last_update_id.Text
                            m_rec.Fields("cstw_last_update_id").Value = cstw_last_update_id.Text
                        End If
                        unit_cost_skey.Text = rec.Fields("unit_cost_skey").Value
                    End If
                End If
            End If
            If m_blnClone = True Then
                'Copy the output_usage data for the unit cost
                strUpdate = "exec sp_copy_output_usage @type = 'U', @FromSkey = '" & m_lngOriginalSkey & "', @ToSkey='" & unit_cost_skey.Text + "', "
                strUpdate = strUpdate + " @last_update_date='" + Format(Now(), "General Date") + "', "
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "', "
                strUpdate = strUpdate + " @last_update_id='1'"
                ExecUpdate strUpdate
                '*** Create records for the long descriptions
                '*** ADDED 6/9/2005 RTD
                CloneLongDescriptions strOriginalCostID, unit_cost_skey.Text
            ElseIf m_blnInsert = True Then
                '*** Create records for the long descriptions
                '*** ADDED 6/9/2005 RTD
                AddNewLongDescriptions unit_cost_skey.Text, book_desc.Text, metric_book_desc.Text
            End If

            If type_code = "M" Then
                'Process changes or deletions
                If m_objGridMap.IsPendingChange Or (m_blnClone = True And m_rec2.RecordCount > 0) Then
                    'If cloning, update the unit_cost_skey in all records in the grid.
                    If m_blnClone = True Or m_blnInsert = True Then
                        m_objGridMap.UnitCostSKey = unit_cost_skey.Text
                        'ADDED 8/18/2005 RTD - CORRECT MATERIAL CLONE PROBLEM
                        m_objGridMap.UnitCostID = unit_cost_id.Text
                        If m_rec2.RecordCount > 0 Then
                            m_rec2.MoveFirst
                            Do Until m_rec2.EOF
                                m_rec2.Fields("unit_cost_skey") = unit_cost_skey.Text
                                m_rec2.MoveNext
                            Loop
                        End If
                    End If
                    If m_blnClone = True Then   'Flag all rows as new
                        m_objGridMap.SetRowStateNew
                    End If
                    blnRet = m_objGridMap.Update
                    If blnRet = False Then
                        m_blnWereErrors = True
                    Else    'If g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
                        'Update the parameter to trigger Material Usage cost update
                        strUpdate = strSaveUpdate + ", @ucd_last_update_id=" + ucd_last_update_id.Text
                        strUpdate = strUpdate + ", @cstw_last_update_id=" + cstw_last_update_id.Text
                        strUpdate = strUpdate + ", @update_material_usage_ind=1"
                        strUpdate = strUpdate + ", @cost_change_ind=1"
                        strUpdate = ReplaceSkey(strUpdate, unit_cost_skey.Text)
                         
                        ExecUpdate strUpdate
                    End If
                ElseIf (m_blnClone = False And m_blnInsert = False) Then    'Update without Material Usage
                    strUpdate = strSaveUpdate
                    If ucd_last_update_id.Text = "" Then ucd_last_update_id.Text = 0
                    If cstw_last_update_id.Text = "" Then cstw_last_update_id.Text = 0
                    strUpdate = strUpdate + ", @ucd_last_update_id=" + ucd_last_update_id.Text + ", "
                    strUpdate = strUpdate + " @cstw_last_update_id=" + cstw_last_update_id.Text
                    'strUpdate = strUpdate + ", @update_material_usage_ind=0"
                    
                    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                    'FORCE "MATERIAL" RECALCULATIONS EVEN WHEN MATERIAL GRID ROW(S)
                    'NOT BEEN
                    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                    If m_rec2.RecordCount > 0 Then
                        strUpdate = strUpdate + ", @update_material_usage_ind=1" '(rlh) 05/10/2010 Force recalc of Material Usage
                    Else
                        strUpdate = strUpdate + ", @update_material_usage_ind=0"
                    End If
                    strUpdate = strUpdate + ", @cost_change_ind=" + CStr(SetCostChange())
                    ExecUpdate strUpdate
                End If
            End If
            m_blnClone = False  ' no longer cloning if we were
            If m_blnWereErrors = False Then
                ' Put latest data into source recordset
                UpdateRecordsetFromForm Me, m_rec
                If pct_ind = 1 Then
                    m_rec.Fields("percent_flag").Value = "Y"
                Else
                    m_rec.Fields("percent_flag").Value = ""
                End If
                'm_rec_unformatfields
                If IsNull(m_rec.Fields("ucd_last_update_id").Value) Then
                    m_rec.Fields("ucd_last_update_id").Value = 1
                'Else
                    'm_rec.Fields("ucd_last_update_id").Value = m_rec.Fields("ucd_last_update_id").Value + 1
                End If
                If IsNull(m_rec.Fields("cstw_last_update_id").Value) Then
                    m_rec.Fields("cstw_last_update_id").Value = 1
                'Else
                    'm_rec.Fields("cstw_last_update_id").Value = m_rec.Fields("cstw_last_update_id").Value + 1
                End If
                ucd_last_update_id.Text = m_rec.Fields("ucd_last_update_id").Value
                cstw_last_update_id.Text = m_rec.Fields("cstw_last_update_id").Value
                If type_code = "M" Then             'And g_intRollupOption = ALWAYS_ROLLUP_MATERIAL Then
                    
                    'Retrieve latest amounts from Material Usage changes.
'''                    strSelect = "select unit_cost_skey , std_mat_cost, std_mat_cost_op, std_total_cost, " + _
'''                    "std_total_cost_op, opn_mat_cost, opn_mat_cost_op, opn_total_cost, opn_total_cost_op, " + _
'''                    "rr_mat_cost, rr_mat_cost_op, rr_total_cost, rr_total_cost_op" + _
'''                    ", metric_mat_cost, metric_mat_cost_op, metric_total_cost, metric_total_cost_op " + _
'''                    "from published_unit_cost_costworks  " + _
'''                    "where unit_cost_skey = " & unit_cost_skey.Text
                    
                    'RESI addition/change - rlh 05/06/2010
                    strSelect = "select unit_cost_skey , std_mat_cost, std_mat_cost_op, std_total_cost, " + _
                    "std_total_cost_op, opn_mat_cost, opn_mat_cost_op, opn_total_cost, opn_total_cost_op, " + _
                    "rr_mat_cost, rr_mat_cost_op, rr_total_cost, rr_total_cost_op" + _
                    ", metric_mat_cost, metric_mat_cost_op, metric_total_cost, metric_total_cost_op " + _
                     ", res_mat_cost, res_mat_cost_op, res_total_cost, res_total_cost_op " + _
                    "from published_unit_cost_costworks  " + _
                    "where unit_cost_skey = " & unit_cost_skey.Text
                    
                    g_objDAL.GetRecordset vbNullString, strSelect, rec
                    If Not (rec.EOF And rec.BOF) Then
                        For i = 0 To rec.Fields.Count
                            m_rec.Fields(rec.Fields(i).Name) = rec.Fields(i).Value
                        Next i
                    End If
                End If
                UpdateFormFromRecordset Me, m_rec
                GetLongDescriptions
                last_update_person.Text = strUserName
                last_update_date.Text = Now
                set_pct_ind
            End If
            If m_intMasterFormat = EXT_MASTERFORMAT_VERSION Then
                ext_unit_cost_id.Text = Format(Compress_String(ext_unit_cost_id.Text), FORMAT_UNIT_COST_SRV)
                unit_cost_id.Text = Format(Compress_String(unit_cost_id.Text), FORMAT_UNIT_COST_04_SRV)
            Else
                ext_unit_cost_id.Text = Format(Compress_String(ext_unit_cost_id.Text), FORMAT_UNIT_COST_04_SRV)
                unit_cost_id.Text = Format(Compress_String(unit_cost_id.Text), FORMAT_UNIT_COST_SRV)
            End If
            If m_blnWereErrors = False Then
                If Not cmdLongDesc.Enabled Then cmdLongDesc.Enabled = True
                
                'code added by mohan Jan 18, 2012: update the Hierarchy tree
                Dim retBlnVal As Boolean
                retBlnVal = MainModule.Update_Tree_With_Unit_Cost_Id(unit_cost_id.Text, alt_unit_cost_id.Text)
                
                Screen.MousePointer = vbNormal
                
                'set all the flags for cost column changes back to false
                SetCostChangedFlagsToFalse

                
                If retBlnVal = False Then
                    MsgBox "Update successful for Unit Cost Details, but there was an error while updating the Tree.", vbExclamation + vbOKOnly
                Else
                    MsgBox "Update successful.", vbInformation + vbOKOnly
                End If
                
                pct_ind_o = pct_ind
            End If
            RebindTDBGridNow
            varSaveBookmark = TDBGrid.Bookmark
            TDBGrid.Refresh
            TDBGrid.Bookmark = varSaveBookmark
            
            
        Else
            If m_intMasterFormat = EXT_MASTERFORMAT_VERSION Then
                ext_unit_cost_id.Text = Format(Compress_String(ext_unit_cost_id.Text), FORMAT_UNIT_COST_SRV)
                unit_cost_id.Text = Format(Compress_String(unit_cost_id.Text), FORMAT_UNIT_COST_04_SRV)
            Else
                ext_unit_cost_id.Text = Format(Compress_String(ext_unit_cost_id.Text), FORMAT_UNIT_COST_04_SRV)
                unit_cost_id.Text = Format(Compress_String(unit_cost_id.Text), FORMAT_UNIT_COST_SRV)
            End If
            Screen.MousePointer = vbNormal
            MsgBox "You must modify a field before updating."
        End If
    End If
    Screen.MousePointer = vbNormal


    Exit Sub
    
Err_Handler:
    Dim errMessage As String
    errMessage = "frmUnitCostGrid:cmdUpdate_Click - " + Err.Description

    Debug.Print errMessage
    MsgBox errMessage
    
End Sub

Private Function GenericErrorMessageForChangedTotalsOnly(ByVal costsSectionName As String, ByVal columnName As String) As String
    
    GenericErrorMessageForChangedTotalsOnly = costsSectionName + " '" + columnName + " Total' was changed without changing other " + costsSectionName + " '" + columnName + "' columns" + vbCrLf

End Function

Private Function CheckOnlyTotalsChanged() As String

    CheckOnlyTotalsChanged = ""
    Dim strMessage As String
    strMessage = ""
    If std_total_cost_changed = True And std_NON_total_cost_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("Bare Costs", "Standard")
    End If
    If rr_total_cost_changed = True And rr_NON_total_cost_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("Bare Costs", "R&R")
    End If
    If opn_total_cost_changed = True And opn_NON_total_cost_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("Bare Costs", "Open")
    End If
    If metric_total_cost_changed = True And metric_NON_total_cost_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("Bare Costs", "Metric")
    End If
    If res_total_cost_changed = True And res_NON_total_cost_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("Bare Costs", "Resi")
    End If
    
    If std_total_cost_op_changed = True And std_NON_total_cost_op_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("O&P", "Standard")
    End If
    If rr_total_cost_op_changed = True And rr_NON_total_cost_op_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("O&P", "R&R")
    End If
    If opn_total_cost_op_changed = True And opn_NON_total_cost_op_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("O&P", "Open")
    End If
    If metric_total_cost_op_changed = True And metric_NON_total_cost_op_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("O&P", "Metric")
    End If
    If res_total_cost_op_changed = True And res_NON_total_cost_op_changed = False Then
        strMessage = strMessage + GenericErrorMessageForChangedTotalsOnly("O&P", "Resi")
    End If
    CheckOnlyTotalsChanged = strMessage

End Function

Private Sub ExecUpdate(strUpdate As String)
'Update the database with the current update sql string.
'If the update fails, display a message, otherwise increment the last update Id
    Dim blnRet As Boolean
    Dim strError As String
    
    On Error Resume Next
    blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)  'Update the Material Usage for unit cost
    If blnRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox strError, vbExclamation
        m_blnWereErrors = True
    Else
        ucd_last_update_id.Text = CInt(ucd_last_update_id.Text) + 1
        cstw_last_update_id.Text = CInt(cstw_last_update_id.Text) + 1
        m_rec.Fields("ucd_last_update_id").Value = ucd_last_update_id.Text
        m_rec.Fields("cstw_last_update_id").Value = cstw_last_update_id.Text
    End If
End Sub

Private Function ReplaceSkey(strString, strSkey As String) As String
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim strTemp As String
    
    iStart = InStr(1, strString, "@unit_cost_skey=")
    If iStart > 0 Then
        iEnd = InStr(iStart, strString, ",")
        strTemp = Left(strString, iStart + 15) + strSkey + Right(strString, Len(strString) - iEnd + 1)
        ReplaceSkey = strTemp
    End If

End Function

Private Sub comment_Change()
    Dim intLength As Integer
    Dim intPosition As Integer
    Dim txtSaveComment As String
    Dim txtNewComment As String
    
    If Len(comment) > 0 Then
        intPosition = comment.SelStart
        If intPosition > 0 Then
            If Asc(Mid(comment, intPosition, 1)) >= 0 And Asc(Mid(comment, intPosition, 1)) <= 31 Then
                intLength = Len(comment)
                txtSaveComment = comment.Text
                MsgBox "Non-printable characters are not allowed in the comments."
                txtNewComment = Left(txtSaveComment, intPosition - 2)
                If intPosition < intLength Then
                    txtNewComment = txtNewComment + Right(txtSaveComment, intLength - intPosition)
                End If
                comment.Text = txtNewComment
                comment.SelStart = intPosition - 2
            End If
        End If
    End If
    
End Sub

Private Sub crew_id_Change()
    crew_id = UCase(crew_id)
    crew_id.SelStart = Len(crew_id)

End Sub

Private Sub crew_id_LostFocus()
'Validate the Crew Qty based on Crew.type_code
    Dim bln_result As Boolean
    If crew_id.DataChanged = True Then
        If Trim(crew_id) = "" Then
            crew_qty = ""
            daily_output = ""
            metric_daily_output = ""
            bln_result = LockField(Me, "crew_qty")
            bln_result = LockField(Me, "daily_output")
            bln_result = LockField(Me, "metric_daily_output")
        Else
            bln_result = UnLockField(Me, "crew_qty")
            bln_result = UnLockField(Me, "daily_output")
            bln_result = UnLockField(Me, "metric_daily_output")
            If validate_crew_id = True Then
                If crew_qty.Locked = False Then
                    crew_qty.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Function validate_crew_id() As Boolean
Dim strSelect As String
Dim rec As New ADODB.RecordSet ' Recordset to hold query results
Dim blnReturn As Boolean
Dim vntMyBookmark As Variant
Dim i As Integer
Dim blnResult As Boolean

validate_crew_id = True
If crew_id = "" Then
    crew_qty = ""
Else
        strSelect = "Select type_code from crew inner join published_crew_rate " + _
        " on crew.crew_skey = published_crew_rate.crew_skey " + _
        " where crew.crew_id='" + crew_id.Text + "'"
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
    ' If it does, copy that data into grid
    If rec.RecordCount = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "Please enter a valid crew. The crew must also be in the Published Crew Rate table."
        m_blnWereErrors = True
        crew_id.SetFocus
        validate_crew_id = False
    Else
        crew_type_code.Text = rec.Fields("type_code").Value
        If rec.Fields("type_code") = "C" Then
            crew_qty.Text = 1
            blnResult = LockField(Me, "crew_qty")
        Else
            blnResult = UnLockField(Me, "crew_qty")
        End If
    End If
    rec.Close
End If
End Function

Private Sub crew_qty_LostFocus()
    validate_crew_qty
End Sub

Private Sub ext_unit_cost_id_GotFocus()
    ext_unit_cost_id.Text = Compress_String(ext_unit_cost_id.Text)
End Sub

Private Sub ext_unit_cost_id_LostFocus()
    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
        ext_unit_cost_id.Text = Format(ext_unit_cost_id.Text, FORMAT_UNIT_COST_SRV)
    Else
        ext_unit_cost_id.Text = Format(ext_unit_cost_id.Text, FORMAT_UNIT_COST_04_SRV)
    End If
End Sub

Private Sub ext_unit_cost_id_Validate(Cancel As Boolean)
    Dim bln_New As Boolean
    Dim sField As String
    Dim sTable As String
    
    VALIDATE_CANCEL = False
    
    If ext_unit_cost_id.Text <> "" Then
        If m_blnInsert Or m_blnClone Then
            bln_New = True
        End If
        If MasterFormat = EXT_MASTERFORMAT_VERSION Then
            ' In MF04 mode, EXT_UNIT_COST_ID is the MF95 ID (unit_cost_id in UCD table)
            sField = "unit_cost_id"
            sTable = "UNIT_COST_DETAIL"
        Else
            ' In MF95 mode, EXT_UNIT_COST_ID is the MF04 ID (unit_cost_id in UCD_EXT table)
            sField = "unit_cost_id"
            sTable = "UNIT_COST_DETAIL_EXT"
        End If
        If Invalid_ID_Format(Compress_String(ext_unit_cost_id), sField, m_rec, bln_New, sTable) = True Then
            Cancel = True
            VALIDATE_CANCEL = True   'rlh 05/16/2008
        End If
    End If
    
'     RLH 01/07/2009  - COMMENTED OUT CODE BELOW - DO NOT FORCE USER TO SPECIFY A MF95 UNIT COST ID !!!

'     RLH 07/01/2008  - force user to specify a MF95 unit cost id when building NEW ucl in MF04 MODE
'    If MasterFormat = EXT_MASTERFORMAT_VERSION And ext_unit_cost_id.Text = "" And Me.type_code = "M" Then
'        VALIDATE_CANCEL = True       'rlh 07/01/2008 You MUST have a MF95 for a (M)aterial/Cost ucl in MF04 PROCESSSING MODE
'        MsgBox ("Invalid MF95 Unit Cost Id: " & Me.ext_unit_cost_id.Text)
'    End If
End Sub

Private Sub Form_Activate()
    OutputView False
End Sub

Private Sub Form_Initialize()
    m_blnInsert = False
    m_blnDeleted = False
End Sub

Private Sub SetCostChangedFlagsToFalse()

    std_total_cost_changed = False
    rr_total_cost_changed = False
    opn_total_cost_changed = False
    metric_total_cost_changed = False
    res_total_cost_changed = False
    std_total_cost_op_changed = False
    rr_total_cost_op_changed = False
    opn_total_cost_op_changed = False
    metric_total_cost_op_changed = False
    res_total_cost_op_changed = False
        
    
    std_NON_total_cost_changed = False
    rr_NON_total_cost_changed = False
    opn_NON_total_cost_changed = False
    metric_NON_total_cost_changed = False
    res_NON_total_cost_changed = False
    std_NON_total_cost_op_changed = False
    rr_NON_total_cost_op_changed = False
    opn_NON_total_cost_op_changed = False
    metric_NON_total_cost_op_changed = False
    res_NON_total_cost_op_changed = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean
    
    'initializing variables for total changes
    form_loading = True
    
    SetCostChangedFlagsToFalse

    
    
    If START_HEIGHT < 8115 Then
        Move START_LEFT, START_TOP, 11775, START_HEIGHT - 120
    Else
        Move START_LEFT, START_TOP, 11775, 8115
    End If
    
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    
    type_code.AddItem "M"
    type_code.AddItem "E"
    If m_blnInsert = False Then
        type_code.AddItem "B"
    End If
    type_code.AddItem "H"

    cmdLongDesc.Enabled = Not m_blnInsert
    
    If Not m_rec.State = adStateClosed Then
        UpdateFormFromRecordset Me, m_rec
        ' *** ADDED 6/6/2005 RTD per TD ***
        GetLongDescriptions
        '
        'format_costs
        g_objDAL.GetRecordset vbNullString, "select unit from unit_of_measure order by unit", rec
        While Not rec.EOF
            If Len(Trim(rec.Fields("unit").Value)) > 0 Then 'Do not allow blank unit, do allow blank metric unit
                unit.AddItem (rec.Fields("unit").Value)
                If Trim(m_rec.Fields("unit")) = Trim(rec.Fields("unit").Value) Then
                    unit.Text = unit.List(unit.NewIndex)
                End If
            End If
            
            If (rec.Fields("unit") <> m_rec("metric_unit")) Then                'rlh
                metric_unit.AddItem (rec.Fields("unit").Value)
                
                If Trim(m_rec.Fields("metric_unit")) = Trim(rec.Fields("unit").Value) Then
                
                '  rlh - 11/16/06 ---------------------------------------------------------
                '  For some reason, when you try o set the combobox text property/value
                '  within the loop, the assignment gets IGNORED !??  (so, I have to do it
                '  after the loop is complete and all of the possible metric units have been
                '  added to the dropdown/combobox ?!
                '---------------------------------------------------------------------------
                
                    metric_unit.Text = metric_unit.List(metric_unit.NewIndex)  'rlh
                End If
            End If                                                             'rlh
            rec.MoveNext
        Wend
        
        'rlh - SELECT/REMEMBER THE "metric unit" from previous session's add/update/save
        metric_unit.AddItem (m_rec("metric_unit"))     'rlh - this adds to list but doesn't override selected text !?
        metric_unit.Text = m_rec("metric_unit")        'rlh

        
        rec.Close
    End If
    
    ' 8/12/2005 UPDATE TITLE BAR WITH MASTERFORMAT VERSION
    ' If we are NOT inserting
    If m_blnInsert = False Then
        ' Lock fields that can't be changed
        unit_cost_id.Locked = True
        unit_cost_id.BackColor = LTGREY
        Me.Caption = Me.Caption & " - MF" & MasterFormat & " - [" & m_rec.Fields("unit_cost_id").Value + "]"
        strOriginalCostID = m_rec.Fields("unit_cost_id").Value
    Else
        ' If we are inserting and not showing data
        ' Set some defaults
        unit_cost_skey.Text = ""
        If Not m_blnRecFlag Then
'            active_status_ind.Value = 1
            Me.Caption = Me.Caption & " - MF" & MasterFormat & " - [New]"
            strOriginalCostID = ""
        Else
            Me.Caption = Me.Caption & " - MF" & MasterFormat & " - [Clone of " & m_rec.Fields("unit_cost_id").Value + "]"
            strOriginalCostID = m_rec.Fields("unit_cost_id").Value
            m_blnClone = True
            m_lngOriginalSkey = m_rec.Fields("unit_cost_skey").Value
        End If
    End If
    If Not m_blnClone Then
        strLast_unit_cost_id = m_rec.Fields("unit_cost_id").Value
    Else
        unit_cost_id.CausesValidation = True
    End If

    ' Make the form show the right fields based on current type_code
    m_type_code = type_code
    type_code_LostFocus
    blnReturn = validate_crew_id   'set properties for crew
    SaveOrigCost
    set_pct_ind

    blnReturn = LockField(Me, "std_labor_hour")
    blnReturn = LockField(Me, "rr_labor_hour")
    blnReturn = LockField(Me, "opn_labor_hour")
    blnReturn = LockField(Me, "metric_labor_hour")
    blnReturn = LockField(Me, "metric_daily_output")
    blnReturn = LockField(Me, "res_labor_hour")         'rlh RESI 05/6/2010
    ColorLockedFields Me
    UpdateBookPreviewLine
        
        'CCD 8.4 (rlh) 04/14/2009
         'ALWAYS disabled!!!  (both MF95 and MF04 unit cost id text boxes)
        
        Me.ext_unit_cost_id.Enabled = False
        Me.unit_cost_id.Enabled = False
        
        'MF 2004 IS ACTIVE
        If MasterFormat = EXT_MASTERFORMAT_VERSION Then
            If MF95_ENABLED = False Then
            
                Me.ext_unit_cost_id.Enabled = False
                
                If m_blnClone_pub = True Then
                     ' In MF04 mode, EXT_UNIT_COST_ID is the MF95 ID (unit_cost_id in UCD table)
                     Me.ext_unit_cost_id.Enabled = True
                     Me.unit_cost_id.Enabled = True
                End If
            
                If m_blnNew_pub = True Then
                     ' In MF95 mode, EXT_UNIT_COST_ID is the MF04 ID (unit_cost_id in UCD_EXT table)
                    Me.ext_unit_cost_id.Enabled = True
                    Me.unit_cost_id.Enabled = True
                End If
            End If
        End If
        
        'MF 1995 IS ACTIVE
       
                
        If MasterFormat = UCD_MASTERFORMAT_VERSION Then
            If MF95_ENABLED = True Then
            
                
                If m_blnClone_pub = True Then
                     ' In MF04 mode, EXT_UNIT_COST_ID is the MF95 ID (unit_cost_id in UCD table)
                     Me.ext_unit_cost_id.Enabled = True
                     Me.unit_cost_id.Enabled = True
                End If
            
                If m_blnNew_pub = True Then
                     ' In MF95 mode, EXT_UNIT_COST_ID is the MF04 ID (unit_cost_id in UCD_EXT table)
                    Me.ext_unit_cost_id.Enabled = True
                    Me.unit_cost_id.Enabled = True
                End If
            End If
        End If
   
        m_blnNew_pub = False    'rlh ccd 8.4 4/15/2009
        m_blnClone_pub = False  'rlh ccd 8.4 4/15/2009

    If type_code.Text = "E" And pct_ind.Value = 0 Then
        cmdShowRounding.Visible = True
    Else
        cmdShowRounding.Visible = False
    End If
    
    'form is done loading
    form_loading = False

End Sub

Private Sub Form_Resize()

    If Me.Height > 6510 Then
        'picGrid.top = Me.Height - picGrid.Height + 100
        picGrid.Top = picTop.Height
        picGrid.Height = Me.Height - picTop.Height + 120
        picGrid.Width = Me.Width
        fraUnitCost.Top = 120
        fraUnitCost.Height = picGrid.Height - 1340
        'TDBGrid.top = 300
        TDBGrid.Height = fraUnitCost.Height - 850
        cmdMatUsageDelete.Top = fraUnitCost.Top + fraUnitCost.Height - cmdMatUsageDelete.Height - 240
        
        cmdUpdate.Top = fraUnitCost.Height + 250
        cmdDelete.Top = cmdUpdate.Top

        lblUpdated.Top = fraUnitCost.Height + 325
        lblUpdatedBy.Top = lblUpdated.Top
        lblSkey.Top = lblUpdated.Top
        lblMasterFormat.Top = lblUpdated.Top
        last_update_date.Top = fraUnitCost.Height + 300
        last_update_person.Top = last_update_date.Top
        unit_cost_skey.Top = last_update_date.Top
        txtMasterFormat.Top = last_update_date.Top
    End If
    
    fraUnitCost.Width = Me.Width - (fraUnitCost.Left * 3)
    TDBGrid.Width = fraUnitCost.Width - (TDBGrid.Left * 2)
    ResizeForm Me

End Sub


Private Sub format_characters_Change()
    UpdateBookPreviewLine
End Sub

Private Sub format_code_Change()
    UpdateBookPreviewLine
End Sub

Private Sub format_code_KeyPress(KeyAscii As Integer)
'FORCE LOWERCASE KEYPRESSES TO UPPERCASE
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii + Asc("A") - Asc("a")
    End If
End Sub

Private Sub indent_code_Change()
    UpdateBookPreviewLine
End Sub

Private Sub index_code_KeyPress(KeyAscii As Integer)
'FORCE LOWERCASE KEYPRESSES TO UPPERCASE
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii + Asc("A") - Asc("a")
    End If
End Sub

Private Sub index_desc_Change()
Dim intLength As Integer
Dim intPosition As Integer
Dim txtSaveindex_desc As String
Dim txtNewindex_desc As String

If Len(index_desc) > 0 Then
    intPosition = index_desc.SelStart
    If intPosition > 0 Then
        If Asc(Mid(index_desc, intPosition, 1)) >= 0 And Asc(Mid(index_desc, intPosition, 1)) <= 31 Then
            intLength = Len(index_desc)
            txtSaveindex_desc = index_desc.Text
            MsgBox "Non-printable characters are not allowed in the index_description."
            txtNewindex_desc = Left(txtSaveindex_desc, intPosition - 2)
            If intPosition < intLength Then
                txtNewindex_desc = txtNewindex_desc + Right(txtSaveindex_desc, intLength - intPosition)
            End If
            index_desc.Text = txtNewindex_desc
            index_desc.SelStart = intPosition - 2
        End If
    End If
End If

End Sub

Private Sub metric_assembly_book_desc_Change()
Dim intLength As Integer
Dim intPosition As Integer
Dim txtSavemetric_assembly_book_desc As String
Dim txtNewmetric_assembly_book_desc As String

If Len(metric_assembly_book_desc) > 0 Then
    intPosition = metric_assembly_book_desc.SelStart
    If intPosition > 0 Then
        If Asc(Mid(metric_assembly_book_desc, intPosition, 1)) >= 0 And Asc(Mid(metric_assembly_book_desc, intPosition, 1)) <= 31 Then
            intLength = Len(metric_assembly_book_desc)
            txtSavemetric_assembly_book_desc = metric_assembly_book_desc.Text
            MsgBox "Non-printable characters are not allowed in the metric_assembly_book_description."
            txtNewmetric_assembly_book_desc = Left(txtSavemetric_assembly_book_desc, intPosition - 2)
            If intPosition < intLength Then
                txtNewmetric_assembly_book_desc = txtNewmetric_assembly_book_desc + Right(txtSavemetric_assembly_book_desc, intLength - intPosition)
            End If
            metric_assembly_book_desc.Text = txtNewmetric_assembly_book_desc
            metric_assembly_book_desc.SelStart = intPosition - 2
        End If
    End If
End If

End Sub

Private Sub metric_book_desc_Change()
Dim intLength As Integer
Dim intPosition As Integer
Dim txtSavemetric_book_desc As String
Dim txtNewmetric_book_desc As String

If Len(metric_book_desc) > 0 Then
    intPosition = metric_book_desc.SelStart
    If intPosition > 0 Then
        If Asc(Mid(metric_book_desc, intPosition, 1)) >= 0 And Asc(Mid(metric_book_desc, intPosition, 1)) <= 31 Then
            intLength = Len(metric_book_desc)
            txtSavemetric_book_desc = metric_book_desc.Text
            MsgBox "Non-printable characters are not allowed in the metric_book_description."
            txtNewmetric_book_desc = Left(txtSavemetric_book_desc, intPosition - 2)
            If intPosition < intLength Then
                txtNewmetric_book_desc = txtNewmetric_book_desc + Right(txtSavemetric_book_desc, intLength - intPosition)
            End If
            metric_book_desc.Text = txtNewmetric_book_desc
            metric_book_desc.SelStart = intPosition - 2
        End If
    End If
End If

End Sub

Private Sub metric_tech_desc_Change()
Dim intLength As Integer
If Len(metric_tech_desc) > 0 Then
    If Asc(Right(metric_tech_desc, 1)) >= 0 And Asc(Right(metric_tech_desc, 1)) <= 31 Then
        intLength = Len(metric_tech_desc)
        MsgBox "Non-printable characters are not allowed in the metric_tech_descs."
        metric_tech_desc.Text = Left(metric_tech_desc.Text, intLength - 2)
        metric_tech_desc.SelStart = intLength - 2
    End If
End If

End Sub

Private Sub metric_total_cost_Change()
    If form_loading = False Then
        metric_total_cost_changed = True
    End If
End Sub

Private Sub metric_total_cost_op_Change()
    If form_loading = False Then
        metric_total_cost_op_changed = True
    End If
End Sub

Private Sub opn_total_cost_Change()
    If form_loading = False Then
        opn_total_cost_changed = True
    End If
End Sub

Private Sub opn_total_cost_op_Change()
    If form_loading = False Then
        opn_total_cost_op_changed = True
    End If
End Sub

Private Sub res_total_cost_Change()
    If form_loading = False Then
        res_total_cost_changed = True
    End If
End Sub

Private Sub res_total_cost_op_Change()
    If form_loading = False Then
        res_total_cost_op_changed = True
    End If
End Sub

Private Sub rr_total_cost_Change()
    If form_loading = False Then
        rr_total_cost_changed = True
    End If
End Sub

Private Sub rr_total_cost_op_Change()
    If form_loading = False Then
        rr_total_cost_op_changed = True
    End If
End Sub


Private Sub std_total_cost_Change()

    If form_loading = False Then
        std_total_cost_changed = True
    End If
    
End Sub

Private Sub std_equip_cost_Change()
    sub_NON_total_cost_changed ("std")
End Sub

Private Sub std_labor_cost_Change()
    sub_NON_total_cost_changed ("std")
End Sub

Private Sub std_mat_cost_Change()
    sub_NON_total_cost_changed ("std")
End Sub

Private Sub rr_equip_cost_Change()
    sub_NON_total_cost_changed ("rr")
End Sub

Private Sub rr_labor_cost_Change()
    sub_NON_total_cost_changed ("rr")
End Sub

Private Sub rr_mat_cost_Change()
    sub_NON_total_cost_changed ("rr")
End Sub

Private Sub opn_equip_cost_Change()
    sub_NON_total_cost_changed ("opn")
End Sub

Private Sub opn_labor_cost_Change()
    sub_NON_total_cost_changed ("opn")
End Sub

Private Sub opn_mat_cost_Change()
    sub_NON_total_cost_changed ("opn")
End Sub

Private Sub metric_equip_cost_Change()
    sub_NON_total_cost_changed ("metric")
End Sub

Private Sub metric_labor_cost_Change()
    sub_NON_total_cost_changed ("metric")
End Sub

Private Sub metric_mat_cost_Change()
    sub_NON_total_cost_changed ("metric")
End Sub

Private Sub res_equip_cost_Change()
    sub_NON_total_cost_changed ("res")
End Sub

Private Sub res_labor_cost_Change()
    sub_NON_total_cost_changed ("res")
End Sub

Private Sub res_mat_cost_Change()
    sub_NON_total_cost_changed ("res")
End Sub

Private Sub std_equip_cost_op_Change()
    sub_NON_total_cost_changed ("std_op")
End Sub

Private Sub std_labor_cost_op_Change()
    sub_NON_total_cost_changed ("std_op")
End Sub

Private Sub std_mat_cost_op_Change()
    sub_NON_total_cost_changed ("std_op")
End Sub

Private Sub rr_equip_cost_op_Change()
    sub_NON_total_cost_changed ("rr_op")
End Sub

Private Sub rr_labor_cost_op_Change()
    sub_NON_total_cost_changed ("rr_op")
End Sub

Private Sub rr_mat_cost_op_Change()
    sub_NON_total_cost_changed ("rr_op")
End Sub

Private Sub opn_equip_cost_op_Change()
    sub_NON_total_cost_changed ("opn_op")
End Sub

Private Sub opn_labor_cost_op_Change()
    sub_NON_total_cost_changed ("opn_op")
End Sub

Private Sub opn_mat_cost_op_Change()
    sub_NON_total_cost_changed ("opn_op")
End Sub

Private Sub metric_equip_cost_op_Change()
    sub_NON_total_cost_changed ("metric_op")
End Sub

Private Sub metric_labor_cost_op_Change()
    sub_NON_total_cost_changed ("metric_op")
End Sub

Private Sub metric_mat_cost_op_Change()
    sub_NON_total_cost_changed ("metric_op")
End Sub

Private Sub res_equip_cost_op_Change()
    sub_NON_total_cost_changed ("res_op")
End Sub

Private Sub res_labor_cost_op_Change()
    sub_NON_total_cost_changed ("res_op")
End Sub

Private Sub res_mat_cost_op_Change()
    sub_NON_total_cost_changed ("res_op")
End Sub


Private Sub sub_NON_total_cost_changed(ByVal whichone As String)

    If form_loading = False Then
        If whichone = "std" Then
            std_NON_total_cost_changed = True
            std_total_cost_changed = False
        ElseIf whichone = "rr" Then
            rr_NON_total_cost_changed = True
        ElseIf whichone = "opn" Then
            opn_NON_total_cost_changed = True
        ElseIf whichone = "metric" Then
            metric_NON_total_cost_changed = True
        ElseIf whichone = "res" Then
            res_NON_total_cost_changed = True
        ElseIf whichone = "std_op" Then
            std_NON_total_cost_op_changed = True
        ElseIf whichone = "rr_op" Then
            rr_NON_total_cost_op_changed = True
        ElseIf whichone = "opn_op" Then
            opn_NON_total_cost_op_changed = True
        ElseIf whichone = "metric_op" Then
            metric_NON_total_cost_op_changed = True
        ElseIf whichone = "res_op" Then
            res_NON_total_cost_op_changed = True
        End If
    End If
    
End Sub

Private Sub std_total_cost_op_Change()
    If form_loading = False Then
        std_total_cost_op_changed = True
    End If
End Sub

Private Sub tech_desc_Change()
Dim intLength As Integer
Dim intPosition As Integer
Dim txtSavetech_desc As String
Dim txtNewtech_desc As String

If Len(tech_desc) > 0 Then
    intPosition = tech_desc.SelStart
    If intPosition > 0 Then
        If Asc(Mid(tech_desc, intPosition, 1)) >= 0 And Asc(Mid(tech_desc, intPosition, 1)) <= 31 Then
            intLength = Len(tech_desc)
            txtSavetech_desc = tech_desc.Text
            MsgBox "Non-printable characters are not allowed in the tech_description."
            txtNewtech_desc = Left(txtSavetech_desc, intPosition - 2)
            If intPosition < intLength Then
                txtNewtech_desc = txtNewtech_desc + Right(txtSavetech_desc, intLength - intPosition)
            End If
            tech_desc.Text = txtNewtech_desc
            tech_desc.SelStart = intPosition - 2
        End If
    End If
End If

End Sub

Private Sub type_code_GotFocus()
    
    If type_code <> "M" Then
        m_type_code = type_code
    End If

End Sub

Private Sub type_code_Validate(Cancel As Boolean)
If IsNumeric(unit_cost_skey) Then   'not numeric in add mode, no need to validate
    Cancel = Not validate_uc_type_code(type_code, CLng(unit_cost_skey))
Else
'    If type_code = "B" Then
'        'type_code.Enabled = False
'        cmdUpdate.Enabled = False
'        MsgBox "Type Code B is not a valid type code for a new record!", vbOK
'    End If
End If
End Sub

Private Sub unit_cost_id_LostFocus()
    strLast_unit_cost_id = unit_cost_id.Text
End Sub

Private Sub unit_cost_id_Validate(Cancel As Boolean)
    Dim bln_New As Boolean
    Dim sTable As String
    
    If strLast_unit_cost_id <> unit_cost_id.Text Or unit_cost_id.Text = "" Then
        If m_blnInsert Or m_blnClone Then
            bln_New = True
        End If
        If MasterFormat = EXT_MASTERFORMAT_VERSION Then
            ' In MF04 mode, UNIT_COST_ID is the MF04 ID (in the UCD_EXT table)
            sTable = "UNIT_COST_DETAIL_EXT"
        End If
        If Invalid_ID_Format(Compress_String(unit_cost_id), "unit_cost_id", m_rec, bln_New, sTable) = True Then
            Cancel = True
        End If
    End If
End Sub

Private Sub FillUsageGrid()

    On Error GoTo Error_Processing
    Dim strSelect As String
    Dim blnReturn As Boolean

        ' Check to see if the mat_id entered exists already
    If m_blnClone = True Then
        'If cloning, set skey to 0 after reading data
        strSelect = "Select mu.mat_skey, mu.unit_cost_skey, mu.unit_qty, mu.input_factor, mu.output_factor, mu.adj_factor, mu.last_update_person, mu.last_update_date, mu.last_update_id, mu.comment, m.mat_id from material_usage as mu, material as m where mu.mat_skey = m.mat_skey and mu.unit_cost_skey = " + str(m_lngOriginalSkey) + " order by m.mat_id"
    Else
        If unit_cost_skey.Text = "" Then
            strSelect = "Select mu.mat_skey, mu.unit_cost_skey as unit_cost_skey, mu.unit_qty, mu.input_factor, mu.output_factor, mu.adj_factor, mu.last_update_person, mu.last_update_date, mu.last_update_id, mu.comment, m.mat_id from material_usage as mu, material as m where mu.mat_skey = m.mat_skey and mu.unit_cost_skey = 0 order by m.mat_id"
        Else
            strSelect = "Select mu.mat_skey, mu.unit_cost_skey, mu.unit_qty, mu.input_factor, mu.output_factor, mu.adj_factor, mu.last_update_person, mu.last_update_date, mu.last_update_id, mu.comment, m.mat_id from material_usage as mu, material as m where mu.mat_skey = m.mat_skey and mu.unit_cost_skey = " + unit_cost_skey.Text + " order by m.mat_id"
        End If
    End If

    ' Use DAL to perform select
    m_rec2.Close
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec2)
    If Not IsNumeric(unit_cost_skey.Text) Then
        unit_cost_skey.Text = 0
    End If
    m_objGridMap.UnitCostSKey = CLng(unit_cost_skey.Text)
    If unit_cost_skey.Text = "" And m_rec2.RecordCount > 0 Then
        m_rec2.MoveFirst
        Do Until m_rec2.EOF
            m_rec2.Fields("unit_cost_skey") = 0
            m_rec2.MoveNext
        Loop
    End If

    m_objGridMap.RecordSet = m_rec2
    If m_blnClone = True Then
        blnReturn = m_objGridMap.SetRowStateNew
    Else
        blnReturn = m_objGridMap.SetRowStateNone
    End If
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec2.RecordCount
Exit_Sub:
Exit Sub

Error_Processing:
If Err = 3704 Then ' object closed, ignore
    Resume Next
Else
    Screen.MousePointer = vbNormal
    MsgBox Error$
    Resume Exit_Sub
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button
    Dim blnPendingChange As Boolean
    Dim bln_New As Boolean

    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        blnPendingChange = IsControlChanged(Me, m_rec)
        If blnPendingChange = True Or m_objGridMap.IsPendingChange Or pct_ind <> pct_ind_o Then
            Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
            If Button = vbYes Then
                m_blnWereErrors = False
                If m_blnInsert Or m_blnClone Then
                    bln_New = True
                End If
                If alt_unit_cost_id <> "" Then
                    m_blnWereErrors = Invalid_ID_Format(Compress_String(alt_unit_cost_id), "alt_unit_cost_id", m_rec, bln_New)
                End If
                If Not m_blnWereErrors Then
                    m_blnWereErrors = Invalid_ID_Format(Compress_String(unit_cost_id), "unit_cost_id", m_rec, bln_New)
                End If
                If Not m_blnWereErrors Then
                    m_blnWereErrors = Not validate_crew_qty
                End If
                If Not m_blnWereErrors Then
                    m_blnWereErrors = Not validate_crew_id
                End If
                If Not m_blnWereErrors Then
                    m_blnWereErrors = Not validate_uc_type_code(type_code, CLng(unit_cost_skey))
                End If
                If m_blnWereErrors Then
                    Cancel = True
                Else
                    cmdUpdate_Click
                    ' If there were errors, cancel the close
                    If m_blnWereErrors Then
                        Cancel = True
                    Else
                        RestoreGridValues
                    End If
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

Private Sub type_code_LostFocus()
Dim blnResult As Boolean
    On Error Resume Next
    
    If type_code.Text = "E" Then
        fraUnitCost.Visible = False
        SSTab1.TabEnabled(2) = True
        pct_ind.Visible = True
    ElseIf type_code.Text = "M" Then
        If m_type_code <> type_code Then
            MsgBox "Make sure you attach at least one material.", vbInformation
            m_type_code = type_code
        End If
        FillUsageGrid
        SSTab1.TabEnabled(2) = True
'        fraUnitCost.Move 120, 3660
        fraUnitCost.Visible = True
        fraUnitCost.Move 120
        pct_ind.Visible = False
    Else
        fraUnitCost.Visible = False
        If SSTab1.Tab = 2 Then
            SSTab1.Tab = 0
        End If
        SSTab1.TabEnabled(2) = False
    End If
    If type_code.Text = "E" Or type_code.Text = "H" Then
        blnResult = LockField(Me, "crew_id")
        blnResult = LockField(Me, "daily_output")
'        blnResult = LockField(Me, "metric_daily_output")
        blnResult = LockField(Me, "crew_qty")
    Else
        blnResult = UnLockField(Me, "crew_id")
        blnResult = UnLockField(Me, "crew_qty")
        blnResult = UnLockField(Me, "daily_output")
'        blnResult = UnLockField(Me, "metric_daily_output")
    End If
    If type_code.Text = "E" Then
        blnResult = UnLockField(Me, "std_mat_cost")
        blnResult = UnLockField(Me, "std_labor_cost")
        blnResult = UnLockField(Me, "std_mat_cost_op")
        blnResult = UnLockField(Me, "std_labor_cost_op")
        blnResult = UnLockField(Me, "std_equip_cost")
        blnResult = UnLockField(Me, "std_equip_cost_op")
        blnResult = UnLockField(Me, "std_total_cost")
        blnResult = UnLockField(Me, "std_total_cost_op")
        blnResult = UnLockField(Me, "opn_mat_cost")
        blnResult = UnLockField(Me, "opn_mat_cost_op")
        blnResult = UnLockField(Me, "opn_labor_cost")
        blnResult = UnLockField(Me, "opn_labor_cost_op")
        blnResult = UnLockField(Me, "opn_equip_cost")
        blnResult = UnLockField(Me, "opn_equip_cost_op")
        blnResult = UnLockField(Me, "opn_total_cost")
        blnResult = UnLockField(Me, "opn_total_cost_op")
        blnResult = UnLockField(Me, "rr_mat_cost")
        blnResult = UnLockField(Me, "rr_mat_cost_op")
        blnResult = UnLockField(Me, "rr_labor_cost")
        blnResult = UnLockField(Me, "rr_labor_cost_op")
        blnResult = UnLockField(Me, "rr_equip_cost")
        blnResult = UnLockField(Me, "rr_equip_cost_op")
        blnResult = UnLockField(Me, "rr_total_cost")
        blnResult = UnLockField(Me, "rr_total_cost_op")
        blnResult = UnLockField(Me, "metric_mat_cost")
        blnResult = UnLockField(Me, "metric_mat_cost_op")
        blnResult = UnLockField(Me, "metric_labor_cost")
        blnResult = UnLockField(Me, "metric_labor_cost_op")
        blnResult = UnLockField(Me, "metric_equip_cost")
        blnResult = UnLockField(Me, "metric_equip_cost_op")
        blnResult = UnLockField(Me, "metric_total_cost")
        blnResult = UnLockField(Me, "metric_total_cost_op")
        blnResult = UnLockField(Me, "metric_unit")
        'RESI
        blnResult = UnLockField(Me, "res_mat_cost")
        blnResult = UnLockField(Me, "res_mat_cost_op")
        blnResult = UnLockField(Me, "res_labor_cost")
        blnResult = UnLockField(Me, "res_labor_cost_op")
        blnResult = UnLockField(Me, "res_equip_cost")
        blnResult = UnLockField(Me, "res_equip_cost_op")
        blnResult = UnLockField(Me, "res_total_cost")
        blnResult = UnLockField(Me, "res_total_cost_op")
        'FMR (In-House)
        blnResult = UnLockField(Me, "inhouse_mat_cost_op")
        blnResult = UnLockField(Me, "inhouse_labor_cost_op")
        blnResult = UnLockField(Me, "inhouse_equip_cost_op")
        blnResult = UnLockField(Me, "inhouse_total_cost_op")
    Else
        blnResult = LockField(Me, "metric_unit")
        blnResult = LockField(Me, "std_mat_cost")
        blnResult = LockField(Me, "std_labor_cost")
        blnResult = LockField(Me, "std_mat_cost_op")
        blnResult = LockField(Me, "std_labor_cost_op")
        blnResult = LockField(Me, "std_equip_cost")
        blnResult = LockField(Me, "std_equip_cost_op")
        blnResult = LockField(Me, "std_total_cost")
        blnResult = LockField(Me, "std_total_cost_op")
        blnResult = LockField(Me, "opn_mat_cost")
        blnResult = LockField(Me, "opn_mat_cost_op")
        blnResult = LockField(Me, "opn_labor_cost")
        blnResult = LockField(Me, "opn_labor_cost_op")
        blnResult = LockField(Me, "opn_equip_cost")
        blnResult = LockField(Me, "opn_equip_cost_op")
        blnResult = LockField(Me, "opn_total_cost")
        blnResult = LockField(Me, "opn_total_cost_op")
        blnResult = LockField(Me, "rr_mat_cost")
        blnResult = LockField(Me, "rr_mat_cost_op")
        blnResult = LockField(Me, "rr_labor_cost")
        blnResult = LockField(Me, "rr_labor_cost_op")
        blnResult = LockField(Me, "rr_equip_cost")
        blnResult = LockField(Me, "rr_equip_cost_op")
        blnResult = LockField(Me, "rr_total_cost")
        blnResult = LockField(Me, "rr_total_cost_op")
        blnResult = LockField(Me, "metric_mat_cost")
        blnResult = LockField(Me, "metric_mat_cost_op")
        blnResult = LockField(Me, "metric_labor_cost")
        blnResult = LockField(Me, "metric_labor_cost_op")
        blnResult = LockField(Me, "metric_equip_cost")
        blnResult = LockField(Me, "metric_equip_cost_op")
        blnResult = LockField(Me, "metric_total_cost")
        blnResult = LockField(Me, "metric_total_cost_op")
        'RESI
        blnResult = LockField(Me, "res_mat_cost")
        blnResult = LockField(Me, "res_mat_cost_op")
        blnResult = LockField(Me, "res_labor_cost")
        blnResult = LockField(Me, "res_labor_cost_op")
        blnResult = LockField(Me, "res_equip_cost")
        blnResult = LockField(Me, "res_equip_cost_op")
        blnResult = LockField(Me, "res_total_cost")
        blnResult = LockField(Me, "res_total_cost_op")
        'FMR
        blnResult = LockField(Me, "inhouse_mat_cost_op")
        blnResult = LockField(Me, "inhouse_labor_cost_op")
        blnResult = LockField(Me, "inhouse_equip_cost_op")
        blnResult = LockField(Me, "inhouse_total_cost_op")
    End If
    If type_code.Text = "H" Then
        blnResult = LockField(Me, "unit")
        metric_unit = ""
        unit = ""
        std_mat_cost = 0
        std_mat_cost_op = 0
        std_labor_cost = 0
        std_labor_cost_op = 0
        std_equip_cost = 0
        std_equip_cost_op = 0
        std_total_cost = 0
        std_equip_cost_op = 0
        std_total_cost = 0
        std_total_cost_op = 0
        opn_mat_cost = 0
        opn_mat_cost_op = 0
        opn_labor_cost = 0
        opn_labor_cost_op = 0
        opn_equip_cost = 0
        opn_equip_cost = 0
        opn_total_cost = 0
        opn_total_cost_op = 0
        rr_mat_cost = 0
        rr_mat_cost_op = 0
        rr_labor_cost = 0
        rr_labor_cost_op = 0
        rr_equip_cost = 0
        rr_equip_cost_op = 0
        rr_total_cost = 0
        rr_total_cost_op = 0
        metric_mat_cost = 0
        metric_mat_cost_op = 0
        metric_labor_cost = 0
        metric_labor_cost_op = 0
        metric_equip_cost = 0
        metric_equip_cost_op = 0
        metric_total_cost = 0
        metric_total_cost_op = 0
        'RESI
        res_mat_cost = 0
        res_mat_cost_op = 0
        res_labor_cost = 0
        res_labor_cost_op = 0
        res_equip_cost = 0
        res_equip_cost_op = 0
        res_total_cost = 0
        res_total_cost_op = 0
        'FMR
        inhouse_mat_cost_op = 0
        inhouse_labor_cost_op = 0
        inhouse_equip_cost_op = 0
        inhouse_total_cost_op = 0
    Else
        blnResult = UnLockField(Me, "unit")
    End If
    Me.Refresh
    DoEvents
    
End Sub

Private Function LockField(frm As Form, fld As String) As Boolean
'7/27/2005 MODIFIED TO USE ALTERNATE LOCK COLOR (IF USER ENABLED)
    frm.Controls(fld).Enabled = False
    frm.Controls(fld).Locked = True
    frm.Controls(fld).ForeColor = LTGREY
    If Not g_blnUseAlternateDisabledColor Then
        frm.Controls(fld).BackColor = vbButtonFace
    Else
        frm.Controls(fld).BackColor = g_intAlternateRowColor
    End If
    
End Function

Private Sub set_pct_ind()
    If m_rec.Fields("percent_flag").Value = "Y" Then
        pct_ind = 1
    Else
        pct_ind = 0
    End If
End Sub

Private Sub GetLongDescriptions()
'RETRIEVE LONG TECHNICAL DESCRIPTIONS FROM DATABASE
'FOR DISPLAY ON 'DESCRIPTION' TAB
'ADDED 6/6/2005 RTD FOR VERSION 7.3.0
'MODIFIED 8/22/2005 RTD FOR VERSION 7.5.0 - SUPPORT UCD_EXT TABLE
    Dim rsTemp As ADODB.RecordSet
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim sUnitCostId As String
    Dim iUnitCostSkey As Long
    Dim sDesc As String
    Dim sTable As String
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    txtLongDescriptionI.Text = ""
    txtLongDescriptionM.Text = ""
    iUnitCostSkey = unit_cost_skey.Text
    
    If MasterFormat = EXT_MASTERFORMAT_VERSION Then
        sTable = "UNIT_COST_DETAIL_EXT"
        sUnitCostId = Compress_String(Me.ext_unit_cost_id.Text)
        LockField Me, "txtLongDescriptionI"
        LockField Me, "txtLongDescriptionM"
        'MODIFIED 9/7/2005 RTD - MASTERFORMAT 2004 IDs ARE NOW SUPPORTED,
        'ALLOW LONG DESC BUTTON TO EDIT LONG DESCRIPTIONS
        'cmdLongDesc.Enabled = False
    Else
         'rlh 05/22/2007  SUPPRESS update of MF-1995 DESCRIPTION FIELDS
         
         'rlh 07/08/2008  (RE)ENABLED DESCRIPTION FIELDS REGARDLESS OF MF SETTING!!!
'''         If Not m_blnInsert Then                'rlh 05/22/2007
'''            LockField Me, "tech_desc"
'''            LockField Me, "metric_tech_desc"
'''
'''            LockField Me, "book_desc"
'''            LockField Me, "metric_book_desc"
'''
'''            LockField Me, "assembly_book_desc"
'''            LockField Me, "metric_assembly_book_desc"
'''         End If                                 'rlh 05/22/2007 - end of my changes
         
        sTable = "UNIT_COST_DETAIL"
        sUnitCostId = Compress_String(Me.unit_cost_id.Text)
    End If

    strSelect = "SELECT long_tech_desc, metric_long_tech_desc " & _
                " FROM " & sTable & _
                " WHERE unit_cost_skey = " & iUnitCostSkey
    'strSELECT = "exec sp_select_attribute_value @min_object_id = '" & sUnitCostId & "'," & _
               "@max_object_id = '" & sUnitCostId & "'," & _
               "@skey_type = 'U', @meas_sys_cd = 'A', @obj_desc_filter = ''"
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsTemp)
    If Not rsTemp.EOF Then
        sDesc = rsTemp.Fields("long_tech_desc") & ""
        txtLongDescriptionI.Text = sDesc
        sDesc = rsTemp.Fields("metric_long_tech_desc") & ""
        txtLongDescriptionM.Text = sDesc
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    Screen.MousePointer = vbNormal
    Exit Sub
    
Err_Handler:
    Screen.MousePointer = vbNormal
    Debug.Print "ERROR RETRIEVING LONG DESCRIPTION FOR " & sUnitCostId & ": " & Err.Description
    Exit Sub
End Sub

Private Function CloneLongDescriptions(ByVal sOldObjCostID As String, ByVal sNewObjSKey As String) As Boolean
'COPY NEW LONG DESCRIPTION RECORDS FROM OLD ITEM
'ADDED 6/9/2005 RTD
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim rsTemp As ADODB.RecordSet
    Dim i As Long
    Dim sRowMSys As String
    Dim sText As String
    Dim sMSys As String
    
    Screen.MousePointer = vbHourglass
    blnReturn = True
    On Error GoTo Err_Handler
    ' Get Long Description records
    ' MODIFIED 9/6/2005 RTD - FIX PROBLEM REPORTED BY G. SPENCER
    ' MASTERFORMAT 2004 AWARE STORED PROC: usp_select_attribute_value_ext
    strSelect = "exec usp_select_attribute_value_ext @min_object_id = '" & sOldObjCostID & "'," & _
        "@max_object_id = '" & sOldObjCostID & "'," & _
        "@skey_type = 'U', @meas_sys_cd = 'A', @obj_desc_filter = ''," & _
        "@master_format = " & MasterFormat
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsTemp)
'04/02/2009  BOB MEWIS failing trying to clone a unit cost line w/o long descriptions!!!
'            No fields named "col_" exist in the returned recordset!!!
    If IsGridColsExist(rsTemp) Then  'rlh 04/02/2009
    ' Cycle through description columns
    Do While Not rsTemp.EOF
        For i = 1 To 8
            If Not IsNull(rsTemp.Fields("col_" & i).Value) Then
                sRowMSys = rsTemp.Fields("row_meas_sys_cd")
                sText = rsTemp.Fields("col_" & i).Value
                sMSys = rsTemp.Fields("msc_" & i).Value
                ' Append data using new SKey
                 
                strSelect = "exec sp_update_object_attribute_value2 " & _
                        "" & sNewObjSKey & ", 'U', " & i & ", NULL, " & _
                        "'" & SQLFixString(sText) & "', " & _
                        "'', '" & sMSys & "', '" & sRowMSys & "', 0"
                g_cnShared.Execute (strSelect)
                If g_cnShared.Errors.Count > 0 Then
                
                Debug.Print g_cnShared.Errors(0).Description
                
                End If
            End If
        Next
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    'Update the UNIT_COST_DETAIL/UNIT_COST_DETAIL_EXT long description fields
    strSelect = "exec sp_update_object_description " & sNewObjSKey & ", 'U', 1"
    g_cnShared.Execute (strSelect)
    End If   'rlh 04/02/2009
    
    Set rsTemp = Nothing
    Screen.MousePointer = vbNormal
    CloneLongDescriptions = blnReturn
    Exit Function
    
Err_Handler:
    Screen.MousePointer = vbNormal
    'Stop   'RLH 04/02/2009  (PER BOB MEWIS ISSUE)
    Debug.Print "ERROR CLONING LONG DESCRIPTION FOR " & sNewObjSKey & ": " & Err.Description
    CloneLongDescriptions = False
    m_blnWereErrors = False  'RLH 04/25/08
    Exit Function
    
End Function

Private Function AddNewLongDescriptions(ByVal sObjSKey As String, ByVal sImpDesc As String, ByVal sMetDesc As String) As Boolean
'INSERT NEW LONG DESCRIPTION RECORDS FOR ITEM
'ADDED 6/9/2005 RTD
    Dim strSelect As String
    Dim blnReturn As Boolean
    Dim sNewText As String
    
    If Not IsNumeric(sObjSKey) Then
        Exit Function
    End If
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    'Insert description
    sNewText = "new item"
    strSelect = "exec sp_update_object_attribute_value2 " & _
                    "" & sObjSKey & ", 'U', 1, NULL, '" & sNewText & "', '', 'A', 'I', 0"
    g_cnShared.Execute (strSelect)
    'Insert Imperial book description in Dimension field
    sNewText = SQLFixString(sImpDesc)
    If sNewText <> "" Then
        strSelect = "exec sp_update_object_attribute_value2 " & _
                    "" & sObjSKey & ", 'U', 4, NULL, '" & sNewText & "', '', 'I', 'I', 0"
        g_cnShared.Execute (strSelect)
    End If
    'Insert Metric book description in Dimension field
    sNewText = SQLFixString(sMetDesc)
    If sNewText <> "" Then
        strSelect = "exec sp_update_object_attribute_value2 " & _
                    "" & sObjSKey & ", 'U', 4, NULL, '" & sNewText & "', '', 'M', 'M', 0"
        g_cnShared.Execute (strSelect)
    End If
    'Update the UNIT_COST_DETAIL field
    strSelect = "exec sp_update_object_description " & sObjSKey & ", 'U', 1"
    g_cnShared.Execute (strSelect)
    blnReturn = True
    AddNewLongDescriptions = blnReturn
    Screen.MousePointer = vbNormal
    Exit Function
    
Err_Handler:
    Screen.MousePointer = vbNormal
    Debug.Print "ERROR ADDING LONG DESCRIPTION FOR sKEY " & sObjSKey & ": " & Err.Description
    blnReturn = False
    AddNewLongDescriptions = blnReturn
    Exit Function
End Function

Public Sub UpdateBookPreviewLine()
'UPDATE FORMATTING PREVIEW AREA
    Dim sTemp1 As String
    Dim sTemp2 As String
    Dim iIndent As Long
    Dim sUCID As String
    
    On Error GoTo Err_Handler
    sUCID = Compress_String(unit_cost_id.Text)
    Shape1.BackColor = vbWhite
    Shape1.BorderColor = vbBlack
    Line4.BorderColor = vbBlack
    lblPreview1.ForeColor = vbBlack
    lblPreview2.ForeColor = vbBlack
    lblPreview3.ForeColor = vbBlack
    If format_code.Text = "H1" Then
        Shape1.BackColor = vbBlack
        Shape1.BorderColor = vbWhite
        Line4.BorderColor = vbWhite
        lblPreview1.ForeColor = vbWhite
        lblPreview2.ForeColor = vbWhite
        lblPreview3.ForeColor = vbWhite
        lblPreview1.FontBold = True
        lblPreview2.FontBold = True
        sTemp1 = Replace(book_desc.Text, "&", "&&")
        If MasterFormat = EXT_MASTERFORMAT_VERSION Then
            lblPreview1.Caption = Left(sUCID, 6)
        Else
            lblPreview1.Caption = Left(sUCID, 5)
        End If
        lblPreview2.Caption = sTemp1
        lblPreview3.Caption = ""
    ElseIf Left(format_code.Text, 1) = "H" Then
        lblPreview1.FontBold = True
        lblPreview2.FontBold = True
        sTemp1 = Replace(book_desc.Text, "&", "&&")
        If MasterFormat = EXT_MASTERFORMAT_VERSION Then
            lblPreview1.Caption = Left(sUCID, 6)
        Else
            lblPreview1.Caption = Left(sUCID, 5)
        End If
        lblPreview2.Caption = sTemp1
        lblPreview3.Caption = ""
    Else
        lblPreview1.FontBold = False
        lblPreview1.Caption = Right(sUCID, 4)
        iIndent = Val(indent_code.Text)
        If iIndent < 0 Then iIndent = 0
        If Val(format_characters.Text) > 0 Then
            sTemp1 = Left(book_desc.Text, Val(format_characters.Text))
            sTemp2 = Mid(book_desc.Text, Val(format_characters.Text) + 1)
            sTemp1 = Replace(sTemp1, "&", "&&")
            sTemp2 = Replace(sTemp2, "&", "&&")
            lblPreview2.Caption = Space(iIndent * 4) & sTemp1
            lblPreview3.Caption = sTemp2
            lblPreview2.FontBold = True
        Else
            sTemp1 = Replace(book_desc.Text, "&", "&&")
            lblPreview2.FontBold = False
            lblPreview2.Caption = Space(iIndent * 4) & sTemp1
            lblPreview3.Caption = ""
        End If
        lblPreview3.Left = lblPreview2.Left + lblPreview2.Width
    End If
    Exit Sub

Err_Handler:
    lblPreview2.Caption = ""
    lblPreview3.Caption = ""
    Exit Sub
    
End Sub

Public Function IsGridColsExist(rs As ADODB.RecordSet) As Boolean
'04/02/2009  BOB MEWIS failing trying to clone a unit cost line w/o long descriptions!!!
'            No fields named "col_" exist in the returned recordset!!!

Dim i As Integer
On Error GoTo ERRLBL

IsGridColsExist = False
For i = 0 To rs.Fields.Count - 1
    If (InStr(rs.Fields(i).Name, "col_") > 0) Then
        IsGridColsExist = True
        Exit Function
    End If
Next
Exit Function
ERRLBL:
    MsgBox ("Error: IsGridColsExist: " & Err.Description)
    Stop
    Resume
End Function
