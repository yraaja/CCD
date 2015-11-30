VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmInfoSource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information Source"
   ClientHeight    =   5265
   ClientLeft      =   1740
   ClientTop       =   1890
   ClientWidth     =   9510
   Icon            =   "frmInfoSource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9510
   Begin VB.TextBox last_update_id 
      Height          =   315
      Left            =   7140
      Locked          =   -1  'True
      TabIndex        =   68
      Tag             =   "N"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   3240
      TabIndex        =   36
      Top             =   4620
      Width           =   1150
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4800
      TabIndex        =   37
      Top             =   4620
      Width           =   1150
   End
   Begin VB.TextBox txtContactId 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   8340
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtCompanyName 
      BackColor       =   &H00C0C0C0&
      DataField       =   "company_name"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4500
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   2535
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   3855
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Company"
      TabPicture(0)   =   "frmInfoSource.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblURL"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFax"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPhone2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPhone1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblZip"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblState"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCity"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblAddress3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblAddress2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblAddress1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCompanyName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblCountryCode"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblContactId"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fax"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "phone2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "zip_code"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "phone1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdGoWeb"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "state_code"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "country_code"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "address3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "company_name"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "city"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "address2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "address1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "URL"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "contact_id"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "&Personal"
      TabPicture(1)   =   "frmInfoSource.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblNickname"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblTitle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblSalutation"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblSuffix"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblMI"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblFirstName"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblLastName"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblEmail"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "nickname"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "middle_initial"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "name_suffix"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "salutation"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "title"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "first_name"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "last_name"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "email"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "&Means Tracking"
      TabPicture(2)   =   "frmInfoSource.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblComments"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblLastUpdateDate"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblLastUpdatePerson"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblTicklerDate"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblSourceCode"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblCreatePerson"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblCreateDate"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblKeyword"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "tickler_date"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "comment"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "last_update_person"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "create_person"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "last_update_date"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "source_code"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "create_date"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "keyword"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "council_assoc_ind"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "City Cost Index"
      TabPicture(3)   =   "frmInfoSource.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label4"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cci_use_ind"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cci_update_cd"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cci_letter_cd"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cci_metro_cd"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cci_contact_nm"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      Begin VB.TextBox cci_contact_nm 
         Height          =   285
         Left            =   -70200
         MaxLength       =   25
         TabIndex        =   78
         Tag             =   "1S"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox cci_metro_cd 
         Height          =   285
         Left            =   -70200
         TabIndex        =   75
         Tag             =   "1S"
         Top             =   2130
         Width           =   735
      End
      Begin VB.TextBox cci_letter_cd 
         Height          =   285
         Left            =   -70200
         TabIndex        =   73
         Tag             =   "1S"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox cci_update_cd 
         Height          =   285
         Left            =   -70200
         MaxLength       =   1
         TabIndex        =   71
         Tag             =   "1S"
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox cci_use_ind 
         Caption         =   "CCI Use Indicator"
         Height          =   375
         Left            =   -70800
         TabIndex        =   70
         Tag             =   "1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox contact_id 
         DataField       =   "contact_id"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "1S"
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox URL 
         DataField       =   "URL"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1500
         MaxLength       =   30
         MouseIcon       =   "frmInfoSource.frx":04B2
         TabIndex        =   13
         Tag             =   "1S"
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox address1 
         DataField       =   "address1"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "1S"
         Top             =   1380
         Width           =   2895
      End
      Begin VB.TextBox address2 
         DataField       =   "address2"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "1S"
         Top             =   1830
         Width           =   2895
      End
      Begin VB.TextBox city 
         DataField       =   "address3"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1500
         MaxLength       =   23
         TabIndex        =   6
         Tag             =   "1S"
         Top             =   2760
         Width           =   2115
      End
      Begin VB.TextBox company_name 
         DataField       =   "company_name"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "1S"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox address3 
         Height          =   315
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "1S"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.ComboBox country_code 
         Height          =   315
         Left            =   7920
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox state_code 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "1"
         Top             =   2760
         Width           =   675
      End
      Begin VB.CheckBox council_assoc_ind 
         Caption         =   "Council Assoc"
         Height          =   315
         Left            =   -68760
         TabIndex        =   28
         Tag             =   "1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox keyword 
         Height          =   315
         Left            =   -73500
         MaxLength       =   120
         TabIndex        =   30
         Tag             =   "1S"
         Top             =   1080
         Width           =   6135
      End
      Begin VB.TextBox create_date 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   -70920
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox email 
         DataField       =   "email"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -73500
         MaxLength       =   80
         TabIndex        =   26
         Tag             =   "1S"
         Top             =   2100
         Width           =   4455
      End
      Begin VB.CommandButton cmdGoWeb 
         Caption         =   "Go"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   14
         Top             =   3240
         Width           =   435
      End
      Begin VB.TextBox source_code 
         DataField       =   "source_code"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -69540
         MaxLength       =   2
         TabIndex        =   27
         Tag             =   "1S"
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox last_update_date 
         BackColor       =   &H00C0C0C0&
         DataField       =   "last_update_date"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -70920
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2460
         Width           =   1935
      End
      Begin VB.TextBox create_person 
         BackColor       =   &H00C0C0C0&
         DataField       =   "create_person"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -73500
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox last_update_person 
         BackColor       =   &H00C0C0C0&
         DataField       =   "last_update_person"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -73500
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2460
         Width           =   1335
      End
      Begin VB.TextBox comment 
         DataField       =   "comments"
         DataSource      =   "Data1"
         Height          =   735
         Left            =   -73500
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   31
         Tag             =   "1S"
         Top             =   1500
         Width           =   6135
      End
      Begin VB.TextBox last_name 
         DataField       =   "last_name"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -73500
         MaxLength       =   20
         TabIndex        =   19
         Tag             =   "1S"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox first_name 
         DataField       =   "first_name"
         DataSource      =   "Data1"
         Height          =   325
         Left            =   -70380
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "1S"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox title 
         DataField       =   "title"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -73500
         MaxLength       =   20
         TabIndex        =   25
         Tag             =   "1S"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox salutation 
         DataField       =   "salutation"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -73500
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "1S"
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox name_suffix 
         DataField       =   "name_suffix"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -67320
         MaxLength       =   5
         TabIndex        =   22
         Tag             =   "1S"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox middle_initial 
         DataField       =   "middle_initial"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -68580
         MaxLength       =   1
         TabIndex        =   21
         Tag             =   "1S"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox nickname 
         DataField       =   "nickname"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -70380
         MaxLength       =   20
         TabIndex        =   24
         Tag             =   "1S"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox phone1 
         DataField       =   "phone1"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5760
         TabIndex        =   10
         Tag             =   "1S"
         Top             =   960
         Width           =   1515
      End
      Begin VB.TextBox zip_code 
         DataField       =   "zip_code"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "1S"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox phone2 
         DataField       =   "fax"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5760
         TabIndex        =   11
         Tag             =   "1S"
         Top             =   1380
         Width           =   1515
      End
      Begin VB.TextBox fax 
         DataField       =   "phone2"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5760
         TabIndex        =   12
         Tag             =   "1S"
         Top             =   1800
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker tickler_date 
         DataField       =   "tickler_date"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   -73500
         TabIndex        =   29
         Tag             =   "1D"
         Top             =   600
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   22806531
         CurrentDate     =   36103
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "CCI Contact Name:"
         Height          =   255
         Left            =   -71760
         TabIndex        =   77
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Metro Code:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   76
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Letter Code:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   74
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Update Code:"
         Height          =   255
         Left            =   -71520
         TabIndex        =   72
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label lblContactId 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact ID:"
         Height          =   255
         Left            =   480
         TabIndex        =   69
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lblKeyword 
         Alignment       =   1  'Right Justify
         Caption         =   "Keywords:"
         Height          =   255
         Left            =   -74700
         TabIndex        =   67
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label lblCreateDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Create Date:"
         Height          =   255
         Left            =   -72000
         TabIndex        =   66
         Top             =   2940
         Width           =   975
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Left            =   -74220
         TabIndex        =   65
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblCountryCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Country:"
         Height          =   255
         Left            =   7200
         TabIndex        =   64
         Top             =   2820
         Width           =   675
      End
      Begin VB.Label lblCreatePerson 
         Alignment       =   1  'Right Justify
         Caption         =   "Created By:"
         Height          =   255
         Left            =   -74460
         TabIndex        =   61
         Top             =   2940
         Width           =   855
      End
      Begin VB.Label lblSourceCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Source Code:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   60
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label lblTicklerDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Tickler Date:"
         Height          =   255
         Left            =   -74520
         TabIndex        =   59
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblLastUpdatePerson 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Updated By:"
         Height          =   255
         Left            =   -74940
         TabIndex        =   58
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblLastUpdateDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Updated:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   57
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblComments 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments:"
         Height          =   255
         Left            =   -74460
         TabIndex        =   56
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Name:"
         Height          =   255
         Left            =   -74460
         TabIndex        =   55
         Top             =   660
         Width           =   855
      End
      Begin VB.Label lblFirstName 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
         Height          =   255
         Left            =   -71280
         TabIndex        =   54
         Top             =   660
         Width           =   795
      End
      Begin VB.Label lblMI 
         Alignment       =   1  'Right Justify
         Caption         =   "MI:"
         Height          =   255
         Left            =   -68940
         TabIndex        =   53
         Top             =   660
         Width           =   255
      End
      Begin VB.Label lblSuffix 
         Alignment       =   1  'Right Justify
         Caption         =   "Suffix:"
         Height          =   255
         Left            =   -67920
         TabIndex        =   52
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lblSalutation 
         Alignment       =   1  'Right Justify
         Caption         =   "Salutation:"
         Height          =   255
         Left            =   -74460
         TabIndex        =   51
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Height          =   255
         Left            =   -74340
         TabIndex        =   50
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label lblNickname 
         Alignment       =   1  'Right Justify
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   -71340
         TabIndex        =   49
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCompanyName 
         Alignment       =   1  'Right Justify
         Caption         =   "Company:"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label lblAddress1 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         Height          =   255
         Left            =   480
         TabIndex        =   47
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblAddress2 
         Alignment       =   1  'Right Justify
         Caption         =   "Suppl Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label lblAddress3 
         Alignment       =   1  'Right Justify
         Caption         =   "Suppl Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label lblCity 
         Alignment       =   1  'Right Justify
         Caption         =   "City:"
         Height          =   255
         Left            =   720
         TabIndex        =   44
         Top             =   2820
         Width           =   615
      End
      Begin VB.Label lblState 
         Alignment       =   1  'Right Justify
         Caption         =   "State:"
         Height          =   255
         Left            =   3840
         TabIndex        =   43
         Top             =   2820
         Width           =   435
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip Code:"
         Height          =   255
         Left            =   5280
         TabIndex        =   42
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label lblPhone1 
         Alignment       =   1  'Right Justify
         Caption         =   "Telephone:"
         Height          =   255
         Left            =   4680
         TabIndex        =   41
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lblPhone2 
         Alignment       =   1  'Right Justify
         Caption         =   "Alt Phone:"
         Height          =   255
         Left            =   4680
         TabIndex        =   40
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblFax 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
         Height          =   255
         Left            =   5040
         TabIndex        =   39
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label lblURL 
         Alignment       =   1  'Right Justify
         Caption         =   "Web Site:"
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   3240
         Width           =   735
      End
   End
   Begin VB.TextBox txtContactName 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Expr1"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   780
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblContactIdHeading 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact Id:"
      Height          =   255
      Left            =   7380
      TabIndex        =   63
      Top             =   180
      Width           =   855
   End
   Begin VB.Label lblCompanyNameHeading 
      Alignment       =   1  'Right Justify
      Caption         =   "Company:"
      Height          =   255
      Index           =   4
      Left            =   3660
      TabIndex        =   62
      Top             =   180
      Width           =   735
   End
   Begin VB.Label lblContactNameHeading 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "frmInfoSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As ADODB.RecordSet
Dim m_blnRecFlag As Boolean ' True if the screen has a data source, False if we are adding new record
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload

' Fills all fields with data
Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If Not m_rec.Fields("contact_id") = "" Then
        m_blnRecFlag = True
    End If
    
    '9/29/2005 RTD - Place Letter Code into Keyword field if blank
    If keyword.Text = "" And cci_letter_cd.Text <> "" Then
        keyword.Text = cci_letter_cd.Text
    End If
    
End Sub

Private Sub cci_use_ind_Click()
    cci_letter_cd.Enabled = cci_use_ind.Value
    cci_update_cd.Enabled = cci_use_ind.Value
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

    strUpdate = "exec sp_delete_information_source "
    strUpdate = strUpdate + "@contact_id='" + Me.Controls("contact_id") + "', "
    strUpdate = strUpdate + "@last_update_person='" + strUserName + "'"
    
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

Private Sub cmdGoWeb_Click()
    Dim FileName, Dummy As String
    Dim BrowserString As String * 255
    Dim BrowserExec As String
    Dim RetVal As Long
    Dim FileNumber As Integer
    
    If URL.Text <> "" Then
        ' First, create a known, temporary HTML file
        BrowserString = Space(255)
        FileName = "C:\temphtm.HTM"
        FileNumber = FreeFile                    ' Get unused file number
        Open FileName For Output As #FileNumber  ' Create temp HTML file
        Write #FileNumber, "<HTML> <\HTML>"      ' Output text
        Close #FileNumber                        ' Close file
        
        ' Then find the application associated with it
        RetVal = FindExecutable(FileName, Dummy, BrowserString)
        BrowserExec = Trim(BrowserString)
        ' If an application is found, launch it!
        If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
            MsgBox "Could not find a Web Browser", vbExclamation, "Browser Not Found"
        Else
            RetVal = ShellExecute(Me.hWnd, "open", BrowserExec, URL.Text, Dummy, 1)
            If RetVal <= 32 Then        ' Error
                MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
            End If
        End If
        Kill FileName                   ' delete temp HTML file
    End If

End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim strError As String
    Dim blnRet As Boolean
    
    m_blnWereErrors = False
    
    ' If we are updating
    If m_blnInsert = False Then
        strUpdate = "exec sp_update_information_source @last_update_id=" + last_update_id.Text + ", "
        BuildStoredProcSQL Me, strUpdate, "1"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    ' If we are inserting
    Else
        strUpdate = "exec sp_insert_information_source "
        BuildStoredProcSQL Me, strUpdate, "1"
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
    End If
    
    blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
    If Not blnRet Then
        MsgBox strError
        m_blnWereErrors = True
    Else
        ' Put latest data into source recordset
        UpdateRecordsetFromForm Me, m_rec
        m_rec.Fields("last_update_id").Value = m_rec.Fields("last_update_id").Value + 1
        last_update_id.Text = m_rec.Fields("last_update_id").Value
        UpdateFormFromRecordset Me, m_rec
        MsgBox "Update successful."
    End If
End Sub

Private Sub contact_id_Change()
    With contact_id
        Dim OldCursorPos As Integer
        If Len(.Text) > 0 Then
            OldCursorPos = .SelStart
            .Text = UCase(.Text)
            .SelStart = OldCursorPos
        End If
    End With
End Sub

Private Sub fax_LostFocus()
    fax.Text = FormatPhoneNumber(fax.Text)
End Sub

Private Sub Form_Activate()
    OutputView False
End Sub

Private Sub Form_Initialize()
    m_blnRecFlag = False
    m_blnDeleted = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Move START_LEFT, START_TOP

    g_objDAL.GetRecordset vbNullString, "select state_code from state_country", rec
    While Not rec.EOF
        State_Code.AddItem (rec.Fields("state_code").Value)
        rec.MoveNext
    Wend
    rec.Close
    g_objDAL.GetRecordset vbNullString, "select country_code from country", rec
    While Not rec.EOF
        country_code.AddItem (rec.Fields("country_code").Value)
        rec.MoveNext
    Wend
    
    ' If we are showing data
'    If m_blnRecFlag = True Then
        If Not m_rec.State = adStateClosed Then
            UpdateFormFromRecordset Me, m_rec
        End If
        txtCompanyName = m_rec.Fields("company_name").Value
        txtContactName = m_rec.Fields("first_name").Value + " " + m_rec.Fields("last_name").Value
        txtContactId = m_rec.Fields("contact_id").Value
        ' Lock fields that can't be changed
    If m_blnInsert = False Then
        contact_id.Locked = True
        contact_id.BackColor = LTGREY
        Me.Caption = Me.Caption + " [" + m_rec.Fields("contact_id").Value + "]"
    Else
        ' If we are inserting and not showing data
        ' Set some defaults
        If Not m_blnRecFlag Then
            Dim dtm As Date
            dtm = Date
            m_rec.Fields("tickler_date").Value = DateAdd("yyyy", 1, dtm)
            tickler_date.Value = DateAdd("yyyy", 1, dtm)
            Me.Caption = Me.Caption + " [New]"
        Else
            Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("contact_id").Value + "]"
        End If
    End If
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
'        last_update_id.Text = 0
'    End If
    
    Tab1.Tab = 0
    
    URL.MaxLength = m_rec.Fields("url").DefinedSize
    ColorLockedFields Me
    
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
            End If
        End If
    End If
End Sub

Public Function FormatPhoneNumber(ByVal sPhoneNumber As String) As String
    Dim sPhone As String
    Dim sAreaCode As String
    
    sPhone = Trim(sPhoneNumber)
    sPhone = RemoveCharacters(sPhone, " +()-#")
    If Len(sPhone) = 10 Then
        sPhone = Mid(sPhone, 1, 3) & "-" & Mid(sPhone, 4, 3) & "-" & Mid(sPhone, 7)
    ElseIf Len(sPhone) = 7 Then
        sAreaCode = QueryRegistryKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Locations\Location1", "AreaCode", "")
        If sAreaCode <> "" Then
            sPhone = sAreaCode & "-" & Mid(sPhone, 1, 3) & "-" & Mid(sPhone, 4)
        Else
            sPhone = Mid(sPhone, 1, 3) & "-" & Mid(sPhone, 4)
        End If
    ElseIf Len(sPhone) = 11 And Left(sPhone, 1) = "1" Then
        sPhone = Mid(sPhone, 2)
        sPhone = Mid(sPhone, 1, 3) & "-" & Mid(sPhone, 4, 3) & "-" & Mid(sPhone, 7)
    Else
    
    End If
    FormatPhoneNumber = sPhone
    
End Function

Private Sub phone1_LostFocus()
    phone1.Text = FormatPhoneNumber(phone1.Text)
End Sub

Private Sub phone2_LostFocus()
    phone2.Text = FormatPhoneNumber(phone2.Text)
End Sub

Private Sub URL_Change()
    cmdGoWeb.Enabled = (URL.Text <> "")
    If URL.Text <> "" Then
        URL.ToolTipText = "Press Ctrl and Click to launch web site..."
    Else
        URL.ToolTipText = ""
    End If
End Sub

Private Sub URL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (Shift And vbCtrlMask) > 0 Then
        'Ctrl-Key is down, make field 'URL' formatted
        URL.MousePointer = vbCustom
        URL.ForeColor = vbBlue
        URL.FontUnderline = True
    Else
        URL.MousePointer = vbDefault
        URL.ForeColor = &H80000008
        URL.FontUnderline = False
    End If
    
End Sub

Private Sub URL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Left-Clicked and CTRL-Key down, launch web page
    If (URL.Text <> "") And (Button And vbLeftButton) > 0 And (Shift And vbCtrlMask) > 0 Then
        LaunchBrowser URL.Text
    End If
    
End Sub
