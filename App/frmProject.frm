VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProject 
   Caption         =   "Project Parameters Maintenance"
   ClientHeight    =   7095
   ClientLeft      =   960
   ClientTop       =   1530
   ClientWidth     =   11145
   Icon            =   "frmProject.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11145
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
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   6600
      Width           =   2310
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
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "S"
      Top             =   6600
      Width           =   1170
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   6255
      Left            =   120
      TabIndex        =   54
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmProject.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbProjkey"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbType"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbAltType"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbCity"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbState"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbCapacity"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbFrame"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbExterior"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbBasement"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbStories"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbTotCost"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbTFA"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbGFA"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbVolume"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbACPct"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbACT"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbOwner"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbArchitect"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbGC"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbDate"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbBaySize"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbUnion"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lbQuality"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbShape"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbComments"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbCountry"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txProjkey"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txCity"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "comboType"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "comboAltType"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "comboState"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txCapacity"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "comboFrame"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "comboExterior"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "comboBasement"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txStories"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txContractor"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txTCost"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txTotalFloorArea"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txGroundFloorArea"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txVolume"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txACPct"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txACTons"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txOwner"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txArchitect"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txDate"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txBaySize"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txUnion"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txComments"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "comboQuality"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "comboShape"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txProjskey"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "comboCountry"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).ControlCount=   53
      Begin VB.ComboBox comboCountry 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txProjskey 
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Text            =   "Project Skey"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox comboShape 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   3480
         Width           =   1935
      End
      Begin VB.ComboBox comboQuality 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txComments 
         Height          =   1695
         Left            =   4440
         MultiLine       =   -1  'True
         TabIndex        =   52
         Text            =   "frmProject.frx":045E
         Top             =   4320
         Width           =   6135
      End
      Begin VB.TextBox txUnion 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         Height          =   285
         Left            =   8640
         TabIndex        =   46
         Text            =   "% Union"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txBaySize 
         Height          =   285
         Left            =   8640
         TabIndex        =   44
         Text            =   "Bay Size"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txDate 
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Text            =   "Bid Date"
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox txArchitect 
         Height          =   285
         Left            =   8640
         TabIndex        =   40
         Text            =   "Architect"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txOwner 
         Height          =   285
         Left            =   8640
         TabIndex        =   38
         Text            =   "Owner"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txACTons 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5880
         TabIndex        =   36
         Text            =   "TC Tons"
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txACPct 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         Height          =   285
         Left            =   5880
         TabIndex        =   34
         Text            =   "AC Pct"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txVolume 
         Height          =   285
         Left            =   5880
         TabIndex        =   32
         Text            =   "Volume"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txGroundFloorArea 
         Height          =   285
         Left            =   5880
         TabIndex        =   30
         Text            =   "G.Area"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txTotalFloorArea 
         Height          =   285
         Left            =   5880
         TabIndex        =   28
         Text            =   "T. Area"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txTCost 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   5880
         TabIndex        =   26
         Text            =   "Cost"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txContractor 
         Height          =   285
         Left            =   8640
         TabIndex        =   42
         Text            =   "Contractor"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txStories 
         Height          =   285
         Left            =   5880
         TabIndex        =   24
         Text            =   "Stories"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox comboBasement 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4440
         Width           =   2895
      End
      Begin VB.ComboBox comboExterior 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3960
         Width           =   2895
      End
      Begin VB.ComboBox comboFrame 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox txCapacity 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Text            =   "Capacity"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox comboState 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox comboAltType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Width           =   2895
      End
      Begin VB.ComboBox comboType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txCity 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "City"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txProjkey 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Project ID"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lbCountry 
         Alignment       =   1  'Right Justify
         Caption         =   "Country"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lbComments 
         Caption         =   "Comments"
         Height          =   255
         Left            =   4440
         TabIndex        =   51
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lbShape 
         Caption         =   "Shape"
         Height          =   255
         Left            =   7200
         TabIndex        =   49
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lbQuality 
         Caption         =   "Quality"
         Height          =   255
         Left            =   7200
         TabIndex        =   47
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label lbUnion 
         Caption         =   "% Union"
         Height          =   255
         Left            =   7200
         TabIndex        =   45
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lbBaySize 
         Caption         =   "Bay Size"
         Height          =   255
         Left            =   7200
         TabIndex        =   43
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lbDate 
         Caption         =   "Bid Date"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label lbGC 
         Caption         =   "General Contractor"
         Height          =   255
         Left            =   7200
         TabIndex        =   41
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbArchitect 
         Caption         =   "Architect"
         Height          =   255
         Left            =   7200
         TabIndex        =   39
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbOwner 
         Caption         =   "Owner"
         Height          =   255
         Left            =   7200
         TabIndex        =   37
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbACT 
         Caption         =   "A/C Tons"
         Height          =   255
         Left            =   4440
         TabIndex        =   35
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lbACPct 
         Caption         =   "A/C Pct"
         Height          =   255
         Left            =   4440
         TabIndex        =   33
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lbVolume 
         Caption         =   "Volume"
         Height          =   255
         Left            =   4440
         TabIndex        =   31
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lbGFA 
         Caption         =   "Ground Floor Area"
         Height          =   255
         Left            =   4440
         TabIndex        =   29
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lbTFA 
         Caption         =   "Tot. Floor Area"
         Height          =   255
         Left            =   4440
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbTotCost 
         Caption         =   "Tot. Cost"
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbStories 
         Caption         =   "No. Stories"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbBasement 
         Caption         =   "Basement"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label lbExterior 
         Caption         =   "Exterior"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lbFrame 
         Caption         =   "Frame"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lbCapacity 
         Caption         =   "Capacity"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label lbState 
         Caption         =   "State"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lbCity 
         Caption         =   "City"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lbAltType 
         Caption         =   "Alt Type"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lbType 
         Caption         =   "Type"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbProjkey 
         Caption         =   "Project ID"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton buUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   9720
      TabIndex        =   53
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label lbllast_update_date 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Updated:"
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   6630
      Width           =   705
   End
   Begin VB.Label lbllast_update_person 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Updated By:"
      Height          =   255
      Left            =   3315
      TabIndex        =   57
      Top             =   6630
      Width           =   1035
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <modulename> frmProject</modulename>
' <functionname>General (Main) </functionname>
'
' <summary> PROJECT PARAMETERS MAINTENANCE
'
'For a given "Project" this window/form permits user to uniformly and systematically enter project data from an array of text and combo boxes including:
'
'1.  Project ID
'2.  Type
'3.  City, State, Country
'4.  Frame
'5.  Exterior
'6.  Basement
'7.  Bid Date
'8.  No. Stories
'9.  Tot Cost
'10. Tot. Floor Area
'11. …
'12. Owner
'13. Architect
'14. General Contractor
'15. …
'
'Add or Change PROJECT (DIV 17) INFORMATION.
'
'Search Criteria:
'
'"   Sort By
'ID              blows up!
'Class
'WARNING: doesn 't work!!!
'"   Year Built:         doesn't work?
'"   State               doesn't work?
'"   Project ID:             works!
'"   Min Max
'o   Cost            doesn't work?
'o   Area            doesn't work?
'
'(Major function buttons)
'
'1.  Search                      (buSearch_Click() )
'Designed not to work w/o the above "search criteria"
'2.  Parameters                  (frmProject)
'3.  Project Analysis                (frmProjectAnalysis)
'4.  Update                  (CProjectMap.Update() /
' m_objGridMap.Update())
'5.  New                     (frmProject.frm)
'6.  Delete                      (CProjectMap.Delete() )
'
'NOTE: SetRow()  is where the fields on the form are set.  (from m_rec)
'
'TABLES:
'"   PROJECT_BUILDING_DETAIL
'Contains the majority of the project data
'"   BUILDING_ELEMENT_RATING
'????
'
'HELPER Class: CProjectMap.Cls
'
'</summary>
'
'<seealso> CProjectMap.cls </seealso>
'<seealso>frmProjectGrid.frm</seealso>
'<seealso>frmProjectRpt.frm</seealso>
'
'
' <datastruct>m_objGridMap</datastruct>
'<datastruct>m_rec</datastruct>
'
' <storedprocedurename>sp_update_project_parameters</storedprocedurename>
'
'
'
'<returns>N/A</returns>
' <exception>Always trap with an accompanying message box</exception>
' <example>
'
'<code>
'(In SetRow():  this code populates the form fields!!!)
'
'SELECT P.*, B.quality_code FROM PROJECT_BUILDING_DETAIL P LEFT OUTER JOIN BUILDING_ELEMENT_RATING B ON P.proj_bldg_skey = B.proj_bldg_skey WHERE P.proj_bldg_skey = 113052
'</code>
'
' <code>  this exec updates the project data/information
'
'exec sp_update_project_parameters  @projkey = 113052,  @type = 'UTILITIES, CIVIL ENGINEERING FACILITIES',  @alttype = 'UTILITIES, CIVIL ENGINEERING FACILITIES',  @city = 'Kingstons', @state = 'MA',  @country = 'United States',  @capacity = 0,  @frame = 'Masonry wall bearing, Masonry bearing',  @exterior = 'Metal Panel & Metal Steel',  @basement = 'Not Specified',  @stories = 3,  @tcost = 0,  @tarea = 0,  @garea = 0,  @volume = 0,  @acp = 0,  @act = 0,  @owner = '',  @architect = '',  @contractor = '',  @biddate = '1/1/2008',  @baysize = 0,  @union = 0,  @quality = 'Not Specified',  @shape = 'Not Specified',  @comments = '"hack" to get year 2008 in...'
'</code>
'</example>
'<permission>Public</Permission>
'<dependson>This component depends on the following
'1.  CProjectMap.cls
'2.  CGridMap.cls
'3.  CCDdal.CRSMDataAccess (
'Access to the DAL (data access layer dll) opened in MainModule_Main() )
'</dependson>




Dim m_projkey As String
Dim sEventSubscriberID As String

Dim cnTemp As New ADODB.Connection
Dim m_rec As New ADODB.RecordSet

Private Sub buUpdate_Click()
    Dim bPass As Boolean
    Dim Button As VbMsgBoxResult
    Dim rec As New ADODB.RecordSet
    Dim strSQL As String, strAltType As String, I As Long
    Dim sTemp As String
    
    bPass = True
    On Error GoTo Err_Handler
    
    ' 9/8/2005 RTD - Validate classification in case this is a new project
    '               (Necessary for CR #1518)
    If comboType.Text = "" Then
        MsgBox "Project Type Classification is a required field.", vbCritical + vbOKOnly
        bPass = False
    End If
    
    ' validate the country combo
    If comboState.Text = "" Then
        MsgBox "State is a required field.", vbOKOnly + vbExclamation
        bPass = False
    Else
        strSQL = "SELECT C.country_name FROM COUNTRY C INNER JOIN STATE_COUNTRY S ON C.country_code = S.country_code WHERE S.state_code = '" & comboState.Text & "'"
        If Not g_objDAL.GetRecordset(vbNullString, strSQL, rec) Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching for available country codes.", vbInformation
        Else
            If comboCountry.Text <> rec("country_name") Then
                MsgBox comboCountry.Text & " is not a valid country for the state code " & comboState.Text & ".", vbInformation
                bPass = False
            End If
            rec.Close
        End If
    End If
    
    ' MODIFIED 9/8/2005 RTD - BLANK VALUE CAUSED RTE (RELATED TO CR#1508)
    If Not IsNumeric(txTotalFloorArea.Text) Then
        MsgBox "Invalid Total Area value specified. Please check your data.", vbExclamation
        txTotalFloorArea.SetFocus
        bPass = False
    ElseIf CLng(txTotalFloorArea.Text) < 1000 Then
        Button = MsgBox("The value you entered for Total Area is less than 1000. Do you want to update the data?", vbYesNo + vbQuestion, "Invalid number for Area")
        ' MODIFIED 9/8/2005 RTD - DON'T SET bPass = TRUE IF IT WAS ALREADY FALSE FROM PREVIOUS TESTS!
        If Button = vbNo Then
            bPass = False
        End If
    End If
    ' MODIFIED 9/8/2005 RTD - BLANK VALUE CAUSED RTE (RELATED TO CR#1508)
    If Not IsNumeric(txTCost.Text) Then
        MsgBox "Invalid Total Cost value specified. Please check your data.", vbExclamation
        txTCost.SetFocus
        bPass = False
    ElseIf CLng(txTCost.Text) < 100000 Then
        Button = MsgBox("The value you entered for Total Cost is less than 100,000. Do you want to update the data?", vbYesNo + vbQuestion, "Invalid number for Cost")
        ' MODIFIED 9/8/2005 RTD - DON'T SET bPass = TRUE IF IT WAS ALREADY FALSE FROM PREVIOUS TESTS!
        If Button = vbNo Then
            bPass = False
        End If
    End If
    If Not IsNumeric(txACPct.Text) Then
        MsgBox "Invalid input for AC Pct. Please enter a numeric value.", vbExclamation
        bPass = False
    ElseIf txACPct.Text > 100 Then
        MsgBox "The value entered for AC Pct must not be greater than 100.", vbExclamation
        bPass = False
    End If
    If Not IsNumeric(txUnion.Text) Then
        MsgBox "Invalid input for % Union. Please enter a numeric value.", vbExclamation
        bPass = False
    ElseIf txUnion.Text > 100 Then
        MsgBox "The value entered for % Union must not be greater than 100.", vbExclamation
        bPass = False
    End If
    '9/8/2005 RTD - VERIFY THAT BID DATE IS FILLED IN (CR#1509)
    If txDate.Text = "" Or Not IsDate("1/" & txDate.Text) Then
        MsgBox "The bid date must be entered in the format MM/YYYY.", vbExclamation
        bPass = False
    End If
    '9/9/2005 RTD - VERIFY VOLUME VALUE IS FILLED IN (CR#1508)
    If txVolume.Text = "" Or Not IsNumeric(txVolume.Text) Then
        MsgBox "Invalid input for Volume. Please enter a numeric value.", vbExclamation
        bPass = False
    End If
    If comboAltType.Text <> "" Then
        strAltType = Trim(Right(comboAltType.Text, Len(comboAltType.Text) - InStr(comboAltType.Text, ")") - 1))
    End If
    
    If bPass Then
        Screen.MousePointer = vbHourglass
        Status ("Updating project parameters...")
        
        strSQL = "exec sp_update_project_parameters "
        strSQL = strSQL & " @projkey = " & txProjskey.Text & ", "
        strSQL = strSQL & " @type = '" & Replace(Trim(Right(comboType.Text, Len(comboType.Text) - InStr(comboType.Text, ")") - 1)), "'", "''") & "', "
        strSQL = strSQL & " @alttype = '" & Replace(strAltType, "'", "''") & "', "
        strSQL = strSQL & " @city = '" & Replace(txCity.Text, "'", "''") & "',"
        strSQL = strSQL & " @state = '" & comboState.Text & "', "
        strSQL = strSQL & " @country = '" & comboCountry.Text & "', "
        strSQL = strSQL & " @capacity = " & txCapacity.Text & ", "
        strSQL = strSQL & " @frame = '" & comboFrame.Text & "', "
        strSQL = strSQL & " @exterior = '" & comboExterior.Text & "', "
        strSQL = strSQL & " @basement = '" & comboBasement.Text & "', "
        strSQL = strSQL & " @stories = " & txStories.Text & ", "
        strSQL = strSQL & " @tcost = " & CLng(txTCost.Text) & ", "
        strSQL = strSQL & " @tarea = " & CLng(txTotalFloorArea.Text) & ", "
        strSQL = strSQL & " @garea = " & CLng(txGroundFloorArea.Text) & ", "
        strSQL = strSQL & " @volume = " & CLng(txVolume.Text) & ", "
        strSQL = strSQL & " @acp = " & txACPct.Text & ", "
        strSQL = strSQL & " @act = " & txACTons.Text & ", "
        strSQL = strSQL & " @owner = '" & Replace(txOwner.Text, "'", "''") & "', "
        strSQL = strSQL & " @architect = '" & Replace(txArchitect.Text, "'", "''") & "', "
        strSQL = strSQL & " @contractor = '" & Replace(txContractor.Text, "'", "''") & "', "
        strSQL = strSQL & " @biddate = '1/" & txDate.Text & "', "
        strSQL = strSQL & " @baysize = " & txBaySize.Text & ", "
        strSQL = strSQL & " @union = " & txUnion.Text & ", "
        strSQL = strSQL & " @quality = '" & comboQuality.Text & "', "
        strSQL = strSQL & " @shape = '" & comboShape.Text & "', "
        strSQL = strSQL & " @comments = '" & Replace(txComments.Text, "'", "''") & "'"
        ' 9/9/2005 RTD - STORED PROC NOW OUTPUTS A RECORDSET CONTAINING THE PROJ_BLDG_SKEY,
        '                USED TO RETRIEVE THE NEW SKEY WHEN ADDING A NEW PROJECT;
        '                CHECK FOR AND DISPLAY ANY DATABASE ERRORS
        Set rec = g_cnShared.Execute(strSQL)
        If g_cnShared.Errors.Count = 0 Then
            ' SET THE PROJECT SKEY SO THAT WE CAN REQUERY
            If m_projkey = 0 Then
                m_projkey = rec.Fields("proj_bldg_skey")
            End If
            ' NOTIFY LISTENING FORMS OF THE UPDATE
            EventSubscriberNotify esnProjectRecordUpdated, m_projkey
            MsgBox "Project " & m_projkey & " Update completed.", vbInformation
        Else
            For I = 0 To g_cnShared.Errors.Count - 1
                sTemp = sTemp & vbCrLf & Space(3) & g_cnShared.Errors.Item(0).Description
            Next
            MsgBox "An error occurred while updating project:" & sTemp, vbCritical
        End If
        ' REQUERY THE DATABASE AND UPDATE FORM
        If m_projkey > 0 Then
            SetRow m_projkey
        End If
        Status ("")
        Screen.MousePointer = vbNormal
    End If
    Set rec = Nothing
    Exit Sub
    
'9/9/2005 RTD
'PREVENT RUN-TIME ERRORS FROM CRASHING THE APPLICATION.
Err_Handler:
    Screen.MousePointer = vbNormal
    MsgBox "An error occurred while updating: " & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
    
End Sub

Private Sub comboState_Click()
    Dim rec As New ADODB.RecordSet
    Dim strSelect, I
    strSelect = "SELECT C.country_name FROM COUNTRY C INNER JOIN STATE_COUNTRY S ON C.country_code = S.country_code WHERE S.state_code = '" & comboState.Text & "'"
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available country code."
    Else
        For I = 0 To comboCountry.listcount - 1
            If comboCountry.List(I) = rec("country_name") Then
                comboCountry.ListIndex = I
            End If
        Next
        rec.Close
    End If
    Set rec = Nothing
End Sub

Private Sub Form_Initialize()
    sEventSubscriberID = EventSubscriberAdd(Me)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    ColorLockedFields Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    If Me.WindowState <> vbMinimized Then
'        Me.Height = 7350
'        Me.Width = 11265
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    '   Disables & hides the sort buttons on the main form.
    If m_rec.State <> adStateClosed Then m_rec.Close
    Set m_rec = Nothing
    HideGridSort
    EventSubscriberRemove sEventSubscriberID
End Sub

' 9/8/2005 RTD - ADDED TO SUPPORT ADDING NEW BUILDINGS
' DIRECTLY FROM THIS FORM (CR#1518)
Public Sub NewRow(ByVal Class_ID As Long)
    Dim strSQL As String
    Dim strSelect As String
    Dim rec As New ADODB.RecordSet
    Dim Index As Integer
    Dim projkey As Long
    
    projkey = 0
    Screen.MousePointer = vbHourglass
    Status ("Loading Project Parameters Form ...")
    
    ' close the m_rec if it's already open
    If m_rec.State <> adStateClosed Then m_rec.Close
    m_projkey = projkey
    Me.Caption = "Parameters - New Project"
    Me.txProjskey = m_projkey
    
    'strSQL = "SELECT P.*, B.quality_code FROM PROJECT_BUILDING_DETAIL P LEFT OUTER JOIN BUILDING_ELEMENT_RATING B ON P.proj_bldg_skey = B.proj_bldg_skey WHERE P.proj_bldg_skey = " & m_projkey
    'Set m_rec = New ADODB.RecordSet
    'm_rec.AddNew
    'm_rec("proj_bldg_id") = 0
    'm_rec.Update
    
    Me.txProjkey.Text = m_projkey
    Me.txCity = ""
    Me.txCapacity = 0
    Me.txStories = 1
    Me.txTCost = 0
    Me.txGroundFloorArea = 0
    Me.txTotalFloorArea = 0
    Me.txVolume = 0
    Me.txACPct = 0
    Me.txACTons = 0
    Me.txOwner = ""
    Me.txArchitect = ""
    Me.txContractor = ""
    Me.txDate = Format(Date, "MM/YYYY")
    Me.txBaySize = 0
    Me.txUnion = 0
    Me.txComments = ""
    '
    '   Fill the type combo box
    strSelect = "SELECT DISTINCT class_id, class_desc FROM CLASSIFICATION WHERE class_system_id = 'F' ORDER BY class_id"
    PopulateCombo comboType, strSelect, "class_id", "facility1_class_id", "class_desc"
    If Class_ID = 0 Then
        comboType.ListIndex = -1
    Else
        comboType.ListIndex = FindComboItemDataIndex(comboType, Class_ID)
    End If
    '
    '   Fill the alt type combo box
    strSelect = "SELECT DISTINCT class_id, class_desc FROM CLASSIFICATION WHERE class_system_id = 'F' ORDER BY class_id"
    PopulateCombo comboAltType, strSelect, "class_id", "", "class_desc"
    comboAltType.ListIndex = -1
    '
    '   Fill the state combo box
    strSelect = "SELECT DISTINCT state_code FROM STATE_COUNTRY ORDER BY state_code"
    PopulateCombo comboState, strSelect, "state_code", "state_code", "state_code"
    comboState.ListIndex = -1
    '
    '   Fill the frame combo box
    strSelect = "SELECT DISTINCT frame_mat_code, frame_mat_desc FROM FRAME_MATERIAL ORDER BY frame_mat_code"
    PopulateCombo comboFrame, strSelect, "frame_mat_code", "frame_mat_code", "frame_mat_desc"
    comboFrame.ListIndex = SendMessage(comboFrame.hWnd, CB_FINDSTRING, 0, "Not Specified")
    '
    '   Fill the exterior combo box
    strSelect = "SELECT DISTINCT exterior_mat_code, exterior_material_desc FROM EXTERIOR_MATERIAL ORDER BY exterior_material_desc"
    PopulateCombo comboExterior, strSelect, "exterior_mat_code", "exterior_mat_code", "exterior_material_desc"
    comboExterior.ListIndex = SendMessage(comboExterior.hWnd, CB_FINDSTRING, 0, "Not Specified")
    '
    '   Fill the basement combo box
    strSelect = "SELECT DISTINCT basement_code, basement_desc FROM BASEMENT_CODE ORDER BY basement_desc"
    PopulateCombo comboBasement, strSelect, "basement_code", "basement_code", "basement_desc"
    comboBasement.ListIndex = SendMessage(comboBasement.hWnd, CB_FINDSTRING, 0, "Not Specified")
    '
    '   Fill the quality combo box
    strSelect = "SELECT DISTINCT quality_code, quality_desc FROM QUALITY_CODE ORDER BY quality_desc"
    PopulateCombo comboQuality, strSelect, "quality_code", "quality_code", "quality_desc"
    comboQuality.ListIndex = SendMessage(comboQuality.hWnd, CB_FINDSTRING, 0, "Not Specified")
    '
    '   Fill the shape combo box
    strSelect = "SELECT DISTINCT shape_code, shape_desc FROM SHAPE_CODE ORDER BY shape_desc"
    PopulateCombo comboShape, strSelect, "shape_code", "shape_code", "shape_desc"
    comboShape.ListIndex = SendMessage(comboShape.hWnd, CB_FINDSTRING, 0, "Not Specified")
    '
    '   Fill the country combo box
    strSelect = "SELECT DISTINCT country_code, country_name FROM COUNTRY ORDER BY country_name"
    PopulateCombo comboCountry, strSelect, "country_code", "country_code", "country_name"
    ' 9/8/2005 RTD - MAKE 'USA' THE DEFAULT COUNTRY (CR#1514)
    If comboCountry.Text = "" Then
        ' Get the ListIndex of "United States" and set the combo
        Index = SendMessage(comboCountry.hWnd, CB_FINDSTRING, 0, "United States")
        If Index >= 0 Then
            comboCountry.ListIndex = Index
        End If
    End If
    
    Status ("")
    Screen.MousePointer = vbNormal
    
End Sub

Public Sub SetRow(ByVal projkey As String)
    Screen.MousePointer = vbHourglass
    Status ("Loading Project Parameters Form ...")
    Dim strSQL As String
    Dim strSelect As String
    Dim rec As New ADODB.RecordSet
    Dim Index As Integer
    
    ' close the m_rec if it's already open
    If m_rec.State <> adStateClosed Then m_rec.Close
    m_projkey = projkey
    Me.Caption = "Parameters - Project " & m_projkey
    Me.txProjskey = m_projkey
    
    strSQL = "SELECT P.*, B.quality_code FROM PROJECT_BUILDING_DETAIL P LEFT OUTER JOIN BUILDING_ELEMENT_RATING B ON P.proj_bldg_skey = B.proj_bldg_skey WHERE P.proj_bldg_skey = " & m_projkey
    
    If Not g_objDAL.GetRecordset(vbNullString, strSQL, m_rec) Then
        MsgBox "An error occurred while searching for project " & m_projkey & "."
        Unload Me
    Else
        '9/8/2005 RTD - TEST EOF TO PREVENT ERRORS WITH NEW PROJECTS
        If Not m_rec.EOF Then
            Me.txProjkey.Text = m_rec("proj_bldg_id")
            Me.txCity = m_rec("city")
            Me.txCapacity = m_rec("proj_bldg_functional_uom_qty")
            Me.txStories = m_rec("upper_floor_qty")
            Me.txTCost = FormatNumber(m_rec("proj_bldg_project_tot_cost"), 0)
            Me.txGroundFloorArea = FormatNumber(m_rec("ground_floor_area"), 0)
            Me.txTotalFloorArea = FormatNumber(m_rec("gross_floor_area"), 0)
            Me.txVolume = FormatNumber(m_rec("volume"), 0)
            Me.txACPct = m_rec("air_conditioned_pct")
            Me.txACTons = m_rec("air_conditioned_volume")
            Me.txOwner = m_rec("owner")
            Me.txArchitect = m_rec("architect")
            Me.txContractor = m_rec("general_contractor")
            '9/8/2005 RTD - HANDLE NULL DATE FIELD (CR#1509)
            If Not IsNull(m_rec("bid_date")) Then
                Me.txDate = Month(m_rec("bid_date")) & "/" & Year(m_rec("bid_date"))
            Else
                Me.txDate = ""
            End If
            Me.txBaySize = m_rec("bay_size")
            Me.txUnion = m_rec("union_pct")
            Me.txComments = m_rec("note")
            '9/8/2005 RTD - ADDED ADDITIONAL FIELDS FOR INFORMATIVE PURPOSES
            '               AND CONSISTENCY WITH OTHER CCD FORMS
            Me.txtlast_update_person = m_rec("last_update_person")
            Me.txtlast_update_date = m_rec("last_update_date")
        End If
        '
        '   Fill the type combo box
        strSelect = "SELECT DISTINCT class_id, class_desc FROM CLASSIFICATION WHERE class_system_id = 'F' ORDER BY class_id"
        PopulateCombo comboType, strSelect, "class_id", "facility1_class_id", "class_desc"
        '
        '   Fill the alt type combo box
        strSelect = "SELECT DISTINCT class_id, class_desc FROM CLASSIFICATION WHERE class_system_id = 'F' ORDER BY class_id"
        PopulateCombo comboAltType, strSelect, "class_id", "facility2_class_id", "class_desc"
        '
        '   Fill the state combo box
        strSelect = "SELECT DISTINCT state_code FROM STATE_COUNTRY ORDER BY state_code"
        PopulateCombo comboState, strSelect, "state_code", "state_code", "state_code"
        '
        '   Fill the frame combo box
        strSelect = "SELECT DISTINCT frame_mat_code, frame_mat_desc FROM FRAME_MATERIAL ORDER BY frame_mat_code"
        PopulateCombo comboFrame, strSelect, "frame_mat_code", "frame_mat_code", "frame_mat_desc"
        '
        '   Fill the exterior combo box
        strSelect = "SELECT DISTINCT exterior_mat_code, exterior_material_desc FROM EXTERIOR_MATERIAL ORDER BY exterior_material_desc"
        PopulateCombo comboExterior, strSelect, "exterior_mat_code", "exterior_mat_code", "exterior_material_desc"
        '
        '   Fill the basement combo box
        strSelect = "SELECT DISTINCT basement_code, basement_desc FROM BASEMENT_CODE ORDER BY basement_desc"
        PopulateCombo comboBasement, strSelect, "basement_code", "basement_code", "basement_desc"
        '
        '   Fill the quality combo box
        strSelect = "SELECT DISTINCT quality_code, quality_desc FROM QUALITY_CODE ORDER BY quality_desc"
        PopulateCombo comboQuality, strSelect, "quality_code", "quality_code", "quality_desc"
        '
        '   Fill the shape combo box
        strSelect = "SELECT DISTINCT shape_code, shape_desc FROM SHAPE_CODE ORDER BY shape_desc"
        PopulateCombo comboShape, strSelect, "shape_code", "shape_code", "shape_desc"
        '
        '   Fill the country combo box
        strSelect = "SELECT DISTINCT country_code, country_name FROM COUNTRY ORDER BY country_name"
        PopulateCombo comboCountry, strSelect, "country_code", "country_code", "country_name"
        ' 9/8/2005 RTD - MAKE 'USA' THE DEFAULT COUNTRY (CR#1514)
        If comboCountry.Text = "" Then
            ' Get the ListIndex of "United States" and set the combo
            Index = SendMessage(comboCountry.hWnd, CB_FINDSTRING, 0, "United States")
            If Index >= 0 Then
                comboCountry.ListIndex = Index
            End If
        End If
    End If
    Status ("")
    Screen.MousePointer = vbNormal
End Sub

Private Function PopulateCombo(objcombo, strSelect, strField1, strField2, strComboText)
    Dim Index As Integer
    Dim rec As New ADODB.RecordSet
    
    On Error Resume Next
    Index = 0
    objcombo.Clear
    If Not g_objDAL.GetRecordset(vbNullString, strSelect, rec) Then
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching for available classification code."
    Else
        While Not rec.EOF
            Index = Index + 1
            If strComboText = "class_desc" Then
                objcombo.AddItem "(" & rec.Fields("class_id") & ") " & rec.Fields("class_desc")
                '9/8/2005 RTD
                'USE ITEMDATA PROPERTY TO STORE CLASS ID FOR LATER SEARCHING
                objcombo.ItemData(objcombo.NewIndex) = Trim(rec.Fields("class_id"))
            Else
                objcombo.AddItem rec.Fields(strComboText)
                
            End If
            ' MODIFIED 8/9/2005 RTD
            ' If strField was not supplied, don't check the recordset
            If strField2 <> "" Then
                If rec.Fields(strField1) = m_rec(strField2) Then
                    objcombo.ListIndex = Index - 1
                End If
            End If
            rec.MoveNext
        Wend
        rec.Close
    End If
    Set rec = Nothing
End Function

