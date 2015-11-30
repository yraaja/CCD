VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{5936A75C-3F42-11D6-AF6B-AA0004005F12}#1.3#0"; "MeansCtrl.ocx"
Begin VB.Form frmMatPriceGrid 
   Caption         =   "Material Price Grid"
   ClientHeight    =   6855
   ClientLeft      =   2205
   ClientTop       =   2625
   ClientWidth     =   11565
   Icon            =   "frmMatPriceGrid.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11565
   Begin VB.TextBox StartMatID 
      Height          =   315
      Left            =   8040
      TabIndex        =   2
      Top             =   720
      Width           =   1515
   End
   Begin VB.CommandButton cmdMultiplier 
      Caption         =   "Multiplier"
      Height          =   315
      Left            =   3720
      TabIndex        =   14
      Top             =   2880
      Width           =   1150
   End
   Begin VB.CommandButton cmdPublishMatPrice 
      Caption         =   "Publish"
      Height          =   315
      Left            =   2520
      TabIndex        =   13
      Top             =   2880
      Width           =   1150
   End
   Begin VB.TextBox altmatid 
      Height          =   315
      Left            =   8040
      TabIndex        =   5
      Top             =   1080
      Width           =   1515
   End
   Begin VB.TextBox EndMatID 
      Height          =   315
      Left            =   9720
      TabIndex        =   4
      Top             =   720
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reports"
      Height          =   855
      Left            =   6240
      TabIndex        =   35
      Top             =   6000
      Width           =   1455
      Begin VB.CommandButton cmdRFQ 
         Caption         =   "Price Quote"
         Height          =   495
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Select individual records, or none for all records.  The report will use the selection criteria specified."
         Top             =   240
         Width           =   995
      End
   End
   Begin VB.TextBox techdesc 
      Height          =   315
      Left            =   8040
      TabIndex        =   8
      Top             =   1800
      Width           =   2235
   End
   Begin VB.TextBox txtTotalChildForms 
      Height          =   375
      Left            =   11160
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdFactor 
      Caption         =   "Factor"
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   2880
      Width           =   1150
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "Clone"
      Height          =   495
      Left            =   10800
      TabIndex        =   25
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   9840
      TabIndex        =   24
      Top             =   6240
      Width           =   855
   End
   Begin VB.CheckBox ckbUse 
      Caption         =   "Use = ""Y"" only"
      Height          =   315
      Left            =   9720
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin ConstructionCostDatabase.DynaTree FormatTree 
      Height          =   2775
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go To"
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   6015
      Begin VB.CommandButton cmdMaterial 
         Caption         =   "Material Maint."
         Height          =   495
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdMaterialPrice 
         Caption         =   "Material Price"
         Height          =   495
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdHistory 
         Caption         =   "History"
         Height          =   495
         Left            =   2070
         TabIndex        =   17
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdMaterialUsage 
         Caption         =   "Material Usage"
         Height          =   495
         Left            =   2940
         TabIndex        =   18
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdInfoSources 
         Caption         =   "Info Sources"
         Height          =   495
         Left            =   3930
         TabIndex        =   19
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdMaterialManufacturer 
         Caption         =   "Material Manufac"
         Height          =   495
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   8880
      TabIndex        =   23
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   7920
      TabIndex        =   22
      Top             =   6240
      Width           =   855
   End
   Begin VB.CheckBox ckbRowWrap 
      Caption         =   "Row Wrap"
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   495
      Left            =   9720
      TabIndex        =   10
      Top             =   2220
      Width           =   1150
   End
   Begin VB.TextBox ManufacturerID 
      Height          =   315
      Left            =   8040
      TabIndex        =   9
      Top             =   2160
      Width           =   1515
   End
   Begin VB.TextBox ContactID 
      Height          =   315
      Left            =   8040
      TabIndex        =   6
      Top             =   1440
      Width           =   1515
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid 
      Height          =   2715
      Left            =   0
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3240
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   4789
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   255
      Left            =   8040
      TabIndex        =   1
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Alt Material ID:"
      Height          =   255
      Left            =   6720
      TabIndex        =   36
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Left            =   9720
      TabIndex        =   3
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Technical Desc:"
      Height          =   255
      Left            =   6720
      TabIndex        =   33
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblRowCount 
      Caption         =   "0 rows returned"
      Height          =   255
      Left            =   6720
      TabIndex        =   31
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6780
      TabIndex        =   29
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Manufacturer ID:"
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact ID:"
      Height          =   255
      Left            =   6720
      TabIndex        =   27
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Material ID:"
      Height          =   255
      Left            =   6720
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   11040
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line1 
      X1              =   6660
      X2              =   6660
      Y1              =   2700
      Y2              =   60
   End
End
Attribute VB_Name = "frmMatPriceGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''<modulename> frmMatPriceGrid</modulename>
'''<functionname>General (Main) </functionname>
'''
''' <summary>
''' Provides u/i permitting user to do the following:
'''
'''(Major function buttons)
'''
'''1.  Display “Material Price” form       (frmMatPrice.frm)
'''2.  Display “Material Maint.” form      (frmMaterial.frm)
'''3.  Display “History”               (frmMatHistoryGrid.frm)
'''4.  Display Material Usage          (frmMatUsageGrid.frm)
'''5.  Display “Info Sources”          (frmInfoSourceGrid)
'''6.  Display “Material Manufac”          (frmMatManufacturerGrid.frm)
'''7.  Display “Price Quote” reports       (PreviewReport() )
'''8.  Update (save) Material price change(s)  (CMatPriceMap.Update() )
'''9.  Update / Save any changes to unit cost related data
'''(m_objGridMap.Update())
'''10. Create a NEW material price line        (frmMatPrice.frm)
'''11. Delete a selected material price line       (m_objGridMap.Delete())
'''12. Clone a selected material price line        (frmMatPrice.frm)
'''
'''(Grid buttons)
'''
'''1.  SEARCH – refresh contents of the material price grid datagrid based upon  search criteria
'''2.  FACTOR      (Apply factor % to row(s))      (m_objGridMap.Factor())
'''3.  PUBLISH     (Rows to “published” table(s))  (cmdUpdate_Click() )
'''4.  MULTIPLIER  (Percent Multiplier)            (TDBGrid.Update() )
''' </summary>
''' <seealso>N/A</seealso>
''' <datastruct>m_rec</datastruct>
''' <storedprocedurename> sp_rfq_rpt_options_current_user
'''</storedprocedurename>
''' <storedprocedurename> sp_rpt_mat_rfq</storedprocedurename>
''' <storedprocedurename> sp_select_material_price</storedprocedurename>
'''
''' <param name="data"> ???a dataset containing all the data for updating ?
'''</param>
''' <param name="someParameter">
'''???? Description of someParameter goes here  updating
''' </param>
''' <returns>N/A</returns>
''' <exception>Always trap with an accompanying message box</exception>
''' <example>
''' <code>
'''exec sp_select_material_price @start_mat_id = 'M030500000000', @end_mat_id = 'M030999999999', @alt_mat_id = '', @manufacturer_id = '', @tech_desc = '', @contact_id = '', @use_ind=0 </code>
''' <code>
'''exec sp_rpt_mat_rfq @start_mat_id = 'M030500000000', @end_mat_id = 'M030999999999', @contact_id = '030PHI', @print_contact_id = '030PHI', @manufacturer_id = '%', @tech_desc = '%', @use_ind = 2, @print_price = 1, @report_option_value_id = 2, @filtered = 0, @use_recip_price = 0, @user_name = 'Hancockrl'
'''</code>
'''</example>
'''<permission>Public</Permission>
'''<dependson>This component depends on the following
'''1.  CMatPriceMap.cls
'''2.  CGridMap.cls
'''3.  CCDdal.CRSMDataAccess (
'''Access to the DAL (data access layer dll) opened in MainModule_Main() )
'''4.  rptRequestForQuote.xml  (in ProgramFiles\RSMeans\Construction Cost Database)
'''</dependson>



Dim m_objGridMap As New CMatPriceMap ' Class to handle grid
Public m_blnFirstSearch As Boolean ' Is this the first search we have made on this screen
Dim m_blnJumpIn As Boolean ' Are we jumping here from another screen
Dim m_rec As New ADODB.RecordSet ' Recordset to hold query results
Dim m_blnDoubleClick As Boolean ' Did a double click just occurr
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload

Dim m_strCurrentFormControl As String

Dim strRFQText As String
Dim strPrintContact As String
Dim blnSuppressPrices As Boolean
Dim blnSuppressAddressee As Boolean
Dim blnUseRecipientPrice As Boolean

Private Function ContactDescription(strContact As String) As String
    Dim rsTemp As ADODB.RecordSet
    Dim strSQL As String
    'Fill the contact description
    
    strSQL = "select company_name from information_source where contact_id = '" + strContact + "'"
    g_objDAL.GetRecordset CONNECT, strSQL, rsTemp
    If Not (rsTemp.EOF And rsTemp.BOF) Then
        ContactDescription = strContact + "     " + rsTemp.Fields("company_name")
    End If
    rsTemp.Close
End Function


Private Function GetPrintContact(ListSelection As cdlgLstSel) As String
    Dim sql As String
    Dim rec As ADODB.RecordSet
    Dim varCurSelectedRow  As Variant
    Dim blnResult As Boolean
    
    'A list of contacts for the material
    'parameters selected will be constructed, and the list selections populated from it.
    'If the grid has selected rows, contacts will be retrieved from it.  If not, use the contacts from the selection critiria (all grid records)
    'If there is only one contact in the list or found in the recordset, it will be used and the display bypassed.
    'Per Mel:  always display selection, 2/9/01
    
    ListSelection.Caption = "Recipient Contact Selection"
    ListSelection.ComboCaption = "Select Recipient:"
    If TDBGrid.SelBookmarks.Count > 0 Then
        For Each varCurSelectedRow In TDBGrid.SelBookmarks
            TDBGrid.Bookmark = varCurSelectedRow
            ListSelection.AddUniqueItem ContactDescription(TDBGrid.Columns("Contact")), 1, 1
        Next varCurSelectedRow
    Else        'No records selected - need to validate/retrieve the contact_id
        If Left(UCase(StartMatID.Text), 1) <> "M" Then
            If Len(StartMatID.Text) = 0 Then
                StartMatID.Text = "%"
            End If
            StartMatID.Text = "M" + StartMatID.Text
        End If
        sql = "select distinct contact_id from material inner join material_price on material.mat_skey = material_price.mat_skey" + _
            " Where contact_id Like '" + IIf(Len(ContactID.Text) = 0, "%", SQLChangeWildcard(ContactID.Text)) + "' and     manufacturer_id like '" + IIf(Len(ManufacturerID.Text) = 0, "%", SQLChangeWildcard(ManufacturerID.Text)) + "'" + _
                " and tech_desc like '" + IIf(Len(techdesc.Text) = 0, "%", SQLChangeWildcard(techdesc.Text)) + _
                "' and  '" + CStr(Format(Now(), "short date")) + "' between start_date and term_date "
        If ckbUse.value = 1 Then
            sql = sql + "  and (use_ind = 1)"
        End If
        If Right(StartMatID, 1) = "*" Or Right(StartMatID, 1) = "%" Or Len(StartMatID.Text) = 0 Or Len(EndMatID) = 0 Then
            sql = sql + " and mat_id like '" + IIf(Len(StartMatID.Text) = 0, "%", SQLChangeWildcard(Compress_String(StartMatID.Text))) + "'"
        Else
            sql = sql + " and mat_id between '" + StartMatID.Text + "' and '" + EndMatID + "'"
        End If
    
        g_objDAL.GetRecordset CONNECT, sql, rec
        If rec.EOF And rec.BOF Then
            MsgBox "Please select a valid contact."
            GoTo Exit_Sub
        Else
            If rec.RecordCount = 0 Then     'invalid
                MsgBox "No contacts found."
            Else
                Do Until rec.EOF
                    ListSelection.AddUniqueItem ContactDescription(rec.Fields("contact_id")), 1, 1
                    rec.MoveNext
                Loop
            End If
            rec.Close
        End If
    End If
    
    If ListSelection.itemCount > 0 Then
        If ListSelection.SetList = True Then
            Screen.MousePointer = vbNormal
            blnResult = ListSelection.ShowList
            Screen.MousePointer = vbHourglass
        End If
    End If
    
    If blnResult = True And ListSelection.itemCount > 0 Then 'Contact selected or only 1 found - if none, ignore
        GetPrintContact = Mid(ListSelection.SingleValue, 1, 6)
    Else
        GetPrintContact = "-1"
    End If

Exit_Sub:

End Function

Private Function GetRFQTextID(ListRFQTextID As cdlgLstSel) As Long
    Dim sql As String
    Dim rec As ADODB.RecordSet
    Dim varCurSelectedRow  As Variant
    Dim blnResult As Boolean
    
    'A list of available body texts for the user will
    ' be constructed, and the list selections populated from it.
    
    ListRFQTextID.Caption = "Letter Text Selection"
    ListRFQTextID.ComboCaption = "Select Text:"
    ListRFQTextID.Check1Caption = "Suppress Prices"
    ListRFQTextID.Check2Caption = "Suppress Addressee"
    
        sql = "exec sp_rfq_rpt_options_current_user"
        g_objDAL.GetRecordset CONNECT, sql, rec
        If rec.EOF And rec.BOF Then
            MsgBox "No texts have been set up for this report.  Please contact the IS department for help."
            GoTo Exit_Sub
        Else
            If rec.RecordCount = 0 Then     'invalid
                MsgBox "No contacts found."
            Else
                Do Until rec.EOF
                    ListRFQTextID.AddUniqueItem rec.Fields("value_description"), 0, rec.Fields("report_option_value_id")
                    rec.MoveNext
                Loop
            End If
            rec.Close
        End If
    
    If ListRFQTextID.itemCount > 0 Then
        If ListRFQTextID.SetList = True Then
            Screen.MousePointer = vbNormal
            blnResult = ListRFQTextID.ShowList()
            Screen.MousePointer = vbHourglass
        End If
    End If
    
    If blnResult = True And ListRFQTextID.itemCount > 0 Then  'Contact selected or only 1 found - if none, ignore
        GetRFQTextID = ListRFQTextID.SingleItemData
    Else
        GetRFQTextID = -1
    End If

Exit_Sub:

End Function

Private Sub GetSettings()
    Screen.MousePointer = vbHourglass
    
    Dim ListSelection As New cdlgLstSel
    Dim ListRFQTextID As New cdlgLstSel
    
    strRFQText = GetRFQTextID(ListRFQTextID)
    
    blnSuppressPrices = ListRFQTextID.Check1Value
    blnSuppressAddressee = ListRFQTextID.Check2Value
    If (blnSuppressAddressee) Or (strRFQText = "-1") Then
        strPrintContact = ""
        blnUseRecipientPrice = False
    Else
        If blnSuppressPrices Or TDBGrid.SelBookmarks.Count > 0 Then
            ListSelection.Check1Caption = ""
        Else
            ListSelection.Check1Caption = "Use Recip. Price vs. Avg"
        End If
        ListSelection.Check2Caption = ""
        strPrintContact = GetPrintContact(ListSelection)
        If Not blnSuppressPrices Then
            blnUseRecipientPrice = ListSelection.Check1Value
        End If
    End If
    
    Set ListSelection = Nothing
    Set ListRFQTextID = Nothing

Exit_Sub:
    Screen.MousePointer = vbNormal
    Set ListSelection = Nothing
End Sub

Public Sub PrintReport()
    PreviewReport
End Sub

Public Sub PreviewReport()
    Dim rtn As Integer
    Dim strStartMatID As String
    Dim strEndMatID As String
    Dim strContactId As String
    Dim strManufacturerID As String
    Dim strTechDesc As String
    Dim iUseInd As Integer
    Dim iFiltered As Integer
    Dim rsRFQ As ADODB.RecordSet
    Dim iSuppressAddressee As Integer
    Dim strSelect As String
    Dim iUseRecipientPrice As Integer
    Dim iPrintPrice  As Integer
    Dim frmRFQ As New frmRFQRpt
    Dim blnReturn As Boolean

    On Error Resume Next
    Screen.MousePointer = vbNormal  '<-- Reset in case some problem leaves as hourglass.
    GetSettings
    
'Pass parms and recordset to the VS View report form
'Recordset will be filtered if the grid has selected records
'   Report Parameters:
'   0   Starting Material ID
'   1   Ending Material ID
'   2   Selected Contact for Prices
'   3   Print Contact
'   4   Manufacturer ID
'   5   Tech Description
'   6   Use Indicator
'   7   Print Prices - Checkbox is to Suppress Prices
'   8   Suppress Addressee
'   9   RFQ Text ID
'   10  Filter Records indicator
'   11
'   12  Use Print Contact Prices

    strStartMatID = StartMatID.Text
    If Len(EndMatID.Text) = 0 And Right(strStartMatID, 1) <> "%" Then
        strStartMatID = strStartMatID + "%"
    End If
    If Len(strStartMatID) = 0 Then
        strStartMatID = "%"
    Else
        If UCase(Left(strStartMatID, 1)) <> "M" Then
            strStartMatID = "M" + SQLChangeWildcard(strStartMatID)
        Else
            strStartMatID = SQLChangeWildcard(strStartMatID)
        End If
    End If
    If Len(EndMatID.Text) = 0 Then
        strEndMatID = "%"
    Else
        If UCase(Left(EndMatID.Text, 1)) <> "M" Then
            strEndMatID = "M" + SQLChangeWildcard(EndMatID.Text)
        Else
            strEndMatID = SQLChangeWildcard(EndMatID.Text)
        End If
    End If
    If Len(ContactID.Text) = 0 Then 'Restrict contact to selected contact
        strContactId = "%"   'by default ?  'rlh
        If Len(strPrintContact) > 0 Then
            strContactId = strPrintContact  'rlh 02/06/2009 (Dave Drain issue...)
        End If
    Else
        strContactId = SQLChangeWildcard(ContactID.Text)
    End If
    If Len(ManufacturerID.Text) = 0 Then
        strManufacturerID = "%"
    Else
        strManufacturerID = SQLChangeWildcard(ManufacturerID.Text)
    End If
    If Len(techdesc.Text) = 0 Then
        strTechDesc = "%"
    Else
        strTechDesc = SQLChangeWildcard(techdesc.Text)
    End If
    If ckbUse.value = 1 Then
        iUseInd = 1
    Else
        iUseInd = 2
    End If
    If TDBGrid.SelBookmarks.Count > 0 Then
        iFiltered = 1
    Else
        iFiltered = 0
    End If
    If strRFQText = "-1" Then        'Cancel
        GoTo Exit_Sub
    End If
    If strPrintContact = "-1" Then
        GoTo Exit_Sub
    End If
    If blnSuppressPrices = True Then 'Suppress prices
        iPrintPrice = 0
    Else
        iPrintPrice = 1
    End If
    If blnSuppressAddressee Then                  'Suppress addressee
        iSuppressAddressee = 1
    Else
        iSuppressAddressee = 0
    End If
    If blnUseRecipientPrice = True Then 'Use recipient prices
        iUseRecipientPrice = 1
    Else
        iUseRecipientPrice = 0
    End If

'Stored Proc parameters:
'    @start_mat_id       varchar(14),
'    @end_mat_id     varchar(14),
'    @contact_id             varchar(7),
'    @print_contact_id       varchar(7),
'    @manufacturer_id        varchar(7),
'    @tech_desc              varchar(76),
'    @use_ind        int,
'    @print_price        int,
'    @report_option_value_id int,
'    @filtered       int,
'    @use_recip_price    int

    strSelect = "exec sp_rpt_mat_rfq "
    If DEBUGON Then Stop
    'strSelect = "exec sp_rpt_mat_rfq_rlh "      'rlh Temporary 04/20/2010
    
    strSelect = strSelect + "@start_mat_id = '" + Compress_String(strStartMatID) + "'"
    strSelect = strSelect + ", @end_mat_id = '" + Compress_String(strEndMatID) + "'"
    strSelect = strSelect + ", @contact_id = '" + strContactId + "'"
    strSelect = strSelect + ", @print_contact_id = '" + strPrintContact + "'"
    strSelect = strSelect + ", @manufacturer_id = '" + strManufacturerID + "'"
    strSelect = strSelect + ", @tech_desc = '" + strTechDesc + "'"
    strSelect = strSelect + ", @use_ind = " + CStr(iUseInd)
    strSelect = strSelect + ", @print_price = " + CStr(iPrintPrice)
    strSelect = strSelect + ", @report_option_value_id = " + strRFQText 'rfq text id
    strSelect = strSelect + ", @filtered = " + CStr(iFiltered)
    strSelect = strSelect + ", @use_recip_price = " + CStr(iUseRecipientPrice)
    strSelect = strSelect + ", @user_name = '" + CStr(strUserName) & "'"     'RLH 01/22/2009 missing signature and incorrect FAX number issue
    
    rsRFQ.Close ' Make sure it is closed
    If TDBGrid.SelBookmarks.Count > 0 Then 'set filter for recordset
        Dim blnfound As Boolean
        Dim varCurSelectedRow As Variant
        Dim strSavevarCurSelectedRow As String
        Dim sFilterClause As String
        Dim sKey As String
        strSavevarCurSelectedRow = TDBGrid.Bookmark
        sFilterClause = ""
        For Each varCurSelectedRow In TDBGrid.SelBookmarks
            TDBGrid.Bookmark = varCurSelectedRow
            sKey = TDBGrid.Columns("mat_skey") & TDBGrid.Columns("Manufacturer") & TDBGrid.Columns("Contact")
            If InStr(1, sKey, sFilterClause) = "" Then 'Not found, add it.
                sFilterClause = sFilterClause & "filter_key = " & "'" & sKey & "' or "
            End If
        Next varCurSelectedRow
        sFilterClause = Left(sFilterClause, Len(sFilterClause) - 3) 'remove last or
    End If

    'TDBGrid.Bookmark = strSavevarCurSelectedRow
    TDBGrid.MoveFirst
    
'    ':::::::::::::::  temporary for debug only ::::::::::::::::::
'    'rlh 01/22/2009
'    Dim tmpSql As String
'    tmpSql = "SELECT SYSTEM_USER"
'    blnReturn = g_objDAL.GetRecordset(vbNullString, tmpSql, rsRFQ)
'    MsgBox ("returned value: " & rsRFQ(0))
'    MsgBox ("user name: " & strUserName)
'    '::::::::::::::   end of temporary debug code :::::::::::::::
    
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rsRFQ)
    If blnReturn = False Then
        MsgBox "An error occurred while retrieving report data."
        GoTo Exit_Sub
    End If
    rsRFQ.Filter = sFilterClause
    With frmRFQ
        .LoadReport rsRFQ, iSuppressAddressee, iPrintPrice, strPrintContact
        .RenderReport
        .Show
    End With

Exit_Sub: Exit Sub
    Screen.MousePointer = vbNormal

End Sub
'@#Public Sub PrintReport()
Public Sub SelectAllRows()
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    
    If m_rec.RecordCount > 0 Then
        m_objGridMap.SelectAllRows
    End If
End Sub

' Handles Row Wrap feature
Private Sub ckbRowWrap_Click()
    m_objGridMap.RowWrap (ckbRowWrap)
End Sub


Private Sub cmdClone_Click()
    On Error GoTo Out
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    Dim rec As ADODB.RecordSet
    TDBGrid.Update

    Set rec = m_objGridMap.CloneRow
    rec.Fields("traces_list_price").value = rec.Fields("list_price").value

    ' Force any changes into recordset from grid
    TDBGrid.Update
    ' Navigate to single-record view
    Dim frm As frmMatPrice
    Set frm = New frmMatPrice
    frm.SetRow rec, True ' Pass the current record into the form
    Set frm.frmCallingForm = Me
    Set frm.tdbCols = Me.TDBGrid.Columns
    Set frm.myTDBGrid = Me.TDBGrid
    txtTotalChildForms.Text = Val(txtTotalChildForms.Text) + 1
    frm.Show
Out:
End Sub

Private Sub cmdDelete_Click()
    m_objGridMap.Delete
End Sub

Private Sub cmdFactor_Click()
    Dim dblFactor As Double
    Dim strComment As String
    Dim intColumns As Integer
    TDBGrid.Update

    If TDBGrid.SelBookmarks.Count > 0 Then
        dblFactor = -1
        intColumns = 0
        dlgFactor.GetFactor dblFactor, strComment, intColumns
        m_objGridMap.Factor dblFactor, strComment, intColumns
    Else
        MsgBox "You must select a row first"
    End If
End Sub

Private Sub cmdHistory_Click()
    TDBGrid.Update
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' Open single record view with data from row selected
    Dim frm As frmMatHistoryGrid
    Set frm = New frmMatHistoryGrid
    frm.JumpIn Compress_String(TDBGrid.Columns("Material ID").CellText(TDBGrid.Bookmark))
End Sub

Private Sub cmdInfoSources_Click()
    Dim sKey As String
    Dim sFilter As String
    Dim varCurSelectedRow As Variant
    
    TDBGrid.Update
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select at least one row.", vbInformation
        Exit Sub
    End If
    
    ' Open grid view with data from rows selected
    Dim frm As frmInfoSourceGrid
    Set frm = New frmInfoSourceGrid
    
    If TDBGrid.SelBookmarks.Count <= 1 Then
        frm.JumpIn TDBGrid.Columns("Contact").CellText(TDBGrid.Bookmark)
    Else
        For Each varCurSelectedRow In TDBGrid.SelBookmarks
            TDBGrid.Bookmark = varCurSelectedRow
            sKey = TDBGrid.Columns("contact")
            If InStr(sFilter, sKey) = 0 Then
                sFilter = sFilter & "," & sKey
            End If
        Next varCurSelectedRow
        sFilter = Mid(sFilter, 2)
        frm.JumpIn sFilter
    End If
    
End Sub

Private Sub cmdMaterial_Click()
    TDBGrid.Update
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' Navigate to single-record view
    Dim frm As frmMaterial
    Dim rec As ADODB.RecordSet
    Set frm = New frmMaterial
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    rec.Fields("last_update_id") = rec.Fields("mat_last_update_id") ' Need to switch ID's - mat uses last_update_id, not mat_last_update_id
    frm.SetRow rec ' Pass the current record into the form
    frm.Show
End Sub

Private Sub cmdMaterialPrice_Click()
    TDBGrid.Update
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    If Val(txtTotalChildForms.Text) > 0 Then
        MsgBox "You have one Single Record Form Open. " + vbCrLf + _
            "Please Close it before opening another.", vbInformation
            TDBGrid.Refresh
        Exit Sub
    End If
    ' Navigate to single-record view
    Dim frm As frmMatPrice
    Dim rec As ADODB.RecordSet
    Set frm = New frmMatPrice
    ' Make copy of recordset
    Set rec = m_rec.Clone
    ' Get the selected row from grid
    rec.Bookmark = TDBGrid.Bookmark
    frm.SetRow rec ' Pass the current record into the form
    Set frm.frmCallingForm = Me
    Set frm.tdbCols = Me.TDBGrid.Columns
    Set frm.myTDBGrid = Me.TDBGrid
    txtTotalChildForms.Text = Val(txtTotalChildForms.Text) + 1
    frm.Show
End Sub

Private Sub cmdMaterialUsage_Click()
    Dim bln_Continue As Boolean
    Dim varCurrentM_recBookmark As Variant
    Dim MaterialID As String

    TDBGrid.Update

    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
    Exit Sub
        
    End If
        
    ' Open spreadsheet view with data from row selected
    Dim frm As frmMatUsageGrid
    Set frm = New frmMatUsageGrid
    frm.strSource = "Material"
    frm.JumpIn Compress_String(TDBGrid.Columns("Material ID").CellText(TDBGrid.Bookmark))
End Sub

Private Sub cmdMultiplier_Click()
    Dim varMultiplier As Variant
    TDBGrid.Update
    If TDBGrid.SelBookmarks.Count = 0 Then
        MsgBox "Please select rows prior to setting the Percent Multiplier."
    Else
        varMultiplier = InputBox("Enter new multipllier:", "Change Percent Multiplier", 1, (Me.Width / 2 + Me.Left) - 2800, Me.Height / 4 + Me.Top)
        If IsNumeric(varMultiplier) Then
            m_objGridMap.SetMultiplier CDbl(varMultiplier)
        Else
            If varMultiplier <> "" Then MsgBox "Please enter a valid multiplier"
        End If
    End If
End Sub

Private Sub cmdNew_Click()
    On Error GoTo Out
    Dim rec As New ADODB.RecordSet
    TDBGrid.Update

    CopyRSFields rec, m_rec
    ' Open empty single record view
    Dim frm As frmMatPrice
    Set frm = New frmMatPrice
    ' Force any changes into recordset from grid
    TDBGrid.Update
    frm.SetRow rec, True
    Set frm.frmCallingForm = Me
    Set frm.tdbCols = Me.TDBGrid.Columns
    Set frm.myTDBGrid = Me.TDBGrid
    txtTotalChildForms.Text = Val(txtTotalChildForms.Text) + 1
    frm.Show
Out:
End Sub

Public Sub Sort(intDir As Integer)
    m_objGridMap.Sort intDir
End Sub


Private Sub cmdPublishMatPrice_Click()
    Dim strUpdate As String
    TDBGrid.Update

    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdUpdate_Click
            ' If there were errors, cancel the search
            If m_blnWereErrors Then
                Exit Sub
            End If
        ElseIf Button = vbCancel Then
            ' Cancel the search
        GoTo Exit_Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    m_objGridMap.Publish

Exit_Sub:
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdRFQ_Click()
    TDBGrid.Update
    If IsNumeric(TDBGrid.Bookmark) Then
        PreviewReport
    Else
        MsgBox "You must search for the records you'd like to print.", vbOKOnly + vbInformation
    End If
End Sub
Public Sub cmdUpdate_Click()
    On Error GoTo Out
    Dim blnRet As Boolean
    Dim vntBookmark As Variant

    vntBookmark = TDBGrid.Bookmark
    TDBGrid.Update
    blnRet = m_objGridMap.Update
    If blnRet = False Then
        m_blnWereErrors = True  'Update failed
    Else
        m_blnWereErrors = False
    End If
    TDBGrid.Bookmark = vntBookmark
Out:
End Sub

Private Sub cmdMaterialManufacturer_Click()
    TDBGrid.Update
    If IsNumeric(TDBGrid.Bookmark) = False Then
        MsgBox "You must select a row."
        Exit Sub
    End If
    ' Open single record view with data from row selected
    Dim frm As frmMatManufacturerGrid
    Set frm = New frmMatManufacturerGrid
    frm.JumpIn TDBGrid.Columns("Manufacturer").CellText(TDBGrid.Bookmark)
End Sub

Private Sub ContactID_LostFocus()
    ContactID = Trim(ContactID)
End Sub


Private Sub EndMatID_Change()
    If EndMatID.SelStart = 1 And Len(EndMatID.Text) > 0 Then
        EndMatID.Text = UCase(Left(EndMatID.Text, 1)) + Right(EndMatID.Text, Len(EndMatID.Text) - 1)
        EndMatID.SelStart = 1
    End If
End Sub

Private Sub EndMatID_LostFocus()
    EndMatID = Trim(EndMatID)
End Sub

Private Sub Form_Deactivate()
    HideGridSort
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = False
    fMainForm.mnuFilePageSetup.Enabled = False
    fMainForm.mnuFilePrint.Enabled = False
    fMainForm.mnuFilePrintPreview.Enabled = False
    m_strCurrentFormControl = Me.ActiveControl.Name
End Sub

Private Sub Form_Initialize()
    Status ("Loading Material Price...")
    Screen.MousePointer = vbHourglass
    m_blnFirstSearch = True
    ' Fill the MasterFormat tree
    'Line of code was changed by Mohan on Jan 05,2012, added "MATERIAL04" to make sure it uses MASTERFORMAT04
    FormatTree.InitData g_cnShared, "MATERIAL04"
    ' Initialize grid
    m_objGridMap.SetGrid TDBGrid
    m_objGridMap.InitGrid
    m_blnJumpIn = False
    Screen.MousePointer = vbNormal
    m_blnFirstSearch = False
End Sub
Private Sub Form_Load()
    Dim strSelect As String
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    StartMatID.Text = "~"
    altmatid.Text = "~"
    cmdSearch_Click
    StartMatID.Text = ""
    altmatid.Text = ""
    Status ("")
End Sub

' Called when coming here from another screen
Public Sub JumpIn(strMatID As String)
    StartMatID.Text = Compress_String(strMatID)
    cmdSearch_Click
End Sub

Private Sub Form_LostFocus()
    TDBGrid.Update
    HideGridSort
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = False
    fMainForm.mnuFilePageSetup.Enabled = False
    fMainForm.mnuFilePrint.Enabled = False
    fMainForm.mnuFilePrintPreview.Enabled = False
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    On Error Resume Next
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width >= 11250 Then
            TDBGrid.Width = Me.Width - 255
            Line2.X2 = Me.Width - 210
        Else
            Me.Width = 11250
        End If
        
        If Me.Height >= 7260 Then
            TDBGrid.Height = Me.Height - 4545
            Frame1.Top = Me.Height - 1260
            Frame2.Top = Me.Height - 1260
            cmdUpdate.Top = Me.Height - 1020
            cmdNew.Top = Me.Height - 1020
            cmdClone.Top = Me.Height - 1020
            cmdDelete.Top = Me.Height - 1020
        Else
            Me.Height = 7260
        End If
    Else
        ShowMinimizedForms
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HideGridSort
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = False
    fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = False
    fMainForm.mnuFilePageSetup.Enabled = False
    fMainForm.mnuFilePrint.Enabled = False
    fMainForm.mnuFilePrintPreview.Enabled = False
End Sub

' Leaf in MasterFormat tree selected.
Private Sub FormatTree_NodeSelected(ByVal strID As String)
    Dim rs As New ADODB.RecordSet
    Dim strSelect As String
    Dim blnReturn As Boolean
    
    On Error Resume Next
    '    If m_blnFirstSearch = True Then
    '        m_blnFirstSearch = False
    '    Else
        If Len(strID) = 13 Then
            StartMatID.Text = strID + "*"
            EndMatID.Text = ""
        Else
            rs.Close ' Make sure it is closed
            'Line of code was changed by Mohan on Jan 05,2012, MASTERFORMAT95_ID_HIERARCHY was changed to MASTERFORMAT04_ID_HIERARCHY
            strSelect = "select mat_id_start, mat_id_end from MASTERFORMAT04_ID_HIERARCHY where hier_id='" + strID + "'"
            blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rs)
            StartMatID.Text = rs.Fields("mat_id_start")
            EndMatID.Text = rs.Fields("mat_id_end")
            ' Clear other boxes
            rs.Close
        End If
        ' Synch text box with tree
        ' Clear other boxes
        ContactID.Text = ""
        ManufacturerID.Text = ""
        techdesc.Text = ""
        ' Kick-off search
        cmdSearch_Click
'    End If
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim dtmToday As Date
    Dim dtmStart As Date
    Dim strStartMatSrch As String
   
    
    'do this comparision first
 '   If MsgBox("you wish to do comparision ?", vbYesNo) = vbYes Then DoComparision
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
        Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
        If Button = vbYes Then
            cmdUpdate_Click
            ' If there were errors, cancel the search
            If m_blnWereErrors Then
                Exit Sub
            End If
        ElseIf Button = vbCancel Then
            ' Cancel the search
            TDBGrid.ReBind
            GoTo Exit_Sub
        Else
            TDBGrid.DataChanged = False
        End If
    End If
    Screen.MousePointer = vbHourglass
    dtmToday = Date
    
    ' Synch tree with text box
    '    If Not StartMatID.Text = "" Then
    '        FormatTree.FocusItem (MaterialID.Text)
    '    End If

    StartMatID = Compress_String(StartMatID)   'rlh CCD 8.4 embedded blanks failure (Chris Babbitt)
    
    If Len(StartMatID.Text) = 0 And Len(EndMatID.Text) = 0 And Len(altmatid.Text) = 0 And Len(ManufacturerID.Text) = 0 And Len(ContactID.Text) = 0 And Len(techdesc.Text) = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "You must enter search criteria before searching."
        GoTo Exit_Sub
    End If
    
    If Not Left(StartMatID.Text, 1) = "M" Then
        If Len(Trim(StartMatID)) = 0 Then
            StartMatID = "M*"
        Else
            StartMatID = "M" + StartMatID
        End If
        
    If Not Left(altmatid.Text, 1) = "M" Then
        If Len(Trim(altmatid)) = 0 Then
            altmatid = "M*"
        Else
            altmatid = "M" + altmatid
        End If
        End If
    End If
    If Len(StartMatID) = 13 And InStr(1, StartMatID, "*") = 0 And Len(EndMatID) = 0 Then
        strStartMatSrch = Compress_String(StartMatID) + "*"
    Else
        strStartMatSrch = Compress_String(StartMatID)
    End If

    'finished
    lblRowCount.Caption = "Working..."
    lblRowCount.Refresh

    strSelect = "exec sp_select_material_price "
    strSelect = strSelect + "@start_mat_id = '" + SQLChangeWildcard(strStartMatSrch) + "'"
    strSelect = strSelect + ", @end_mat_id = '" + Compress_String(SQLChangeWildcard(EndMatID)) + "'"
    strSelect = strSelect + ", @alt_mat_id = '" + Compress_String(SQLChangeWildcard(altmatid)) + "'"
    strSelect = strSelect + ", @manufacturer_id = '" + SQLChangeWildcard(ManufacturerID.Text) + "'"
    strSelect = strSelect + ", @tech_desc = '" + SQLFixString(SQLChangeWildcard(techdesc.Text)) + "'"
    strSelect = strSelect + ", @contact_id = '" + SQLChangeWildcard(ContactID.Text) + "'"
    If ckbUse.value = 1 Then
        strSelect = strSelect + ", @use_ind=1"
    Else
        strSelect = strSelect + ", @use_ind=0"
    End If
    
    m_rec.Close ' Make sure it is closed
    m_rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
    dtmStart = Now
    
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, m_rec)
    If blnReturn = False Then
        MsgBox "An error occurred while searching."
       
        lblRowCount.Caption = "0 rows returned."
        GoTo Exit_Sub
    End If
    
    ' Pass recordset to handler class
    m_objGridMap.RecordSet = m_rec
    If m_rec.RecordCount > 0 Then
        lblRowCount.Caption = str(m_rec.RecordCount) + " rows returned in " + str(DateDiff("s", dtmStart, Now)) + " seconds"
    Else
        lblRowCount.Caption = "0 rows returned."
    End If
    
    ' If the upper bound was hit, inform user
    If m_rec.RecordCount = MAX_RECORDS And m_rec.State = adStateOpen Then
        MsgBox "The search returned the maximum number of records allowed. More records may be available."
    End If
    m_objGridMap.SetMenuBar
    ' Reset the grid contents
    TDBGrid.Bookmark = Null
    TDBGrid.ReBind
    TDBGrid.ApproxCount = m_rec.RecordCount
  
Exit_Sub:
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ' before proceeding to check if changes were done,
    ' check if any child forms are being opened.
    ' if so, do not allow him to go here
    If Val(txtTotalChildForms.Text) > 0 Then
        MsgBox "Please close the Single Record Forms that you have initiated," + vbCrLf + "before closing this form.", vbInformation, "Cannot Close Now"
        Cancel = True
        Exit Sub
    End If
    
    ' Check if there are pending changes
    If m_objGridMap.IsPendingChange = True Then
        Dim Button
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
End Sub

Private Sub ManufacturerID_LostFocus()
    ManufacturerID = Trim(ManufacturerID)
End Sub

Private Sub StartMatID_Change()
    Dim blnReturn As Boolean
    If InStr(1, StartMatID.Text, "*") > 0 Then
        blnReturn = LockField(Me, "EndMatID")
    Else
        blnReturn = UnLockField(Me, "EndMatID")
    End If
    If StartMatID.SelStart = 1 And Len(StartMatID.Text) > 0 Then
        StartMatID.Text = UCase(Left(StartMatID.Text, 1)) + Right(StartMatID.Text, Len(StartMatID.Text) - 1)
        StartMatID.SelStart = 1
    End If
End Sub

Private Sub StartMatID_LostFocus()
    StartMatID = Trim(StartMatID)
End Sub


Private Sub TDBGrid_Change()
    With TDBGrid
        If .Columns(.Col).Caption = "Tech Desc" Or .Columns(.Col).Caption = "Metric Tech Desc" Then
            If Len(.Text) > 75 Then .Text = Left(.Text, 75)
        End If
    End With
End Sub

Private Sub TDBGrid_DblClick()
    ' Signal that double-click has occurred
    m_blnDoubleClick = True
End Sub

Private Sub TDBGrid_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
    TDBGrid.DataChanged = False
End Sub

Private Sub TDBGrid_GotFocus()
    TDBGrid.TabStop = True
End Sub

Private Sub TDBGrid_KeyPress(KeyAscii As Integer)
   Dim i As Integer
   With TDBGrid
    If .Columns(.Col).Caption = "Tech Desc" Or .Columns(.Col).Caption = "Metric Tech Desc" Then
        If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then
            ' do nothing. allow the key
        Else
            If .SelLength <= 0 And Len(.Text) + 1 > 75 Then KeyAscii = 0
        End If
    End If
    End With
End Sub

Private Sub TDBGrid_LostFocus()
    TDBGrid.TabStop = False
End Sub

Private Sub TDBGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If this is the mouse-up form a double click
    If m_blnDoubleClick Then
        ' Make sure it is the left button
        If Button = vbLeftButton Then
            m_blnDoubleClick = False
            ' Same function as clicking Material Price button, open single record view
            cmdMaterialPrice_Click
        End If
    Else
        If Button = vbRightButton And IsNumeric(TDBGrid.Bookmark) Then
            Dim strErrorMsg As String
            strErrorMsg = m_objGridMap.GetError(TDBGrid.Bookmark)
            If Len(strErrorMsg) > 0 Then
                MsgBox strErrorMsg
            End If
        End If
    End If
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    Dim ctl As Control
    If Me.WindowState <> vbMinimized Then
        If Len(m_strCurrentFormControl) > 0 Then
            For Each ctl In Me.Controls
                If ctl.Name = m_strCurrentFormControl Then
                    ctl.SetFocus
                    Exit For
                End If
            Next ctl
        End If
        ShowGridSort
        fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Enabled = True
        fMainForm.tbToolBar.Buttons.Item(tbrPRINT).Visible = True
        fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = True
        fMainForm.tbToolBar.Buttons.Item(tbrPREVIEW).Visible = True
        fMainForm.mnuFilePageSetup.Enabled = True
        fMainForm.mnuFilePrint.Enabled = True
        fMainForm.mnuFilePrintPreview.Enabled = True
    
        OutputView False
        m_objGridMap.SetMenuBar
    End If
End Sub

Private Sub DoComparision()
    On Error GoTo myErr
    Dim i As Integer
    '*** APEX Migration Utility Code Change ***
    'Dim tdb As TrueOleDBGrid60.Columns
'*** APEX Migration Utility Code Change ***
'    Dim tdb As TrueOleDBGrid70.Columns
    Dim tdb As TrueOleDBGrid80.Columns
    Set tdb = TDBGrid.Columns
    m_rec.MoveFirst
    Debug.Print " True Grid Values "
    For i = 0 To tdb.Count - 1
        Debug.Print tdb.Item(i).Caption;
        Debug.Print " = ";
        Debug.Print tdb.Item(i).value
    Next i
    i = 0
    Debug.Print " Actual Record : "
    Debug.Print m_rec.Source
    For i = 0 To m_rec.Fields.Count - 1
        Debug.Print m_rec.Fields(i).Name;
        Debug.Print " = ";
        If Not IsNull(m_rec.Fields(i).value) Then
            Debug.Print m_rec.Fields(i).value
        Else
            Debug.Print " Null "
        End If
    Next i
    Exit Sub
myErr:
    Debug.Print Err.Description
    m_rec.MoveFirst
    Resume
End Sub

Private Sub techdesc_LostFocus()
    techdesc = Trim(techdesc)
End Sub

