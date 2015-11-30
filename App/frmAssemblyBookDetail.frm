VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAssemblyBookDetail 
   Caption         =   "Assembly Book Detail Maintenance"
   ClientHeight    =   4545
   ClientLeft      =   1965
   ClientTop       =   495
   ClientWidth     =   9135
   Icon            =   "frmAssemblyBookDetail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   9135
   Begin VB.TextBox alt_assembly_book_id 
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
      Left            =   3240
      TabIndex        =   1
      Tag             =   "1S"
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox asbly_last_update_id 
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   40
      Tag             =   "1N"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox assembly_book_skey 
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Tag             =   "1N"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox last_update_id 
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
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   26
      Tag             =   "1N"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox assembly_skey 
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "1N"
      Top             =   3480
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3480
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
      Left            =   4620
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "S"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   4020
      Width           =   1150
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   1440
      TabIndex        =   17
      Top             =   4020
      Width           =   1150
   End
   Begin VB.ComboBox type_code 
      Height          =   315
      ItemData        =   "frmAssemblyBookDetail.frx":0442
      Left            =   7920
      List            =   "frmAssemblyBookDetail.frx":0455
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "1S"
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox assembly_id 
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
      Left            =   5520
      TabIndex        =   2
      Tag             =   "2N"
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox assembly_book_id 
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
      Left            =   960
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   60
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Book Detail"
      TabPicture(0)   =   "frmAssemblyBookDetail.frx":0468
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label32"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label30"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label28"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "metric_unit"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "index_code"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "unit"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "book_qty"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "labor_hour"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "calculation_factor"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "coml_ind"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "resi_ind"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "metric_calculation_factor"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Descriptions"
      TabPicture(1)   =   "frmAssemblyBookDetail.frx":0484
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "metric_section_head_desc"
      Tab(1).Control(1)=   "index_desc"
      Tab(1).Control(2)=   "book_desc"
      Tab(1).Control(3)=   "section_head_desc"
      Tab(1).Control(4)=   "metric_book_desc"
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(6)=   "Label31"
      Tab(1).Control(7)=   "Label42"
      Tab(1).Control(8)=   "Label45"
      Tab(1).Control(9)=   "Label46"
      Tab(1).ControlCount=   10
      Begin VB.TextBox metric_section_head_desc 
         Height          =   285
         Left            =   -73020
         MaxLength       =   75
         TabIndex        =   15
         Tag             =   "1S"
         Top             =   2124
         Width           =   6795
      End
      Begin VB.TextBox metric_calculation_factor 
         Height          =   315
         Left            =   5880
         TabIndex        =   10
         Tag             =   "1N"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox resi_ind 
         Caption         =   "&Residential Use"
         Height          =   255
         Left            =   6240
         TabIndex        =   44
         Tag             =   "1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox coml_ind 
         Caption         =   "&Commercial Use"
         Height          =   255
         Left            =   4680
         TabIndex        =   43
         Tag             =   "1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox calculation_factor 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Tag             =   "1N"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox index_desc 
         Height          =   285
         Left            =   -73020
         MaxLength       =   35
         TabIndex        =   16
         Tag             =   "1S"
         Top             =   2520
         Width           =   2595
      End
      Begin VB.TextBox labor_hour 
         Height          =   315
         Left            =   5880
         TabIndex        =   8
         Tag             =   "1S"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox book_qty 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Tag             =   "1S"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox unit 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "1S"
         Top             =   840
         Width           =   1215
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
         Left            =   2040
         TabIndex        =   11
         Tag             =   "1S"
         Top             =   2280
         Width           =   435
      End
      Begin VB.TextBox book_desc 
         Height          =   525
         Left            =   -73020
         MaxLength       =   255
         TabIndex        =   12
         Tag             =   "1S"
         Top             =   465
         Width           =   6795
      End
      Begin VB.TextBox section_head_desc 
         Height          =   285
         Left            =   -73020
         MaxLength       =   75
         TabIndex        =   14
         Tag             =   "1S"
         Top             =   1731
         Width           =   6795
      End
      Begin VB.TextBox metric_book_desc 
         Height          =   525
         Left            =   -73020
         MaxLength       =   255
         TabIndex        =   13
         Tag             =   "1S"
         Top             =   1098
         Width           =   6795
      End
      Begin VB.TextBox metric_unit 
         Height          =   315
         Left            =   5880
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "1S"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Metric Sect Hdg Desc:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   46
         Top             =   2190
         Width           =   1755
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Metric Calculation Factor:"
         Height          =   255
         Left            =   3840
         TabIndex        =   45
         Top             =   1860
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Calculation Factor:"
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   1860
         Width           =   1335
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "Index Desc:"
         Height          =   255
         Left            =   -73980
         TabIndex        =   39
         Top             =   2580
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Labor Hours:"
         Height          =   255
         Left            =   4140
         TabIndex        =   36
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Quantity:"
         Height          =   255
         Left            =   660
         TabIndex        =   35
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Metric Unit:"
         Height          =   255
         Left            =   4740
         TabIndex        =   34
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit:"
         Height          =   255
         Left            =   1380
         TabIndex        =   33
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Index Code:"
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   2340
         Width           =   915
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "Imperial Book Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   735
         Width           =   1635
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Section Heading Desc:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   1780
         Width           =   1755
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Metric Book Desc:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   1368
         Width           =   1635
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Alt Book ID:"
      Height          =   255
      Left            =   2160
      TabIndex        =   42
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Assembly Book Skey:"
      Height          =   255
      Left            =   6120
      TabIndex        =   38
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "Book ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Assembly Skey:"
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   3540
      Width           =   1335
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   3540
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Type Code:"
      Height          =   255
      Left            =   6960
      TabIndex        =   20
      Top             =   90
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Assembly ID:"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   90
      Width           =   915
   End
End
Attribute VB_Name = "frmAssemblyBookDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As ADODB.RecordSet
Dim m_rec2 As New ADODB.RecordSet   'Header/Footer grid
Dim m_recAssembly As New ADODB.RecordSet

Dim m_blnRecFlag As Boolean ' True if a populated RecordSet was passed, then we show data
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
Dim m_blnClone As Boolean  'Indicate if clone is in progress
Dim m_blnDeleted As Boolean ' Indicates if the data has been deleted, used in QueryUnload
Dim strLastassembly_book_id As String ' Holds last assembly_book_id so we know if it changed
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
Dim strLast_assembly_book_id As String
Dim m_orig_assembly_book_id As String
Dim m_lngOriginalSkey As Long
Dim m_intKey As Long

Const BACKSPACE = 8
Const UNSELECTED = -1
Public Sub PrintReport()

End Sub

Public Sub PreviewReport()

End Sub
Private Function CheckEntryErrors() As Boolean
    Dim bln_New As Boolean
    Dim I As Integer
    Dim strError As String
    Dim strItem As String
    Dim varBookmarks() As Variant
    
    CheckEntryErrors = False
    If m_blnInsert Or m_blnClone Then
        bln_New = True
    End If
    If Len(Trim(assembly_id)) > 0 Then
        If Invalid_Assembly_id_Format(assembly_id, "assembly_id", m_rec, _
            bln_New, ConvertAssemblySkey(assembly_skey)) = True And CheckEntryErrors = False Then
            CheckEntryErrors = True
        End If
    End If
    If resi_ind.Value = 0 And coml_ind.Value = 0 And CheckEntryErrors = False Then
        MsgBox "The Commercial Use indicator or Residential Use indicator must be checked."
        CheckEntryErrors = True
    End If
End Function


Private Sub RebindTDBGridNow()
    Dim oldRow As Variant
    On Error Resume Next
    oldRow = myTDBGrid.Bookmark
    myTDBGrid.Refresh
    myTDBGrid.Bookmark = oldRow
End Sub

' Fills all fields with data
Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If m_rec.Fields("assembly_book_id") <> "" Then 'not new, must be clone
'    If Not (m_rec.Fields("assembly_book_skey") = 0 Or m_rec.Fields("assembly_book_skey") = "") Then
        m_blnRecFlag = True
    End If
    
'    If m_rec.Fields("type_code") = "D" Then
'    book_desc.Enabled = False
'    metric_book_desc.Enabled = False
'    End If
    
End Sub

Private Function valid_assembly_id() As Boolean
    On Error Resume Next
    If strLast_assembly_book_id <> assembly_id.Text Then
        Dim strSELECT As String
        Dim blnReturn As Boolean
        Dim rec As New ADODB.RecordSet
        strLast_assembly_book_id = assembly_id.Text

        ' Validate the entered assembly ID and retrieve the skey.
        strSELECT = "Select assembly_skey, last_update_id,unit, metric_unit,book_desc,metric_book_desc from assembly_detail where assembly_id='" + assembly_id.Text + "'"
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
        If rec.RecordCount > 0 Then
            assembly_skey.Text = rec.Fields("assembly_skey")
            asbly_last_update_id = rec.Fields("last_update_id")
            unit = rec.Fields("unit")
            metric_unit = rec.Fields("metric_unit")
            If Len(book_desc) = 0 Or type_code = "D" Then
                book_desc = rec.Fields("book_desc")
            End If
            If Len(metric_book_desc) = 0 Or type_code = "D" Then
                metric_book_desc = rec.Fields("metric_book_desc")
            End If
            valid_assembly_id = True
        Else
            MsgBox "Invalid Assembly Id."
            valid_assembly_id = False
        End If
        rec.Close
    Else
        valid_assembly_id = True
    End If

End Function

Private Sub assembly_book_id_Validate(Cancel As Boolean)
    Dim strSELECT As String
    Dim blnReturn As Boolean
    Dim ctr As Control
    Dim rec As New ADODB.RecordSet
    On Error Resume Next
    If strLast_assembly_book_id <> assembly_book_id.Text Then
        strLast_assembly_book_id = assembly_book_id.Text
    
        ' Validate the entered assembly ID and retrieve the skey.
        strSELECT = "Select * from assembly_book_detail where assembly_book_id='" + assembly_book_id.Text + "'"
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
        ' If it does
        If rec.RecordCount > 0 Then
            MsgBox "The assembly book id already exists and may not be used."
            Cancel = True
        End If
    End If

End Sub


Private Sub assembly_id_Change()

Static intIDLength As Integer
Dim strSELECT As String
Dim blnReturn As Boolean
Static blnChanging As Boolean
Static strAssemblyId As String

On Error GoTo Error_Processing

If m_intKey = BACKSPACE Then
    m_intKey = UNSELECTED
    intIDLength = Len(assembly_id.Text)     'Initial length
    If assembly_id <> "" Then
        assembly_id = Left(assembly_id, intIDLength - 1)
    End If
    GoTo Exit_Sub
End If

If blnChanging = False Then
    'Open Assembly recordset - used to autopopulate assembly Id
    strSELECT = "Select * from assembly_detail where assembly_id like '" + assembly_id.Text + "%'"
    m_recAssembly.MaxRecords = 1
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, m_recAssembly)
    If m_recAssembly.RecordCount = 0 Then
        If intIDLength > 0 Then
            m_recAssembly.Close
            strAssemblyId = Left(assembly_id, intIDLength)
            strSELECT = "Select * from assembly_detail where assembly_id like '" + strAssemblyId + "%'"
            m_recAssembly.MaxRecords = 1
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, m_recAssembly)
            strAssemblyId = m_recAssembly.Fields("assembly_id")
            blnChanging = True
            m_recAssembly.Close
            assembly_id = strAssemblyId
        End If
        m_recAssembly.Close
    Else
        strAssemblyId = m_recAssembly.Fields("assembly_id")
        blnChanging = True
        m_recAssembly.Close
        intIDLength = Len(assembly_id.Text)     'Initial length
        assembly_id = strAssemblyId
    End If
Else
    assembly_id.SelStart = intIDLength
    If Len(assembly_id) > intIDLength Then
        assembly_id.SelLength = Len(assembly_id) - intIDLength
    End If
    blnChanging = False
End If

Exit_Sub:
If m_recAssembly.State = adStateOpen Then
    m_recAssembly.Close
End If

Exit Sub

Error_Processing:
If Err <> 3705 And Err <> 3704 Then
    MsgBox Error$
End If
blnChanging = False
GoTo Exit_Sub
Resume 0
End Sub

Private Sub assembly_id_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8      'backspace
        m_intKey = BACKSPACE
End Select
    
End Sub


Private Sub assembly_id_LostFocus()
        Dim strSELECT As String
        Dim blnReturn As Boolean
        Dim rec As New ADODB.RecordSet
        Dim I As Integer
        strSELECT = "Select assembly_skey, type_code, assembly_skey, unit, metric_unit, book_desc, metric_book_desc, last_update_id from assembly_detail where assembly_id='" + assembly_id.Text + "'"
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec)
        If rec.RecordCount > 0 Then
            assembly_skey.Text = rec.Fields("assembly_skey")
            asbly_last_update_id = rec.Fields("last_update_id")
            assembly_skey = rec.Fields("assembly_skey")
            For I = 0 To unit.ListCount
                If unit.List(I) = rec.Fields("unit") Then
                    unit.ListIndex = I
                    Exit For
                End If
            Next I
            If rec.Fields("type_code") = "M" Then
                Dim rec2 As New ADODB.RecordSet
                strSELECT = "select labor_hour from published_assembly_cost where assembly_skey = " + CStr(rec.Fields("assembly_skey"))
                blnReturn = g_objDAL.GetRecordset(CONNECT, strSELECT, rec2)
                If Not rec2.EOF Then
                       labor_hour = rec2.Fields("labor_hour")
                End If
                rec2.Close
                Set rec2 = Nothing
            End If
            metric_unit = rec.Fields("metric_unit")
            If type_code = "D" Then     'Assembly Book Detail type code = D
                book_desc = rec.Fields("book_desc")
                metric_book_desc = rec.Fields("metric_book_desc")
            End If
            asbly_last_update_id = rec.Fields("last_update_id")
        End If
        rec.Close
        Set rec = Nothing
End Sub


Private Sub assembly_id_Validate(Cancel As Boolean)
If assembly_id <> "" Then
    If valid_assembly_id() = False Then
        Cancel = True
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

    strUpdate = "exec sp_delete_assembly_book_detail "
    strUpdate = strUpdate + "@assembly_book_skey=" + str(Me.Controls("assembly_book_skey"))
    
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
    Dim I As Integer
    ReDim tdbOldCols(tdbCols.Count - 1)
    For I = 0 To tdbCols.Count - 1
        tdbOldCols(I) = tdbCols.Item(I).Value
    Next
End Sub

Private Sub RestoreGridValues()
    ' this restores the grid back to its positioin if the user did not choose to save
    Dim I As Integer
    On Error Resume Next
    If m_blnInsert = False Then
        For I = 1 To tdbCols.Count - 1
            If tdbCols.Item(I).Value <> tdbOldCols(I) Then
                tdbCols.Item(I).Value = tdbOldCols(I)
            End If
        Next I
        'myTDBGrid.RefetchRow
        myTDBGrid.DataChanged = False
        myTDBGrid.RefreshRow
        DoEvents
    End If
End Sub

Private Sub cmdUpdate_Click()
    Dim blnRet As Boolean
    Dim blnUpdateBookDetail As Boolean
    Dim bln_Update_Grid As Boolean
    Dim ctr As Control
    Dim fld As ADODB.Field
    Dim rec As New ADODB.RecordSet
    Dim strError As String
    Dim strPercent_flag As String
    Dim strSELECT As String
    Dim strUpdate As String
    Dim strSaveUpdate As String
    Dim intStart As Integer
    Dim varSaveBookmark As Variant
    Dim I As Integer
    
    On Error GoTo Error_Processing
    
    m_blnWereErrors = CheckEntryErrors()
If m_blnWereErrors = False Then
    
    Screen.MousePointer = vbHourglass
    
    Dim recClone As ADODB.RecordSet
    Set recClone = m_rec.Clone
    recClone.AddNew
    UpdateRecordsetFromForm Me, recClone
    On Error Resume Next
    
    For Each fld In m_rec.Fields
        ' If the value changed
        If Not fld.Value = recClone.Fields(fld.Name).Value Or ((IsNull(fld.Value) Or fld.Value = "") Xor (recClone.Fields(fld.Name).Value = "")) Then
            Set ctr = Nothing
            Set ctr = Me.Controls(fld.Name)
            If Not ctr Is Nothing Then
                ' See what table the field is from
                If Left(Me.Controls(fld.Name).Tag, 1) = 1 Then
                    blnUpdateBookDetail = True
                End If
            End If
        End If
    Next

    On Error GoTo Error_Processing

    ' Undo the changes made by the UpdateRecordsetFromForm call above
    recClone.CancelUpdate
    recClone.Close
    Set recClone = Nothing

    If blnUpdateBookDetail = True Then
        strUpdate = "exec sp_update_assembly_book_detail "
        If last_update_id.Text = "" Then last_update_id.Text = 0
        If asbly_last_update_id.Text = "" Then asbly_last_update_id.Text = 0
        BuildStoredProcSQL Me, strUpdate, 1, m_rec
        strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
        strUpdate = strUpdate + ", @assembly_id='" + "'"
        m_blnWereErrors = False
        blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
        If blnRet = False Then
            MsgBox strError
            m_blnWereErrors = True
        Else
            last_update_id.Text = CInt(last_update_id.Text) + 1
            If m_blnClone = True Then       'Copy the output data for the assembly detail
                'need to retrieve the skey of the newly added book detail record
                strSELECT = "select assembly_book_skey from assembly_book_detail where assembly_book_id = '" + _
                assembly_book_id.Text + "'"
                blnRet = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
                If blnRet = False Then
                    MsgBox strError
                    m_blnWereErrors = True
                Else
                    strUpdate = "exec sp_copy_output_usage @type = 'A', @FromSkey = '" & m_lngOriginalSkey & _
                    "', @ToSkey='" & CStr(rec.Fields("assembly_book_skey")) + "', "
                    strUpdate = strUpdate + " @last_update_date='" + Format(Now(), "General Date") + "', "
                    strUpdate = strUpdate + " @last_update_person='" + strUserName + "', "
                    strUpdate = strUpdate + " @last_update_id='1'"
                    blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                    If blnRet = False Then
                        MsgBox strError
                        m_blnWereErrors = True
                    End If
                End If
                rec.Close
            End If
            'Process changes or deletions
            UpdateRecordsetFromForm Me, m_rec
            MsgBox "Update successful."
            RebindTDBGridNow
        End If
    Else
        If blnUpdateBookDetail = False Then
            MsgBox "You must modify a field before updating."
        End If
    End If
End If
Exit_Sub:
Screen.MousePointer = vbNormal
Exit Sub

Error_Processing:

MsgBox Error$
Resume Exit_Sub
Resume 0
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
    Dim blnReturn As Boolean
    
    Move START_LEFT, START_TOP, 9255, 4905
    
    If Not m_rec.State = adStateClosed Then
        UpdateFormFromRecordset Me, m_rec
        g_objDAL.GetRecordset vbNullString, "select unit from unit_of_measure order by unit", rec
        While Not rec.EOF
            unit.AddItem (rec.Fields("unit").Value)
            If Trim(m_rec.Fields("unit")) = Trim(rec.Fields("unit").Value) Then
                unit.Text = unit.List(unit.NewIndex)
            End If
            rec.MoveNext
        Wend
        rec.Close
    End If

    blnReturn = LockField(Me, "Metric_Unit")
    ' If we are NOT inserting
    If m_blnInsert = False Then
        ' Lock fields that can't be changed
        blnReturn = LockField(Me, "assembly_book_id")
'        blnReturn = LockField(Me, "assembly_id")
        Me.Caption = Me.Caption + " [" + m_rec.Fields("assembly_book_id").Value + "]"
    Else
        ' If we are inserting and not showing data
        ' Set some defaults
        If Not m_blnRecFlag Then
            Me.Caption = Me.Caption + " [New]"
            calculation_factor = 1
            metric_calculation_factor = 1
        Else
            Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("assembly_book_id").Value + "]"
            m_blnClone = True
            m_orig_assembly_book_id = m_rec.Fields("assembly_book_id").Value
            m_lngOriginalSkey = m_rec.Fields("assembly_book_skey").Value
            assembly_book_skey.Text = "0"   'Clear skey for clone
        End If
    End If
    If Not m_blnClone Then
        strLastassembly_book_id = m_rec.Fields("assembly_book_id").Value
    End If
    type_code_LostFocus
    ColorLockedFields Me

End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub index_code_Change()
    ConvertUcase Me.Controls("index_code")
End Sub

Public Sub ConvertUcase(ctlControl As Control)
    Dim intSelStart As Integer
    Dim intSelLen As Integer
    
    intSelStart = ctlControl.SelStart
    intSelLen = ctlControl.SelLength
    
    ctlControl = UCase(ctlControl)
    ctlControl.SelStart = intSelStart
    ctlControl.SelLength = intSelLen

End Sub

Private Sub index_code_Validate(Cancel As Boolean)
    If index_code <> "IX" And index_code <> "JX" Then
        MsgBox "Please enter 'JX' or 'IX' for the index code."
        Cancel = True
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


Private Sub metric_section_head_desc_Validate(Cancel As Boolean)
Dim intLength As Integer
If Len(metric_section_head_desc) > 0 Then
    If Asc(Right(metric_section_head_desc, 1)) >= 0 And Asc(Right(metric_section_head_desc, 1)) <= 31 Then
'        intLength = Len(metric_section_head_desc)
        MsgBox "Non-printable characters are not allowed in the metric section head description."
'        metric_section_head_desc.Text = left(metric_section_head_desc.Text, intLength - 2)
'        metric_section_head_desc.SelStart = intLength - 2
        Cancel = True
    End If
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
                m_blnWereErrors = False
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

Private Sub section_head_desc_Validate(Cancel As Boolean)
Dim intLength As Integer
If Len(section_head_desc) > 0 Then
    If Asc(Right(section_head_desc, 1)) >= 0 And Asc(Right(section_head_desc, 1)) <= 31 Then
'        intLength = Len(section_head_desc)
        MsgBox "Non-printable characters are not allowed in the section head description."
        Cancel = True
'        section_head_desc.Text = left(section_head_desc.Text, intLength - 2)
'        section_head_desc.SelStart = intLength - 2
    End If
End If

End Sub

Private Sub type_code_LostFocus()
    If type_code.Text <> "H" Then
        LockField Me, "section_head_desc"
        LockField Me, "metric_section_head_desc"
    Else
        UnLockField Me, "section_head_desc"
        UnLockField Me, "metric_section_head_desc"
    End If
End Sub


