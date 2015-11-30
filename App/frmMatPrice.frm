VERSION 5.00
Begin VB.Form frmMatPrice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Price"
   ClientHeight    =   7470
   ClientLeft      =   2265
   ClientTop       =   2175
   ClientWidth     =   8400
   Icon            =   "frmMatPrice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8400
   Begin VB.TextBox mat_skey 
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
      Height          =   315
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   74
      Tag             =   "N"
      Top             =   6300
      Width           =   1215
   End
   Begin VB.TextBox matprice_last_update_person 
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
      Height          =   315
      Left            =   4380
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   6300
      Width           =   1215
   End
   Begin VB.TextBox matprice_last_update_date 
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
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   6300
      Width           =   1635
   End
   Begin VB.TextBox alt_mat_id 
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
      MaxLength       =   13
      TabIndex        =   1
      Tag             =   "1S"
      Top             =   60
      Width           =   1395
   End
   Begin VB.Frame Frame4 
      Caption         =   "Material Price Indicators"
      Height          =   555
      Left            =   180
      TabIndex        =   64
      Top             =   2820
      Width           =   7875
      Begin VB.CheckBox traces_ind 
         Caption         =   "TRACES"
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Tag             =   "2"
         Top             =   180
         Width           =   975
      End
      Begin VB.CheckBox factor_ind 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   15
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox use_ind 
         Caption         =   "Use"
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Tag             =   "2"
         Top             =   180
         Width           =   615
      End
      Begin VB.CheckBox wst_use_ind 
         Caption         =   "West"
         Height          =   315
         Left            =   1260
         TabIndex        =   13
         Tag             =   "2"
         Top             =   180
         Width           =   735
      End
      Begin VB.CheckBox update_ind 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6900
         TabIndex        =   18
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox estimated_ind 
         Caption         =   "Estimated"
         Height          =   315
         Left            =   4140
         TabIndex        =   16
         Tag             =   "2"
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox update_status_code 
         Height          =   315
         Left            =   6300
         TabIndex        =   17
         Tag             =   "2S"
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label28 
         Caption         =   "Update"
         Height          =   195
         Left            =   7155
         TabIndex        =   67
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label27 
         Caption         =   "Factor"
         Height          =   195
         Left            =   3495
         TabIndex        =   66
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Update Stat:"
         Height          =   255
         Left            =   5340
         TabIndex        =   65
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Material Price"
      Height          =   915
      Left            =   180
      TabIndex        =   56
      Top             =   1860
      Width           =   7875
      Begin VB.TextBox traces_price 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox list_price 
         Height          =   315
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   9
         Tag             =   "2N"
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox traces_list_price 
         Height          =   315
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   10
         Tag             =   "2N"
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox pct_multiplier 
         Height          =   315
         Left            =   6480
         TabIndex        =   11
         Tag             =   "2N"
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox material_price 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "TRACES Price:"
         Height          =   255
         Left            =   2580
         TabIndex        =   63
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "List Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "TRACES List:"
         Height          =   255
         Left            =   2580
         TabIndex        =   60
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Pct Mult:"
         Height          =   255
         Left            =   5460
         TabIndex        =   59
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Material Price:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contact"
      Height          =   555
      Left            =   180
      TabIndex        =   52
      Top             =   4380
      Width           =   7875
      Begin VB.TextBox contact_name 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox company_name 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox contact_id 
         Height          =   315
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   25
         Tag             =   "2S"
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   5280
         TabIndex        =   70
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Company:"
         Height          =   255
         Left            =   2160
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact ID:"
         Height          =   255
         Left            =   60
         TabIndex        =   53
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.TextBox matprice_last_update_id 
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   50
      Tag             =   "N"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox mat_last_update_id 
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   49
      Tag             =   "N"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox usage_unit 
      Height          =   315
      Left            =   6900
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "1S"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox purchase_unit 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "1S"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4320
      TabIndex        =   29
      Top             =   6840
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manufacturer"
      Height          =   915
      Left            =   180
      TabIndex        =   40
      Top             =   3420
      Width           =   7875
      Begin VB.TextBox commodity_code 
         Height          =   315
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   21
         Tag             =   "2S"
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox manufacturer_id 
         Height          =   315
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "2S"
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox catalog_num 
         Height          =   315
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   22
         Tag             =   "2S"
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox page_num 
         Height          =   315
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   24
         Tag             =   "2S"
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox model_name 
         Height          =   315
         Left            =   3840
         MaxLength       =   25
         TabIndex        =   20
         Tag             =   "2S"
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox item_num 
         Height          =   315
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   23
         Tag             =   "2S"
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Comm Code:"
         Height          =   255
         Left            =   5460
         TabIndex        =   51
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Manufact ID:"
         Height          =   255
         Left            =   180
         TabIndex        =   45
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Cat Num:"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Num:"
         Height          =   255
         Left            =   2700
         TabIndex        =   43
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Page Num:"
         Height          =   255
         Left            =   5460
         TabIndex        =   42
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Model:"
         Height          =   255
         Left            =   2700
         TabIndex        =   41
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   2820
      TabIndex        =   28
      Top             =   6840
      Width           =   1150
   End
   Begin VB.TextBox latest_price_update_comment 
      Height          =   315
      Left            =   1260
      MaxLength       =   80
      TabIndex        =   26
      Tag             =   "2S"
      Top             =   5460
      Width           =   6855
   End
   Begin VB.TextBox comment 
      Height          =   315
      Left            =   1260
      MaxLength       =   255
      TabIndex        =   27
      Tag             =   "2S"
      Top             =   5880
      Width           =   6855
   End
   Begin VB.TextBox term_date 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3780
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   5040
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
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   30
      Tag             =   "D"
      Top             =   5040
      Width           =   1035
   End
   Begin VB.TextBox purchase_usage_conv_factor 
      Height          =   315
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   6
      Tag             =   "1N"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox metric_tech_desc 
      Height          =   315
      Left            =   1260
      MaxLength       =   75
      TabIndex        =   4
      Tag             =   "1S"
      Top             =   900
      Width           =   6915
   End
   Begin VB.TextBox tech_desc 
      Height          =   315
      Left            =   1260
      MaxLength       =   75
      TabIndex        =   3
      Tag             =   "1S"
      Top             =   480
      Width           =   6915
   End
   Begin VB.CheckBox active_status_ind 
      Caption         =   "Active"
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Tag             =   "1"
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox mat_id 
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
      Left            =   1260
      MaxLength       =   17
      TabIndex        =   0
      Tag             =   "1S"
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Caption         =   "Material Skey:"
      Height          =   255
      Left            =   5640
      TabIndex        =   71
      Top             =   6360
      Width           =   1035
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "Alt Mat ID:"
      Height          =   255
      Left            =   3180
      TabIndex        =   68
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated:"
      Height          =   255
      Left            =   60
      TabIndex        =   48
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Updated By:"
      Height          =   255
      Left            =   3000
      TabIndex        =   47
      Top             =   6360
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      X1              =   120
      X2              =   8280
      Y1              =   1780
      Y2              =   1780
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Usage Unit:"
      Height          =   255
      Left            =   5700
      TabIndex        =   46
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "Latest Price Comment:"
      Height          =   435
      Left            =   240
      TabIndex        =   39
      Top             =   5460
      Width           =   915
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "Comment:"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   5940
      Width           =   915
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "End Date:"
      Height          =   255
      Left            =   2760
      TabIndex        =   37
      Top             =   5100
      Width           =   915
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date:"
      Height          =   255
      Left            =   60
      TabIndex        =   36
      Top             =   5100
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Conversion Factor:"
      Height          =   255
      Left            =   2580
      TabIndex        =   35
      Top             =   1380
      Width           =   1395
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Purchase Unit:"
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Metric Tech Desc:"
      Height          =   435
      Left            =   240
      TabIndex        =   33
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tech Desc:"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Material ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmMatPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_rec As ADODB.RecordSet
Dim m_blnRecFlag As Boolean ' True if a populated RecordSet was passed, then we show data
Dim m_blnWereErrors As Boolean ' True if the Update had errors, used in QueryUnload
Dim m_blnInsert As Boolean ' Tells if we are doing an insert or update
Dim strLast_mat_id As String ' Holds last mat_id so we know if it changed
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
Public UserIsEditing As Boolean

' Fills all fields with data
Public Sub SetRow(rec As ADODB.RecordSet, Optional blnInsert As Boolean = False)
    Set m_rec = rec
    m_blnInsert = blnInsert
    ' If we are inserting/cloning
    If m_blnInsert Then
        ' Do this so OriginalValue will be set to the values copied into the row
        m_rec.UpdateBatch
    End If
    If Not m_rec.Fields("mat_skey") = 0 Then
        m_blnRecFlag = True
    End If

End Sub


Private Sub active_status_ind_Click()
    ChangeInGrid "Active", active_status_ind.Value
End Sub

Private Sub alt_mat_id_Change()
Dim intSelStart As Integer
Dim intSelLen As Integer

ChangeInGrid "Alt Mat Id", alt_mat_id

If alt_mat_id.DataChanged Then
    intSelStart = alt_mat_id.SelStart
    intSelLen = alt_mat_id.SelLength
    If UCase(Left(alt_mat_id, 1)) <> "M" Then
        alt_mat_id = "M" + alt_mat_id
        intSelStart = intSelStart + 1
    Else
        alt_mat_id = UCase(alt_mat_id)
    End If
    alt_mat_id.SelStart = intSelStart
    alt_mat_id.SelLength = intSelLen
End If


End Sub

Private Sub alt_mat_id_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub alt_mat_id_Validate(Cancel As Boolean)

If alt_mat_id <> "" Then
    If Invalid_mat_id_Format(Compress_String(alt_mat_id), "alt_mat_id", m_rec) = True Then
        Cancel = True
    End If
End If

End Sub

Private Sub catalog_num_Change()
    ChangeInGrid "Cat Num", catalog_num
End Sub

Private Sub catalog_num_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    Dim strUpdate As String
    Dim blnRet As Boolean
    Dim strError As String
    Dim strSelect As String
    Dim rec As ADODB.RecordSet
    
    strSelect = "select count(*) as NbrMatsUsed from material_usage as mu where mat_skey = " + CStr(m_rec.Fields("mat_skey"))
    g_objDAL.GetRecordset CONNECT, strSelect, rec
    If Not rec.EOF Then
        If rec.Fields("NbrMatsUsed") > 0 Then
            MsgBox str(rec.Fields("NbrMatsUsed")) + " material usage record(s) exist.  The material may not be deleted."
        Else
            Dim varButton
            varButton = MsgBox("Are you sure you want to delete?", vbYesNo + vbCritical)
            If varButton = vbNo Then
                Exit Sub
            End If
        
            strUpdate = "exec sp_delete_material_price "
            strUpdate = strUpdate + "@mat_skey=" + str(Me.Controls("mat_skey")) + ","
            strUpdate = strUpdate + " @contact_id='" + Me.Controls("contact_id") + "',"
            strUpdate = strUpdate + " @manufacturer_id='" + Me.Controls("manufacturer_id") + "',"
            strUpdate = strUpdate + " @start_date='" + Me.Controls("start_date") + "',"
            strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
            
            blnRet = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
            If Not blnRet Then
                MsgBox strError
            Else
                MsgBox "Delete successful."
                m_rec.Delete
                RebindTDBGridNow
                Unload Me
            End If
        End If
    End If
    rec.Close
    Set rec = Nothing
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
    Dim blnUpdateMat As Boolean
    Dim blnUpdateMatPrice As Boolean
    Dim strSelect As String
    Dim rec As New ADODB.RecordSet
    Dim recClone As ADODB.RecordSet
    Dim strMsg As String
    
    m_blnWereErrors = False
    mat_id = Compress_String(mat_id)
    ' if we are updating
    If m_blnInsert = False Then
        Set recClone = m_rec.Clone
        recClone.AddNew
        UpdateRecordsetFromForm Me, recClone ' m_rec
        For Each fld In m_rec.Fields
            ' If the value changed
'            If Not fld.OriginalValue = fld.Value Or ((IsNull(fld.OriginalValue) Or fld.Value = "") Xor (fld.OriginalValue = "")) Then
            If Trim(fld.OriginalValue) <> Trim(recClone.Fields(fld.Name).Value) Or _
                        (IsNull(fld.Value) Or Trim(fld.Value = "") And (Trim(recClone.Fields(fld.Name).Value <> ""))) Then
                Set ctr = Nothing
                Set ctr = Me.Controls(fld.Name)
                If Not ctr Is Nothing Then
                    ' See what table the field is from
                    ' Mark the table we should update
                    If Left(Me.Controls(fld.Name).Tag, 1) = 1 Then
                        blnUpdateMat = True
                    ElseIf Left(Me.Controls(fld.Name).Tag, 1) = 2 Then
                        blnUpdateMatPrice = True
                    End If
                End If
            End If
        Next
        ' Undo the changes made by the UpdateRecordsetFromForm call above
        recClone.CancelUpdate
        recClone.Close
        Set recClone = Nothing
        If blnUpdateMatPrice And blnUpdateMat Then
            blnRet = False
            strUpdate = "exec sp_update_material_and_price @mat_skey=" + mat_skey.Text + ", @matprice_last_update_id=" + matprice_last_update_id.Text + ", @mat_last_update_id=" + mat_last_update_id.Text + ", "
            BuildStoredProcSQL Me, strUpdate, 1
            BuildStoredProcSQL Me, strUpdate, 2
            strUpdate = strUpdate + " @old_contact_id='" + m_rec.Fields("contact_id").OriginalValue + "', "
            strUpdate = strUpdate + " @old_manufacturer_id='" + m_rec.Fields("manufacturer_id").OriginalValue + "', "
            strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
            strUpdate = strUpdate + " @factor_ind=" + str(factor_ind.Value) + ", "
            strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
            blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
            If blnRet = False Then
                strMsg = strError
                m_blnWereErrors = True
            Else
                ' Put latest data into source recordset
                UpdateRecordsetFromForm Me, m_rec
                m_rec.Fields("matprice_last_update_id").Value = m_rec.Fields("matprice_last_update_id").Value + 1
                matprice_last_update_id.Text = m_rec.Fields("matprice_last_update_id").Value
                m_rec.Fields("mat_last_update_id").Value = m_rec.Fields("mat_last_update_id").Value + 1
                mat_last_update_id.Text = m_rec.Fields("mat_last_update_id").Value
                strSelect = "select start_date from material_price where mat_skey=" + str(m_rec.Fields("mat_skey")) + " and last_update_id=" + str(m_rec.Fields("matprice_last_update_id"))
                blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                If blnRet Then
                    m_rec.Fields("start_date") = rec.Fields("start_date")
                End If
                UpdateFormFromRecordset Me, m_rec
                strMsg = "Update successful."
                RebindTDBGridNow
           End If
        ElseIf blnUpdateMatPrice Then
            blnRet = False
            strUpdate = "exec sp_update_material_price @mat_skey=" + mat_skey.Text + ", @matprice_last_update_id=" + matprice_last_update_id.Text + ", "
            BuildStoredProcSQL Me, strUpdate, 2
            strUpdate = strUpdate + " @old_contact_id='" + m_rec.Fields("contact_id").OriginalValue + "', "
            strUpdate = strUpdate + " @old_manufacturer_id='" + m_rec.Fields("manufacturer_id").OriginalValue + "', "
            strUpdate = strUpdate + " @start_date='" + str(m_rec.Fields("start_date").Value) + "', "
            strUpdate = strUpdate + " @factor_ind=" + str(factor_ind.Value) + ", "
            strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
            blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
            If blnRet = False Then
                strMsg = strError
                 m_blnWereErrors = True
            Else
                ' Put latest data into source recordset
                UpdateRecordsetFromForm Me, m_rec
                m_rec.Fields("matprice_last_update_id").Value = m_rec.Fields("matprice_last_update_id").Value + 1
                matprice_last_update_id.Text = m_rec.Fields("matprice_last_update_id").Value
                strSelect = "select start_date from material_price where mat_skey=" + str(m_rec.Fields("mat_skey")) + " and last_update_id=" + str(m_rec.Fields("matprice_last_update_id"))
                blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                If blnRet Then
                    m_rec.Fields("start_date") = rec.Fields("start_date")
                End If
                UpdateFormFromRecordset Me, m_rec
               strMsg = "Update successful."
                RebindTDBGridNow
           End If
        ElseIf blnUpdateMat Then
            blnRet = False
'''            Dim tmpNext As Integer
'''            If (mat_skey.Text = old_mat_skey_text) Then
'''                tmpNext = CInt(mat_last_update_id.Text)
'''                tmpNext = tmpNext + 1
'''                strUpdate = "exec sp_update_material @mat_skey=" + mat_skey.Text + ", @last_update_id=" + CStr(tmpNext) + ", "
              strUpdate = "exec sp_update_material @mat_skey=" + mat_skey.Text + ", @last_update_id=" + mat_last_update_id.Text + ", "  'LEGACY
'''    ''          strUpdate = "exec sp_update_material @mat_skey=" + mat_skey.Text + ", @last_update_id=" + CStr(m_rec.Fields("mat_last_update_id").Value + 1) + ", "
'''            Else
'''                strUpdate = "exec sp_update_material @mat_skey=" + mat_skey.Text + ", @last_update_id=" + mat_last_update_id.Text + ", "
'''            End If
            BuildStoredProcSQL Me, strUpdate, 1
            strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
            Debug.Print strUpdate
            blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
            If blnRet = False Then
                strMsg = strError
                m_blnWereErrors = True
            Else
                ' Put latest data into source recordset
'                old_mat_skey_text = mat_skey.Text       'rlh - save
                
                UpdateRecordsetFromForm Me, m_rec
                m_rec.Fields("mat_last_update_id").Value = m_rec.Fields("mat_last_update_id").Value + 1
                mat_last_update_id.Text = m_rec.Fields("mat_last_update_id").Value
                UpdateFormFromRecordset Me, m_rec
                strMsg = "Update successful."
                RebindTDBGridNow
            End If
        Else
            strMsg = "You must modify a field before updating."
            GoTo Exit_sub
        End If
    ' If we are inserting (or cloning)
    Else
        If IsControlChanged(Me, m_rec) Then
            Set recClone = m_rec.Clone
            recClone.AddNew
            UpdateRecordsetFromForm Me, recClone
            ' If mat_skey is set, then just insert price and maybe update material
            If Len(mat_skey.Text) > 0 Then
                ' Check if we need to update material
                For Each fld In m_rec.Fields
                    ' If the value changed
                    If Not fld.Value = recClone.Fields(fld.Name).Value Or (fld.Value = "" Xor (recClone.Fields(fld.Name).Value = "")) Then
                        Set ctr = Nothing
                        Set ctr = Me.Controls(fld.Name)
                        If Not ctr Is Nothing Then
                            ' See what table the field is from
                             ' If it is from Material
                            If Left(Me.Controls(fld.Name).Tag, 1) = 1 Then
                                blnRet = False
                                strUpdate = "exec sp_update_material @mat_skey=" + mat_skey.Text + ", @last_update_id=" + mat_last_update_id.Text + ", "
                                BuildStoredProcSQL Me, strUpdate, 1
                                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                                If blnRet = False Then
                                    strMsg = strError
                                    m_blnWereErrors = True
                                Else
                                   ' Put latest data into source recordset
                                   UpdateRecordsetFromForm Me, m_rec
                                   m_rec.Fields("mat_last_update_id").Value = m_rec.Fields("mat_last_update_id").Value + 1
                                   mat_last_update_id.Text = m_rec.Fields("mat_last_update_id").Value
                                   UpdateFormFromRecordset Me, m_rec
                                End If
                                Exit For
                            End If
                        End If
                    End If
                Next
                
                ' Undo the changes made by the UpdateRecordsetFromForm call above
                recClone.CancelUpdate
                recClone.Close
                Set recClone = Nothing
                
                ' Now insert price
                strUpdate = "exec sp_insert_material_price @mat_skey=" + mat_skey.Text + ", "
                BuildStoredProcSQL Me, strUpdate, 2
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    strMsg = strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("matprice_last_update_id").Value = m_rec.Fields("matprice_last_update_id").Value + 1
                    matprice_last_update_id.Text = m_rec.Fields("matprice_last_update_id").Value
                    strSelect = "select start_date from material_price where mat_skey=" + str(m_rec.Fields("mat_skey")) + " and last_update_id=" + str(m_rec.Fields("matprice_last_update_id") - 1)
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                End If
                If m_blnWereErrors = False Then
                    strMsg = "Update successful."
                    RebindTDBGridNow
                End If
            ' Insert material and price
            Else
                blnRet = False
                strUpdate = "exec sp_insert_material_and_price "
                BuildStoredProcSQL Me, strUpdate, 1
                BuildStoredProcSQL Me, strUpdate, 2
                strUpdate = strUpdate + " @last_update_person='" + strUserName + "'"
                blnRet = g_objDAL.ExecQuery(vbNullString, strUpdate, strError)
                If blnRet = False Then
                    strMsg = strError
                    m_blnWereErrors = True
                Else
                    ' Put latest data into source recordset
                    UpdateRecordsetFromForm Me, m_rec
                    m_rec.Fields("mat_last_update_id").Value = m_rec.Fields("mat_last_update_id").Value + 1
                    mat_last_update_id.Text = m_rec.Fields("mat_last_update_id").Value
                    m_rec.Fields("matprice_last_update_id").Value = m_rec.Fields("matprice_last_update_id").Value + 1
                    matprice_last_update_id.Text = m_rec.Fields("matprice_last_update_id").Value
                    strSelect = "select start_date from material_price where mat_skey=" + str(m_rec.Fields("mat_skey")) + " and last_update_id=" + str(m_rec.Fields("matprice_last_update_id"))
                    blnRet = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
                    If blnRet Then
                        m_rec.Fields("start_date") = rec.Fields("start_date")
                    End If
                    UpdateFormFromRecordset Me, m_rec
                    strMsg = "Update successful."
                    RebindTDBGridNow
                    
                    'code added by mohan Jan 18, 2012: update the Hierarchy tree
                    Dim retBlnVal As Boolean
                    retBlnVal = MainModule.Update_Tree_With_Unit_Cost_Id(mat_id.Text, alt_mat_id.Text)
                    If retBlnVal = False Then
                        strMsg = "Update successful for Material Price, but there was an error while updating the Tree."
                    End If
                    
                    
                End If
            End If
        Else
            strMsg = "You must modify a field before updating."
        End If
    End If

Exit_sub:
    'Line of code was changed by Mohan on Jan 05,2012: changed FORMAT_MATERIAL_SRV to FORMAT_MATERIAL_04_SRV
    mat_id = Format(mat_id, FORMAT_MATERIAL_04_SRV)
    MsgBox strMsg

End Sub

Private Sub comment_Change()
    ChangeInGrid "Comment", comment
End Sub

Private Sub comment_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub commodity_code_Change()
    ChangeInGrid "Comm Code", commodity_code
End Sub

Private Sub commodity_code_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub contact_id_Change()
    Dim mPos As Integer
    If Len(contact_id) > 0 Then
        mPos = contact_id.SelStart
        contact_id = UCase(contact_id)
        contact_id.SelStart = mPos
        ChangeInGrid "Contact", contact_id
    End If
End Sub

Private Sub contact_id_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub contact_id_Validate(Cancel As Boolean)
    Dim rec As ADODB.RecordSet
    If Not Len(contact_id.Text) = 0 Then
        contact_id.Text = UCase(contact_id.Text)
        ' inserted the above 1 line on 8/13/99 siva
        g_objDAL.GetRecordset vbNullString, "select company_name, first_name, last_name from information_source where contact_id = '" + contact_id.Text + "'", rec
        If rec.RecordCount = 0 Then
            MsgBox "You must enter a valid Contact ID."
            SendKeys ("{HOME}+{END}")
            rec.Close
            Cancel = True
        Else
            company_name.Text = rec.Fields("company_name").Value
            contact_name.Text = rec.Fields("first_name")
            If Len(contact_name.Text) > 0 Then contact_name.Text = contact_name.Text + " "
            contact_name.Text = contact_name.Text + rec.Fields("last_name")
            rec.Close
        End If
    End If
End Sub

Private Sub estimated_ind_Click()
    ChangeInGrid "Estimated", estimated_ind.Value
End Sub

Private Sub factor_ind_Click()
    ChangeInGrid "Factor", factor_ind.Value
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    UserIsEditing = True
    frmCallingForm.Enabled = False
    OutputView False
End Sub

Private Sub Form_Initialize()
    m_blnInsert = False
    Set m_rec = Nothing
    UserIsEditing = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim ctr As Control
    Dim rec As ADODB.RecordSet
    Dim strSelect As String
    Dim blnReturn As Boolean
    
    Move START_LEFT, START_TOP
    strLast_mat_id = ""
    
    g_objDAL.GetRecordset CONNECT, "select unit from unit_of_measure order by unit", rec
    While Not rec.EOF
        purchase_unit.AddItem (rec.Fields("unit").Value)
        usage_unit.AddItem (rec.Fields("unit").Value)
        rec.MoveNext
    Wend
    rec.Close
    strSelect = "select count(*) as NbrMatsUsed from material_usage as mu where mat_skey = " + CStr(m_rec.Fields("mat_skey").Value)
    g_objDAL.GetRecordset CONNECT, strSelect, rec
    If Not rec.EOF Then
        If rec.Fields("NbrMatsUsed") > 0 Then
            blnReturn = LockField(Me, "active_status_ind")
        End If
    End If
    rec.Close
    
    ' Load data into form
    If Not m_rec.State = adStateClosed Then
        UpdateFormFromRecordset Me, m_rec
    End If
    
    'Line of code was changed by Mohan on Jan 05,2012: changed FORMAT_MATERIAL_SRV to FORMAT_MATERIAL_04_SRV
    mat_id = Format(Compress_String(mat_id), FORMAT_MATERIAL_04_SRV)
    strLast_mat_id = m_rec.Fields("mat_id").Value
    
    ' Build contact name
    contact_name.Text = m_rec.Fields("first_name")
    If Len(contact_name.Text) > 0 Then contact_name.Text = contact_name.Text + " "
    contact_name.Text = contact_name.Text + m_rec.Fields("last_name")
    
    ' If we are NOT inserting
    If m_blnInsert = False Then
        ' Lock fields that can't be changed
        mat_id.Locked = True
        mat_id.Enabled = False
        mat_id.BackColor = LTGREY
        ' Set caption
        Me.Caption = Me.Caption + " [" + m_rec.Fields("mat_id").Value + "]"
    Else
        ' If we are inserting and not showing data
        active_status_ind.Value = 1
        m_rec.Fields("active_status_ind").Value = True
        blnReturn = LockField(Me, "active_status_ind")
        blnReturn = LockField(Me, "update_status_code")
        ' Set some defaults
        'Retrieve the updt_status from the domain_tbl for new records
        g_objDAL.GetRecordset CONNECT, "select update_status_code = domain_value from domain_tbl where domain_name = 'PAPER_CLIP'", rec
        If Not rec.EOF Then
            m_rec.Fields("update_status_code") = rec.Fields("update_status_code")
            update_status_code = m_rec.Fields("update_status_code")
        End If
        If Not m_blnRecFlag Then
            pct_multiplier.Text = 100
            m_rec.Fields("pct_multiplier").Value = 100
            purchase_usage_conv_factor.Text = 1
            m_rec.Fields("purchase_usage_conv_factor").Value = 1
            wst_use_ind.Value = 0
            m_rec.Fields("wst_use_ind").Value = 0
            estimated_ind.Value = 0
            m_rec.Fields("estimated_ind").Value = 0
            traces_ind.Value = 0
            m_rec.Fields("traces_ind").Value = 0
            update_ind.Value = 0
            m_rec.Fields("update_ind").Value = 0
            factor_ind.Value = 0
            m_rec.Fields("factor_ind").Value = 0
            ' This is a new record
            Me.Caption = Me.Caption + " [New]"
            use_ind.Value = 1
            use_ind.Enabled = False 'Required field
        Else
            ' This means we are cloning
            Me.Caption = Me.Caption + " [Clone of " + m_rec.Fields("mat_id").Value + "]"
            use_ind.Enabled = True    'If the material does exist, unlock the use indicator
        End If
        m_rec.Fields("use_ind").Value = True
    End If
    Store_Grid_Old_Values
    ColorLockedFields Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmCallingForm.txtTotalChildForms.Text = Val(frmCallingForm.txtTotalChildForms.Text) - 1
    frmCallingForm.Enabled = True
End Sub

Private Sub item_num_Change()
    ChangeInGrid "Item Num", item_num
End Sub

Private Sub item_num_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub latest_price_update_comment_Change()
    ChangeInGrid "Latest Price Comment", latest_price_update_comment
End Sub

Private Sub latest_price_update_comment_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub list_price_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub list_price_KeyPress(KeyAscii As Integer)
    ' added 8/12/99 siva
    If CheckNumericField(list_price, KeyAscii, list_price.SelStart, list_price.SelLength, 2) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub list_price_Validate(Cancel As Boolean)
    CheckValueForNumber list_price.Text, Cancel
End Sub

Private Sub manufacturer_id_Change()
    Dim mPos As Integer
    If Len(manufacturer_id) > 0 Then
        mPos = manufacturer_id.SelStart
        manufacturer_id = UCase(manufacturer_id)
        manufacturer_id.SelStart = mPos
        ChangeInGrid "Manufacturer", manufacturer_id
    End If
End Sub

Private Sub manufacturer_id_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub manufacturer_id_Validate(Cancel As Boolean)
    Dim rec As ADODB.RecordSet
    If Not Len(manufacturer_id.Text) = 0 Then
        g_objDAL.GetRecordset vbNullString, "select count(manufacturer_id) from material_manufacturer where manufacturer_id = '" + manufacturer_id.Text + "'", rec
        If rec.Fields(0).Value = 0 Then
            MsgBox "You must enter a valid Manufacturer ID."
            Cancel = True
        End If
        rec.Close
    End If
End Sub

Private Sub mat_id_Change()
Dim intSelStart As Integer
Dim intSelLen As Integer

If mat_id.DataChanged Then
    intSelStart = mat_id.SelStart
    intSelLen = mat_id.SelLength
    If UCase(Left(Trim(mat_id), 1)) <> "M" Then
        mat_id = "M" + mat_id
        intSelStart = intSelStart + 1
    Else
        mat_id = UCase(mat_id)
    End If
    mat_id.SelStart = intSelStart
    mat_id.SelLength = intSelLen
End If

End Sub

Private Sub mat_id_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub mat_id_Validate(Cancel As Boolean)

    If Invalid_mat_id_Format(Compress_String(mat_id), "mat_id", m_rec) = True Then
        Cancel = True
    End If

End Sub

Private Sub mat_last_update_id_Change()
    ChangeInGrid "mat_last_update_id", mat_last_update_id
End Sub

Private Sub mat_skey_Change()
    ChangeInGrid "mat_skey", mat_skey
End Sub

Private Sub matprice_last_update_date_Change()
    ChangeInGrid "Update Date", matprice_last_update_date
End Sub

Private Sub matprice_last_update_id_Change()
    ChangeInGrid "matprice_last_update_id", matprice_last_update_id
End Sub

Private Sub matprice_last_update_person_Change()
    ChangeInGrid "Update Person", matprice_last_update_person
End Sub

Private Sub metric_tech_desc_Change()
    ChangeInGrid "Metric Tech Desc", metric_tech_desc
End Sub

Private Sub metric_tech_desc_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub model_name_Change()
    ChangeInGrid "Model", model_name
End Sub

Private Sub model_name_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub page_num_Change()
    ChangeInGrid "Page Num", page_num
End Sub

Private Sub page_num_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub pct_multiplier_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub pct_multiplier_KeyPress(KeyAscii As Integer)
    If CheckNumericField(pct_multiplier, KeyAscii, pct_multiplier.SelStart, pct_multiplier.SelLength, 5) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub purchase_unit_Click()
    ChangeInGrid "Purch Unit", purchase_unit.Text
End Sub

Private Sub purchase_usage_conv_factor_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub purchase_usage_conv_factor_KeyPress(KeyAscii As Integer)
    If CheckNumericField(purchase_usage_conv_factor, KeyAscii, purchase_usage_conv_factor.SelStart, purchase_usage_conv_factor.SelLength, 5) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub start_date_Change()
    ChangeInGrid "Start Date", start_date
End Sub

Private Sub tech_desc_Change()
    ChangeInGrid "Tech Desc", tech_desc.Text
End Sub

Private Sub tech_desc_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub term_date_Change()
    ChangeInGrid "Term Date", term_date
End Sub

Private Sub traces_ind_Click()
    ChangeInGrid "TRACES", traces_ind.Value
End Sub

Private Sub traces_list_price_GotFocus()
    SendKeys ("{HOME}+{END}")
End Sub

Private Sub traces_list_price_KeyPress(KeyAscii As Integer)
    If CheckNumericField(traces_list_price, KeyAscii, traces_list_price.SelStart, traces_list_price.SelLength, 2) = False Then
        KeyAscii = 0
   End If
End Sub

Private Sub traces_list_price_Validate(Cancel As Boolean)
    If Not Len(traces_list_price.Text) = 0 Then
        CheckValueForNumber traces_list_price.Text, Cancel
    End If
End Sub

Private Sub pct_multiplier_Validate(Cancel As Boolean)
    CheckValueForNumber pct_multiplier.Text, Cancel
End Sub

Private Sub purchase_usage_conv_factor_Validate(Cancel As Boolean)
    CheckValueForNumber purchase_usage_conv_factor.Text, Cancel
End Sub

Private Sub mat_id_LostFocus()
    If mat_id.Locked = False And Not strLast_mat_id = mat_id.Text Then
        Dim strSelect As String
        Dim blnReturn As Boolean
        Dim ctr As Control
        Dim rec As New ADODB.RecordSet
        strLast_mat_id = mat_id.Text
        
        ' Check to see if the mat_id entered exists already
        strSelect = "Select * from Material where mat_id='" + Compress_String(mat_id) + "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rec)
        ' If it does, copy that data into fields
        If rec.RecordCount > 0 Then
            use_ind.Enabled = True    'If the material does exist, unlock the use indicator
            Dim fld As ADODB.Field
            For Each fld In rec.Fields
                m_rec.Fields(fld.Name).Value = fld.Value
            Next
            For Each ctr In Me.Controls
                If Left(ctr.Tag, 1) = "1" Then
                    ' Check type of control
                    If TypeOf ctr Is TextBox Then
                        If Not IsNull(rec.Fields(ctr.Name)) Then
                            ctr.Text = rec.Fields(ctr.Name)
                        Else
                            ctr.Text = ""
                        End If
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
            mat_skey.Text = rec.Fields("mat_skey")
            'adding this following 1 line, siva 8/17/99
            mat_last_update_id.Text = rec.Fields("last_update_id")
        Else
            use_ind.Value = 1   'New material, requires use ind
            use_ind.Enabled = False    'If the material does exist, unlock the use indicator
            ' Only blank out fields if we are not inserting
            If m_blnInsert = False Then
                For Each ctr In Me.Controls
                    If Left(ctr.Tag, 1) = "1" And Not ctr.Name = "mat_id" Then
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
            mat_skey.Text = ""
            mat_last_update_id.Text = 1
        End If
'        m_rec.Close
    End If
    If Me.ActiveControl.Name = "cmdUpdate" Then
        cmdUpdate_Click     'Finish update
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Button
    Dim blnPendingChange As Boolean
    If Me.ActiveControl.Name = "mat_id" Then mat_id_LostFocus
    ' Only go through this if the close wasn't invoked from code
    If Not UnloadMode = vbFormCode Then
        blnPendingChange = IsControlChanged(Me, m_rec)
        If blnPendingChange = True Then
            Button = MsgBox("Do you want to save your changes?", vbYesNoCancel)
            If Button = vbYes Then
                If alt_mat_id <> "" Then
                    If Invalid_mat_id_Format(alt_mat_id, "alt_mat_id", m_rec) = True Then
                        m_blnWereErrors = True
                    End If
                End If
                If Invalid_mat_id_Format(Compress_String(mat_id), "mat_id", m_rec) = True Then
                    m_blnWereErrors = True
                End If
                If m_blnWereErrors = False Then
                    cmdUpdate_Click
                End If
                ' If there were errors, cancel the close
                If m_blnWereErrors Then
                    Cancel = True
                End If
            ElseIf Button = vbNo Then
'                RestoreGridValues
            ElseIf Button = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        Else
'            RestoreGridValues
        End If
    End If
End Sub

Private Sub pct_multiplier_Change()
    If Trim(Me.pct_multiplier.Text) <> "." Then
        RecalcMaterialPrice
        RecalcTRACESPrice
        ChangeInGrid "Pct Mult", pct_multiplier
    End If
End Sub

Private Sub list_price_Change()
    If Trim(Me.list_price.Text) <> "." Then
        RecalcMaterialPrice
        traces_list_price = list_price
        ChangeInGrid "Price", list_price
    End If
End Sub

Private Sub traces_list_price_Change()
    If Trim(Me.traces_list_price.Text) <> "." Then
        RecalcTRACESPrice
        ChangeInGrid "TRACES Price", traces_list_price
    End If
End Sub

Private Sub purchase_usage_conv_factor_Change()
    If Trim(Me.purchase_usage_conv_factor.Text) <> "." Then
        RecalcMaterialPrice
        RecalcTRACESPrice
        ChangeInGrid "Conv", purchase_usage_conv_factor
    End If
End Sub

Private Sub RecalcMaterialPrice()
    If Len(list_price) > 0 And Len(pct_multiplier) > 0 And Len(purchase_usage_conv_factor) > 0 And Val(purchase_usage_conv_factor) > 0 Then
        list_price.Text = Change_Format_To_Numbers(list_price.Text, 2)
        purchase_usage_conv_factor.Text = Change_Format_To_Numbers(purchase_usage_conv_factor.Text, 5)
        pct_multiplier.Text = Change_Format_To_Numbers(pct_multiplier.Text, 5)
        ' added above 3 lines 8/12/99 siva
        material_price.Text = Format(Trim(str(CDbl(list_price.Text) * CDbl(pct_multiplier.Text) / CDbl(purchase_usage_conv_factor.Text) / 100)), "#,###,##0.00")
    End If
End Sub

Private Sub RecalcTRACESPrice()
    If Len(traces_list_price) > 0 And Len(pct_multiplier) > 0 And Len(purchase_usage_conv_factor) > 0 And Val(purchase_usage_conv_factor) > 0 Then
        traces_list_price.Text = Change_Format_To_Numbers(traces_list_price.Text, 2)
        traces_price.Text = Format(Trim(str(CDbl(traces_list_price.Text) * CDbl(pct_multiplier.Text) / CDbl(purchase_usage_conv_factor.Text) / 100)), "#,###,##0.00")
    End If
End Sub

Public Function Change_Format_With_Commas(myOldTxt As String, TotDecimals As Integer) As String
' this function changes the format from regular numbers to commas
' added on 8/24/99 siva
    Dim i As Integer, myNewTxt As String
    myNewTxt = "#,###,##0."
    For i = 1 To TotDecimals
        myNewTxt = myNewTxt + "0"
    Next
    Change_Format_With_Commas = Format(myOldTxt, myNewTxt)
End Function

Private Sub Store_Grid_Old_Values()
    Dim i As Integer
    ReDim tdbOldCols(tdbCols.count - 1)
    For i = 0 To tdbCols.count - 1
        tdbOldCols(i) = tdbCols.Item(i).Value
    Next
End Sub

Private Sub RestoreGridValues()
    ' this restores the grid back to its positioin if the user did not choose to save
    Dim i As Integer
    If m_blnInsert = False Then
        For i = 1 To tdbCols.count - 1
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

Private Sub update_ind_Click()
    ChangeInGrid "Update", update_ind.Value
End Sub

Private Sub update_status_code_Change()
    ChangeInGrid "Updt Stat", update_status_code
End Sub

Private Sub usage_unit_Click()
    ChangeInGrid "Use Unit", usage_unit.Text
End Sub

Private Sub use_ind_Click()
    ChangeInGrid "Use", use_ind.Value
End Sub

Private Sub wst_use_ind_Click()
    ChangeInGrid "West Use", wst_use_ind.Value
End Sub

Private Sub ChangeInGrid(myItem As String, mVariable As Variant)
    If m_blnInsert = False And UserIsEditing = True Then
        If tdbCols.Item("Material ID") = mat_id.Text Then
            tdbCols.Item(myItem).Value = mVariable
        End If
    End If
End Sub

Private Sub RebindTDBGridNow()
    Dim oldRow As Variant
    oldRow = myTDBGrid.Bookmark
    myTDBGrid.ReBind
    myTDBGrid.Bookmark = oldRow
End Sub

