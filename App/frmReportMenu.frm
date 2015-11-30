VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReportMenu 
   Caption         =   "Reports Menu"
   ClientHeight    =   6195
   ClientLeft      =   3900
   ClientTop       =   1995
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportMenu.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   9735
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build"
      Height          =   450
      Left            =   4560
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   1110
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportMenu.frx":0442
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportMenu.frx":082A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportMenu.frx":0C16
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportMenu.frx":0FDF
            Key             =   "Leaf"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   450
      Left            =   8400
      TabIndex        =   3
      Top             =   5520
      Width           =   1110
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Enabled         =   0   'False
      Height          =   450
      Left            =   7080
      TabIndex        =   2
      Top             =   5520
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Parameters"
      Height          =   4335
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9763
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblReportName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblReportName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   1410
   End
   Begin VB.Label lblTreeCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Report:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmReportMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iReportID As Long

Private m_rec As ADODB.RecordSet    ' Recordset to hold query results
Private m_parm_rec As New ADODB.RecordSet   ' Recordset to hold parameters
'
' Required so that State combo box can receive events to update City combo box
Private WithEvents cboStateCode As ComboBox
Attribute cboStateCode.VB_VarHelpID = -1
'

Private Function GetSqlCode() As String
    Dim strSELECT As String
    Dim sParamName As String
    Dim sValue As String

    If m_parm_rec.BOF And m_parm_rec.EOF Then
        ' NO PARAMETERS
        GetSqlCode = ""
        Exit Function
    End If

    strSELECT = ""
    m_parm_rec.MoveFirst
    Do While Not m_parm_rec.EOF
        sParamName = m_parm_rec.Fields("parameter_name")
        sValue = SQLFixString(m_parm_rec.Fields("new_value") & "")
        sValue = SQLChangeWildcard(sValue)
        If m_parm_rec.Fields("parameter_data_type") <> "number" Then
            sValue = "'" & sValue & "'"
        Else
            If sValue = "" Then sValue = 0
        End If
        strSELECT = strSELECT & "," & sParamName & "=" & sValue
        m_parm_rec.MoveNext
    Loop
    GetSqlCode = Mid(strSELECT, 2)

End Function

Private Function VerifyParameters() As Boolean
    Dim ctl As Control
    Dim sError As String
    Dim sParamName As String
    Dim sControlName As String
    
    If m_parm_rec.BOF And m_parm_rec.EOF Then
        ' NO PARAMETERS
        VerifyParameters = True
        Exit Function
    End If
    
    m_parm_rec.MoveFirst
    Do While Not m_parm_rec.EOF
        If m_parm_rec.Fields("parameter_visible") Then
            sParamName = m_parm_rec.Fields("parameter_name")
            sControlName = Mid(sParamName, 2)
            sControlName = Replace(sControlName, "_", "")
            Select Case m_parm_rec.Fields("parameter_appearance")
            Case 1
                sControlName = "txt" & sControlName
            Case 2
                sControlName = "chk" & sControlName
            Case Else
                sControlName = "cbo" & sControlName
            End Select
            Set ctl = Me.Controls(sControlName)
            If m_parm_rec.Fields("parameter_appearance") > 10 And m_parm_rec.Fields("parameter_data_type") = "number" Then
                ' Numeric Combo Box
                If ctl.ListIndex >= 0 Then
                    m_parm_rec.Fields("New_Value") = ctl.ItemData(ctl.ListIndex)
                Else
                    m_parm_rec.Fields("New_Value") = ""
                End If
                m_parm_rec.Update
            ElseIf m_parm_rec.Fields("parameter_appearance") = 2 Then
                ' CheckBox VALUE
                m_parm_rec.Fields("New_Value") = ctl.Value
                m_parm_rec.Update
            Else
                ' TEXT
                m_parm_rec.Fields("New_Value") = ctl.Text
                m_parm_rec.Update
            End If
            If m_parm_rec.Fields("parameter_required") Then
                If m_parm_rec.Fields("New_Value") = "" Then
                    sError = sError & "     • " & m_parm_rec.Fields("parameter_label") & " is a required field." & vbCrLf
                End If
            End If
        End If
        m_parm_rec.MoveNext
    Loop
    If sError = "" Then
        VerifyParameters = True
    Else
        VerifyParameters = False
        MsgBox "Please correct the following errors:" & vbCrLf & vbCrLf & sError, vbCritical
    End If
    
End Function

Private Sub LoadReportTree()
    Dim blnReturn As Boolean
    Dim strSELECT As String
    Dim rsTree As New ADODB.RecordSet
    Dim strCategory As String
    Dim strParent As String
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    lblReportName.Caption = ""
    With TreeView1
        ' Reset the Tree; Add Root Node [K0]
        .Nodes.Clear
        .Nodes.Add , , "K0", "Reports", "Closed", "Open"
        strParent = "K0"
        ' Load the Tree
        strSELECT = "SELECT * FROM REPORT_MASTER" & _
                    " WHERE (Report_Category IS NOT NULL)" & _
                    " ORDER BY Report_Category, Report_Name"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rsTree)
        If blnReturn = False Then
            MsgBox "An error occurred while searching:" & vbCrLf & g_objDAL.LastErrorDescription, vbCritical
            Screen.MousePointer = vbNormal
            Exit Sub
        Else
            strCategory = ""
            Do While Not rsTree.EOF
                If rsTree.Fields("report_category") <> strCategory Then
                    'Create a "folder" node for the new category
                    strParent = "K" & Replace(rsTree.Fields("report_category"), " ", "_")
                    .Nodes.Add "K0", tvwChild, strParent, rsTree.Fields("report_category"), "Closed", "Open"
                End If
                .Nodes.Add strParent, tvwChild, "K" & rsTree.Fields("report_id"), rsTree.Fields("report_name"), "Leaf"
                strCategory = rsTree.Fields("report_category") & ""
                rsTree.MoveNext
            Loop
        End If
        ' Finish Up
        rsTree.Close
        .Nodes("K0").Expanded = True
        .Nodes("K0").EnsureVisible
    End With
    Screen.MousePointer = vbNormal
    Exit Sub
    
Err_Handler:
    Screen.MousePointer = vbNormal
    MsgBox "An error occurred while loading Report Tree:" & vbCrLf & Err.Description, vbCritical
    Exit Sub
    
End Sub

Private Sub RemoveReportParameters()
'DELETE ALL CONTROLS CONTAINED IN FRAME1
    Dim ctl As Control
    
    On Error Resume Next
    For Each ctl In Me.Controls
        If ctl.Container.Name = "Frame1" Then
            Me.Controls.Remove ctl.Name
        End If
    Next
    DoEvents
    
End Sub

Private Function LoadReportParameters(iReportID As Long) As Boolean
'LOAD THE REPORT'S USER PARAMETERS INTO THE DETAIL FRAME
    Dim blnReturn As Boolean
    Dim strSELECT As String
    Dim ctl As Control
    Dim sParamName As String
    Dim sControlName As String
    Dim iTop As Long
    Dim iDlgUnit As Long
    
    On Error GoTo 0
    Screen.MousePointer = vbHourglass
    iDlgUnit = TreeView1.Left
    Status "Loading Report Options...."
    'Dynamically Add Controls to Frame1
    RemoveReportParameters
    Frame1.Visible = False
    strSELECT = "exec usp_select_report_parameters @report_id=" & iReportID
    ' Use DAL to perform select
    Set m_parm_rec = New ADODB.RecordSet
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, m_parm_rec)
    If blnReturn = False Then
        '8/16/2005 RTD - Added new DAL property LastErrorDescription
        Screen.MousePointer = vbNormal
        MsgBox "An error occurred while searching:" & vbCrLf & g_objDAL.LastErrorDescription, vbCritical
        Exit Function
    Else
        iTop = iDlgUnit * 2
        Do While Not m_parm_rec.EOF
            If m_parm_rec.Fields("parameter_visible") Then
                sParamName = m_parm_rec.Fields("parameter_name")
                sControlName = Mid(sParamName, 2)
                sControlName = Replace(sControlName, "_", "")
                ' Add Label Control
                Set ctl = Me.Controls.Add("VB.Label", "lbl" & sControlName, Frame1)
                ctl.Caption = m_parm_rec.Fields("parameter_label") & ":"
                ctl.Move iDlgUnit * 2, iTop + 40, 800, 285
                ctl.AutoSize = True
                ctl.Visible = True
                ' Add Control
                Select Case m_parm_rec.Fields("parameter_appearance")
                Case 1  ' text box
                    Set ctl = Me.Controls.Add("VB.TextBox", "txt" & sControlName, Frame1)
                    ctl.Move 2400, iTop, 1600, 285
                    ctl.Text = m_parm_rec.Fields("parameter_default_value") & ""
                Case 2  ' bit/boolean
                    Set ctl = Me.Controls.Add("VB.CheckBox", "chk" & sControlName, Frame1)
                    ctl.Move 2400, iTop - 80
                    ctl.Value = 0 & m_parm_rec.Fields("parameter_default_value")
                Case 3  ' date
                    Set ctl = Me.Controls.Add("MSComCtl2.DTPicker", "txt" & sControlName, Frame1)
                    ctl.Move 2400, iTop, 1600, 315
                    If IsNull(m_parm_rec.Fields("parameter_default_value")) Then
                        ctl.Value = Date
                    Else
                        ctl.Value = m_parm_rec.Fields("parameter_default_value")
                    End If
                Case 10 ' class system id
                    Set ctl = Me.Controls.Add("VB.ComboBox", "cbo" & sControlName, Frame1)
                    'ctl.Style = 2
                    ctl.AddItem "MF"
                    ctl.AddItem "R1"
                    ctl.AddItem "U2"
                    ctl.Move 2400, iTop, 1600
                    ctl.Text = m_parm_rec.Fields("parameter_default_value") & ""
                Case 11 ' geographic selection
                    Set ctl = Me.Controls.Add("VB.ComboBox", "cbo" & sControlName, Frame1)
                    ctl.AddItem "1 - Primary Cities (316)"
                    ctl.ItemData(ctl.NewIndex) = 1
                    ctl.AddItem "2 - National Average (30)"
                    ctl.ItemData(ctl.NewIndex) = 2
                    ctl.AddItem "3 - CCI Cities (727)"
                    ctl.ItemData(ctl.NewIndex) = 3
                    ctl.AddItem "4 - All Cities (731)"
                    ctl.ItemData(ctl.NewIndex) = 4
                    ctl.Move 2400, iTop, 2600
                    If IsNumeric(m_parm_rec.Fields("parameter_default_value")) Then
                        ctl.ListIndex = FindComboItemDataIndex(ctl, m_parm_rec.Fields("parameter_default_value"))
                    End If
                Case 12 ' MasterFormat
                    Set ctl = Me.Controls.Add("VB.ComboBox", "cbo" & sControlName, Frame1)
                    ctl.AddItem "1995"
                    ctl.ItemData(ctl.NewIndex) = 1995
                    ctl.AddItem "2004"
                    ctl.ItemData(ctl.NewIndex) = 2004
                    ctl.Move 2400, iTop, 1600
                    ctl.ListIndex = FindComboItemDataIndex(ctl, g_intMasterFormat)
                Case 20 ' Quarters
                    Set ctl = Me.Controls.Add("VB.ComboBox", "cbo" & sControlName, Frame1)
                    ctl.Move 2400, iTop, 1600
                    LoadComboBox ctl, m_parm_rec.Fields("parameter_sql"), False
                    ctl.Text = g_sQuarterID
                Case 22 ' State combo box
                    If (sControlName = "statecode") Then
                        Set cboStateCode = Me.Controls.Add("VB.ComboBox", "cbo" & sControlName, Frame1)
                        Set ctl = cboStateCode
                    ElseIf (sControlName = "state") Then
                        Set cboStateCode = Me.Controls.Add("VB.ComboBox", "cbo" & sControlName, Frame1)
                        Set ctl = cboStateCode
                    Else
                        Set ctl = Me.Controls.Add("VB.ComboBox", "cbo" & sControlName, Frame1)
                    End If
                    ctl.Move 2400, iTop, 1600
                    LoadComboBox ctl, m_parm_rec.Fields("parameter_sql"), (m_parm_rec.Fields("parameter_data_type") = "number")
                    If m_parm_rec.Fields("parameter_data_type") = "number" Then
                        ctl.ListIndex = FindComboItemDataIndex(ctl, 0 & m_parm_rec.Fields("parameter_default_value"))
                    Else
                        ctl.Text = m_parm_rec.Fields("parameter_default_value") & ""
                    End If
                Case 20, 21, 23 To 99   ' combo box
                    Set ctl = Me.Controls.Add("VB.ComboBox", "cbo" & sControlName, Frame1)
                    ctl.Move 2400, iTop, 2400
                    LoadComboBox ctl, m_parm_rec.Fields("parameter_sql"), (m_parm_rec.Fields("parameter_data_type") = "number")
                    If m_parm_rec.Fields("parameter_data_type") = "number" Then
                        ctl.ListIndex = FindComboItemDataIndex(ctl, 0 & m_parm_rec.Fields("parameter_default_value"))
                    Else
                        ctl.Text = m_parm_rec.Fields("parameter_default_value") & ""
                    End If
                End Select
                ctl.Tag = sParamName
                ctl.Visible = True
                iTop = iTop + (iDlgUnit * 2)
            End If
            m_parm_rec.MoveNext
        Loop
    End If
    Frame1.Visible = True
    DoEvents
    Status ""
    Screen.MousePointer = vbNormal
    Exit Function
    
Err_Handler:
    Status ""
    Screen.MousePointer = vbNormal
    MsgBox "An error occurred while loading Report Parameters:" & vbCrLf & Err.Description, vbCritical
    Exit Function
    
End Function

Private Sub LoadComboBox(ctl As ComboBox, strSELECT As String, blnUseItemData As Boolean)
    Dim blnReturn As Boolean
    Dim rec As New ADODB.RecordSet
    
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
    If blnReturn Then
        Do While Not rec.EOF
            ctl.AddItem rec.Fields(0)
            If blnUseItemData Then
                ctl.ItemData(ctl.NewIndex) = rec.Fields(1)
            End If
            rec.MoveNext
        Loop
    End If
    rec.Close
    
End Sub

Private Sub cboStateCode_Change()
    'cboStateCode_Click
End Sub

Private Sub cboStateCode_Click()
    Dim cmbCity As ComboBox
    
    On Error Resume Next
    Set cmbCity = Me.Controls("cbolocid")
    If Not (cmbCity Is Nothing) Then
        LoadCities cmbCity, cboStateCode.Text
    Else
        Set cmbCity = Me.Controls("cbocity")
        If Not (cmbCity Is Nothing) Then
            LoadCities cmbCity, cboStateCode.Text
        End If
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Set frmReportMenu = Nothing
End Sub

Private Sub cmdPreview_Click()
    PreviewReport
End Sub

Private Sub Form_Activate()
    ShowToolbarIcons True
End Sub

Private Sub Form_Deactivate()
    ShowToolbarIcons False
End Sub

Private Sub Form_Initialize()
    Status ("Loading Reports Menu...")
    Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Load()
    Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
    LoadReportTree
    Screen.MousePointer = vbDefault
    Status ("")
End Sub

Private Sub Form_Resize()
    Dim iDlgUnit As Long
    
    On Error Resume Next
    iDlgUnit = TreeView1.Left
    cmdClose.Left = Me.Width - cmdClose.Width - iDlgUnit * 2
    cmdClose.Top = Me.Height - cmdClose.Height - iDlgUnit * 3
    cmdPreview.Left = cmdClose.Left - cmdPreview.Width - iDlgUnit
    cmdPreview.Top = cmdClose.Top
    cmdBuild.Left = TreeView1.Left + TreeView1.Width + iDlgUnit
    cmdBuild.Top = cmdClose.Top
    Frame1.Left = TreeView1.Left + TreeView1.Width + iDlgUnit
    Frame1.Width = Me.Width - Frame1.Left - iDlgUnit * 2
    Frame1.Height = cmdClose.Top - Frame1.Top - iDlgUnit
    TreeView1.Height = cmdClose.Top + cmdClose.Height - TreeView1.Top
    
End Sub

Private Sub ShowToolbarIcons(bShowIcons As Boolean)
    
    On Error GoTo Err_Handler
    With fMainForm
        .tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Enabled = False
        .tbToolBar.Buttons.Item(tbrEXPORT).Visible = False
        .tbToolBar.Buttons.Item(tbrEXPORTDATA).Enabled = False
        .tbToolBar.Buttons.Item(tbrEXPORTDATA).Visible = False
        .tbToolBar.Buttons.Item(tbrEXPORTDATA + 1).Visible = False
        .mnuFilePageSetup.Enabled = bShowIcons
        .mnuFilePrint.Enabled = bShowIcons
        .mnuFileSaveAs.Enabled = False
        .mnuFilePrintPreview.Enabled = bShowIcons
    End With
    Exit Sub

Err_Handler:
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowToolbarIcons False
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'A TREE NODE HAS BEEN SELECTED, UPDATE DETAILS FRAME
    Dim iReportID As Long
    Dim sNodeKey As String
    
    sNodeKey = Node.Key
    If IsNumeric(Mid(sNodeKey, 2)) Then
        iReportID = Mid(sNodeKey, 2)
        m_iReportID = iReportID
        If (iReportID > 0) Then
            lblReportName.Caption = Node.Text
            LoadReportParameters iReportID
            cmdPreview.Enabled = True
        Else
            lblReportName.Caption = ""
            RemoveReportParameters
            cmdPreview.Enabled = False
        End If
    Else
        m_iReportID = 0
        lblReportName.Caption = ""
        RemoveReportParameters
        cmdPreview.Enabled = False
    End If
    
End Sub

Public Sub PreviewReport()
    Dim strSELECT As String
    Dim strParams As String
    Dim blnReturn As Boolean
    Dim rec As ADODB.RecordSet
    Dim fPreviewWindow As New frmReportPreview
    
    'On Error GoTo Err_Handler
    If m_iReportID = 0 Then
        Exit Sub
    End If
    If VerifyParameters Then
        Screen.MousePointer = vbHourglass
        strParams = GetSqlCode
        Set m_rec = New ADODB.RecordSet
        strSELECT = "SELECT * FROM REPORT_MASTER WHERE Report_ID = " & m_iReportID
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
        If blnReturn = False Then
            Screen.MousePointer = vbNormal
            MsgBox "An error occurred while searching:" & vbCrLf & g_objDAL.LastErrorDescription, vbCritical
            Exit Sub
        Else
            Status "Retrieving data from Database...."
            strSELECT = "EXEC " & rec.Fields("report_stored_proc") & " " & strParams
            blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, m_rec)
            If blnReturn Then
                If m_rec.EOF Then
                    ' USER SELECTIONS PRODUCED 0 RECORDS
                    Status ""
                    Screen.MousePointer = vbNormal
                    MsgBox "Your selections produced no results. Please check your criteria and try again.", vbExclamation
                    Set fPreviewWindow = Nothing
                    Exit Sub
                Else
                    ' DISPLAY REPORT PREVIEW WINDOW
                    Status "Opening " & rec.Fields("report_name") & " report...."
                    fPreviewWindow.Move START_LEFT, START_TOP, START_WIDTH, START_HEIGHT
                    fPreviewWindow.ReportName = rec.Fields("report_file_def_name")
                    fPreviewWindow.ReportFile = rec.Fields("report_file_name")
                    fPreviewWindow.ConnectString = g_cnShared
                    fPreviewWindow.RecordSet = m_rec
                    fPreviewWindow.RenderReport
                    fPreviewWindow.Show
                End If
            Else
                Status ""
                Screen.MousePointer = vbNormal
                MsgBox "An error occurred while retrieving data:" & vbCrLf & g_objDAL.LastErrorDescription, vbCritical
                Exit Sub
            End If
        End If
    End If
    Status ""
    Screen.MousePointer = vbDefault
    Exit Sub
    
Err_Handler:
    Screen.MousePointer = vbNormal
    MsgBox "An error occurred while loading the data and report preview:" & vbCrLf & Err.Description, vbCritical
    Exit Sub
    
End Sub

