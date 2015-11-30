VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmReportPreview 
   Caption         =   "Report Preview"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   5805
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Align           =   1  'Align Top
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      _cx             =   10239
      _cy             =   6800
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   19.2234848484848
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   1
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   7
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VSReport8LibCtl.VSReport VSReport1 
      Left            =   120
      Top             =   4080
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
End
Attribute VB_Name = "frmReportPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsData As ADODB.RecordSet
Private bShowNavBarPrinter As Boolean
Private bDirectToPrinter As Boolean
Private sConnectString As String
Private sRecordSource As String
Private sReportName As String
Private sReportFile As String
Private bAllowExport As Boolean
Private bAllowFax As Boolean
Private sReportFilename As String
Private sOpenEvent As String
Private oReportFileFormat As FileFormatSettings

Private Const MAX_PATH As Integer = 255
Private Declare Function apiGetTempDir Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Function GetTempFolder() As String
'RETURN THE CURRENT USER'S TEMP FILE FOLDER
'06/30/2005 RTD

    Dim strTempDir As String
    Dim lngx As Long
    
    strTempDir = String$(MAX_PATH, 0)
    lngx = apiGetTempDir(MAX_PATH, strTempDir)
    If lngx <> 0 Then
        strTempDir = Left$(strTempDir, lngx)
        If Right(strTempDir, 1) <> "\" Then
            strTempDir = strTempDir & "\"
        End If
    Else
        strTempDir = ""
    End If
    GetTempFolder = strTempDir
    
End Function

' ********************************************************************

Public Property Get ReportName() As String
Attribute ReportName.VB_Description = "Returns/sets the name of the report definition in the XML file"
    ReportName = sReportName
End Property
Public Property Let ReportName(NewValue As String)
    sReportName = NewValue
    Me.Caption = "Print Preview - " & sReportName
End Property

Public Property Get AllowExport() As Boolean
Attribute AllowExport.VB_Description = "Returns/Sets whether the Export toolbar button is available"
Attribute AllowExport.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowExport = bAllowExport
End Property
Public Property Let AllowExport(NewValue As Boolean)
    bAllowExport = NewValue
    With fMainForm
        .tbToolBar.Buttons.Item(tbrEXPORT).Enabled = bAllowExport
        .tbToolBar.Buttons.Item(tbrEXPORT).Visible = bAllowExport
        .mnuFileSaveAs.Enabled = bAllowExport
    End With
End Property

Public Property Get AllowFax() As Boolean
Attribute AllowFax.VB_Description = "Returns/Sets whether the Fax toolbar button is available"
Attribute AllowFax.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowFax = bAllowFax
End Property
Public Property Let AllowFax(NewValue As Boolean)
    bAllowFax = NewValue
    With fMainForm
        .tbToolBar.Buttons.Item(tbrFAX).Enabled = bAllowFax
        .tbToolBar.Buttons.Item(tbrFAX).Visible = bAllowFax
        .mnuFileFax.Enabled = bAllowFax
    End With
End Property

Public Property Get DirectToPrinter() As Boolean
    DirectToPrinter = bDirectToPrinter
End Property
Public Property Let DirectToPrinter(NewValue As Boolean)
    bDirectToPrinter = NewValue
    VSPrinter1.Preview = Not bDirectToPrinter
End Property

Public Property Get ShowPrintIconInNavBar() As Boolean
Attribute ShowPrintIconInNavBar.VB_Description = "Returns/sets whether the print icon is visible in the child form's client area"
    DirectToPrinter = bShowNavBarPrinter
End Property
Public Property Let ShowPrintIconInNavBar(NewValue As Boolean)
    bShowNavBarPrinter = NewValue
    If bShowNavBarPrinter Then
        VSPrinter1.NavBar = vpnbTopPrint
    Else
        VSPrinter1.NavBar = vpnbTop
    End If
End Property

Public Property Get ExportFormat() As FileFormatSettings
    ExportFormat = oReportFileFormat
End Property
Public Property Let ExportFormat(NewValue As FileFormatSettings)
    oReportFileFormat = NewValue
End Property

Public Property Get ExportFilename() As String
    ExportFilename = sReportFilename
End Property
Public Property Let ExportFilename(NewValue As String)
    sReportFilename = NewValue
End Property

Public Property Let ReportFile(XMLFilename As String)
Attribute ReportFile.VB_Description = "Returns/sets the XML file that contains the report definition"
Attribute ReportFile.VB_ProcData.VB_Invoke_PropertyPut = ";Data"
    Dim sFile As String
    
    On Error GoTo Err_Handler
    sFile = XMLFilename
    If InStr(sFile, "\") = 0 Then
        sFile = App.Path & "\" & sFile
    End If
    sReportFile = sFile
    VSReport1.Load sFile, sReportName
    Exit Property

Err_Handler:
    MsgBox Err.Description, vbCritical

End Property

Public Property Get RecordSource() As String
    RecordSource = sRecordSource
End Property
Public Property Let RecordSource(NewValue As String)
    sRecordSource = NewValue
    VSReport1.DataSource.ConnectionString = sConnectString
    VSReport1.DataSource.RecordSource = sRecordSource
End Property

Public Property Let RecordSet(rs As ADODB.RecordSet)
    'MODIFIED 9/25/2005 RTD - USE A CLONE OF THE ORIGINAL RECORDSET
    '                         TO PREVENT REPORT GROUPING/SORTING FROM
    '                         MODIFYING THE ORIGINAL RECORDSET'S SORTING
    Set rsData = rs.Clone
    VSReport1.DataSource.RecordSet = rsData
End Property

Public Property Get RecordCount() As Long
    If Not rsData Is Nothing Then
        RecordCount = rsData.RecordCount
    Else
        RecordCount = -1
    End If
End Property

Public Property Get ConnectString() As String
Attribute ConnectString.VB_Description = "Returns/Sets the report's database connection string"
Attribute ConnectString.VB_ProcData.VB_Invoke_Property = ";Data"
    ConnectString = sConnectString
End Property
Public Property Let ConnectString(NewValue As String)
    sConnectString = NewValue
    VSReport1.DataSource.ConnectionString = sConnectString
End Property

Public Property Get OpenEvent() As String
    OpenEvent = sOpenEvent
End Property
Public Property Let OpenEvent(NewValue As String)
    sOpenEvent = NewValue
End Property

' ********************************************************************

Public Sub RenderReport()
    'RENDER REPORT ONTO PRINTER PREVIEW
    On Error Resume Next
    SetOpenEvent
    VSReport1.Render VSPrinter1
End Sub

Public Sub PrintReport()
    'PRINT REPORT TO PRINTER
    On Error Resume Next
    VSPrinter1.PrintDoc
End Sub

Public Sub PreviewReport()
    'REFRESH PREVIEW DATA
    On Error Resume Next
    
    If Not rsData Is Nothing Then
        VSReport1.DataSource.RecordSet = rsData
    End If
    VSReport1.Render VSPrinter1
    
End Sub

Public Sub ExportReport()
    'EXPORT REPORT TO FILE
    Dim sExt As String
    On Error GoTo Err_Handler
    
    If sReportFilename = "" Then
        ' Get Filename from user
        With CommonDialog1
            SetFilters CommonDialog1
            .CancelError = False
            .FileName = sReportName
            .DialogTitle = "Save Report as..."
            .ShowSave
        End With
        If CommonDialog1.FileName <> "" Then
            sReportFilename = CommonDialog1.FileName
            sExt = "." & CommonDialog1.DefaultExt
            If Right(sReportFilename, Len(sExt)) <> sExt Then
                sReportFilename = sReportFilename & sExt
            End If
            Status "Exporting Report..."
            If Not rsData Is Nothing Then
                VSReport1.DataSource.RecordSet = rsData
            End If
            Screen.MousePointer = vbHourglass
            VSReport1.RenderToFile sReportFilename, oReportFileFormat
            Screen.MousePointer = vbDefault
            sReportFilename = ""
            Status ""
        End If
    Else
        Status "Exporting Report..."
        If Not rsData Is Nothing Then
            VSReport1.DataSource.RecordSet = rsData
        End If
        Screen.MousePointer = vbHourglass
        VSReport1.RenderToFile sReportFilename, oReportFileFormat
        Screen.MousePointer = vbDefault
        Status ""
    End If
    Exit Sub
    
Err_Handler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub

End Sub

Public Sub FaxReport()
'SEND REPORT TO FAX RECIPIENT
    Dim sReportFilename As String
    Dim outlookHelper As CFaxAndEMail
    Dim recipient As String
    Dim recipientFaxNumber As String
    Dim strContactId As String
    Dim oContact As CInfoSource

    MsgBox "This feature is not yet implemented."
    Exit Sub
    
    On Error GoTo Err_Handler
    
    ' GET RECIPIENT
    strContactId = ""
    Set oContact = New CInfoSource
    oContact.GetContactInfo (strContactId)
    recipient = oContact.ContactName
    recipientFaxNumber = oContact.faxNumber
    
    ' CREATE TEMPORARY PDF FILE
    Status "Creating Report..."
    sReportFilename = GetTempFolder & sReportName & ".pdf"
    If Not rsData Is Nothing Then
        VSReport1.DataSource.RecordSet = rsData
    End If
    Screen.MousePointer = vbHourglass
    VSReport1.RenderToFile sReportFilename, vsrPDF
    Screen.MousePointer = vbDefault
    
    ' FAX REPORT
    Status "Faxing Report..."
    Screen.MousePointer = vbHourglass
    Set outlookHelper = New CFaxAndEMail
    Call outlookHelper.SendFax(recipient, oContact.CompanyName, recipientFaxNumber, sReportName, sReportFilename)
    Set outlookHelper = Nothing
    Screen.MousePointer = vbDefault
    MsgBox ("Your report was faxed to " & recipient & " at " & recipientFaxNumber)
    Status ""
    Exit Sub

Err_Handler:
    Status ""
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub

End Sub

' ********************************************************************

Private Sub SetFilters(ByRef cd As CommonDialog)
    Dim sExt As String
    
    Select Case oReportFileFormat
    Case vsrHTML, vsrHTMLDrillDown, vsrHTMLPaged
        sExt = "htm"
    Case vsrPDF, vsrPDFUncompressed
        sExt = "pdf"
    Case vsrRTF
        sExt = "rtf"
    Case vsrText
        sExt = "txt"
    Case Else
        sExt = ""
    End Select
    If sExt <> "" Then
        cd.Filter = UCase(sExt) & " Files|*." & sExt
        cd.DefaultExt = sExt
    End If

End Sub

Private Sub Form_Activate()
    
    ShowToolbarIcons True

End Sub

Private Sub Form_Deactivate()
    
    ShowToolbarIcons False
    
End Sub

Private Sub Form_Initialize()
    
    ShowPrintIconInNavBar = True
    bDirectToPrinter = False
    bAllowExport = True
    bAllowFax = False
    oReportFileFormat = vsrPDF
    sReportFile = ""
    sReportName = "Print Preview"
    sReportFilename = ""
    sConnectString = g_cnShared.ConnectionString
    sOpenEvent = ""
    VSReport1.DataSource.ConnectionString = sConnectString
    
End Sub

Private Sub Form_Load()
    Me.Caption = "Print Preview - " & sReportName
End Sub

Private Sub Form_Resize()
    VSPrinter1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowToolbarIcons False
    Set frmReportPreview = Nothing
End Sub

Private Sub SetOpenEvent()
    ' Add a custom Open Event to the Report
    Dim sCurrentEvent As String
    
    If sOpenEvent <> "" Then
        sCurrentEvent = VSReport1.OnOpen
        If sCurrentEvent = "" Then
            VSReport1.OnOpen = sOpenEvent
        Else
            VSReport1.OnOpen = sOpenEvent & vbCrLf & vbCrLf & sCurrentEvent
        End If
    End If

End Sub

Private Sub ShowToolbarIcons(bShowIcons As Boolean)
    
    On Error GoTo Err_Handler
    With fMainForm
        .tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Enabled = bAllowExport And bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Visible = bAllowExport And bShowIcons
        .tbToolBar.Buttons.Item(tbrFAX).Enabled = bAllowFax And bShowIcons
        .tbToolBar.Buttons.Item(tbrFAX).Visible = bAllowFax And bShowIcons
        .mnuFilePageSetup.Enabled = bShowIcons
        .mnuFilePrint.Enabled = bShowIcons
        .mnuFileSaveAs.Enabled = bAllowExport And bShowIcons
        .mnuFileFax.Enabled = bAllowFax And bShowIcons
        .mnuFilePrintPreview.Enabled = False
    End With
    Exit Sub

Err_Handler:
    Exit Sub
    
End Sub

