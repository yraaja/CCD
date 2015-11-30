VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLongDescRpt 
   Caption         =   "Long Description Report Reports"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   Icon            =   "frmLongDescRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   7245
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Align           =   1  'Align Top
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7245
      _cx             =   12779
      _cy             =   8996
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
      MarginBottom    =   1080
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
      Zoom            =   27.0833333333333
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
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VSReport8LibCtl.VSReport rptLongDescription 
      Left            =   120
      Top             =   5280
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
Attribute VB_Name = "frmLongDescRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsData As ADODB.RecordSet
Private sConnectString As String
Private sReportName As String
Private sReportFile As String
Private sReportFilename As String
Private oReportFileFormat As FileFormatSettings

' ********************************************************************
Public Property Get ReportName() As String
    ReportName = sReportName
End Property
Public Property Let ReportName(NewValue As String)
    sReportName = NewValue
    Me.Caption = sReportName
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

Public Property Let ReportFile(XMLFilename As String, ReportName As String)
    Dim sFile As String
    
    sFile = XMLFilename
    If InStr(sFile, "\") = 0 Then
        sFile = App.Path & "\" & sFile
    End If
    sReportFile = sFile
    rptLongDescription.Load sFile, ReportName
    
End Property

Public Property Let RecordSet(rs As ADODB.RecordSet)
    Set rsData = rs
End Property

Public Property Get RecordCount() As Long
    If Not rsData Is Nothing Then
        RecordCount = rsData.RecordCount
    Else
        RecordCount = -1
    End If
End Property

Public Property Get ConnectString() As String
    ConnectString = sConnectString
End Property
Public Property Let ConnectString(NewValue As String)
    sConnectString = NewValue
End Property

' ********************************************************************

Public Sub RenderReport()
    'RENDER REPORT ONTO PRINTER PREVIEW
    On Error Resume Next
    rptLongDescription.Render VSPrinter1
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
        rptLongDescription.DataSource.RecordSet = rsData
    End If
    rptLongDescription.Render VSPrinter1
    
End Sub

Public Sub ExportReport()
    'EXPORT REPORT TO FILE
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
            Status "Exporting Report..."
            If Not rsData Is Nothing Then
                rptLongDescription.DataSource.RecordSet = rsData
            End If
            rptLongDescription.RenderToFile sReportFilename, oReportFileFormat
            sReportFilename = ""
            Status ""
        End If
    Else
        Status "Exporting Report..."
        If Not rsData Is Nothing Then
            rptLongDescription.DataSource.RecordSet = rsData
        End If
        rptLongDescription.RenderToFile sReportFilename, oReportFileFormat
        Status ""
    End If
    Exit Sub
    
Err_Handler:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub

End Sub

Public Sub LoadReport(ByVal rs As ADODB.RecordSet)
    Dim sEvent As String
    
    Me.ReportName = "Long Description Report"
    Set rsData = rs
    With rptLongDescription
        .Load App.Path & "\rptLongDescription.xml", "Long Description Report"
        .DataSource.ConnectionString = g_cnShared.ConnectionString
        .DataSource.RecordSet = rs
        'sEvent = "suppress_addressee = " & CStr(iSuppressAddressee) & vbCrLf _
            & "print_price = " & CStr(iPrintPrice)
        'rptLongDescription.OnOpen = sEvent
        '.DoEvents = True
        '.Render VSPrinter1
    End With

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

    oReportFileFormat = vsrPDF
    sReportName = "Print Preview"
    sReportFilename = ""
    sConnectString = g_cnShared.ConnectionString
    rptLongDescription.DataSource.ConnectionString = sConnectString
    
End Sub

Private Sub Form_Load()
    Me.Caption = sReportName
End Sub

Private Sub Form_Resize()
    VSPrinter1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowToolbarIcons False
End Sub

Private Sub ShowToolbarIcons(bShowIcons As Boolean)
    
    With fMainForm
        .tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Visible = bShowIcons
        .mnuFilePageSetup.Enabled = bShowIcons
        .mnuFilePrint.Enabled = bShowIcons
        .mnuFileSaveAs.Enabled = bShowIcons
        .mnuFilePrintPreview.Enabled = False
    End With
    
End Sub

