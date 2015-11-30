VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmRFQRpt 
   Caption         =   "Request for Quote Report"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   Icon            =   "frmRFQRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   8220
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Align           =   1  'Align Top
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8220
      _cx             =   14499
      _cy             =   8705
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
      Zoom            =   26.0416666666667
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
   Begin VSReport8LibCtl.VSReport rptRequestForQuote 
      Left            =   120
      Top             =   5040
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
Attribute VB_Name = "frmRFQRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsData As ADODB.RecordSet
Private ReportType As String
Private sReportName As String
Private sReportFilename As String
Private oReportFileFormat As FileFormatSettings
Private strContactId As String
Private oContact As CInfoSource

Private Const MAX_PATH As Integer = 255
Private Declare Function apiGetTempDir Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Sub LoadReport(ByVal rs As ADODB.RecordSet, _
                        iSuppressAddressee As Integer, _
                        iPrintPrice As Integer, strInfoSourceContactId As String)
    Dim sEvent As String
    Dim user As IADsUser
    
    'MsgBox ("Email Address: " & GetEmailFromUserNamesTab)
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    'GET user info from LDAP
    '
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Set user = Me.UserInfo(strUserName)
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    
    strContactId = strInfoSourceContactId
    sReportName = "Request for Quote"
    Set rsData = rs
    
'    rsData("email") = "xxxxx@xxxxx.com"
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::
    'Set EMAIL ADDRESS (from Active Directory Query)
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::
    rsData("email") = user.EMailAddress
    rsData("user_phone") = FormatPhoneNumber(user.TelephoneNumber)
    
    With rptRequestForQuote
        .Load App.Path & "\rptRequestForQuote.xml", "Request for Quote"
        .DataSource.ConnectionString = g_cnShared.ConnectionString
'        rsData(0) = "xxxx"      'Checking to see if I can update field(s) in the recordset after the fact!
        '::::::::::::::::::: DISPLAY ALL FIELD NAMES :::::::::::::::::::::::::::
'        Dim i As Integer
'        For i = 0 To rsData.Fields.Count - 1
'            Debug.Print (rsData.Fields(i).Name)
'        Next
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        .DataSource.RecordSet = rsData
        sEvent = "suppress_addressee = " & CStr(iSuppressAddressee) & vbCrLf & _
                 "print_price = " & CStr(iPrintPrice)
        .OnOpen = sEvent
        '.DoEvents = True
        '.Render VSPrinter1
    End With

End Sub

Public Function GetEmailFromUserNamesTab() As String
Dim strSelect As String
Dim blnReturn As Boolean
Dim rs As ADODB.RecordSet

On Error GoTo ERRLBL

strSelect = "SELECT * FROM USER_NAMES WHERE user_id='" & strUserName & "'"
 ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSelect, rs)
    If blnReturn = False Then
        MsgBox "An error occurred while retrieving user email address."
       
        'lblRowCount.Caption = "0 rows returned."
        GoTo ERRLBL
    End If
    GetEmailFromUserNamesTab = Trim(rs("user_email"))
    Exit Function
ERRLBL:
    MsgBox ("(Error)GetEmailFromUserNamesTab: " + Err.Description)
    
End Function
Public Function UserInfo(LoginName As String) As IADsUser
'PURPOSE: Display information that is available in
'the Active Directory about a given user

'PARAMETER: Login Name for user

'RETURNS: String with selected information about
'user, or empty string if there is no such
'login on the current domain

'REQUIRES: Windows 2000 ADSI, LDAP Provider
'Proper Security Credentials.

'EXAMPLE: msgbox UserInfo("Administrator")

Dim conn As New ADODB.Connection
Dim rs As ADODB.RecordSet
Dim oRoot As IADs
Dim oDomain As IADs
Dim sBase As String
Dim sFilter As String
Dim sDomain As String

Dim sAttribs As String
Dim sDepth As String
Dim sQuery As String
Dim sAns As String

Dim user As IADsUser



On Error GoTo ErrHandler:

Set UserInfo = Nothing

'Get user Using LDAP/ADO.  There is an easier way
'to bind to a user object using the WinNT provider,
'but this way is a better for educational purposes
Set oRoot = GetObject("LDAP://rootDSE")             'Original code

'work in the default domain
sDomain = oRoot.Get("defaultNamingContext")        'Original code
'sDomain = "GC://DC=b2b,DC=regn,DC=net"            'Taken from Chenchen's setting

Set oDomain = GetObject("LDAP://" & sDomain)       'Original code
sBase = "<" & oDomain.ADsPath & ">"

'Only get user name requested
sFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName={0})(name=" _
  & LoginName & "))"  'this kinda works
  
sFilter = "(&(objectCategory=person)(sAMAccountName=" & LoginName & "))"
 
sAttribs = "sAMAccountName, cn, ADsPath, objectClass"
sDepth = "subTree"

sQuery = sBase & ";" & sFilter & ";" & sAttribs & ";" & sDepth

conn.Open _
  "Data Source=Active Directory Provider;Provider=ADsDSOObject"

Set rs = conn.Execute(sQuery)


If Not rs.EOF Then
Dim i As Integer
For i = 0 To rs.RecordCount - 1
'    Set user = GetObject(rs("adsPath"))
Set user = GetObject(rs(2).value)   'AdsPath
Set UserInfo = user
'
'    With user

    'if the attribute is not stored in AD,
    'an error will occur.  Therefore, this
    'will return data only from populated attributes
   
    
    On Error Resume Next

        Debug.Print rs(0)
        Debug.Print rs(1)
        Debug.Print rs(2)
        Debug.Print "---------------------"
        Debug.Print user.FirstName
        Debug.Print user.EMailAddress
        Debug.Print user.faxNumber
        Debug.Print user.TelephoneNumber
'    End If

rs.MoveNext
Next
End If

ErrHandler:

On Error Resume Next
If Not rs Is Nothing Then
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End If

If Not conn Is Nothing Then
    If conn.State <> 0 Then conn.Close
    Set conn = Nothing
End If

Set oRoot = Nothing
Set oDomain = Nothing
End Function
Public Sub RenderReport()
    On Error Resume Next
    rptRequestForQuote.Render VSPrinter1
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
        rptRequestForQuote.DataSource.RecordSet = rsData
    End If
    rptRequestForQuote.Render VSPrinter1
    
End Sub

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

Private Sub ShowToolbarIcons(bShowIcons As Boolean)
    
    On Error GoTo Err_Handler
    With fMainForm
        .tbToolBar.Buttons.Item(tbrPRINT).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPRINT).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrPREVIEW).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrEXPORT).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrFAX).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrFAX).Visible = bShowIcons
        .tbToolBar.Buttons.Item(tbrEMAIL).Enabled = bShowIcons
        .tbToolBar.Buttons.Item(tbrEMAIL).Visible = bShowIcons
        .mnuFilePageSetup.Enabled = bShowIcons
        .mnuFilePrint.Enabled = bShowIcons
        .mnuFileSaveAs.Enabled = bShowIcons
        .mnuFileFax.Enabled = bShowIcons
        .mnuFilePrintPreview.Enabled = False
    End With
    Exit Sub

Err_Handler:
    Exit Sub
    
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
End Sub

Private Sub Form_Resize()
    VSPrinter1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowToolbarIcons False
    Set frmReportPreview = Nothing
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
                rptRequestForQuote.DataSource.RecordSet = rsData
            End If
            Screen.MousePointer = vbHourglass
            rptRequestForQuote.RenderToFile sReportFilename, oReportFileFormat
            Screen.MousePointer = vbDefault
            sReportFilename = ""
            Status ""
        End If
    Else
        Status "Exporting Report..."
        If Not rsData Is Nothing Then
            rptRequestForQuote.DataSource.RecordSet = rsData
        End If
        Screen.MousePointer = vbHourglass
        rptRequestForQuote.RenderToFile sReportFilename, oReportFileFormat
        Screen.MousePointer = vbDefault
        Status ""
    End If
    Exit Sub
    
Err_Handler:
    Screen.MousePointer = vbDefault
    MsgBox "Error during ExportReport():" & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub

End Sub

Public Sub FaxReport()
    '*************************************************
    ' Save the current Report to Hard Drive. Then
    ' fax the report as an attachment
    '*************************************************
    
    Dim outlookHelper As CFaxAndEMail
    Dim recipient As String
    Dim recipientFaxNumber As String
    
    On Error GoTo Err_Handler
    
    If Not SaveReportToDisk Then
        MsgBox "The application failed to create a fax image of the report. Send Fax aborted.", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    If oContact Is Nothing Then
        Set oContact = New CInfoSource
        oContact.GetContactInfo (strContactId)
    End If
    
    recipient = oContact.ContactName
    recipientFaxNumber = oContact.faxNumber
    
    
    'Send Fax
    Status "Sending Fax ..."
    Screen.MousePointer = vbHourglass
    
    Set outlookHelper = New CFaxAndEMail
    Call outlookHelper.SendFax(recipient, oContact.CompanyName, recipientFaxNumber, sReportName, sReportFilename)
    Set outlookHelper = Nothing
    WriteToLog "FAX", oContact.ContactID, oContact.CompanyName, oContact.ContactName, oContact.ContactLastName, oContact.ContactFirstName, recipientFaxNumber
    Screen.MousePointer = vbDefault
    MsgBox "Your report was faxed to " & recipient & " at " & recipientFaxNumber & ".", vbInformation
    Status ""
    Exit Sub
    
Err_Handler:
    Set outlookHelper = Nothing
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
        
End Sub

Public Sub MailReport()
'*************************************************
' Save the current Report to Hard Drive. Then
' email the report as an attachment
'*************************************************
    
    Dim outlookHelper As CFaxAndEMail
    Dim recipientEmail As String
    
    On Error GoTo Err_Handler
    
    If Not SaveReportToDisk Then
        MsgBox "The application failed to create a PDF image of the report. Send Mail aborted.", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    'Get the EMail address for the supplier
    If oContact Is Nothing Then
        Set oContact = New CInfoSource
        oContact.GetContactInfo (strContactId)
    End If
    
    'Send Email
    recipientEmail = oContact.EMailAddress
    If (recipientEmail = "") Then
        'ask for email address
        recipientEmail = InputBox("Please enter the EMail Address:")
    End If
    
    If (recipientEmail = "") Then
        Exit Sub  'User decided not to provide the email addess
    End If
    
    Status "Sending Email ..."
    Screen.MousePointer = vbHourglass
    
    Set outlookHelper = New CFaxAndEMail
    Call outlookHelper.SendEMail(recipientEmail, sReportName, "", sReportFilename)
    Set outlookHelper = Nothing
    WriteToLog "EMAIL", oContact.ContactID, oContact.CompanyName, oContact.ContactName, oContact.ContactLastName, oContact.ContactFirstName, recipientEmail
    Screen.MousePointer = vbDefault
    MsgBox "Your report was sent to " & recipientEmail & ".", vbInformation
    Status ""
    Exit Sub
    
Err_Handler:
    Set outlookHelper = Nothing
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub

Private Function SaveReportToDisk() As Boolean
'*************************************************
' Save the current Report to Hard Drive. This will
' be used for sending the report in EMail and Fax
' as an attachment
'*************************************************

    On Error GoTo Err_Handler
    
    'Save the Report as PDF
    'MODIFIED TO USE USER TEMP FOLDER INSTEAD OF "C:\"
    'NEEDED TO CORRECT PERMISSIONS PROBLEM ON LIMITED USER ACCOUNTS
    '06/30/2005 RTD
    sReportFilename = GetTempFolder & sReportName & ".pdf"
    
    Status "Exporting Report..."
    If Not rsData Is Nothing Then
        rptRequestForQuote.DataSource.RecordSet = rsData
    End If
    Screen.MousePointer = vbHourglass
    rptRequestForQuote.RenderToFile sReportFilename, oReportFileFormat
    Screen.MousePointer = vbDefault
    Status ""
    SaveReportToDisk = True
    Exit Function
    
Err_Handler:
    Screen.MousePointer = vbDefault
    SaveReportToDisk = False
    MsgBox "SaveReportToDisk() unable to create PDF export:" & Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Function
End Function

Private Sub GetInformationByContactId(sInputContactId As String)
'*************************************************
' Get the email address for the contact Id
'*************************************************

    On Error GoTo Err_Handler
    
    'Save the Report as PDF
    'MODIFIED TO USE USER TEMP FOLDER INSTEAD OF "C:\"
    '06/30/2005 RTD
    sReportFilename = GetTempFolder & sReportName & ".pdf"
    
    Status "Exporting Report..."
    If Not rsData Is Nothing Then
        rptRequestForQuote.DataSource.RecordSet = rsData
    End If
    Screen.MousePointer = vbHourglass
    rptRequestForQuote.RenderToFile sReportFilename, oReportFileFormat
    Screen.MousePointer = vbDefault
    Status ""
    Exit Sub
    
Err_Handler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub

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

