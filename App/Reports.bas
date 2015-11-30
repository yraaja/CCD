Attribute VB_Name = "Reports"
Option Explicit


'''/// <modulename> Reports</modulename>
'''/// <functionname>General (Main) </functionname>
'''
'''/// <summary>
'''/// Provides common subs/functions:
'''1.  CCI_ADMIN()
'''Sets the stored procedure to be invoked for "CCI ADMIN" menu selection and passes to …. ExecStoredProcSelectedQuarter… next
'''2.  ExecStoredProcSelectedQuarter(stored procedure name)
'''
'''This function is used extensively in support of CCI functionality where there may/may not be a supporting windows/form driver
'''The following is a list of CCI (and other) functions supported here:
'''
'''"   Extend Max Term Date        (sp_labor_extend_term_date_rlh)
'''"   Clone CCI Material Prices       (sp_clone_pub_cci_material_price)
'''"   Clone Material/Equipment Prices         (sp_clone_pub_cci_mat_equ_price_rlh)
'''"   Generate Masterformat Exception Report (with PERCENTAGE settings)                                   (sp_pub_cci_index_exception_report_qn_with_fuel_rlh)
'''"   Labor Out-of-Date (Report)                                   (sp_report_labor_rate_out_of_date)
'''"   Dollar Listing                 (sp_report_pub_cci_index_masterformat_rpt_with_fuel)
'''"   Extend Quarter Date         (sp_extend_quarter_dates_rlh)
'''"   Publish Quarterly Labor Rates           (sp_update_published_cci_labor_rate_rlh)
'''"   Build Labor Rates Grid "Report" table   (sp_build_cci_labor_rates_allcities_grid)
'''"   Generate Masterformat Index     (sp_rollup_pub_cci_index_masterformat_with_fuel_rlh)
'''"   Generate UNIFORMAT Index        (sp_rollup_pub_cci_index_uniformat_with_fuel_rlh)
'''
'''"   CCI_Export_Excel            vanilla generic export to excel from a recordset
'''"   CCI_Mail_Export         export to excel a "mailing" report
'''"   CCI_Out_Of_Date_Export  export to excel from Labor-Out-of_date report
'''"   GetQuarterID            Populates a listbox with quarter-ids over several years
'''"   SaveFile                Save to either "csv" or "excel" to a filename parameter from a recordset parameter
'''
'''HELPER CLASS: N/A
'''########################################################################
'''NOTE:  MOST of the stored procedures are executed from "CCI ADMIN" menu items found on the menu tree on the left nav bar and a couple of "LABOR" menu items.
'''#######################################################################
'''
'''/// </summary>
'''/// <seealso>frmNavTree.frm</seealso>
'''
'''/// <datastruct>m_rec</datastruct>
'''/// <storedprocedurename> sp_labor_extend_term_date_rlh </storedprocedurename>
'''/// <storedprocedurename> sp_clone_pub_cci_material_price </storedprocedurename>
'''/// <storedprocedurename> sp_clone_pub_cci_mat_equ_price_rlh </storedprocedurename>
'''/// <storedprocedurename> sp_clone_pub_cci_material_price </storedprocedurename>
'''/// <storedprocedurename> sp_pub_cci_index_exception_report_qn_with_fuel_rlh </storedprocedurename>
'''/// <storedprocedurename> sp_report_labor_rate_out_of_date </storedprocedurename>
'''/// <storedprocedurename> sp_extend_quarter_dates_rlh
'''</storedprocedurename>
'''/// <storedprocedurename> sp_update_published_cci_labor_rate_rlh </storedprocedurename>
'''/// <storedprocedurename> sp_build_cci_labor_rates_allcities_grid </storedprocedurename>
'''/// <storedprocedurename> sp_rollup_pub_cci_index_masterformat_with_fuel_rlh </storedprocedurename>
'''/// <storedprocedurename> sp_rollup_pub_cci_index_uniformat_with_fuel_rlh </storedprocedurename>
'''/// <storedprocedurename> </storedprocedurename>
'''
'''
'''/// <returns>N/A</returns>
'''/// <exception>Always trap with an accompanying message box</exception>
'''/// <example>
'''/// <code>
'''exec usp_select_unit_cost_ext_rlh2 @start_unit_cost_id = '030100000000', @end_unit_cost_id = '030499999999', @alt_unit_cost_id = '', @tech_desc = '', @master_format=2004
'''/// </code>
'''/// <code>sp_report_labor_rate_out_of_date_rlh '2006Q3'
'''///</code>
'''/// <code>
'''sp_labor_extend_term_date_rlh '1/1/2010'
'''///</code>
'''/// <code>
'''exec sp_extend_quarter_dates_rlh 73,'2011-01-01 00:00:00.000','2011-03-31 00:00:00.000','2011Q1'
'''///</code>
'''/// <code>
'''exec SP_CLONE_PUB_CCI_MAT_EQU_PRICE_RLH '2006Q3'
'''///</code>
'''/// <code>
'''exec sp_update_published_cci_labor_rate_rlh '2006Q3'
'''///</code>
'''/// <code>
'''exec [dbo].[sp_rollup_pub_cci_index_masterformat_with_fuel_rlh] '2006Q3'
'''///</code>
'''/// <code>
'''exec SP_ROLLUP_PUB_CCI_INDEX_UNIFORMAT_WITH_FUEL_rlh '2008Q1' ///</code>
'''/// <code>
'''exec sp_pub_cci_index_exception_report_qn_with_fuel_rlh  '2009Q3',0.00,1.10,0.90,1.10,0.90,1.10
'''///</code>
'''/// <code>
'''exec sp_select_cci_mailing_list @quarter =  '2008Q1'
'''///</code>
'''/// <code>
'''
'''///</code>
'''
'''
'''///</example>
'''///<permission>Public</Permission>
'''///<dependson>
'''///</dependson>



Dim m_data_rec As New ADODB.RecordSet       'rlh 03/04/2010
Dim m_rec As New ADODB.RecordSet
Public PCT1 As String
Public PCT2 As String
Public PCT3 As String
Public PCT4 As String
Public PCT5 As String
Public PCT6 As String

Public Function CCI_Mail_Export()
    Dim ListQuarters As New cdlgLstSel
    Dim dlgStatus As New dlgStatus
    Dim sQtr As String
    Dim rec As New ADODB.RecordSet
    Dim blnResult As Boolean
    Dim fldMailList As ADODB.Field
    Dim sRec As String
    Dim strSelect As String
    Dim sFile As String
    Dim sRecordCount As String
    Dim lngRecordCount As Long
    Dim lBufferLen As Long
    Dim lpBuffer As Long
    Dim iResult As Integer
    Dim strTempPath As String
    Dim lngTempPath As Long
    
    On Error GoTo Error_Processing
    
    'MODIFIED 8/17/2005 RTD - EXPORT FILE SHOULD DEFAULT TO 'MY DOCUMENTS' FOLDER
    'TO CONFORM TO WINDOWS STANDARDS AND TO PREVENT SECURITY PROBLEMS
    'WITH LIMITED USER ACCOUNTS

    'sFile = CurDir() & "\"
    sFile = GetSpecialFolderLocation(CSIDL_PERSONAL)
    sQtr = GetQuarterID(ListQuarters, "Current Quarter:")

    If sQtr <> "-1" Then
        strSelect = "sp_select_cci_mailing_list @quarter = '" + Right(sQtr, 1) + "'"
        'rec.MaxRecords = MAX_RECORDS ' Set the maximum number to bring back
        ' Use g_objDAL to perform select
        blnResult = g_objDAL.GetRecordset(vbNullString, strSelect, rec)
        If blnResult = False Then
            MsgBox "An error occurred while searching:" & vbCrLf & g_objDAL.LastErrorDescription, vbCritical
            Exit Function
        End If
        If rec.RecordCount > 0 Then
            sRecordCount = CStr(rec.RecordCount)
            'sFile = sFile & "RSM CCI Address " & sQtr & ".csv"
            sFile = sFile & "RSM CCI Address " & sQtr & ".xls"
            If sFile = "" Then
                Exit Function
            End If
            'NOTE FOR FUTURE RELEASE
            'THIS SHOULD BE REPLACED BY THE WINDOWS SAVE/OPEN COMMON DIALOG
            'Stop
            sFile = InputBox(sRecordCount & " records will be exported." & vbCrLf & "Enter the path and file name to export:", "Mail Export File Name", sFile)
        Else
            MsgBox "There are no records to export. Please make another selection.", , "Mailing Label Export"
        End If
        If sFile <> "" And rec.RecordCount > 0 Then
       
            lngRecordCount = 0
            dlgStatus.prgOverallStatus.Min = 1
            dlgStatus.prgOverallStatus.Max = rec.RecordCount
            dlgStatus.lblTitle = "Mailing Label File Export"
            dlgStatus.Caption = "Mailing Label Export"
            dlgStatus.lblText = sRecordCount & " records will be exported to " & sFile & ". Press Esc or Cancel to abort."
            dlgStatus.cmdOK.Visible = False
            'dlgStatus.Visible = True
            DoEvents
                 
'            Call ExportData(rec)       'rlh 03/04/2010
            Set m_data_rec = rec.Clone  'rlh 03/04/2010
            
            Call SaveFile(sFile, rec)   'rlh 03/04/2010

            'dlgStatus.timUnload.Interval = 3000 'Show message for 30 seconds.
        End If
        rec.Close
        Set rec = Nothing
    End If
    
Exit_Sub:
    Screen.MousePointer = vbNormal
    Exit Function
    
Error_Processing:
    If Err = 91 Then
        Resume Next
    Else
        MsgBox Error$
        rec.Close
        Set rec = Nothing
        Resume Exit_Sub
        Resume 0
    End If

End Function

Public Function CCI_Out_Of_Date_Export(rec As ADODB.RecordSet)
'rlh 03/07/2010 CCD/CCI for Labor Out-of-date support

    Dim ListQuarters As New cdlgLstSel
    Dim dlgStatus As New dlgStatus
    Dim sQtr As String
    Dim blnResult As Boolean
    Dim fldMailList As ADODB.Field
    Dim sRec As String
    Dim strSelect As String
    Dim sFile As String
    Dim sRecordCount As String
    Dim lngRecordCount As Long
    Dim lBufferLen As Long
    Dim lpBuffer As Long
    Dim iResult As Integer
    Dim strTempPath As String
    Dim lngTempPath As Long
    
    On Error GoTo Error_Processing
    
    If DEBUGON Then Stop
    
    'MODIFIED 8/17/2005 RTD - EXPORT FILE SHOULD DEFAULT TO 'MY DOCUMENTS' FOLDER
    'TO CONFORM TO WINDOWS STANDARDS AND TO PREVENT SECURITY PROBLEMS
    'WITH LIMITED USER ACCOUNTS

    'sFile = CurDir() & "\"
    sFile = GetSpecialFolderLocation(CSIDL_PERSONAL)
    
        If rec.RecordCount > 0 Then
            sRecordCount = CStr(rec.RecordCount)
            'sFile = sFile & "RSM CCI Address " & sQtr & ".csv"
            sFile = sFile & "Labor Out-of-Date Report " & sQtr & ".xls"
            If sFile = "" Then
                Exit Function
            End If
            'NOTE FOR FUTURE RELEASE
            'THIS SHOULD BE REPLACED BY THE WINDOWS SAVE/OPEN COMMON DIALOG
            If DEBUGON Then
                Stop
            End If
            sFile = InputBox(sRecordCount & " records will be exported." & vbCrLf & "Enter the path and file name to export:", "Mail Export File Name", sFile)
        Else
            MsgBox "There are no records to export. Please make another selection.", , "Mailing Label Export"
        End If
        If sFile <> "" And rec.RecordCount > 0 Then

            DoEvents
                
            Set m_data_rec = rec.Clone  'rlh 03/04/2010
            
            Call SaveFile(sFile, rec)   'rlh 03/04/2010
           
        End If
        rec.Close
        Set rec = Nothing
    'End If
    
Exit_Sub:
    Screen.MousePointer = vbNormal
    Exit Function
    
Error_Processing:
    If Err = 91 Then
        Resume Next
    Else
        MsgBox Error$
        rec.Close
        Set rec = Nothing
        Resume Exit_Sub
        Resume 0
    End If

End Function

Public Function CCI_Export_Excel(rec As ADODB.RecordSet, excelFile As String)
'rlh 03/07/2010 CCD/CCI for Labor Out-of-date support

    Dim ListQuarters As New cdlgLstSel
    Dim dlgStatus As New dlgStatus
    Dim sQtr As String
    Dim blnResult As Boolean
    Dim fldMailList As ADODB.Field
    Dim sRec As String
    Dim strSelect As String
    Dim sFile As String
    Dim sRecordCount As String
    Dim lngRecordCount As Long
    Dim lBufferLen As Long
    Dim lpBuffer As Long
    Dim iResult As Integer
    Dim strTempPath As String
    Dim lngTempPath As Long
    
    On Error GoTo Error_Processing
    
    'MODIFIED 8/17/2005 RTD - EXPORT FILE SHOULD DEFAULT TO 'MY DOCUMENTS' FOLDER
    'TO CONFORM TO WINDOWS STANDARDS AND TO PREVENT SECURITY PROBLEMS
    'WITH LIMITED USER ACCOUNTS
    
    If DEBUGON Then Stop
    
    'sFile = CurDir() & "\"
    sFile = GetSpecialFolderLocation(CSIDL_PERSONAL)
    
        If rec.RecordCount > 0 Then
            sRecordCount = CStr(rec.RecordCount)
            'sFile = sFile & "RSM CCI Address " & sQtr & ".csv"
            sFile = sFile & excelFile & sQtr & ".xls"
            If sFile = "" Then
                Exit Function
            End If
            'NOTE FOR FUTURE RELEASE
            'THIS SHOULD BE REPLACED BY THE WINDOWS SAVE/OPEN COMMON DIALOG
            If DEBUGON Then
                Stop
            End If
            sFile = InputBox(sRecordCount & " records will be exported." & vbCrLf & "Enter the path and file name to export:", "Mail Export File Name", sFile)
        Else
            MsgBox "There are no records to export. Please make another selection.", , "Mailing Label Export"
        End If
        If sFile <> "" And rec.RecordCount > 0 Then

            DoEvents
                
            Set m_data_rec = rec.Clone  'rlh 03/04/2010
            
            Call SaveFile(sFile, rec)   'rlh 03/04/2010
           
        End If
        rec.Close
        Set rec = Nothing
    'End If
    
Exit_Sub:
    Screen.MousePointer = vbNormal
    Exit Function
    
Error_Processing:
    If Err = 91 Then
        Resume Next
    Else
        MsgBox Error$
        rec.Close
        Set rec = Nothing
        Resume Exit_Sub
        Resume 0
    End If

End Function
Public Sub ExportData(m_rec As ADODB.RecordSet)
'
'    If m_rec.RecordCount > 0 Then
'        Dim fExport As New frmExport
'
'        'fExport.SetRow TDBGrid, m_rec
'        fExport.title = "CCI Mailing Labels"
'        fExport.Show
'    Else
'        MsgBox "Please choose or search for a CCI Mailing Label.", vbInformation + vbOKOnly
'    End If
    
End Sub
Public Sub SaveFile(sFileName As String, m_rec As ADODB.RecordSet)    'rlh 03/04/2010
'    Dim sFileName As String
    Dim sFileExtension As String
    
    
    On Error Resume Next
    
'    With CommonDialog1
'        .Filter = "CSV File (*.csv)|*.csv|Excel File (*.xls)|*.xls|HTML File (*.htm)|*.htm|XML File (*.xml)|*.xml"
'        .FilterIndex = 2
'        .DefaultExt = ".htm"
'        .CancelError = True
'        .FileName = m_Title
'        .DialogTitle = "Export data to file..."
'        .ShowSave
'    End With
'    If (CommonDialog1.FileName <> "") And (Err.Number = 0) Then
'        sFileName = CommonDialog1.FileName
        sFileExtension = GetFileExtension(sFileName)
        Select Case sFileExtension
        Case "csv"
            If ExportToCsv(sFileName, m_rec) Then
                MsgBox "Data was successfully exported to file:" & vbCrLf & sFileName, vbInformation
            End If
'        Case "htm", "html"
'            If ExportToHtml(sFilename) Then
'                MsgBox "Data was successfully exported to file:" & vbCrLf & sFilename, vbInformation
'            End If
'        Case "xml"
'            If ExportToXml(sFilename) Then
'                MsgBox "Data was successfully exported to file:" & vbCrLf & sFilename, vbInformation
'            End If
        Case "xls"
            If ExportToExcel(sFileName) Then
                MsgBox "Data was successfully exported to file:" & vbCrLf & sFileName, vbInformation
            End If
        Case ""
            
        Case Else
            MsgBox "Unsupported export type '" & sFileExtension & "'", vbExclamation
        End Select
'    End If
    
End Sub
Private Function GetFileExtension(sFileName As String) As String 'rlh 03/04/2010
    Dim P As Long
    
    P = InStrRev(sFileName, ".")
    If P > 0 Then
        GetFileExtension = Mid(sFileName, P + 1)
    Else
        GetFileExtension = ""
    End If

End Function
Private Function ExportToCsv(sFileName As String, m_rec As ADODB.RecordSet) As Boolean
    'rlh 03/04/2010  Copied from frmExport to handle "CCI Mailing Labels"
    Dim f As Long
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    f = FreeFile
    Open sFileName For Output As #f
    m_rec.MoveFirst
'''    Do While Not m_rec.EOF
'''        'If m_rec.Fields("Export") Then
'''            Write #f, m_rec.Fields("Name").Value;
'''        'End If
'''        m_rec.MoveNext
'''    Loop
    Write #f,
    m_data_rec.MoveFirst
    Do While Not m_data_rec.EOF
        m_rec.MoveFirst
        Do While Not m_rec.EOF
            'If m_rec.Fields("Export") Then
                Write #f, m_data_rec.Fields(m_rec.Fields("Field").Value).Value & "";
           ' End If
            m_rec.MoveNext
        Loop
        Write #f,
        m_data_rec.MoveNext
    Loop
    Close #f
    Screen.MousePointer = vbDefault
    ExportToCsv = True
    Exit Function
    
Err_Handler:
    ExportToCsv = False
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation
    Exit Function
    
End Function

Public Function ExportToExcel(sFileName As String) As Boolean
    Dim f As Long
    Dim i As Long           'rlh 03/05/2010
    Dim tmpZip As String    'rlh 03/23/2010
    
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    f = FreeFile
    Open sFileName For Output As #f
    
    ' Build the header row 1st
   m_data_rec.MoveFirst
   
    For i = 0 To m_data_rec.Fields.Count - 1
        Print #f, m_data_rec.Fields(i).Name & vbTab;
    Next
    
    Print #f,
    m_data_rec.MoveFirst
    Do While Not m_data_rec.EOF
        
        For i = 0 To m_data_rec.Fields.Count - 1
            Select Case m_data_rec.Fields(i).Name
            Case "zip"  'retain all 5 digits
                'If Mid(m_data_rec(i), 1, 1) = 0 Then Stop
                tmpZip = Format(m_data_rec(i), "0000#")
                tmpZip = "=" & """" & m_data_rec(i) & """"  'rlh this is the trick, right here!!!
                Print #f, tmpZip & "" & vbTab;
            Case Else
                Print #f, m_data_rec.Fields(i).Value & "" & vbTab;
            End Select
        Next
          
        Print #f,
        m_data_rec.MoveNext
    Loop
    Close #f

    Screen.MousePointer = vbDefault
    ExportToExcel = True
    Exit Function
    
Err_Handler:
    ExportToExcel = False
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation
    Exit Function
    
End Function

Public Function CloneCCIEquipmentRate()
    ExecStoredProcSelectedQuarter "sp_clone_pub_cci_equipment_rate"
End Function

Public Function CCI_Admin(Index As Integer)
    Dim sStoredProcName As String
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    '   CHECK ROLE OF USER FOR AUTHORIZATION
    '
    ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    Select Case Index
    Case 2, 3, 4, 5, 6, 7, 8, 9, 10
'    Case 10 'labor rate out-of-date
'    Case 11 'Labor Extend Term Date
    Case Else
    '     0 - Extend Term Dates
    '     1 - Clone Quarterly Mat/Equ Prices
    '     11 - Labor Extend Term Date
        If (modCommon.CheckUserAuth()) Then
             
            Else
                MsgBox ("Unauthorized access.  Please see CCD/CCI Administrator")
                Exit Function
        End If
    End Select
    
    If DEBUGON Then Stop  'rlh 03/09/2010
    
    If Index = 9 Then
        CCI_Mail_Export
    Else
        Select Case Index
        Case 0 'EXTEND QUARTER DATES         rlh 02/26/2010
            If DEBUGON Then Stop
             sStoredProcName = "sp_extend_quarter_dates_rlh"
        Case 1
             sStoredProcName = "SP_CLONE_PUB_CCI_MAT_EQU_PRICE_RLH"
        Case 2
'             sStoredProcName = "SP_UPDATE_PUB_CCI_LABOR_RATE"
             sStoredProcName = "SP_UPDATE_PUBLISHED_CCI_LABOR_RATE_RLH"  'rlh 02/27/2010
        Case 3
             'rlh - changed to "with_fuel" as per ksr
             sStoredProcName = "SP_REPORT_PUB_CCI_MATERIAL_EQUIPMENT_WITH_FUEL"
        Case 4  'Generate Masterformat Index
            If DEBUGON Then Stop
             sStoredProcName = "SP_ROLLUP_PUB_CCI_INDEX_MASTERFORMAT_WITH_FUEL_rlh"
        Case 5 'Generate UNIFORMAT Index
            If DEBUGON Then Stop
'             sStoredProcName = "SP_ROLLUP_PUB_CCI_INDEX_UNIFORMAT"
                sStoredProcName = "sp_rollup_pub_cci_index_uniformat_with_fuel_rlh"     'rlh 02/27/2010
        Case 6  'Generate RESIDENTIAL Index
'             sStoredProcName = "SP_ROLLUP_PUB_CCI_INDEX_RESIDENTIAL"
                sStoredProcName = "SP_EXTRACT_PUB_CCI_INDEX_RESIDENTIAL_rlh"  'rlh 02/27/2010
        Case 7
             sStoredProcName = "SP_REPORT_PUB_CCI_CSIFORMAT_SUM_MAP_RPT"
        Case 8
            If DEBUGON Then Stop
             sStoredProcName = LCase("sp_pub_cci_index_exception_report_Qn_with_fuel_rlh")
        Case 10 'LABOR RATE OUT-OF-DATE          rlh 02/26/2010
            If DEBUGON Then Stop
             sStoredProcName = "sp_report_labor_rate_out_of_date_rlh"
        Case 11 'LABOR EXTEND TERM DATE          rlh 03/07/2010
             sStoredProcName = "sp_labor_extend_term_date_rlh"
        
        End Select
        
        If DEBUGON Then Stop
        
        ExecStoredProcSelectedQuarter sStoredProcName
    End If
End Function

' 10/04/2005 RTD - FUNCTION NOW RETURNS TRUE IF SUCCESSFULL
'                - (PREVIOUS RETURNED NOTHING)
Public Function ExecStoredProcSelectedQuarter(sProcedureName As String) As Boolean
    Dim ListQuarters As New cdlgLstSel
    Dim strSelectedQtr As String
    Dim strUpdate As String
    Dim strError As String
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rsTemp As RecordSet
    Dim sMsg As String
    Dim iResult As Integer
    Dim lCount As Long
    Dim strParameters As String     ' 10/04/2005 RTD - FOR ADDITIONAL SP PARAMETERS
    Dim sFile As String             ' 03/07/2010 rlh - CCD/CCI
    Dim ans As Variant              ' rlh 03/09/2010
    
    On Error GoTo Error_Processing
    If DEBUGON Then
        Stop  'rlh 02/27/2010
    End If
    Screen.MousePointer = vbHourglass
    strParameters = ""
    
    'HACK JOB!!! (rlh) 03/07/2010
    Select Case sProcedureName
    Case "sp_labor_extend_term_date_rlh", "sp_labor_extend_term_date", "sp_extend_quarter_dates_rlh", _
         "sp_build_cci_labor_rates_allcities_grid", "sp_build_cci_labor_rates_anytown_grid"
        If DEBUGON Then Stop        'rlh 03/09/2010
      strSelectedQtr = ""
    Case Else
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::
        ' QUARTER DATE DROP-DOWN IS DISPLAYED HERE
        ':::::::::::::::::::::::::::::::::::::::::::::::::::::
        
        strSelectedQtr = GetQuarterID(ListQuarters, "Current Quarter:")
    End Select
    
    If strSelectedQtr <> "-1" Then
        iResult = vbOK
        Select Case LCase(sProcedureName)
        Case "sp_clone_pub_cci_material_price"
            strSelect = "SELECT COUNT(*) count_rcds FROM PUBLISHED_CCI_MATERIAL_PRICE mp inner join quarter_date qd on mp.qtr_dt_skey = qd.qtr_dt_skey WHERE QUARTER_ID = '" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If blnReturn = False Then
                MsgBox "An error occurred retrieving the count of Material Price records:" & vbCrLf & g_objDAL.LastErrorDescription
            Else
                If Not (rsTemp.EOF And rsTemp.BOF) Then
                    lCount = rsTemp.Fields("count_rcds")
                    If lCount > 0 Then
                        sMsg = "PUBLISHED_CCI_MATERIAL_PRICE already has " & CStr(lCount) & " rows created for quarter_id = '" & strSelectedQtr & "' - Are you sure you want to delete and re-create these rows?"
                        iResult = MsgBox(sMsg, vbOKCancel, "Cloning Date Selection")
                    End If
                End If
            End If
            rsTemp.Close
            Set rsTemp = Nothing
        Case "sp_clone_pub_cci_mat_equ_price_rlh"
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            ' CLONE Quarterly Material/Equipment Prices
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            ans = MsgBox("New data will be created for Quarter Id: " & strSelectedQtr, vbOKCancel, "Clone to New Quarter")
            Select Case ans
            Case vbOK
            Case vbCancel
                Exit Function
            End Select
            
            If DEBUGON Then Stop
            
            '1st check  (does current quarter exist ?)
            
            'strSELECT = "SELECT COUNT(*) AS mycount FROM PUBLISHED_CCI_MATERIAL_PRICE"
            strSelect = "SELECT  qtr_dt_skey FROM QUARTER_DATE WHERE quarter_id ='" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp.EOF = False Then
            Else
                ans = MsgBox("Current Quarter ID, " & strSelectedQtr & ", Does Not Exist- Extend Quarter Date Table")
                Exit Function
            End If
            
            '2nd check  (does previous quarter exist)
            strSelect = "SELECT  qtr_dt_skey FROM QUARTER_DATE WHERE qtr_dt_skey =" & (rsTemp(0) - 1)
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp.EOF = False Then
            Else
                ans = MsgBox("Previous quarter data does not exist", vbOKCancel, "Overwrite Check")
                Select Case ans
                    Case vbOK
                    Case vbCancel
                    Exit Function
                End Select
            End If
            
            ' last check
            strSelect = "SELECT COUNT(*) count_rcds FROM PUBLISHED_CCI_MATERIAL_PRICE er inner join quarter_date qd on er.qtr_dt_skey = qd.qtr_dt_skey WHERE QUARTER_ID = '" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If blnReturn = False Then
                MsgBox "An error occurred retrieving the count of CCI Material records:" & vbCrLf & g_objDAL.LastErrorDescription
            Else
                If Not (rsTemp.EOF And rsTemp.BOF) Then
                    lCount = rsTemp.Fields("count_rcds")
                    If lCount > 0 Then
                        sMsg = "PUBLISHED_CCI_MATERIAL_PRICE already has " & CStr(lCount) & " rows created for quarter_id = '" & strSelectedQtr & "' - Are you sure you want to delete and re-create these rows?"
                        iResult = MsgBox(sMsg, vbOKCancel, "Cloning Date Selection")
                    Else
                        Call MsgBox("No data exists for Quarter Id: " & strSelectedQtr & " Please choose another quarter date", vbOKOnly, "Invalid Quarter Id Selection")
                        Exit Function
                    End If
  
                End If
            End If
            rsTemp.Close
            Set rsTemp = Nothing
        
        Case "sp_pub_cci_index_exception_report_qn_with_fuel_rlh"
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
        'Generate Masterformat Index-Exception Report (with PERCENTAGE settings)
        '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            If DEBUGON Then Stop
            '::::::::::::::::::::::::::::::::::::::::::::::::::
            'STEP #1  (CREATE the report table)
            '::::::::::::::::::::::::::::::::::::::::::::::::::
            'Stop
            strUpdate = "exec  sp_report_pub_cci_index_masterformat_rpt_WITH_FUEL" & " '" & strSelectedQtr & "'"
            strUpdate = strUpdate & ",1"  'qtr_ind
            g_cnShared.Execute strUpdate    'Allow long-running procedures
            
            '::::::::::::::::::::::::::::::::::::::::::::::::::
            'STEP #2  RUN THE REPORT
            '::::::::::::::::::::::::::::::::::::::::::::::::::
            'Prompt user for user inputs (percentage pairs)
            Screen.MousePointer = vbNormal
            dlgCCIADminMFExpRpt.Show (vbModal)
            If Mode Then
            
            
                '::::::::::::::::::::::::::::::::::::::::::::::::::
                '::
                ':: 1ST BUILD THE REPORT TABLE
                '::
                '::::::::::::::::::::::::::::::::::::::::::::::::::
                
                strUpdate = "exec SP_REPORT_PUB_CCI_INDEX_MASTERFORMAT_RPT_WITH_FUEL_RLH " & " '" & strSelectedQtr & "'"
                
                Screen.MousePointer = vbHourglass
                g_cnSharedLong.Execute strUpdate    'Allow long-running procedures
                Screen.MousePointer = vbNormal
                
                '::::::::::::::::::::::::::::::::::::::::
                
                strParameters = PCT1 & ","
                strParameters = strParameters & PCT2 & ","
                strParameters = strParameters & PCT3 & ","
                strParameters = strParameters & PCT4 & ","
                strParameters = strParameters & PCT5 & ","
                strParameters = strParameters & PCT6
                
                strUpdate = "exec " & sProcedureName & " '" & strSelectedQtr & "'," & strParameters
                
                strSelect = strUpdate
                blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
                If blnReturn = False Then
                    MsgBox "An error occurred retrieving Masterformate index-exception report data:" & vbCrLf & g_objDAL.LastErrorDescription
                Else
                    If Not (rsTemp.EOF And rsTemp.BOF) Then
                        lCount = rsTemp.RecordCount
                        If lCount > 0 Then
                            sMsg = "Do You want to save results/report to excel spreadsheet?"
                            iResult = MsgBox(sMsg, vbOKCancel, "Save To Excel Prompt")
                            Select Case iResult
                            Case vbOK
                                 Call Reports.CCI_Export_Excel(rsTemp, "MasterFormatExceptions")
                            Case vbCancel
                            
                            End Select
                        End If
                    Else
                        MsgBox ("No records returned from Master index-exception report")
                        
                    End If
                End If
            End If
            Exit Function
        Case "sp_report_pub_cci_index_masterformat_rpt", "sp_report_pub_cci_index_masterformat_rpt_with_fuel_rlh"
            ' 10/04/2005 RTD - ADDED BECAUSE STORED PROC REQUIRES AN ADDITIONAL PARAMETER
            '                  THAT WAS NOT SUPPLIED BY THE CCD APP, CAUSING AN ERROR.
            '                  @QTR_IND=1 == USE CURRENT QUARTER INDEXES
            '
'            strParameters = ",3"        'rlh changed qtr_in to "3" as per ksr instructions
            
            If DEBUGON Then Stop
        
            strUpdate = "exec " & sProcedureName & " '" + strSelectedQtr + "'" + strParameters
            strSelect = strUpdate
            
        Case "sp_report_cci_detail_rpt"
            ' 10/04/2005 RTD - ADDED BECAUSE STORED PROC REQUIRES AN ADDITIONAL PARAMETER
            '                  @CITY_IND=4 == ALL CCI CITIES
            strParameters = ",4"
'        Case "sp_report_pub_cci_csiformat_sum_map_with_fuel"
'            ' 10/04/2005 RTD - ADDED BECAUSE STORED PROC REQUIRES AN ADDITIONAL PARAMETER
'            '                  @CITY_IND=4 == ALL CCI CITIES
'            strParameters = ",3"
        
        Case "sp_report_labor_rate_out_of_date", _
        "sp_report_labor_rate_out_of_date_rlh"
        
        If DEBUGON Then Stop
        
            strUpdate = "exec " & sProcedureName & " '" + strSelectedQtr + "'" + strParameters
            strSelect = strUpdate
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If blnReturn = False Then
                MsgBox "An error occurred retrieving labor rate out-of-date:" & vbCrLf & g_objDAL.LastErrorDescription
            Else
                If Not (rsTemp.EOF And rsTemp.BOF) Then
                    lCount = rsTemp.RecordCount
                    If lCount > 0 Then
                        sMsg = "Do You want to save results/report to excel spreadsheet?"
                        iResult = MsgBox(sMsg, vbOKCancel, "Save To Excel Prompt")
                        Select Case iResult
                        Case vbOK
                             Call Reports.CCI_Out_Of_Date_Export(rsTemp)
                        Case vbCancel
                        
                        End Select
                    End If
                Else
                    MsgBox ("No Labor data found out of date")
                End If
            End If
'            rsTemp.Close
'            Set rsTemp = Nothing
            Exit Function
        'Case "sp_report_pub_cci_index_masterformat_rpt_rlh", "sp_report_pub_cci_index_masterformat_rpt"
        Case "sp_report_pub_cci_index_masterformat_rpt_with_fuel"
            If DEBUGON Then Stop
            strParameters = ",3"    'Historical Index
        Case "sp_labor_extend_term_date", "sp_labor_extend_term_date_rlh"
            
            If DEBUGON Then Stop
            
            Dim tmpstr As String
            tmpstr = CStr(Month(Date)) & "/" & CStr(Day(Date)) & "/" & CStr(Year(Date))
            tmpstr = InputBox("Please add new max term date as follows mm/dd/yyyy", "New Max Term Date")
            
            strParameters = CDate(tmpstr)
            strUpdate = "exec " & sProcedureName & " '" & strParameters & "'"
            
            ans = MsgBox("Continue update of Labor Term Dates for another 365 days using the following max labor term date: " & strParameters & "?", vbOKCancel, "Labor Term Date Extension")
            Select Case ans
            Case vbOK
                '##################################################
                'Run Extend MAX Term Date stored procedure
                '##################################################
                
                 g_cnShared.Execute strUpdate
                If blnReturn = False Then
                    MsgBox "An error occurred attempting to reset MAX Term Dates:" & vbCrLf & g_objDAL.LastErrorDescription & vbCrLf & " Please run Out-Of-Date Report"
                Else
                    If Not (rsTemp.EOF And rsTemp.BOF) Then
                         MsgBox ("MAX Term dates were extended using max term date of: " & strParameters)
                    End If
                End If
                Exit Function
            Case vbCancel
                Exit Function
            End Select
            
             g_cnShared.Execute strUpdate    'Allow long-running procedures
              MsgBox ("Applicable Labor Term Dates have been extended another 365 days -Completed") 'rlh 02/27/2010
            iResult = vbCancel
        Case "sp_extend_quarter_dates_rlh"
            
            If DEBUGON Then Stop
            
            '-- Get row associated with highest skey
            strSelect = "SELECT * FROM quarter_date WHERE qtr_dt_skey = (Select max(qtr_dt_skey) from quarter_date)"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If blnReturn = False Then
                MsgBox "An error occurred retrieving the (last) Quarter Date Record:" & vbCrLf & g_objDAL.LastErrorDescription
            Else
                Dim startdate As String
                Dim enddate As String
                Dim tmpstrtdt_yyyy As String
                Dim tmpstrtdt_mm As String
                Dim tmpstrtdt_dd As String
                
                Dim nxtyyyy As String
                Dim nxtmm As String
                Dim nxtdd As String
                
                Dim parm1 As String
                Dim parm2 As String
                Dim parm3 As String
                Dim parm4 As String
                
                Dim tmpenddt_yyyy As String
                Dim tmpenddt_mm As String
                Dim tmpenddt_dd As String
                Dim qtrid As String
                
                Dim tmpenddt As String
                Dim tmptimestr As String
                Dim ary() As String
                Dim ary2() As String
                
                ary = Split(rsTemp("start_date"), "/")
                tmpstrtdt_yyyy = ary(2)
                tmpstrtdt_mm = ary(0)
                
                
                
                ary2 = Split(rsTemp("term_date"), "/")
                tmpenddt_yyyy = ary2(2)
                tmpenddt_mm = ary2(0)
                
                tmptimestr = "00:00:00.000"
                
                parm1 = CStr(CInt(rsTemp("qtr_dt_skey")) + 1)
                parm2 = tmpstrtdt_yyyy
                parm3 = tmpenddt_yyyy
                
                nxtyyyy = tmpstrtdt_yyyy
                'SET STARTING MONTH
                Select Case tmpstrtdt_mm
                    Case "01", "1"
                        nxtmm = "04"
                    Case "04", "4"
                        nxtmm = "07"
                    Case "07", "7"
                        nxtmm = "10"
                    Case "10"
                        nxtmm = "01"
                        nxtyyyy = CStr(CInt(tmpstrtdt_yyyy) + 1)
                    Case Else
                End Select
                
                parm2 = nxtyyyy & "-" & nxtmm & "-01" & " " & tmptimestr
                
                nxtyyyy = tmpenddt_yyyy
                
                'SET ENDING MONTH
                Select Case tmpenddt_mm
                    Case "03", 3
                        nxtmm = "06"
                        nxtdd = "30"
                    Case "06", 6
                        nxtmm = "09"
                        nxtdd = "30"
                    Case "09", 9
                        nxtmm = "12"
                        nxtdd = "31"
                    Case "12"
                        nxtmm = "03"
                        nxtdd = "31"
                        nxtyyyy = CStr(CInt(tmpstrtdt_yyyy) + 1)
                    Case Else
                End Select
                
                parm3 = nxtyyyy & "-" & nxtmm & "-" & nxtdd & " " & tmptimestr
                
                'SET QUARTER ID
               
                Dim qtrid_yyyy As String
                Dim qtrid_qn As String
                Dim nxtqn As String
                
                qtrid = rsTemp("quarter_id").Value
                qtrid_yyyy = Mid(qtrid, 1, 4)
                qtrid_qn = Mid(qtrid, 5, 2)
                
               
                Select Case qtrid_qn
                    Case "Q1"
                        nxtqn = "Q2"
                    Case "Q2"
                        nxtqn = "Q3"
                    Case "Q3"
                        nxtqn = "Q4"
                    Case "Q4"
                        nxtqn = "Q1"
                        qtrid_yyyy = CStr(CInt(qtrid_yyyy) + 1)
                End Select
                
                parm4 = qtrid_yyyy & nxtqn
                
                tmpstr = vbCrLf & "Please scrutinize extended quarter parameters.  Thank you." & vbCrLf
                tmpstr = tmpstr & vbCrLf & "Old qtr key: " & parm1 & vbTab
                tmpstr = tmpstr & vbCrLf & "Old Start Date: " & rsTemp("start_date") & vbTab & " New Start Date: " & parm2
                tmpstr = tmpstr & vbCrLf & "Old End Date:   " & rsTemp("term_date") & vbTab & " New End Date  : " & parm3
                tmpstr = tmpstr & vbCrLf & "Old Quarter id: " & rsTemp("quarter_id") & vbTab & " New Quarter id: " & parm4



                'SCREEN THE PARAMETERS to be passed to the stored procedure
                MsgBox (tmpstr)
                
                ans = MsgBox("Do you wish to continue to extend quarter dates?", vbOKCancel, "Extend Quarter Dates")
                Select Case ans
                Case vbOK
                Case vbCancel
                    Exit Function
                End Select
                
                
                strUpdate = "exec " & sProcedureName & " "
                strUpdate = strUpdate & CInt(parm1)
                strUpdate = strUpdate & ",'" & parm2 & "'"
                strUpdate = strUpdate & ",'" & parm3 & "'"
                strUpdate = strUpdate & ",'" & parm4 & "'"
                
                
                If DEBUGON Then Stop
                g_cnShared.Execute strUpdate    'Allow long-running procedures
                MsgBox ("Completed") 'rlh 02/27/2010
                Exit Function
            End If
           Case "sp_update_published_cci_labor_rate_rlh"
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::
            ':: PUBLISH QUARTERLY LABOR RATES
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::
            
            If DEBUGON Then Stop
            
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            'RETIRED CHECK: 1ST check if the Quarter Id is retired
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            strSelect = "SELECT  * FROM PUBLISHED_CCI_QUARTERS_RETIRED WHERE quarter_id ='" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp.EOF = False Then
                MsgBox ("Current Quarter Id: " & strSelectedQtr & " is retired")
                Exit Function
            Else
            End If
            
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            ':: Rerun Labor Rates Out-of-Date
            ':: If anything is yet out of date, quit, and alert user
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            strUpdate = "exec " & "sp_report_labor_rate_out_of_date_rlh" & " '" + strSelectedQtr + "'" + strParameters
            strSelect = strUpdate
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If blnReturn = False Then
                MsgBox "An error occurred retrieving labor rate out-of-date:" & vbCrLf & g_objDAL.LastErrorDescription
            Else
                If Not (rsTemp.EOF And rsTemp.BOF) Then
                    lCount = rsTemp.RecordCount
                    If lCount > 0 Then
                        sMsg = CStr(lCount) & " Labor Rates were found to be out of date.  You can not continue"
                        iResult = MsgBox(sMsg, vbOKOnly, "Labor Rate Out of Date Pre-run Failure")

                        Exit Function
                    End If
                Else
                    'MsgBox ("No Labor data found out of date")
                End If
            End If
            
            ':::::::::::::::::::::::::::::::::::::::::::::::::
            'Get qtr_dt_skey for current Quarter_Id
            ':::::::::::::::::::::::::::::::::::::::::::::::::
            Dim qtr_dt_skey As Integer
            strSelect = "SELECT  qtr_dt_skey FROM QUARTER_DATE WHERE quarter_id ='" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            qtr_dt_skey = rsTemp(0)
            
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            ':: Check for existence of "published" data on published_cci_labor_rate
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            
            strSelect = "SELECT count(*) from published_cci_labor_rate WHERE qtr_dt_skey=" & qtr_dt_skey
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp(0) > 0 Then
                ans = MsgBox("Data is already published for the requested quarter id: " & strSelectedQtr & " Continue?", vbOKCancel, "Existing Published Data")
                Select Case ans
                Case vbOK
                Case vbCancel
                    Exit Function
                End Select
            End If
            
            
           Case "sp_build_cci_labor_rates_allcities_grid", "sp_build_cci_labor_rates_anytown_grid"
           '#############################################################################
           '## BUILD REPORT TABLE GRID AUTOMATICALLY IN Labor Rates Grid - CmdSearch_Click
           '#############################################################################
            If DEBUGON Then Stop
            strSelectedQtr = MainModule.strLaborSelectedQtr
            
            DoEvents
            strUpdate = "exec " & sProcedureName & " '" + strSelectedQtr + "'" + strParameters
             Screen.MousePointer = vbHourglass   'rlh 03/04/2010
            g_cnShared.Execute strUpdate    'Allow long-running procedures
            ExecStoredProcSelectedQuarter = True
            Screen.MousePointer = vbNormal  'rlh 03/04/2010
            Exit Function
            
            
            
            Case "sp_rollup_pub_cci_index_masterformat_with_fuel_rlh"
            
            If DEBUGON Then Stop
            ':::::::::::::::::::::::::::::::::::::::::::::::::
            'GENERATE MASTERFORMAT INDEX
            ':::::::::::::::::::::::::::::::::::::::::::::::::
            
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            'RETIRED CHECK: 1ST check if the Quarter Id is retired
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            strSelect = "SELECT  * FROM PUBLISHED_CCI_QUARTERS_RETIRED WHERE quarter_id ='" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp.EOF = False Then
                MsgBox ("Current Quarter Id: " & strSelectedQtr & " is retired")
                Exit Function
            Else
                
            End If
            ':::::::::::::::::::::::::::::::::::::::::::::::::
            'Get qtr_dt_skey for current Quarter_Id
            ':::::::::::::::::::::::::::::::::::::::::::::::::
           
            strSelect = "SELECT  qtr_dt_skey FROM QUARTER_DATE WHERE quarter_id ='" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            qtr_dt_skey = rsTemp(0)
            
            'Check Materials for data at requested quarter id
            strSelect = "select count(*) from PUBLISHED_CCI_material_price where qtr_dt_skey=" & qtr_dt_skey
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp(0) = 0 Then
                 MsgBox ("No Data exists for MATERIAL PRICES for this quarter: " & strSelectedQtr)
               
                    Exit Function
               
            End If
            
            'Check Equipment Rates for data at requested quarter id
            
            strSelect = "select count(*) from PUBLISHED_CCI_equipment_rate where qtr_dt_skey=" & qtr_dt_skey
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp(0) = 0 Then
                 MsgBox ("No Data exists for EQUIPMENT RATES for this quarter: " & strSelectedQtr)
                
                    Exit Function
               
            End If
            'Check Labor Rates for data at requested quarter id
            strSelect = "select count(*) from PUBLISHED_CCI_labor_rate where qtr_dt_skey=" & qtr_dt_skey
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp(0) = 0 Then
                 ans = MsgBox("No Data exists for LABOR for this quarter: " & strSelectedQtr)
                
                    Exit Function
               
            End If
            
            'Next check  (does GIVEN quarter exist)
            strSelect = "select count(*) from PUBLISHED_CCI_INDEX_WITH_FUEL where Quarter_id='" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp(0) > 0 Then
                 ans = MsgBox("Data already exists for this index for this quarter: " & strSelectedQtr & "' Continue?", vbOKCancel, "Index Overwrite Check")
                Select Case ans
                Case vbOK
                Case vbCancel
                    Exit Function
                End Select
                
                iResult = vbOK
               
            End If
            
            Case "sp_rollup_pub_cci_index_uniformat_with_fuel_rlh"
            'Generate UNIFORMAT Index
            If DEBUGON Then Stop
            
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            'RETIRED CHECK: 1ST check if the Quarter Id is retired
            '::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            strSelect = "SELECT  * FROM PUBLISHED_CCI_QUARTERS_RETIRED WHERE quarter_id ='" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            If rsTemp.EOF = False Then
                MsgBox ("Current Quarter Id: " & strSelectedQtr & " is retired")
                Exit Function
            Else
                
            End If
            ':::::::::::::::::::::::::::::::::::::::::::::::::
            'Get qtr_dt_skey for current Quarter_Id
            ':::::::::::::::::::::::::::::::::::::::::::::::::
           
            strSelect = "SELECT  qtr_dt_skey FROM QUARTER_DATE WHERE quarter_id ='" & strSelectedQtr & "'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            qtr_dt_skey = rsTemp(0)
            
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            'Make sure that the Masterformat (MF) exists (from which the UF will be built!)
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            
            strSelect = "select count(*) from PUBLISHED_CCI_INDEX_WITH_FUEL where Quarter_id='" & strSelectedQtr & "'"
            strSelect = strSelect & " AND class_system_id='MF'"
            blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
            
            If rsTemp(0) > 0 Then
            
                ans = MsgBox("(MF) Masterformat data does exist on the index table for quarter id, " & strSelectedQtr, vbOKCancel, "INDEX EXISTS CHECK")
                Select Case ans
                    Case vbOK
                    Case vbCancel
                        Exit Function
                End Select

            End If
                
            
        End Select
        
        If iResult = vbOK Then
            ans = MsgBox("Continue processing?", vbOKCancel, "Continue Processing")
            Select Case ans
            Case vbOK
            Case vbCancel
                Exit Function
            End Select
            DoEvents
            strUpdate = "exec " & sProcedureName & " '" + strSelectedQtr + "'" + strParameters
            If DEBUGON Then
                Stop
            End If
            
            
            
            Screen.MousePointer = vbHourglass   'rlh 03/04/2010
            g_cnShared.Execute strUpdate    'Allow long-running procedures
            'blnReturn = g_objDAL.ExecQuery(CONNECT, strUpdate, strError)
           
          
            
            
            ExecStoredProcSelectedQuarter = True
            If UCase(sProcedureName) = "SP_REPORT_PUB_CCI_LABOR_RATE" Then
                ' 9/12/2005 RTD - ALSO RUN NEW SP TO UPDATE THE NO WORKER'S COMP TABLE
                strUpdate = "EXEC SP_REPORT_PUB_CCI_LABOR_RATE_NOWC_RLH '" + strSelectedQtr + "'"
                g_cnSharedLong.Execute strUpdate    'Allow long-running procedures
            
            End If
            
             MsgBox ("Completed") 'rlh 02/27/2010
             
            Screen.MousePointer = vbNormal  'rlh 03/04/2010
        End If
    End If
    
Exit_Sub:
    Screen.MousePointer = vbNormal
    Exit Function

Error_Processing:
    Screen.MousePointer = vbNormal
    ExecStoredProcSelectedQuarter = False
    MsgBox "Specified Date: " & tmpstr & ": " & Err.Description, vbCritical
    'Resume

End Function

Public Function CloneCCIMatPrice()
    ExecStoredProcSelectedQuarter "sp_clone_pub_cci_material_price"
End Function

Public Function GetQuarterID(ListQuarterID As cdlgLstSel, Optional sComboCaption As String) As String
    Dim sql As String
    Dim rec As ADODB.RecordSet
    Dim varCurSelectedRow  As Variant
    Dim blnResult As Boolean
    Dim sValue As String
    
    'A list of available quarters
    ' be constructed, and the list selections populated from it.
    If Len(sComboCaption) > 0 Then
        ListQuarterID.ComboCaption = sComboCaption
    Else
        ListQuarterID.ComboCaption = "Select Quarter:"
    End If
    ListQuarterID.Caption = "Quarter Selection"
    
    sql = "select qtr_dt_skey, quarter_id from quarter_date order by quarter_id"
    g_objDAL.GetRecordset CONNECT, sql, rec
    If rec.EOF And rec.BOF Then
        MsgBox "No quarter date records have been set up. Please contact the IS department for help."
        GoTo Exit_Sub
    Else
        If rec.RecordCount = 0 Then     'invalid
            MsgBox "No contacts found."
        Else
            Do Until rec.EOF
                ListQuarterID.AddUniqueItem rec.Fields("quarter_id"), 0, rec.Fields("qtr_dt_skey")
                rec.MoveNext
            Loop
        End If
        rec.Close
    End If
    
    If ListQuarterID.itemCount > 0 Then
        If ListQuarterID.SetList = True Then
            ListQuarterID.SingleValue = g_sQuarterID
            Screen.MousePointer = vbNormal
            blnResult = ListQuarterID.ShowList()
            Screen.MousePointer = vbHourglass
        End If
    End If
    
    If blnResult = True And ListQuarterID.itemCount > 0 Then  'Quarter selected or only 1 found - if none, ignore
        GetQuarterID = ListQuarterID.SingleValue
    Else
        GetQuarterID = -1
    End If

Exit_Sub:

End Function

Public Sub LaborRateParentHelperEditPrintPreview()
    ' MODIFIED TO USE NEW VSREPORT, 5/25/2005 RTD
    Dim OptionDialog As New cdlgLstSel
    Dim var As Variant
    Dim blnOnlyCurrent As Boolean
    Dim blnOnlyEstimated As Boolean
    
    On Error GoTo Exit_Sub
    OptionDialog.Check1Caption = "Select Current Only?"
    OptionDialog.Check2Caption = "Select Estimated Only?"
    OptionDialog.Caption = "Labor Rate Edit"
    OptionDialog.SelectType = 4
    OptionDialog.SetList
    If OptionDialog.ShowList Then
        blnOnlyCurrent = OptionDialog.Check1Value
        blnOnlyEstimated = OptionDialog.Check2Value
        LaborRatePrintPreview Abs(blnOnlyCurrent), Abs(blnOnlyEstimated)
    End If

Exit_Sub:
    Set OptionDialog = Nothing

End Sub

Public Sub LaborRatePrintPreview(ByVal bCurrentOnly As Integer, ByVal bEstimateOnly As Integer)
    Dim fPreviewWindow As New frmReportPreview
    Dim sRecordSource As String
    
    sRecordSource = "exec sp_rpt_select_labor_exceptions @only_current=" & bCurrentOnly & ", @only_estimated=" & bEstimateOnly & ""
    fPreviewWindow.ReportName = "Labor Rate Helper Edit"
    fPreviewWindow.ReportFile = "rptLaborExceptions.xml"
    fPreviewWindow.RecordSource = sRecordSource
    fPreviewWindow.RenderReport
    fPreviewWindow.Show
    
End Sub

Public Sub BookFormatPrintPreview(ByVal sStartUnitCost As String, ByVal sEndUnitCost As String, ByVal sMasterFormatVersion As String, Optional sBookEdition As String = "")
' Print the "Book Preview" report
' Supply the Start and End Unit-Cost-ID and their MasterFormat version
' 10/11/2005 RTD - MODIFIED TO USE NEW BOOK PREVIEW DIALOG (SUPPORTS MORE BOOK FORMATS)
    Dim fPreviewWindow As New frmReportPreview
    Dim fBookFormat As New dlgBookFormat
    Dim sRecordSource As String
    Dim sReportName As String
    Dim sEdition As String
    Dim iBookID As Long
    
    Status "Selecting Book Format..."
    If sBookEdition = "" Then
        ' PROMPT USER FOR A FORMAT/ID FOR THE BOOK PREVIEW
        fBookFormat.AllowEditing = False
        fBookFormat.UnitCostIdStart = sStartUnitCost
        fBookFormat.UnitCostIdEnd = sEndUnitCost
        fBookFormat.MasterFormatVersion = sMasterFormatVersion
        fBookFormat.Show vbModal
        If fBookFormat.Result = vbCancel Then
            Status ""
            Set fPreviewWindow = Nothing
            Set fBookFormat = Nothing
            Exit Sub
        Else
            sReportName = fBookFormat.XMLReportName
            iBookID = fBookFormat.bookid
        End If
    Else
        iBookID = 1
        sEdition = sBookEdition
        sReportName = sEdition & " Book Format"
    End If
    If (sReportName <> "") Then
        Status "Generating Book Preview..."
        Screen.MousePointer = vbHourglass
        sRecordSource = "exec usp_select_book_format @start_book_unit_cost_id = '" & sStartUnitCost & _
                            "', @end_book_unit_cost_id = '" & sEndUnitCost & _
                            "', @master_format_version = '" & sMasterFormatVersion & _
                            "', @book_output_id = " & iBookID
        fPreviewWindow.ReportName = sReportName
        fPreviewWindow.ReportFile = fBookFormat.XMLFileName
        fPreviewWindow.ConnectString = g_cnShared.ConnectionString
        fPreviewWindow.RecordSource = sRecordSource
        fPreviewWindow.RenderReport
        fPreviewWindow.Show
    End If
    Set fPreviewWindow = Nothing
    Set fBookFormat = Nothing
    Screen.MousePointer = vbDefault
    Status ""
    
End Sub

Public Sub BookFormatPrintPreviewRS(ByVal rs As ADODB.RecordSet, ByVal sMasterFormatVersion As String)
    Dim sStartUnitCost As String
    Dim sEndUnitCost As String
    Dim sFieldName As String
    
    sFieldName = "unit_cost_id"
    rs.MoveFirst
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(sFieldName)) Then
            sStartUnitCost = rs.Fields(sFieldName)
            If Left(sStartUnitCost, 1) = "M" Then sStartUnitCost = Mid(sStartUnitCost, 2)
            rs.MoveLast
            sEndUnitCost = rs.Fields(sFieldName)
            If Left(sEndUnitCost, 1) = "M" Then sEndUnitCost = Mid(sEndUnitCost, 2)
            BookFormatPrintPreview sStartUnitCost, sEndUnitCost, sMasterFormatVersion
        End If
    End If
    
End Sub

Public Sub MatPriceDiv1_14PrintPreview()
    ' MODIFIED TO USE NEW VSREPORT, 5/25/2005 RTD
    Dim fPreviewWindow As New frmReportPreview
    
    fPreviewWindow.ReportName = "Material Report for Divs 1-14"
    fPreviewWindow.ReportFile = "rptMaterialReport.xml"
    fPreviewWindow.RenderReport
    fPreviewWindow.Show

    'Forms(0).rptMatPrice1_14.ConnectionString = g_cnShared.ConnectionString
    'Forms(0).rptMatPrice1_14.PrintPreview

End Sub

Public Sub MatPriceDiv15_16PrintPreview()
    ' MODIFIED TO USE NEW VSREPORT, 5/25/2005 RTD
    Dim fPreviewWindow As New frmReportPreview
    
    fPreviewWindow.ReportName = "Material Report for Divs 15-16"
    fPreviewWindow.ReportFile = "rptMaterialReport.xml"
    fPreviewWindow.RenderReport
    fPreviewWindow.Show
    
    'Forms(0).rptMatPrice15_16.ConnectionString = g_cnShared.ConnectionString
    'Forms(0).rptMatPrice15_16.PrintPreview
    
End Sub

Public Function CommercialEstimatePreview(ByVal sRecordSource As String)
    Dim fPreviewWindow As New frmReportPreview

    fPreviewWindow.ReportName = "Commercial Summary"
    fPreviewWindow.ReportFile = "rptSummaryEstimate.xml"
    fPreviewWindow.ConnectString = g_cnShared.ConnectionString
    fPreviewWindow.RecordSource = sRecordSource
    fPreviewWindow.RenderReport
    fPreviewWindow.Show
    
End Function

Public Function ResidentialEstimatePreview(ByVal sRecordSource As String)
    Dim fPreviewWindow As New frmReportPreview

    fPreviewWindow.ReportName = "Residential Summary"
    fPreviewWindow.ReportFile = "rptSummaryEstimate.xml"
    fPreviewWindow.ConnectString = g_cnShared.ConnectionString
    fPreviewWindow.RecordSource = sRecordSource
    fPreviewWindow.RenderReport
    fPreviewWindow.Show
    
End Function
