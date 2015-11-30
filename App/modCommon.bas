Attribute VB_Name = "modCommon"
Option Explicit

Private cooEventSubscribers    As New Collection
Private mnSubscriberCount      As Long

Public Enum EEventSubscriberNotifyType
    esnBuildingRecordUpdated
    esnModelRecordUpdated
    esnProjectRecordUpdated         ' 9/9/2005 RTD - New Event for Projects
    esnUserRecordupdated            ' 9/13/2005 RTD - New Event for Users
End Enum

Public Sub LoadCities(cmbCity As ComboBox, State As String, Optional strCity As String)
    Dim strSelect As String
    Dim rsTemp As RecordSet
    Dim blnReturn As Boolean
    
    'Load Cities
    Screen.MousePointer = vbHourglass
    cmbCity.Clear
    If State > "" Then
'        strSELECT = "select city, loc_id from location where location.state_code = '" + State + "'  order by city"
        
        'rlh 02/26/2010 - Add "AnyTown" to the city drop-down
        strSelect = "select city, loc_id from location where location.state_code = '" + State + "' "
'        strSELECT = strSELECT & " OR city='Anytown'"
        strSelect = strSelect & " order by city"
        
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "An error occurred loading Cities." & vbCrLf & g_objDAL.LastErrorDescription, vbExclamation
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                Do Until rsTemp.EOF
                    If strCity > "" Then
                        If strCity = rsTemp![City] Then
                            cmbCity.Text = ConvertCase(rsTemp![City])
                        End If
                    End If
                    cmbCity.AddItem ConvertCase(rsTemp![City])
                    cmbCity.ItemData(cmbCity.NewIndex) = rsTemp![loc_id]
                    rsTemp.MoveNext
                Loop
            End If
        End If
        rsTemp.Close
    
    End If
    Screen.MousePointer = vbDefault

End Sub

Public Function CheckUserAuth() As Boolean   'rlh 03/04/2010
Dim strSelect As String
    Dim rsTemp As RecordSet
    Dim blnReturn As Boolean
    
    'Check user ROLE authorization
    Screen.MousePointer = vbHourglass
    CheckUserAuth = False
  
'        strSELECT = "select city, loc_id from location where location.state_code = '" + State + "'  order by city"
        
        'rlh 02/26/2010 - Add "AnyTown" to the city drop-down
        strSelect = "select user_role from user_names where user_id='" & strUserName & "'"
        ' Use DAL to perform select
        blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
        If blnReturn = False Then
            MsgBox "(CCI User Authorization): An error occurred finding User Id." & vbCrLf & g_objDAL.LastErrorDescription, vbExclamation
        Else
            If Not (rsTemp.EOF And rsTemp.BOF) Then
                Do Until rsTemp.EOF
                    If rsTemp("user_role") = 128 Then
                        CheckUserAuth = True
                    End If
                    rsTemp.MoveNext
                Loop
            End If
        End If
        rsTemp.Close
    
  
    Screen.MousePointer = vbDefault

End Function
Public Function ConvertCase(strText As String) As String
'REWRITTEN 9/8/2005 RTD
'USE STRCONV FUNCTION; TAKES INTO ACCOUNT SPECIAL CASES
    Dim strTemp As String
    Dim P As Long, i As Long
    Dim aTokens() As Variant
    Dim strToken As String
    
    aTokens = Array("-", ".", ",", "(", "[", "/", "O'", " Mc")
    strTemp = StrConv(strText, vbProperCase) & " "
    For i = LBound(aTokens) To UBound(aTokens)
        strToken = aTokens(i)
        Do While InStr(strTemp, strToken) > 0
            P = InStr(strTemp, strToken)
            strTemp = Left(strTemp, P - 1) & "~" & UCase(Mid(strTemp, P + Len(strToken), 1)) & Mid(strTemp, P + Len(strToken) + 1)
        Loop
        strTemp = Replace(strTemp, "~", strToken)
    Next
    ConvertCase = Trim(strTemp)

End Function

Public Function EventSubscriberAdd(frmTarget As Form) As String
    On Error Resume Next
    mnSubscriberCount = mnSubscriberCount + 1
    cooEventSubscribers.Add frmTarget, "Subscriber" & mnSubscriberCount
    EventSubscriberAdd = "Subscriber" & mnSubscriberCount
End Function

Public Sub EventSubscriberRemove(sKey As String)
    On Error Resume Next
    cooEventSubscribers.Remove sKey
End Sub

Public Sub EventSubscriberNotify(eNotifyType As EEventSubscriberNotifyType, sAffectedRecordIdentifier As String)
    Dim frmTarget As Form
        
    On Error Resume Next
    For Each frmTarget In cooEventSubscribers
        frmTarget.EventNotify eNotifyType, sAffectedRecordIdentifier
    Next frmTarget
End Sub

Public Sub HiliteTextBox(ctlTextBox As VB.TextBox)
    With ctlTextBox
        .SelStart = 0
        .SelLength = 65535
    End With
End Sub

Public Function RemoveCharacters(sTarget As String, sBadChars As String) As String
    Dim sReturn         As String
    Dim nCurrentChar    As Long

    sReturn = sTarget
    '
    '   This will remove every bad char that is in the sBadChars
    '   string which is formmatted/passed in like = "\?/|[]{}*&><"":|."
    '
    For nCurrentChar = 1 To Len(sBadChars)
        sReturn = RemoveCharacter(sReturn, Mid$(sBadChars, nCurrentChar))
    Next
    RemoveCharacters = sReturn
End Function

Public Function RemoveCharacter(sTarget As String, sBadChar As String) As String
    Dim nBadCharLocation    As Long
    Dim sBad                As String

    sBad = Left$(sBadChar, 1)
    Do
        nBadCharLocation = InStr(sTarget, sBad)
        If nBadCharLocation > 0 Then
            '
            '   Remove the character entirely
            '
            sTarget = Left$(sTarget, nBadCharLocation - 1) & _
                            Mid$(sTarget, nBadCharLocation + 1)
        End If
    Loop While nBadCharLocation > 0
    
    RemoveCharacter = sTarget
End Function

Public Function RetrieveCurrentQuarter() As String
    
    Dim blnReturn As Boolean
    Dim strSelect As String
    Dim rsTemp As RecordSet

    On Error Resume Next

'Fill current quarter
    strSelect = "SELECT quarter_id FROM QUARTER_DATE where '" & Format(Now(), "Short Date") & "' between start_date and term_date ORDER BY quarter_id"
    ' Use DAL to perform select
    blnReturn = g_objDAL.GetRecordset(CONNECT, strSelect, rsTemp)
    If blnReturn = False Then
        MsgBox "An error occurred retrieving the current Quarter."
    Else
        If Not (rsTemp.EOF And rsTemp.BOF) Then
            RetrieveCurrentQuarter = rsTemp.Fields("quarter_id")
        End If
    End If
    rsTemp.Close
    Set rsTemp = Nothing

End Function

' Returns true if the rollup flag is 1, else false.
Public Function GetDomainTableValue(strDomainName As String) As String

    Dim rec As ADODB.RecordSet
    Dim strSelect As String
    
    GetDomainTableValue = ""

    strSelect = "select domain_value from domain_tbl where domain_name = '" & strDomainName & "'"
    g_objDAL.GetRecordset CONNECT, strSelect, rec
    
    If Not rec.EOF Then
        GetDomainTableValue = rec.Fields("domain_value")
    End If
    
    rec.Close
    Set rec = Nothing
    
End Function


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



Public Function CostRoundingFromStoredProc( _
ByRef std_mat_cost As Double, ByRef std_labor_cost As Double, ByRef std_equip_cost As Double, ByRef std_total_cost As Double, _
ByRef rr_mat_cost As Double, ByRef rr_labor_cost As Double, ByRef rr_equip_cost As Double, ByRef rr_total_cost As Double, _
ByRef opn_mat_cost As Double, ByRef opn_labor_cost As Double, ByRef opn_equip_cost As Double, ByRef opn_total_cost As Double, _
ByRef metric_mat_cost As Double, ByRef metric_labor_cost As Double, ByRef metric_equip_cost As Double, ByRef metric_total_cost As Double, _
ByRef res_mat_cost As Double, ByRef res_labor_cost As Double, ByRef res_equip_cost As Double, ByRef res_total_cost As Double, _
ByRef std_mat_cost_op As Double, ByRef std_labor_cost_op As Double, ByRef std_equip_cost_op As Double, ByRef std_total_cost_op As Double, _
ByRef rr_mat_cost_op As Double, ByRef rr_labor_cost_op As Double, ByRef rr_equip_cost_op As Double, ByRef rr_total_cost_op As Double, _
ByRef opn_mat_cost_op As Double, ByRef opn_labor_cost_op As Double, ByRef opn_equip_cost_op As Double, ByRef opn_total_cost_op As Double, _
ByRef metric_mat_cost_op As Double, ByRef metric_labor_cost_op As Double, ByRef metric_equip_cost_op As Double, ByRef metric_total_cost_op As Double, _
ByRef res_mat_cost_op As Double, ByRef res_labor_cost_op As Double, ByRef res_equip_cost_op As Double, ByRef res_total_cost_op As Double _
) As String
    
    On Error GoTo ErrHandler

    Dim cmd As New ADODB.Command
    Dim param As ADODB.Parameter
    cmd.CommandText = "sp_round_cost_values_exception_rules"
    cmd.CommandType = CommandTypeEnum.adCmdStoredProc

    Set param = cmd.CreateParameter("std_mat_cost", adCurrency, adParamInputOutput, , std_mat_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("std_labor_cost", adCurrency, adParamInputOutput, , std_labor_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("std_equip_cost", adCurrency, adParamInputOutput, , std_equip_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("std_total_cost", adCurrency, adParamInputOutput, , std_total_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("rr_mat_cost", adCurrency, adParamInputOutput, , rr_mat_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("rr_labor_cost", adCurrency, adParamInputOutput, , rr_labor_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("rr_equip_cost", adCurrency, adParamInputOutput, , rr_equip_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("rr_total_cost", adCurrency, adParamInputOutput, , rr_total_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("opn_mat_cost", adCurrency, adParamInputOutput, , opn_mat_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("opn_labor_cost", adCurrency, adParamInputOutput, , opn_labor_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("opn_equip_cost", adCurrency, adParamInputOutput, , opn_equip_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("opn_total_cost", adCurrency, adParamInputOutput, , opn_total_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("metric_mat_cost", adCurrency, adParamInputOutput, , metric_mat_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("metric_labor_cost", adCurrency, adParamInputOutput, , metric_labor_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("metric_equip_cost", adCurrency, adParamInputOutput, , metric_equip_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("metric_total_cost", adCurrency, adParamInputOutput, , metric_total_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("res_mat_cost", adCurrency, adParamInputOutput, , res_mat_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("res_labor_cost", adCurrency, adParamInputOutput, , res_labor_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("res_equip_cost", adCurrency, adParamInputOutput, , res_equip_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("res_total_cost", adCurrency, adParamInputOutput, , res_total_cost)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("std_mat_cost_op", adCurrency, adParamInputOutput, , std_mat_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("std_labor_cost_op", adCurrency, adParamInputOutput, , std_labor_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("std_equip_cost_op", adCurrency, adParamInputOutput, , std_equip_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("std_total_cost_op", adCurrency, adParamInputOutput, , std_total_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("rr_mat_cost_op", adCurrency, adParamInputOutput, , rr_mat_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("rr_labor_cost_op", adCurrency, adParamInputOutput, , rr_labor_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("rr_equip_cost_op", adCurrency, adParamInputOutput, , rr_equip_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("rr_total_cost_op", adCurrency, adParamInputOutput, , rr_total_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("opn_mat_cost_op", adCurrency, adParamInputOutput, , opn_mat_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("opn_labor_cost_op", adCurrency, adParamInputOutput, , opn_labor_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("opn_equip_cost_op", adCurrency, adParamInputOutput, , opn_equip_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("opn_total_cost_op", adCurrency, adParamInputOutput, , opn_total_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("metric_mat_cost_op", adCurrency, adParamInputOutput, , metric_mat_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("metric_labor_cost_op", adCurrency, adParamInputOutput, , metric_labor_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("metric_equip_cost_op", adCurrency, adParamInputOutput, , metric_equip_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("metric_total_cost_op", adCurrency, adParamInputOutput, , metric_total_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("res_mat_cost_op", adCurrency, adParamInputOutput, , res_mat_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("res_labor_cost_op", adCurrency, adParamInputOutput, , res_labor_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("res_equip_cost_op", adCurrency, adParamInputOutput, , res_equip_cost_op)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("res_total_cost_op", adCurrency, adParamInputOutput, , res_total_cost_op)
    cmd.Parameters.Append param

    Set cmd.ActiveConnection = g_cnShared
    cmd.Execute

    std_mat_cost = cmd("std_mat_cost").Value
    std_labor_cost = cmd("std_labor_cost").Value
    std_equip_cost = cmd("std_equip_cost").Value
    std_total_cost = cmd("std_total_cost").Value
    rr_mat_cost = cmd("rr_mat_cost").Value
    rr_labor_cost = cmd("rr_labor_cost").Value
    rr_equip_cost = cmd("rr_equip_cost").Value
    rr_total_cost = cmd("rr_total_cost").Value
    opn_mat_cost = cmd("opn_mat_cost").Value
    opn_labor_cost = cmd("opn_labor_cost").Value
    opn_equip_cost = cmd("opn_equip_cost").Value
    opn_total_cost = cmd("opn_total_cost").Value
    metric_mat_cost = cmd("metric_mat_cost").Value
    metric_labor_cost = cmd("metric_labor_cost").Value
    metric_equip_cost = cmd("metric_equip_cost").Value
    metric_total_cost = cmd("metric_total_cost").Value
    res_mat_cost = cmd("res_mat_cost").Value
    res_labor_cost = cmd("res_labor_cost").Value
    res_equip_cost = cmd("res_equip_cost").Value
    res_total_cost = cmd("res_total_cost").Value
    std_mat_cost_op = cmd("std_mat_cost_op").Value
    std_labor_cost_op = cmd("std_labor_cost_op").Value
    std_equip_cost_op = cmd("std_equip_cost_op").Value
    std_total_cost_op = cmd("std_total_cost_op").Value
    rr_mat_cost_op = cmd("rr_mat_cost_op").Value
    rr_labor_cost_op = cmd("rr_labor_cost_op").Value
    rr_equip_cost_op = cmd("rr_equip_cost_op").Value
    rr_total_cost_op = cmd("rr_total_cost_op").Value
    opn_mat_cost_op = cmd("opn_mat_cost_op").Value
    opn_labor_cost_op = cmd("opn_labor_cost_op").Value
    opn_equip_cost_op = cmd("opn_equip_cost_op").Value
    opn_total_cost_op = cmd("opn_total_cost_op").Value
    metric_mat_cost_op = cmd("metric_mat_cost_op").Value
    metric_labor_cost_op = cmd("metric_labor_cost_op").Value
    metric_equip_cost_op = cmd("metric_equip_cost_op").Value
    metric_total_cost_op = cmd("metric_total_cost_op").Value
    res_mat_cost_op = cmd("res_mat_cost_op").Value
    res_labor_cost_op = cmd("res_labor_cost_op").Value
    res_equip_cost_op = cmd("res_equip_cost_op").Value
    res_total_cost_op = cmd("res_total_cost_op").Value


CostRoundingFromStoredProc = ""

    Exit Function

ErrHandler:
    CostRoundingFromStoredProc = Err.Description



End Function


Public Function ReplaceCharactersForFormat(ByRef strText As String) As String
    
    Dim strUnwanted As String
    Dim strRepl As String
    strRepl = "0"
    strUnwanted = "0123456789"
    Dim i As Integer
    Dim ch As String
    
    Dim beforeDotStr As String
    Dim afterDotStr As String
    Dim splitDotStr
    
    splitDotStr = Split(strText, ".")
    If UBound(splitDotStr) - LBound(splitDotStr) + 1 = 2 Then
        If splitDotStr(1) <> "" Then
            afterDotStr = splitDotStr(1)
            beforeDotStr = splitDotStr(0)
            Dim zeros As Integer
            zeros = Len(splitDotStr(1))
            If (zeros < 2) Then
                zeros = 2
            End If
            afterDotStr = ""
            For i = 1 To zeros
                ' Replace the i-th unwanted character.
                afterDotStr = afterDotStr + "0" ' Replace(afterDotStr, Mid$(afterDotStr, i, 1), "0")
            Next
            For i = 1 To Len(splitDotStr(0))
                ' Replace the i-th unwanted character.
                If (Mid$(beforeDotStr, i, 1)) <> "0" Then
                    beforeDotStr = Replace(beforeDotStr, Mid$(beforeDotStr, i, 1), "#")
                End If
            Next
            ReplaceCharactersForFormat = "#" + beforeDotStr + "." + afterDotStr
        Else
            ReplaceCharactersForFormat = strText
        End If
    Else
        
            beforeDotStr = splitDotStr(0)
            
            For i = 1 To Len(splitDotStr(0))
                ' Replace the i-th unwanted character.
                If (Mid$(beforeDotStr, i, 1)) <> "0" Then
                    beforeDotStr = Replace(beforeDotStr, Mid$(beforeDotStr, i, 1), "#")
                End If
            Next
        ReplaceCharactersForFormat = beforeDotStr + ".00"
    End If
    
End Function

