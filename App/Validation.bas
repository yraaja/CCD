Attribute VB_Name = "Validation"
Option Explicit

Public Function AssemblyUCSortRequired(lngSkey As Long) As Boolean
    Dim strSELECT As String
    Dim rec As ADODB.RecordSet
    Dim blnReturn As Boolean

    strSELECT = "select count(*) as system_books from assembly_book_detail where type_code = 'S' and  assembly_skey = " + CStr(lngSkey)
    blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
    If Not (rec.BOF And rec.EOF) Then
        If rec.Fields("system_books") > 0 Then
            AssemblyUCSortRequired = True
        End If
    End If
    rec.Close
    Set rec = Nothing
End Function

Public Function Compress_String(strExpandedString As String) As String
'REMOVE SPACES AND DECIMAL POINTS FROM STRING
'REWRITE 6/20/2005 RTD - USE VB6 REPLACE FUNCTION INSTEAD OF LOOP
'COMPRESS PERIOD/DECIMAL-POINT AS WELL (FOR MASTERFORMAT 2004)
    Dim i As Integer
    Dim iBlankPosition As Integer
    
    Compress_String = strExpandedString
    Compress_String = Replace(Compress_String, ".", "")
    Compress_String = Replace(Compress_String, " ", "")
    
    'added by cje
    'Compress_String = Replace(Compress_String, "~", "%")

End Function

Public Function ConvertAssemblySkey(assembly_skey As Long) As Long
    Dim lngAssembly_Skey As Long

    If IsNumeric(assembly_skey) Then
        ConvertAssemblySkey = CLng(assembly_skey)
    Else
        ConvertAssemblySkey = 0
    End If

End Function

Public Function StringDataModified(strValue1 As String, strValue2 As String, strFieldName As String) As Boolean
'Return a boolean indicating whether the value was changed, allowing for nulls, etc
'If the field name is a formatted ID, blanks for formatting will be compressed

    Dim strTestValue1 As String
    Dim strTestValue2 As String
    
    strTestValue1 = CStr(IIf(IsNull(strValue1), "", strValue1))
    strTestValue2 = CStr(IIf(IsNull(strValue2), "", strValue2))
    
    If strFieldName = "mat_id" Or _
        strFieldName = "alt_mat_id" Or _
        strFieldName = "unit_cost_id" Or _
        strFieldName = "alt_unit_cost_id" Or _
        strFieldName = "ext_unit_cost_id" Or _
        strFieldName = "assembly_id" Or _
        strFieldName = "alt_assembly_id" Then
            strTestValue1 = Compress_String(strTestValue1)
            strTestValue2 = Compress_String(strTestValue2)
    End If
    
    If strTestValue1 <> strTestValue2 Then
        StringDataModified = True
    End If

End Function

Public Function StringVerifiedData(strValue As String, strFieldName As String) As String
'Return a boolean indicating whether the value was changed, allowing for nulls, etc
'If the field name is a formatted ID, blanks for formatting will be compressed
    
    On Error GoTo Errlbl    'rlh
    
    StringVerifiedData = CStr(IIf(IsNull(strValue), "", strValue))
    
    If strFieldName = "mat_id" Or _
        strFieldName = "alt_mat_id" Or _
        strFieldName = "unit_cost_id" Or _
        strFieldName = "alt_unit_cost_id" Or _
        strFieldName = "ext_unit_cost_id" Or _
        strFieldName = "assembly_id" Or _
        strFieldName = "alt_assembly_id" Then
            StringVerifiedData = Compress_String(strValue)
    End If
   
    Exit Function   'rlh
Errlbl:
    MsgBox "VALIDATION: StringVerifiedData: " & Err.Description 'rlh
    

End Function

Public Function validate_uc_type_code(strType_code As String, lngUnitCostSkey As Long) As Boolean
    Dim rec As New ADODB.RecordSet
    Dim strSELECT As String
    Dim blnReturn As Boolean
    
    Screen.MousePointer = vbHourglass
    validate_uc_type_code = True
    If strType_code = "H" Then 'H is invalid for UC in use on the UC usage table
        strSELECT = "select count(*) as uc_usage_count from unit_cost_usage" + _
        " where unit_cost_skey = " + CStr(lngUnitCostSkey)
        blnReturn = g_objDAL.GetRecordset(vbNullString, strSELECT, rec)
        If Not blnReturn Then
            MsgBox "Database Error validating the Unit Cost Usage"
            validate_uc_type_code = False
        Else
            If rec.Fields("uc_usage_count") > 0 Then
                MsgBox "This unit cost is in use on the Unit Cost Usage table.  Please remove all instances before changing to type H."
                validate_uc_type_code = False
            End If
        End If
    End If
    Screen.MousePointer = vbNormal

End Function

Public Function AsblyUCGridError_AssemblyID(strAssemblyId As String, strUnitCostId As String, Optional intMasterFormatVersion As Long = UCD_MASTERFORMAT_VERSION) As String
    Dim recAssembly As New ADODB.RecordSet ' Recordset to hold query results
    Dim recUnitCost As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnAssembly As Boolean
    Dim blnUnitCost As Boolean
    Dim strSELECT As String
    
'Validate the assembly/unitcost id combination.  the unit cost may not have been entered yet - no error
'   Errors:     Blank Assembly ID - Assembly ID Required
'               Invalid Assembly ID
'               E Type Assembly for non-E type Unit Cost
'               M Type Assembly for non-M/B type Unit Cost

    AsblyUCGridError_AssemblyID = Empty
    
    If Len(strAssemblyId) = 0 Then    'Assembly ID Required
        AsblyUCGridError_AssemblyID = "The Assembly ID must be entered."
    Else
        strSELECT = "Select assembly_skey, type_code " + _
                    "from Assembly_detail where assembly_id='" + strAssemblyId + "'"
            ' Use DAL to perform select
        blnAssembly = g_objDAL.GetRecordset(CONNECT, strSELECT, recAssembly)
        If blnAssembly = False Then  'error opening recordset
            AsblyUCGridError_AssemblyID = "Database Error encountered validating the Assembly ID"
        Else
        ' Check to see if the assembly_id entered exists already
            If recAssembly.RecordCount = 0 Then 'Not found
                AsblyUCGridError_AssemblyID = "The Assembly ID " + strAssemblyId + " does not exist."
            Else
                If Len(strUnitCostId) > 0 Then
                    'UPDATED 8/5/2005 RTD - SUPPORT MASTERFORMAT VERSION 2004
                    Select Case intMasterFormatVersion
                    Case EXT_MASTERFORMAT_VERSION
                        strSELECT = "SELECT ucdx.unit_cost_skey, ucd.type_code " + _
                            "FROM unit_cost_detail_ext ucdx, unit_cost_detail ucd " + _
                            "WHERE ucdx.unit_cost_skey = ucd.unit_cost_skey " + _
                            "AND ucdx.unit_cost_id='" + strUnitCostId + "'"
                    Case Else
                        strSELECT = "SELECT unit_cost_skey, type_code " + _
                            "FROM unit_cost_detail " + _
                            "WHERE unit_cost_id='" + strUnitCostId + "'"
                    End Select
                    ' Use DAL to perform select
                    blnUnitCost = g_objDAL.GetRecordset(CONNECT, strSELECT, recUnitCost)
                    If blnUnitCost = False Then  'Unit Cost recordset open
                        AsblyUCGridError_AssemblyID = "Database Error encountered validating the Unit Cost ID"
                    Else
                        If recUnitCost.RecordCount = 0 Then
                            AsblyUCGridError_AssemblyID = "The Unit Cost ID " + strUnitCostId + " does not exist."
                        Else
                            If recUnitCost.Fields("type_code") = "E" Then
                                If recAssembly.Fields("type_code") <> "E" Then
                                    AsblyUCGridError_AssemblyID = "Only Type E Assemblies are valid with unit cost " + strUnitCostId + "."
                                End If
                            ElseIf recUnitCost.Fields("type_code") = "M" Or recUnitCost.Fields("type_code") = "B" Then     'Not E, must be M
                                If recAssembly.Fields("type_code") <> "M" Then
                                    AsblyUCGridError_AssemblyID = "Only Type M Assemblies are valid with unit cost " + strUnitCostId + "."
                                End If
                            Else
                                AsblyUCGridError_AssemblyID = "Unit Cost type " + recUnitCost.Fields("type_code") + " may not be assigned to an Assembly."
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function AsblyUCGridError_UnitCostID(strUnitCostId As String, strAssemblyId As String, _
        Optional strAssemblyType As String = Empty, Optional intMasterFormatVersion As Long = UCD_MASTERFORMAT_VERSION) As String
    
    Dim recAssembly As New ADODB.RecordSet ' Recordset to hold query results
    Dim recUnitCost As New ADODB.RecordSet ' Recordset to hold query results
    Dim blnAssembly As Boolean
    Dim blnUnitCost As Boolean
    Dim strSELECT As String
    Dim strTestAssemblyType As String
    
'Validate the assembly/unitcost id combination.  the assembly ID may not have been entered yet - no error
'   Errors:     Blank unit cost ID - ID Required
'               Invalid Unit Cost ID
'               E Type Unit Cost for non-E type Assembly
'               M/B Type Unit Cost for non-M type Assembly

    AsblyUCGridError_UnitCostID = Empty
    
    If Len(strUnitCostId) = 0 Then    'unit cost ID Required
        AsblyUCGridError_UnitCostID = "The Assembly ID must be entered."
    Else
        'UPDATED 8/5/2005 RTD - SUPPORT MASTERFORMAT VERSION 2004
        Select Case intMasterFormatVersion
        Case EXT_MASTERFORMAT_VERSION
            strSELECT = "SELECT ucdx.unit_cost_skey, ucd.type_code " + _
                "FROM unit_cost_detail_ext ucdx, unit_cost_detail ucd " + _
                "WHERE ucdx.unit_cost_skey = ucd.unit_cost_skey " + _
                "AND ucdx.unit_cost_id='" + strUnitCostId + "'"
        Case Else
            strSELECT = "SELECT unit_cost_skey, type_code " + _
                "FROM unit_cost_detail " + _
                "WHERE unit_cost_id='" + strUnitCostId + "'"
        End Select
        ' Use DAL to perform select
        blnUnitCost = g_objDAL.GetRecordset(CONNECT, strSELECT, recUnitCost)
        If blnUnitCost = False Then  'Unit Cost recordset open
            AsblyUCGridError_UnitCostID = "Database Error encountered validating the Unit Cost ID"
        Else
            If recUnitCost.RecordCount = 0 Then
                AsblyUCGridError_UnitCostID = "The Unit Cost ID " + strUnitCostId + " does not exist."
            Else
                If Len(strAssemblyId) > 0 Then
                    strSELECT = "Select assembly_skey, type_code " + _
                                "from Assembly_detail where assembly_id='" + strAssemblyId + "'"
                        ' Use DAL to perform select
                    blnAssembly = g_objDAL.GetRecordset(CONNECT, strSELECT, recAssembly)
                    If blnAssembly = False Then  'error opening recordset
                        AsblyUCGridError_UnitCostID = "Database Error encountered validating the Assembly ID"
                    Else
                    ' Check to see if the assembly_id entered exists already
                        If recAssembly.RecordCount = 0 Then   'Not found, allow for add mode
                            strTestAssemblyType = strAssemblyType
                        Else
                            If IsEmpty(strAssemblyType) Or strAssemblyType = "" Then
                                strTestAssemblyType = recAssembly.Fields("type_code")
                            Else
                                strTestAssemblyType = strAssemblyType
                            End If
                        End If
                        If strTestAssemblyType = "E" Then
                            If recUnitCost.Fields("type_code") <> "E" Then
                                AsblyUCGridError_UnitCostID = "Only Type E Unit Costs are valid with assembly " + strAssemblyId + "."
                            End If
                        Else    'Not E, must be M
                            If recUnitCost.Fields("type_code") <> "M" And recUnitCost.Fields("type_code") <> "B" Then
                                AsblyUCGridError_UnitCostID = "Only Type M or B unit costs are valid with assembly " + strAssemblyId + "."
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

