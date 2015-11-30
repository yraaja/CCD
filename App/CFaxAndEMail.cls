VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFaxAndEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'==============================================
'Description:
'This module is a helper class, facilitates
'sending fax and email from the application
'---------------------------------------------
'Dependencies:
'SafeOutLook Library
'Microsoft Outlook 10.0 object Library
'Outlook 2002
'---------------------------------------------
'Author: AK Dash
'Date: 05/25/2005
'Change Description: Initial Version
'AKD 04/24/2006 Added Removal of Outlook contact after successful faxing
'==============================================
Option Explicit

Dim objOutlook As Outlook.Application
Dim objNamespace As Outlook.NameSpace
Dim olMail As Outlook.MailItem
Dim strErrorMessage As String

Public Function SendEMail(EMailAddress As String, emailSubject As String, emailBody As String, emailAttachment As String) As Boolean
'***************************************************************
' Send email to the email address with attachment (if provided)
'***************************************************************
    Dim SafeEmail As Redemption.SafeMailItem
    Dim objEMail As Outlook.MailItem
    Dim returnValue As Integer
    
    'Initialization
    returnValue = True
    strErrorMessage = ""
   
    'Check EMail Address
    If (Trim(EMailAddress) = "") Then
        strErrorMessage = "Email address can not be empty"
        GoTo errorHandler
    End If
   
    On Error GoTo errorHandler
   
    If objOutlook Is Nothing Then
        Set objOutlook = New Outlook.Application
        Set objNamespace = objOutlook.GetNamespace("MAPI")
        Call objNamespace.Logon
    End If

    Set SafeEmail = New Redemption.SafeMailItem
    Set objEMail = Application.CreateItem(olMailItem)  'Create an Email Message
    SafeEmail.Item = objEMail
    SafeEmail.Recipients.Add EMailAddress
    SafeEmail.Subject = emailSubject
    SafeEmail.Body = emailBody
    
    'Add the attachment to the message
    If (Trim(emailAttachment) > "") Then
        SafeEmail.Attachments.Add emailAttachment
    End If
    
    SafeEmail.Send
    objNamespace.Logoff
    Set objNamespace = Nothing
    Set objOutlook = Nothing
    
    SendEMail = returnValue
    Exit Function
    
errorHandler:
    returnValue = False
    SendEMail = returnValue
    If (strErrorMessage = "") Then
        strErrorMessage = "CFaxAndEMail:SendEMail" & "- " & Err.Number & "- " & Err.Description
    End If
    Call Err.Raise(1001, "CFaxAndEMail.SendEMail", strErrorMessage)
End Function

Public Function SendFax(faxRecipient As String, recipientCompany As String, faxNumber As String, faxNote As String, faxAttachment As String) As Boolean
'***************************************************************
' Send Fax and attachment to the recipient
'***************************************************************
    Dim SafeFax As Redemption.SafeMailItem
    Dim objFax As Outlook.MailItem
    Dim returnValue As Integer
    
    'rlh  4/23/2010
    If Trim(faxRecipient) = "" Then
        MsgBox ("(ERROR)Missing Recipient name.  Please update Information Sources for this company: " & recipientCompany & " Thank you.")
        Exit Function
    End If
    
    'Initialization
    returnValue = True
    strErrorMessage = ""
   
   'Check Fax Number Address
   If (Trim(faxNumber) = "") Then
        strErrorMessage = "Fax Number can not be empty"
        GoTo errorHandler
    End If
   
   On Error GoTo errorHandler
   
   If objOutlook Is Nothing Then
        Set objOutlook = New Outlook.Application
        Set objNamespace = objOutlook.GetNamespace("MAPI")
       
        Call objNamespace.Logon
        'Call objNamespace.Logon("RSMeans (RBI-US RCD)", "", False, True)
    End If
    
    'Clean the contact Information
    Call modLog.WriteToLogRlh("RemoveContact #1", "Clean the contact Information")
    Call RemoveContact(faxRecipient, False)
    
    faxNumber = FormatFaxNumber(faxNumber)
    
   'Add Contact
    Call modLog.WriteToLogRlh("Add Contact", "before")
    Call AddContact(faxRecipient, recipientCompany, faxNumber, False)
    Call modLog.WriteToLogRlh("Add Contact", "after")
    
    Set SafeFax = New Redemption.SafeMailItem
    Set objFax = Application.CreateItem(olMailItem)  'Create an Email Message
    
    objFax.SentOnBehalfOfName = "RSMeansBOS@reedbusiness.com"
    objFax.SentOnBehalfOfName = "See Attached"  'Weird.  It completed all fields on the cover letter w/this entry!!??
   
    

    SafeFax.Item = objFax
        
    SafeFax.Recipients.Add faxRecipient
    SafeFax.Subject = faxNote
    SafeFax.Recipients.ResolveAll
   
    'Add the attachment to the
    If (Trim(faxAttachment) > "") Then
        SafeFax.Attachments.Add faxAttachment
    End If
    'TO BE USED TO DEBUG ON ANOTHER WORKSTATION
    'Moved WriteToLog in to Debug on another workstation - rlh 04/28/2010
    'WriteToLog "FAX", "CONTACT IDXXX", recipientCompany, faxRecipient, "Last Name", "First Name", faxNumber
    Call modLog.WriteToLogRlh("SafeFax.Send", "before")
    SafeFax.Send
     Call modLog.WriteToLogRlh("SafeFax.Send", "after")
    
    'Clean the contact Information Added
    Call modLog.WriteToLogRlh("RemoveContact #2 (after send)", "Clean the contact Information Added")
    Call RemoveContact(faxRecipient, True)
    
'    objNamespace.Logoff
'    Set objNamespace = Nothing
'    Set objOutlook = Nothing

    SendFax = returnValue
    Exit Function
    
errorHandler:
    returnValue = False
    SendFax = returnValue
    
    Set objNamespace = Nothing
    Set objOutlook = Nothing
    
    If (strErrorMessage = "") Then
        strErrorMessage = "CFaxAndEMail:SendFax" & "- " & Err.Number & "- " & Err.Description
    End If
    Call Err.Raise(1002, "CFaxAndEMail.SendFax", strErrorMessage)
    
End Function

Public Function AddContact(ContactName As String, contactCompany As String, contactFaxNumber As String, destroyObjects As Boolean) As Boolean
    '***************************************************************
    ' Add the recipient to the Contact List
    '***************************************************************
    Dim returnValue As Integer
    Dim contactItems
    
    'Initialization
    returnValue = True
    strErrorMessage = ""
   
   'Check Contact name and Fax Number
   If ((Trim(ContactName) = "") Or (Trim(contactFaxNumber) = "")) Then
        strErrorMessage = "Both Contact Name and Fax Number should be provided."
       GoTo errorHandler
    End If
   
   On Error GoTo errorHandler
   
    If objOutlook Is Nothing Then
        Set objOutlook = New Outlook.Application
        Set objNamespace = objOutlook.GetNamespace("MAPI")
        Call objNamespace.Logon
    End If
    
    Set contactItems = objNamespace.GetDefaultFolder(olFolderContacts).Items
    
    With contactItems.Add
        .FullName = ContactName
        .CompanyName = contactCompany
        .BusinessFaxNumber = contactFaxNumber
        .Save
    End With
    Set contactItems = Nothing

    If (destroyObjects = False) Then
        'don't clean the objects
    Else
        objNamespace.Logoff
        Set objNamespace = Nothing
        Set objOutlook = Nothing
    End If
    
    AddContact = returnValue
    Exit Function
    
errorHandler:
    returnValue = False
    AddContact = returnValue
    If (strErrorMessage = "") Then
        strErrorMessage = "CFaxAndEMail:AddContact" & "- " & Err.Number & "- " & Err.Description
    End If
    Call Err.Raise(1003, "CFaxAndEMail.AddContact", strErrorMessage)
End Function

Public Function RemoveContact(ContactName As String, destroyObjects As Boolean) As Boolean
    '***************************************************************
    ' Remove the Contact from Address Book
    '***************************************************************
    Dim returnValue As Boolean
    Dim contactItems As Items
    Dim contactItem As Outlook.contactItem
    Dim faxContact As Outlook.contactItem
    
    'MsgBox ("Inside RemoveContact")
    Call modLog.WriteToLogRlh("RemoveContact: Inside Contact", "")
    
    'Initialization
    returnValue = True
    strErrorMessage = ""
   
   'Check Contact name
   If (Trim(ContactName) = "") Then
        strErrorMessage = "Contact Name should be provided."
       GoTo errorHandler
    End If
   
   On Error GoTo errorHandler
   
    If objOutlook Is Nothing Then
        Set objOutlook = New Outlook.Application
        Set objNamespace = objOutlook.GetNamespace("MAPI")
        Call objNamespace.Logon
    End If
    'MsgBox ("Section #1")
    Call modLog.WriteToLogRlh("RemoveContact: Section #1", "")
    Set contactItems = objNamespace.GetDefaultFolder(olFolderContacts).Items
    ' Restrict the returned collection to contacts only otherwise you will get a type mismatch when
    ' trying to assign, say, a distribution list to the contactItem object.
    Set contactItems = contactItems.Restrict("[MessageClass] = 'IPM.Contact'")
    
    '06/28/2005 RTD - CHANGED ' TO "" IN CASE NAME HAS AN APOSTROPHE (e.g., Miles O'Brien)
    Call modLog.WriteToLogRlh("RemoveContact: Doing a find: ", ContactName)
    
    Set faxContact = contactItems.Find("[FullName] = """ & ContactName & """")
    
    Call modLog.WriteToLogRlh("RemoveContact: Section #2", "")
    If faxContact Is Nothing Then
        Call modLog.WriteToLogRlh("fax contact was empty. exiting function", "")
        Exit Function
    End If
    
    Dim itemCount As Integer
    Dim contactCount As Integer
    Dim itemFound As Boolean
    
    itemCount = 0
    itemFound = False

    contactCount = contactItems.Count
    'MsgBox ("Just before Do While Loop")'
    
    Call modLog.WriteToLogRlh("Just before Do While Loop", "")
    
    Do While itemCount < contactCount
        If itemCount = 0 Then
            Call modLog.WriteToLogRlh("Setting contact(GetFirst): itemCount", CStr(itemCount))
            Set contactItem = contactItems.GetFirst
        Else
            Call modLog.WriteToLogRlh("Setting contact(GetNext): itemCount", CStr(itemCount))
            Set contactItem = contactItems.GetNext
        End If
        
        itemCount = itemCount + 1
        
        Call modLog.WriteToLogRlh("contactItem.FullName", contactItem.FullName)
        If (contactItem.FullName = faxContact.FullName) Then
            itemFound = True
            Call modLog.WriteToLogRlh("itemFound", "true")
            Exit Do
        End If
    Loop
    'MsgBox ("Just after Do While Loop")
    Call modLog.WriteToLogRlh("Just after Do While Loop", "")
    
    
    If itemFound = True Then
        Call modLog.WriteToLogRlh("itemFound", "true")
         
        Call modLog.WriteToLogRlh("(BEFORE)Attempting to remove contact item: ", CStr(itemCount))
        contactItems.Remove (itemCount)
         
        Call modLog.WriteToLogRlh("(AFTER)Attempting to remove contact item: ", CStr(itemCount))
    End If
    
    Set contactItems = Nothing
    
    If (destroyObjects = False) Then
        'don't clean the objects
    Else
        objNamespace.Logoff
        Set objNamespace = Nothing
        Set objOutlook = Nothing
    End If
    
    RemoveContact = returnValue
    'MsgBox ("Normal Exit of Function")
    
    Call modLog.WriteToLogRlh("(Normal Exit of Function", "")
    Exit Function
    
errorHandler:
    returnValue = False
    RemoveContact = returnValue
    If (strErrorMessage = "") Then
        strErrorMessage = "CFaxAndEMail:RemoveContact" & "- " & Err.Number & "- " & Err.Description
    End If
    Call Err.Raise(1004, "CFaxAndEMail:RemoveContact", strErrorMessage)
    
End Function

Public Function FormatFaxNumber(inputFaxNumber As String) As String
 '***************************************************************
 'Format the input fax number to fit the OutLook fax number format
 '***************************************************************
 Dim outputFaxNumber As String
 Dim tempFaxNumber As String
 
 outputFaxNumber = inputFaxNumber
 
 tempFaxNumber = Replace(inputFaxNumber, "-", "")
 tempFaxNumber = Replace(tempFaxNumber, " ", "")
 tempFaxNumber = Replace(tempFaxNumber, "(", "")
 tempFaxNumber = Replace(tempFaxNumber, ")", "")
 
 If (Len(tempFaxNumber) = 10) Then
    outputFaxNumber = "+1 (" & Mid(tempFaxNumber, 1, 3) & ") " & Right(tempFaxNumber, 7)
 End If
 
 FormatFaxNumber = outputFaxNumber
 Exit Function

End Function

Public Property Let ErrorMessage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
    strErrorMessage = vData
End Property

Public Property Get ErrorMessage() As String
'used when retrieving value of a property, on the right side of an assignment.
    ErrorMessage = strErrorMessage
End Property

