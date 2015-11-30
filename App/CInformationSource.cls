VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInfoSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarContactId As String 'local copy
Private mvarContactName As String 'local copy
Private mvarContactLastName As String 'local copy
Private mvarContactFirstName As String 'local copy
Private mvarFaxNumber As String 'local copy
Private mvarCompanyName As String 'local copy
Private mvarPhoneNumber As String 'local copy
Private mvarEMailAddress As String 'local copy

Public Sub GetContactInfo(ByVal strContactId As String)
'*************************************************
'Get the Contact Information from the database
'*************************************************
    Dim sql As String
    Dim rec As ADODB.RecordSet
    
    mvarContactId = strContactId
    
    sql = "SELECT * FROM Information_Source WHERE Contact_Id ='" & strContactId & "'"
    g_objDAL.GetRecordset CONNECT, sql, rec
    
    If rec.EOF Then      'invalid contact_id
        mvarCompanyName = ""
        mvarPhoneNumber = ""
        mvarContactName = ""
        mvarEMailAddress = ""
        mvarFaxNumber = ""
        MsgBox "No contacts found for contact id '" & strContactId & "'."
    Else
        rec.MoveFirst
        If Not IsNull(rec.Fields("Fax")) Then
            mvarFaxNumber = Trim(rec.Fields("Fax"))
        Else
            mvarFaxNumber = ""
        End If
        If Not IsNull(rec.Fields("Email")) Then
            mvarEMailAddress = Trim(rec.Fields("Email"))
        Else
            mvarEMailAddress = ""
        End If
        If Not IsNull(rec.Fields("Company_Name")) Then
            mvarCompanyName = Trim(rec.Fields("Company_Name"))
        Else
            mvarCompanyName = ""
        End If
        If Not IsNull(rec.Fields("phone1")) Then
            mvarPhoneNumber = Trim(rec.Fields("phone1"))
        Else
            mvarPhoneNumber = ""
        End If
        
        mvarContactFirstName = Trim(rec.Fields("first_name") & "")
        mvarContactLastName = Trim(rec.Fields("last_name") & "")
        mvarContactName = Trim(rec.Fields("first_name") & "") & " "
        If (IsNull(rec.Fields("middle_initial")) Or (Trim(rec.Fields("middle_initial") & "") = "")) Then
            mvarContactName = mvarContactName & Trim(rec.Fields("last_name") & "")
        Else
            mvarContactName = mvarContactName & Trim(rec.Fields("middle_initial") & "") & " " & Trim(rec.Fields("last_name") & "")
        End If
        If Not (IsNull(rec.Fields("name_suffix")) Or (Trim(rec.Fields("name_suffix") & "") = "")) Then
            mvarContactName = mvarContactName & " " & Trim(rec.Fields("name_suffix"))
        End If
    End If
    
    'MODIFIED 7/8/2005 RTD TO CORRECT PROBLEM WITH MISSING LAST_NAME
    mvarContactName = Trim(mvarContactName)
    
    rec.Close
    Exit Sub
    
Exit_Sub:

End Sub

Public Property Let EMailAddress(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EMailAddress = "5"
    mvarEMailAddress = vData
End Property

Public Property Get EMailAddress() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EMailAddress
    EMailAddress = mvarEMailAddress
End Property



Public Property Let PhoneNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PhoneNumber = 5
    mvarPhoneNumber = vData
End Property

Public Property Get PhoneNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PhoneNumber
    PhoneNumber = mvarPhoneNumber
End Property


Public Property Let CompanyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompanyName = "5"
    mvarCompanyName = vData
End Property

Public Property Get CompanyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompanyName
    CompanyName = mvarCompanyName
End Property



Public Property Let faxNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FaxNumber = 5
    mvarFaxNumber = vData
End Property


Public Property Get faxNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FaxNumber
    faxNumber = mvarFaxNumber
End Property



Public Property Let ContactName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactName = 5
    mvarContactName = vData
End Property

Public Property Get ContactName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactName
    ContactName = mvarContactName
End Property

Public Property Let ContactFirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactName = 5
    mvarContactFirstName = vData
End Property

Public Property Get ContactFirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactName
    ContactFirstName = mvarContactFirstName
End Property

Public Property Let ContactLastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactName = 5
    mvarContactLastName = vData
End Property

Public Property Get ContactLastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactName
    ContactLastName = mvarContactLastName
End Property

Public Property Let ContactID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactId = 5
    mvarContactId = vData
End Property

Public Property Get ContactID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactId
    ContactID = mvarContactId
End Property

