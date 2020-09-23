VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'*******************************************************************
   ' Distribution List Functions
   '-------------------------------------------------------------------
   '
   ' Functions and Procedures in this Module
   '*******************************************************************
   ' 1. AddressBookExists: Determines if a specific AddressList Exists
   '                       - for example, PAB, GAL, and etc.
   ' 2. AddDLToPAB:        Creates a new Distribution List (DL).
   ' 3. AddUserToDL:       Used to add a new recipient to a DL.
   ' 4. AddUserToPAB:      Add a new user to the MAPI address list.
   ' 5. DeleteUserFromDL:  Removes a recipient from a DL. Also contains
   '                       logic to check if a specific recipient is on
   '                       a DL.
   ' 6. GetDLOwner:        Get the Owner of a Distribution List.
   '
   '*******************************************************************
   'REMEMBER: You still need the pertinent rights to be able to
   'manipulate MAPI AddressList objects.
   '
   '*******************************************************************
   Public objSession As MAPI.Session
   
'local variable(s) to hold property value(s)
Private mvarNewUserName As String 'local copy
Private mvarDistributionList As String 'local copy
Private mvarEmail As String 'local copy
Public Sub Clear()
 mvarNewUserName = ""
 mvarDistributionList = ""
 mvarEmail = ""
End Sub

Public Property Let Email(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Email = 5
    mvarEmail = vData
End Property


Public Property Get Email() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Email
    Email = mvarEmail
End Property



Public Property Let DistributionList(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DistributionList = 5
    mvarDistributionList = vData
End Property


Public Property Get DistributionList() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DistributionList
    DistributionList = mvarDistributionList
End Property



Public Property Let NewUserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewUserName = 5
    mvarNewUserName = vData
End Property


Public Property Get NewUserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewUserName
    NewUserName = mvarNewUserName
End Property



   '*******************************************************************
   ' A conglomerate function that demonstrates use of several of the
   ' other functions in this module. Look through these functions for
   ' the string "TO DO:". The functions will work as is, but will not
   ' yield meaningful content until they are customized as noted on
   ' these lines.
   '*******************************************************************
   Public Sub Main()
     Dim oAddressList As AddressList
     Dim oAddressEntry As AddressEntry
     Dim oNewUserAddressEntry As AddressEntry
     If MapiLogon() = True Then
        Set oAddressList = AddressBookExists("Personal Address Book")
        If Not oAddressList Is Nothing Then
          'TO DO: You will want to customize the string values
          'in the next two lines.
           Set oNewUserAddressEntry = AddUserToPAB(oAddressList, mvarNewUserName)
           Set oAddressEntry = AddDLToPAB(oAddressList, mvarDistributionList)

           If Not oAddressEntry Is Nothing Then
             'TO DO: Recipient to be added on next line should be
             'customized.
             If AddUserToDL(oAddressEntry, mvarNewUserName) = True Then
                MsgBox "Added new user and DL, and new member in the DL"
             Else
                MsgBox "Adding a user to the DL didn't work"
             End If
           Else
               MsgBox "Custom DL wasn't added to the Personal Address Book"
           End If
           Exit Sub
          Else
            MsgBox "There is no PAB Service installed on this machine"
            Exit Sub
        End If
       End If
     End Sub

   '*******************************************************************
   ' FUNCTION     - AddressBookExists
   '-------------------------------------------------------------------
   ' PARAMETERS   - sAddressBookName as string and can be one of two
   '                values:
   '               "Personal Address Book" or "Global AddressList".
   ' DESCRIPTION  - Determines if the Personal Address Book Service is
   '                installed on the current computer.
   ' RETURN VALUE - An AddressList object pointing to the addresslist or
   '                nothing if it does not exist.
   '*******************************************************************
   Public Function AddressBookExists(sAddressBookName As String) _
   As AddressList
      Dim oAddressList As AddressList
     
      For Each oAddressList In objSession.AddressLists
      If oAddressList.Name = sAddressBookName Then
      Set AddressBookExists = oAddressList
      Exit Function
      End If
      Next oAddressList
      Set AddressBookExists = Nothing
      End Function

   '*******************************************************************
   ' FUNCTION     - AddDLToPAB
   '-------------------------------------------------------------------
   ' PARAMETERS   - oAddressList as AddressList, sDLName As String (The
   '                name of the new Distribution List to create.)
   ' DESCRIPTION  - Add a new Distribution List to the MAPI address list
   '                specified in first parameter.
   ' RETURN VALUE - An AddressEntry object pointing to the address entry
   '                or nothing if it does not exist.
   '*******************************************************************
   Public Function AddDLToPAB(oAddressList As AddressList, _
   sDLName As String) As AddressEntry
      Dim oAddressEntries As AddressEntries
      Dim oNewAddressEntry As AddressEntry
      On Error GoTo Trp_AddDLToPAB:
      Set oAddressEntries = oAddressList.AddressEntries
      Set oNewAddressEntry = oAddressEntries.Add("MAPIPDL", sDLName)
      oNewAddressEntry.Update
      Set AddDLToPAB = oNewAddressEntry
      Exit Function
Trp_AddDLToPAB:
      'Set Trp_AddDLToPAB = Nothing
      Exit Function
   End Function

   '*******************************************************************
   ' FUNCTION     - AddUserToDL
   '-------------------------------------------------------------------
   ' PARAMETERS   - oDL As AddressEntry (This is the object reference to
   '                the Distribution List the user is being added to.),
   '                sUsername As String (The name of the new member to
   '                be added to the Distribution List.)
   ' DESCRIPTION  - Adds a new recipient to a Distribution List.
   ' RETURN VALUE - A Boolean value specifying whether or not adding the
   '                user was successful.
   '*******************************************************************
   Public Function AddUserToDL(oDL As AddressEntry, sUserName As String) _
   As Boolean
   Dim oNewMember As AddressEntry
   On Error GoTo Trp_AddUserToDL:
   'TO DO: Customize for your new DL member.
   Set oNewMember = oDL.Members.Add("SMTP", mvarNewUserName)

   'You need to fill in all your own data here. The data below is
    'for sample purposes. This is not an inclusive list of potential
    'fields, and of the fields represented here, not all must be
    'populated for a viable address entry.
    With oNewMember
      'TO DO: Obviously all of these should be modified.
      .Address = mvarEmail
      ''Generic Comments Field.
      '.Fields(ActMsgPR_COMMENT) = "My notes here."
      ''Business Generic Fields.
      '.Fields(ActMsgPR_COMPANY_NAME) = "Microsoft Corporation"
      '.Fields(ActMsgPR_DEPARTMENT_NAME) = "Microsoft Technical Support"
      '.Fields(ActMsgPR_MANAGER_NAME) = "John Doe"
      '.Fields(ActMsgPR_ASSISTANT) = "Jane Doe"
      '.Fields(ActMsgPR_GIVEN_NAME) = "Bill"
      '.Fields(ActMsgPR_MIDDLE_NAME) = "A."
      '.Fields(ActMsgPR_SURNAME) = "Smith"
      '.Fields(ActMsgPR_TITLE) = "Mr."
      '.Fields(ActMsgPR_EMAIL_ADDRESS) = "mymail@Microsoft.com"
      ''Business Address Fields.
      '.Fields(ActMsgPR_OFFICE_LOCATION) = "Building 13"
      '.Fields(ActMsgPR_BUSINESS_ADDRESS_STREET) = "1 Microsoft Way"
      '.Fields(ActMsgPR_BUSINESS_ADDRESS_CITY) = "Redmond"
      '.Fields(ActMsgPR_BUSINESS_ADDRESS_STATE_OR_PROVINCE) = "WA"
      '.Fields(ActMsgPR_BUSINESS_ADDRESS_POSTAL_CODE) = "98052"
      '.Fields(ActMsgPR_BUSINESS_ADDRESS_COUNTRY) = "USA"
      ''Telephone Number Fields.
      '.Fields(ActMsgPR_BUSINESS_FAX_NUMBER) = "425-555-0329"
      '.Fields(ActMsgPR_BUSINESS_TELEPHONE_NUMBER) = "425-555-8080"
      '.Fields(ActMsgPR_BUSINESS2_TELEPHONE_NUMBER) = "425-555-8081"
      '.Fields(ActMsgPR_CALLBACK_TELEPHONE_NUMBER) = "425-555-8082"
      '.Fields(ActMsgPR_CAR_TELEPHONE_NUMBER) = "206-555-0000"
      '.Fields(ActMsgPR_ASSISTANT_TELEPHONE_NUMBER) = "425-555-8083"
      '.Fields(ActMsgPR_COMPANY_MAIN_PHONE_NUMBER) = "425-882-8080"
      '.Fields(ActMsgPR_MOBILE_TELEPHONE_NUMBER) = "425-555-8084"
      '.Fields(ActMsgPR_PAGER_TELEPHONE_NUMBER) = "N/A"
      '.Fields(ActMsgPR_PRIMARY_FAX_NUMBER) = "425-555-8085"
      ''Home detail fields.
      '.Fields(ActMsgPR_HOME_ADDRESS_STREET) = "1234 My Street"
      '.Fields(ActMsgPR_HOME_ADDRESS_CITY) = "MyTown"
      '.Fields(ActMsgPR_HOME_ADDRESS_STATE_OR_PROVINCE) = "WA"
      '.Fields(ActMsgPR_HOME_ADDRESS_COUNTRY) = "USA"
      '.Fields(ActMsgPR_HOME_FAX_NUMBER) = "425-555-0001"
      '.Fields(ActMsgPR_HOME_TELEPHONE_NUMBER) = "425-555-0002"
      '.Fields(ActMsgPR_HOME2_TELEPHONE_NUMBER) = "N/A"

      .Update
     End With

     AddUserToDL = True
     Exit Function
Trp_AddUserToDL:
     AddUserToDL = False
     Exit Function
   End Function

   '*********************************************************************
   ' FUNCTION     - AddUserToPAB
   '---------------------------------------------------------------------
   ' PARAMETERS   - oAddressList as AddressList, sDLName As String (The
   '                name of the new Distribution List to create.)
   ' DESCRIPTION  - Add a new user to the MAPI address list specified in
   '                first parameter.
   ' RETURN VALUE - A As AddressEntry object pointing to the address entry
   '                or nothing if it does not exist.
   '*********************************************************************
   Public Function AddUserToPAB(oAddressList As AddressList, sUserName _
    As String) As AddressEntry
    Dim oAddressEntries As AddressEntries
    Dim oNewAddressEntry As AddressEntry
    On Error GoTo Trp_AddDLToPAB:
    Set oAddressEntries = oAddressList.AddressEntries
    Set oNewAddressEntry = oAddressEntries.Add("SMTP", sUserName)
    'TO DO: Change address in next line.
    oNewAddressEntry.Address = mvarEmail
    oNewAddressEntry.Update
    Set AddUserToPAB = oNewAddressEntry
    Exit Function
Trp_AddDLToPAB:
    'Set Trp_AddDLToPAB = Nothing
    Exit Function
    End Function

   '*********************************************************************
   ' PROCEDURE    - DeleteUserFromDL
   '---------------------------------------------------------------------
   ' PARAMETERS   - None
   ' DESCRIPTION  - Remove a member from a Distribution List, which the
   '                user selects from the Phonebook.
   '*********************************************************************
   Public Sub DeleteUserFromDL()
      'Procedure Level Variables.
       Dim oAddressEntries As AddressEntries
       Dim oAddressList As AddressList
       Dim oRecipients As Recipients
       Dim oRecipient As Recipient
       Dim oAddressEntry As AddressEntry
       Dim oMember As AddressEntry
       'Get the DL from the AddressBook.
       Set oRecipients = objSession.AddressBook(Title:="Select Attendees")
       If oRecipients.Count = 0 Or oRecipients.Count > 1 Then
          MsgBox "Please Select just one member of the Distribution List"
          Exit Sub
       End If
       Set oRecipient = oRecipients.Item(1)
       Set oAddressEntry = oRecipient.AddressEntry
       Set oAddressEntries = oAddressEntry.Members
       'Cycle through the DL until you find the user you seek.
       For Each oMember In oAddressEntry.Members
          If oMember.IsSameAs(oRecipient) Then
             'Or you could say If oMember.Name = "TheRecipsDisplayName"
             oMember.Delete
             Set oMember = Nothing
          End If
       Next oMember
    End Sub

   '*********************************************************************
   ' FUNCTION     - GetDLOwner
   '---------------------------------------------------------------------
   ' PARAMETERS   - oSubordinate As AddressEntry (The object reference to
   '                the Owner of the DL.)
   ' DESCRIPTION  - Gets an address entry's Owner. This is relevant only
   '                to the Global Address List since you are always the
   '                owner of Personal Address Book (PAB) entries.
   ' RETURN VALUE - A String value of the owners name.
   '*********************************************************************

    Public Function GetDLOwner(oSubordinate As AddressEntry) As String
     ' IMPORTANT: This code assumes that the Owner of the
     ' Distribution list is stored in PR_EMS_AB_EXTENSION_ATTRIBUTE_3.
     ' This is not always where the Owner is stored, if it is stored
     ' at all. You should verify where the Owner is stored on your
     ' system and change the Field value that is being retrieved by
     ' this code.
      GetDLOwner = oSubordinate.Fields(&H802F001E).Value
    End Function

   '*********************************************************************
   ' FUNCTION     - MapiLogon
   '---------------------------------------------------------------------
   ' PARAMETERS   - None.
   ' DESCRIPTION  - Creates and Logs onto a MAPI Session according to the
   '                parameters passed when calling the Logon method. The
   '                sample below passes no parameters and uses all default
   '                values.
   ' RETURN VALUE - Boolean indicating success.
   '*********************************************************************
   Public Function MapiLogon() As Boolean
     'Create a session and log on -- username and password in profile
      Set objSession = CreateObject("MAPI.Session")
      'Change the parameters to valid values for your configuration.
       objSession.Logon
       MapiLogon = True
    End Function





