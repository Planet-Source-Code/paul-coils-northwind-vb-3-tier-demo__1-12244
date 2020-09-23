VERSION 1.0 CLASS
   BEGIN
   MultiUse = -1            'True
   Persistable = 0          'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "cgLevel", "1-V1PLC"
Attribute VB_Ext_KEY = "cgDate", "18-Sep-2000 13:19:13"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6", "Yes"
Attribute VB_Ext_KEY = "Top_Level", "No"
Option Explicit

'Class constants
Const C_CLASSNAME = "clsCustomer"

'Date format for database inserts
Const C_DATE_FORMAT As String = "dd-mmm-yyyy"

'Special object properties
Private sIndexKey As String

'Internal object variables
Private sCustomerID As String
Private sCompanyName As String
Private sContactName As Variant
Private sContactTitle As Variant
Private sAddress As Variant
Private sCity As Variant
Private sRegion As Variant
Private sPostalCode As Variant
Private sCountry As Variant
Private sPhone As Variant
Private sFax As Variant

Friend Property Let IndexKey(ByVal vData As String)

   'Allow objects within this project to set this option using friend
   sIndexKey = vData

End Property

Public Property Get IndexKey() As String
Attribute Item.VB_UserMemId = 0

   'Any objects can read this value
   IndexKey = sIndexKey

End Property

Public Property Get CustomerID() As String
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   CustomerID = sCustomerID
End Property

Public Property Let CustomerID(ByVal vData As String)
   sCustomerID = vData
End Property

Public Property Get CompanyName() As String
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   CompanyName = sCompanyName
End Property

Public Property Let CompanyName(ByVal vData As String)
   sCompanyName = vData
End Property

Public Property Get ContactName() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   ContactName = sContactName
End Property

Public Property Let ContactName(ByVal vData As Variant)
   sContactName = vData
End Property

Public Property Get ContactTitle() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   ContactTitle = sContactTitle
End Property

Public Property Let ContactTitle(ByVal vData As Variant)
   sContactTitle = vData
End Property

Public Property Get Address() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   Address = sAddress
End Property

Public Property Let Address(ByVal vData As Variant)
   sAddress = vData
End Property

Public Property Get City() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   City = sCity
End Property

Public Property Let City(ByVal vData As Variant)
   sCity = vData
End Property

Public Property Get Region() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   Region = sRegion
End Property

Public Property Let Region(ByVal vData As Variant)
   sRegion = vData
End Property

Public Property Get PostalCode() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   PostalCode = sPostalCode
End Property

Public Property Let PostalCode(ByVal vData As Variant)
   sPostalCode = vData
End Property

Public Property Get Country() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   Country = sCountry
End Property

Public Property Let Country(ByVal vData As Variant)
   sCountry = vData
End Property

Public Property Get Phone() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   Phone = sPhone
End Property

Public Property Let Phone(ByVal vData As Variant)
   sPhone = vData
End Property

Public Property Get Fax() As Variant
Attribute AssignedToDate.VB_Description = "The purpose of this property is to.."
   Fax = sFax
End Property

Public Property Let Fax(ByVal vData As Variant)
   sFax = vData
End Property

Public Function Update() As Boolean

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim oBusCustomers As NorthwindBus.clsBusCustomers

   'By default function fails
   Update = False

   'Create object
   Set oBusCustomers = New NorthwindBus.clsBusCustomers

   'Update database
   Call oBusCustomers.Update( sCustomerID, sCompanyName, sContactName, sContactTitle, sAddress, sCity, sRegion,  _
      sPostalCode, sCountry, sPhone, sFax )

   'Success and exit
   Update = True
   Set oBusCustomers = Nothing
   Exit Function

ErrorHandler:

   'Bubble up error
   Set oBusCustomers = Nothing
   Call RaiseError(C_CLASSNAME, "Update")

End Function

Public Function Insert() As Boolean

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim oBusCustomers As NorthwindBus.clsBusCustomers

   'By default function fails
   Insert = False

   'Create object
   Set oBusCustomers = New NorthwindBus.clsBusCustomers

   'Insert into database
   Call oBusCustomers.Insert( sCustomerID, sCompanyName, sContactName, sContactTitle, sAddress, sCity, sRegion,  _
      sPostalCode, sCountry, sPhone, sFax )

   'Success and exit
   Insert = True
   Set oBusCustomers = Nothing
   Exit Function

ErrorHandler:

   'Bubble up error
   Set oBusCustomers = Nothing
   Call RaiseError(C_CLASSNAME, "Insert")

End Function

Public Function Delete() As Boolean

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim oBusCustomers As NorthwindBus.clsBusCustomers

   'By default function fails
   Delete = False

   'Create object
   Set oBusCustomers = New NorthwindBus.clsBusCustomers

   'Delete from database
   Call oBusCustomers.Delete( sCustomerID )

   'Success and exit
   Delete = True
   Set oBusCustomers = Nothing
   Exit Function

ErrorHandler:

   'Bubble up error
   Set oBusCustomers = Nothing
   Call RaiseError(C_CLASSNAME, "Delete")

End Function

Public Function Populate(sCustomerID As String) As Boolean

   'Error handling
   On Error GoTo ErrorHandler

   'Recordset 
   Dim RS As ADODB.Recordset
   Dim oBusCustomers As NorthwindBus.clsBusCustomers

   'By default function fails
   Populate = False

   'Create object
   Set oBusCustomers = New NorthwindBus.clsBusCustomers

   'Read Data
   Set RS = oBusCustomers.Populate( sCustomerID )
   Set oBusCustomers = Nothing

   If RS Is Nothing Then
      Exit Function
   End If

   'Check for EOF
   If RS.EOF = False Then

      'Assign properties
      sCustomerID = RS("CustomerID")
      sCompanyName = RS("CompanyName")
      sContactName = RS("ContactName")
      sContactTitle = RS("ContactTitle")
      sAddress = RS("Address")
      sCity = RS("City")
      sRegion = RS("Region")
      sPostalCode = RS("PostalCode")
      sCountry = RS("Country")
      sPhone = RS("Phone")
      sFax = RS("Fax")

   End If
   
   'Clean up
   RS.Close
   Set RS = Nothing

   'Success and exit
   Populate = True
   Exit Function

ErrorHandler:

   'Clean up
   Set oBusCustomers = Nothing
   RS.Close
   Set RS = Nothing
   Call RaiseError(C_CLASSNAME, "Populate")

End Function

