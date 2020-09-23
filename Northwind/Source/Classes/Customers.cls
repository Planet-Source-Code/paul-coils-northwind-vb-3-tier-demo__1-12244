VERSION 1.0 CLASS
   BEGIN
   MultiUse = -1            'True
   Persistable = 0          'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "cgLevel", "1-V1PLC"
Attribute VB_Ext_KEY = "cgDate", "18-Sep-2000 13:19:13"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6", "Yes"
Attribute VB_Ext_KEY = "Collection", "UserDefinedCollection"
Attribute VB_Ext_KEY = "Member0", "clsCustomer"
Attribute VB_Ext_KEY = "Top_Level", "Yes"
Option Explicit

'Class constants
Const C_CLASSNAME = "clsCustomers"

'Collection of items
Private colCustomers As Collection 

Public Sub PopulateAll()

   'Recordset 
   Dim RS As ADODB.Recordset

   'Create object
   Dim oBusCustomers As NorthwindBus.clsBusCustomers
   Set oBusCustomers = New NorthwindBus.clsBusCustomers

   'Read Data
   Set RS = oBusCustomers.PopulateAll()
   Set oBusCustomers = Nothing

   'Load all records
   Call Load( RS )
   Set RS = Nothing

End Sub

Private Sub Load(RS As ADODB.Recordset)

   'Varibles
   Dim oItem As clsCustomer
   
   'Fields
   Dim fldCustomerID As ADODB.Field
   Dim fldCompanyName As ADODB.Field
   Dim fldContactName As ADODB.Field
   Dim fldContactTitle As ADODB.Field
   Dim fldAddress As ADODB.Field
   Dim fldCity As ADODB.Field
   Dim fldRegion As ADODB.Field
   Dim fldPostalCode As ADODB.Field
   Dim fldCountry As ADODB.Field
   Dim fldPhone As ADODB.Field
   Dim fldFax As ADODB.Field

   'Error handling
   On Error GoTo ErrorHandler

   If RS Is Nothing Then
      Exit Sub
   End If

   'Bind fields to recordset
   Set fldCustomerID = RS("CustomerID")
   Set fldCompanyName = RS("CompanyName")
   Set fldContactName = RS("ContactName")
   Set fldContactTitle = RS("ContactTitle")
   Set fldAddress = RS("Address")
   Set fldCity = RS("City")
   Set fldRegion = RS("Region")
   Set fldPostalCode = RS("PostalCode")
   Set fldCountry = RS("Country")
   Set fldPhone = RS("Phone")
   Set fldFax = RS("Fax")

   'Loop for each record
   Do While Not RS.EOF

      'Build item
      Set oItem = New clsCustomer

      'Assign properties
      oItem.CustomerID = fldCustomerID
      oItem.CompanyName = fldCompanyName
      oItem.ContactName = fldContactName
      oItem.ContactTitle = fldContactTitle
      oItem.Address = fldAddress
      oItem.City = fldCity
      oItem.Region = fldRegion
      oItem.PostalCode = fldPostalCode
      oItem.Country = fldCountry
      oItem.Phone = fldPhone
      oItem.Fax = fldFax

      'Add item to collection
      Call Me.Add( oItem )
      
      'Next record
      RS.MoveNext

   Loop

   'Exit
   Exit Sub

ErrorHandler:

   Call RaiseError(C_CLASSNAME, "Load")

End Sub

Public Function Delete(vIndexKey As Variant) As Boolean

   Dim oItem As clsCustomer

   'Current item
   Set oItem = colCustomers(vIndexKey)

   'Delete current item from disk and from collection
   oItem.Delete
   Set oItem = Nothing
   Call colCustomers.Remove(vIndexKey)

   Delete = True

End Function

Public Function Insert(oItem As clsCustomer) As Boolean

   'Add item to collection
   Call Me.Add(oItem)

   'Insert item
   Call oItem.Insert

End Function

Public Function Update(vIndexKey As Variant) As Boolean

   Dim oItem As clsCustomer

   'Current item
   Set oItem = colCustomers(vIndexKey)

   'Update item
   Call oItem.Update

End Function

Private Sub Class_Initialize()

   'Create collections when class created
   Set colCustomers = New Collection

End Sub

Private Sub Class_Terminate()

   'Destroy collections when class terminated
   Set colCustomers = Nothing

End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags ="40"

   'Enumerate collection to enable the For...Each syntax
   Set NewEnum = colCustomers.[_NewEnum]

End Property

Public Property Get Count() As Long

   'Number of items in collection
   Count = colCustomers.Count

End Property

Public Property Get Item(vIndexKey As Variant) As clsCustomer
Attribute Item.VB_UserMemId = 0

   'Used to reference an element in the collection
   Set Item = colCustomers(vIndexKey)

End Property

Public Function Exists( vIndexKey as Variant) As Boolean

   'Error handling
   On Error GoTo KeyDoesNotExist

   'Attempt to reference item
   Dim oItem As clsCustomer
   Set oItem = colCustomers( vIndexKey ) 

   'If referenced ok - no error
   Exists = True
   Exit Function

KeyDoesNotExist:
   Err.Clear
   Exists = False

End Function

Public Function Add(oItem As clsCustomer, Optional sIndexKey As String) As clsCustomer

   'No key has been defined for this classes collection
   If Len(sIndexKey) = 0 Then
      oItem.IndexKey = oItem.CustomerID
   Else
      oItem.IndexKey = sIndexKey
   End If

   'Add item
   colCustomers.Add oItem, oItem.IndexKey

   'Return the object
   Set Add = oItem
   Exit Function

End Function

