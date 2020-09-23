VERSION 1.0 CLASS
   BEGIN
   MultiUse = -1            'True
   Persistable = 0          'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 3   'UsesTransaction
END
Attribute VB_Name = "clsBusCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Class constants
Const C_CLASSNAME = "clsBusCustomers"

'Date format for database inserts
Const C_DATE_FORMAT As String = "dd-mmm-yyyy"

Public Function Populate(ByVal sCustomerID As String) As ADODB.Recordset

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim oDbCustomers As NorthwindData.clsDbCustomers

   'Create data-layer object
   Set oDbCustomers = CtxCreateObject("NorthwindData.clsDbCustomers")

   'Read Data & Clear Object
   Set Populate = oDbCustomers.Populate( sCustomerID )
   
   'Commit/exit
   Set oDbCustomers = Nothing
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Set oDbCustomers = Nothing
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "Populate")

End Function

Public Function PopulateAll() As ADODB.Recordset

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim oDbCustomers As NorthwindData.clsDbCustomers

   'Create data-layer object
   Set oDbCustomers = CtxCreateObject("NorthwindData.clsDbCustomers")

   'Read Data & Clear Object
   Set PopulateAll = oDbCustomers.PopulateAll()
   
   'Commit/exit
   Set oDbCustomers = Nothing
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Set oDbCustomers = Nothing
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "PopulateAll")

End Function

Public Function Delete(ByVal sCustomerID As String)

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim oDbCustomers As NorthwindData.clsDbCustomers

   'Create data-layer object
   Set oDbCustomers = CtxCreateObject("NorthwindData.clsDbCustomers")

   'Delete from database
   Call oDbCustomers.Delete( sCustomerID )
   
   'Commit/exit
   Set oDbCustomers = Nothing
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Set oDbCustomers = Nothing
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "Delete")

End Function

Public Function Insert( ByVal sCustomerID As String, ByVal sCompanyName As String, ByVal sContactName As Variant, ByVal sContactTitle As Variant,  _
   ByVal sAddress As Variant, ByVal sCity As Variant, ByVal sRegion As Variant, ByVal sPostalCode As Variant,  _
   ByVal sCountry As Variant, ByVal sPhone As Variant, ByVal sFax As Variant)

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim oDbCustomers As NorthwindData.clsDbCustomers

   'Create data-layer object
   Set oDbCustomers = CtxCreateObject("NorthwindData.clsDbCustomers")

   'Insert into database
   Call oDbCustomers.Insert( sCustomerID, sCompanyName, sContactName, sContactTitle, sAddress, sCity, sRegion,  _
      sPostalCode, sCountry, sPhone, sFax )
   
   'Commit/exit
   Set oDbCustomers = Nothing
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Set oDbCustomers = Nothing
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "Insert")

End Function

Public Function Update(ByVal sCustomerID As String, ByVal sCompanyName As String, ByVal sContactName As Variant, ByVal sContactTitle As Variant,  _
   ByVal sAddress As Variant, ByVal sCity As Variant, ByVal sRegion As Variant, ByVal sPostalCode As Variant,  _
   ByVal sCountry As Variant, ByVal sPhone As Variant, ByVal sFax As Variant)

   'Error handling
   On Error GoTo ErrorHandler

   'Variables
   Dim oDbCustomers As NorthwindData.clsDbCustomers

   'Create data-layer object
   Set oDbCustomers = CtxCreateObject("NorthwindData.clsDbCustomers")

   'Update database
   Call oDbCustomers.Update( sCustomerID, sCompanyName, sContactName, sContactTitle, sAddress, sCity, sRegion,  _
      sPostalCode, sCountry, sPhone, sFax )
   
   'Commit/exit
   Set oDbCustomers = Nothing
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Set oDbCustomers = Nothing
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "Update")

End Function

