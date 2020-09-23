VERSION 1.0 CLASS
   BEGIN
   MultiUse = -1            'True
   Persistable = 0          'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 3   'UsesTransaction
END
Attribute VB_Name = "clsDbCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "cgLevel", "1-V1PLC"
Attribute VB_Ext_KEY = "cgDate", "18-Sep-2000 13:19:13"
Option Explicit

'Class constants
Const C_CLASSNAME = "clsDbCustomers"

'Date format for database inserts
Const C_DATE_FORMAT As String = "dd-mmm-yyyy"

'Data access object
Private moData As clsDataAccess

Private Sub Class_Initialize()

   'Create data access wrapper
   Set moData = New clsDataAccess

End Sub

Private Sub Class_Terminate()

   'Clean up data access wrapper
   Set moData = Nothing

End Sub

Public Function Delete( ByVal sCustomerID As String)

   'Error handling
   On Error GoTo ErrorHandler

   'Delete from database
   Call moData.Execute("NW_DelCustomer", "", QO_UseStoredProc, _
            Array("@CustomerID", adWChar, 5, sCustomerID ))
   
   'Commit/exit
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "Delete")

End Function

Public Function Insert( ByVal sCustomerID As String, ByVal sCompanyName As String, ByVal sContactName As Variant, ByVal sContactTitle As Variant,  _
   ByVal sAddress As Variant, ByVal sCity As Variant, ByVal sRegion As Variant, ByVal sPostalCode As Variant,  _
   ByVal sCountry As Variant, ByVal sPhone As Variant, ByVal sFax As Variant)

   'Error handling
   On Error GoTo ErrorHandler

   'Add to database
   Call moData.Execute("NW_InsCustomer", "", QO_UseStoredProc, _
            Array("@CustomerID", adWChar, 5, sCustomerID ), _
            Array("@CompanyName", adVarWChar, 40, sCompanyName ), _
            Array("@ContactName", adVarWChar, 30, ParamField( sContactName, False )), _
            Array("@ContactTitle", adVarWChar, 30, ParamField( sContactTitle, False )), _
            Array("@Address", adVarWChar, 60, ParamField( sAddress, False )), _
            Array("@City", adVarWChar, 15, ParamField( sCity, False )), _
            Array("@Region", adVarWChar, 15, ParamField( sRegion, False )), _
            Array("@PostalCode", adVarWChar, 10, ParamField( sPostalCode, False )), _
            Array("@Country", adVarWChar, 15, ParamField( sCountry, False )), _
            Array("@Phone", adVarWChar, 24, ParamField( sPhone, False )), _
            Array("@Fax", adVarWChar, 24, ParamField( sFax, False )))
   
   'Commit/exit
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "Insert")

End Function

Public Function Update( ByVal sCustomerID As String, ByVal sCompanyName As String, ByVal sContactName As Variant, ByVal sContactTitle As Variant,  _
   ByVal sAddress As Variant, ByVal sCity As Variant, ByVal sRegion As Variant, ByVal sPostalCode As Variant,  _
   ByVal sCountry As Variant, ByVal sPhone As Variant, ByVal sFax As Variant)

   'Error handling
   On Error GoTo ErrorHandler

   'Update database
   Call moData.Execute("NW_UpdCustomer", "", QO_UseStoredProc, _
            Array("@CustomerID", adWChar, 5, sCustomerID ), _
            Array("@CompanyName", adVarWChar, 40, sCompanyName ), _
            Array("@ContactName", adVarWChar, 30, ParamField( sContactName, False )), _
            Array("@ContactTitle", adVarWChar, 30, ParamField( sContactTitle, False )), _
            Array("@Address", adVarWChar, 60, ParamField( sAddress, False )), _
            Array("@City", adVarWChar, 15, ParamField( sCity, False )), _
            Array("@Region", adVarWChar, 15, ParamField( sRegion, False )), _
            Array("@PostalCode", adVarWChar, 10, ParamField( sPostalCode, False )), _
            Array("@Country", adVarWChar, 15, ParamField( sCountry, False )), _
            Array("@Phone", adVarWChar, 24, ParamField( sPhone, False )), _
            Array("@Fax", adVarWChar, 24, ParamField( sFax, False )))
   
   'Commit/exit
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "Update")

End Function

Public Function Populate( ByVal sCustomerID As String) As ADODB.Recordset

   'Error handling
   On Error GoTo ErrorHandler

   'Return records from database
   Set Populate = moData.GetRecordset("NW_SelCustomer", "", QO_UseStoredProc, _
            Array("@CustomerID", adWChar, 5, sCustomerID ))
   
   'Commit/exit
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "Populate")

End Function

Public Function PopulateAll() As ADODB.Recordset

   'Error handling
   On Error GoTo ErrorHandler

   'Load all records from database
   Set PopulateAll = moData.GetRecordset("NW_SelAllCustomers", "", QO_UseStoredProc)
   
   'Commit/exit
   Call CtxSetComplete
   Exit Function

ErrorHandler:

   'Clean up
   Call CtxSetAbort
   Call RaiseError(C_CLASSNAME, "PopulateAll")

End Function

