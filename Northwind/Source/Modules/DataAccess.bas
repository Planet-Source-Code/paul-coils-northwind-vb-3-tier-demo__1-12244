Attribute VB_Name = "modDataAccess"
Option Explicit

'Database connection constant
Public Const C_DB_CONNECTION As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Northwind;Data Source=PAULC\PAULC"

'Data access object
Public oDataAccess As New clsDataAccess
