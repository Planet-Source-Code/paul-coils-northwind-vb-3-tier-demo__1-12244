Attribute VB_Name = "modDbLibrary"
Option Explicit

'--- Helper function to build queries for database - uses the SQL keyword
'--- NULL without quotes (by default) if the variable is of zero length or
'--- contains a NULL value.  Blank strings may optionally be left untouched
Public Function ParamField(vData As Variant, Optional bBlankAsNull As Boolean = True, Optional vDefault As Variant = "NULL") As Variant

   'If data is null
   If IsNull(vData) Or (vData = vbNullString And bBlankAsNull) Then
      ParamField = vDefault
   Else
      ParamField = vData
   End If

End Function
