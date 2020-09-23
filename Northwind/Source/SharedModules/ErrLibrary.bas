Attribute VB_Name = "modErrLibrary"
Option Explicit

Public Sub RaiseError(sModule As String, sFunction As String, Optional sAlternateDesc As Variant)

   'If alternate error description provided, use instead of Err.Description.
   'It is often preferred to provide a more user friendly error description
   'and just log the actual technical message.  If you are calling this method
   'from a lower layer you may wish to retain the technical description and
   'do the translation in application layer.
   
   'Raise error back to the client with sAlternateDesc or Err.Description
   Err.Raise Err.Number, SetErrSource(sModule, sFunction), IIf(IsMissing(sAlternateDesc), Err.Description, sAlternateDesc)
   
End Sub
   
Public Function GetErrorTextFromResource(lErrorNumber As Long) As String

   'Get the string from a resource file
   GetErrorTextFromResource = LoadResString(lErrorNumber)

End Function

Function SetErrSource(sModule As String, sFunction As String) As String

   'Returns the source of the error -
   'Format:   [Library.Class] FunctionName [on ComputerName Version Major.Minor.Revision]
   'Example:  [JobTracking.Jobs] Insert [on FALCON Version 1.0.0]
   SetErrSource = Err.Source & "[" & sModule & "]  " & sFunction & _
      " [on " & GetComputerName() & " Version " & GetVersionNumber() & "]"

End Function
   
Function GetVersionNumber() As String

   'Return application version
   GetVersionNumber = App.Major & "." & App.Minor & "." & App.Revision

End Function
