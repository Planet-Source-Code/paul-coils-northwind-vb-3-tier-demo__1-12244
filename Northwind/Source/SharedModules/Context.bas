Attribute VB_Name = "modContext"
Option Explicit

'These context wrapping functions are provided so that the project will
'run within Visual Basic without having to compile and register the layer
'objects as DLL's in MTS/COM+

'You have the option to generate the code with or without these wrappers.

Public Sub CtxSetAbort()

   Dim oContext As ObjectContext
   Set oContext = GetObjectContext

   'If MTS function available
   If Not (oContext Is Nothing) Then
      oContext.SetAbort
   End If

   Set oContext = Nothing

End Sub

Public Sub CtxSetComplete()

   Dim oContext As ObjectContext
   Set oContext = GetObjectContext

   'If MTS function available
   If Not (oContext Is Nothing) Then
      oContext.SetComplete
   End If

   Set oContext = Nothing

End Sub

Public Function CtxCreateObject(ByVal sProgID As String) As Object

   Dim oContext As ObjectContext
   Set oContext = GetObjectContext

   'If MTS function available
   If Not (oContext Is Nothing) Then
      Set CtxCreateObject = oContext.CreateInstance(sProgID)
      Set oContext = Nothing
   Else
      Set CtxCreateObject = CreateObject(sProgID)
   End If

End Function

