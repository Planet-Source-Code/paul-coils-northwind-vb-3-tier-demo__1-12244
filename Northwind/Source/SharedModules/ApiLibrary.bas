Attribute VB_Name = "modApiLibrary"
Option Explicit

'---
'--- Windows API Call Declarations
'---

'API Call to get the computer name
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" _
     (ByVal lpBuffer As String, nSize As Long) As Long

'---
'--- Windows API Call to return the computer name
'---
Function GetComputerName() As String

   'Set or retrieve the name of the computer.
   Dim sBuffer As String
   Dim lLen As Long

   'Pad string with spaces
   sBuffer = Space(255 + 1)
   lLen = Len(sBuffer)

   If CBool(GetComputerNameAPI(sBuffer, lLen)) Then
      GetComputerName = Left(sBuffer, lLen)
   Else
      GetComputerName = ""
   End If

End Function

