Attribute VB_Name = "modFormLibrary"
Option Explicit

'Form state (dirty/clean)
Public Enum FormChangeState
   Clean
   Dirty
End Enum

Public Sub SortListView(lvwListView As ListView, iIndex As Integer)

   lvwListView.Sorted = True
   
   'Toggle sort direction
   If lvwListView.SortKey = iIndex Then
      If lvwListView.SortOrder = lvwAscending Then
      lvwListView.SortOrder = lvwDescending
      Else
      lvwListView.SortOrder = lvwAscending
      End If
   Else
      'Sort by new column
      lvwListView.SortKey = iIndex
   End If

End Sub

Public Sub ValidateData(bAlphaUpper As Boolean, bAlphaLower As Boolean, bNumbers As Boolean, KeyAscii As Integer, Optional bForceCase As Boolean = False, Optional sAdditional = "")

   Dim sChar, sValidList

   'Init value
   sValidList = "

   'If forcing specific case
   If bForceCase = True Then
      If bAlphaUpper = True Then
         sChar = UCase(Chr(KeyAscii))
      Else
         sChar = LCase(Chr(KeyAscii))
      End If
   Else
      sChar = Chr(KeyAscii)
   End If

   'Determine strings to validate
   If bAlphaUpper = True Then
      sValidList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
   End If

   'Determine strings to validate
   If bAlphaLower = True Then
      sValidList = sValidList & "abcdefghijklmnopqrstuvwxyz"
   End If

   'Numbers allowed
   If bNumbers = True Then
      sValidList = sValidList & "01234567890"
   End If

   'Any additional characters
   sValidList = sValidList & sAdditional

   'Filter out unwanted charactes (allow backspace)
   If InStr(sValidList & Chr(vbKeyBack), sChar) = 0 Then
      KeyAscii = 0
   Else
      KeyAscii = Asc(sChar)
   End If

End Sub

