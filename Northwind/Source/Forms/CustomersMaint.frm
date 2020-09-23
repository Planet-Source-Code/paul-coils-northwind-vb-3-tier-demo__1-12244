VERSION 5.00
Begin VB.Form frmCustomersMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Maintenance"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   LinkTopic       =   "frmCustomersMaint"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   2100
      MaxLength       =   5
      TabIndex        =   0
      Top             =   60
      Width           =   2565
   End
   Begin VB.TextBox txtCompanyName 
      Height          =   285
      Left            =   6735
      MaxLength       =   40
      TabIndex        =   1
      Top             =   60
      Width           =   2565
   End
   Begin VB.TextBox txtContactName 
      Height          =   285
      Left            =   2100
      MaxLength       =   30
      TabIndex        =   2
      Top             =   420
      Width           =   2565
   End
   Begin VB.TextBox txtContactTitle 
      Height          =   285
      Left            =   6735
      MaxLength       =   30
      TabIndex        =   3
      Top             =   420
      Width           =   2565
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   2100
      MaxLength       =   60
      TabIndex        =   4
      Top             =   780
      Width           =   2565
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   6735
      MaxLength       =   15
      TabIndex        =   5
      Top             =   780
      Width           =   2565
   End
   Begin VB.TextBox txtRegion 
      Height          =   285
      Left            =   2100
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1140
      Width           =   2565
   End
   Begin VB.TextBox txtPostalCode 
      Height          =   285
      Left            =   6735
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1140
      Width           =   2565
   End
   Begin VB.TextBox txtCountry 
      Height          =   285
      Left            =   2100
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1500
      Width           =   2565
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   6735
      MaxLength       =   24
      TabIndex        =   9
      Top             =   1500
      Width           =   2565
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   2100
      MaxLength       =   24
      TabIndex        =   10
      Top             =   1860
      Width           =   2565
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3690
      TabIndex        =   22
      Top             =   2430
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   5040
      TabIndex        =   23
      Top             =   2430
      Width           =   1275
   End
   Begin VB.Label lblCustomerID 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Customerid"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   90
      Width           =   1800
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Company Name"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4800
      TabIndex        =   12
      Top             =   90
      Width           =   1800
   End
   Begin VB.Label lblContactName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Contact Name"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   450
      Width           =   1800
   End
   Begin VB.Label lblContactTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Contact Title"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4800
      TabIndex        =   14
      Top             =   450
      Width           =   1800
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   810
      Width           =   1800
   End
   Begin VB.Label lblCity 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4800
      TabIndex        =   16
      Top             =   810
      Width           =   1800
   End
   Begin VB.Label lblRegion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Region"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1170
      Width           =   1800
   End
   Begin VB.Label lblPostalCode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Postal Code"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4800
      TabIndex        =   18
      Top             =   1170
      Width           =   1800
   End
   Begin VB.Label lblCountry 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Country"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   1530
      Width           =   1800
   End
   Begin VB.Label lblPhone 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phone"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4800
      TabIndex        =   20
      Top             =   1530
      Width           =   1800
   End
   Begin VB.Label lblFax 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   1890
      Width           =   1800
   End
End
Attribute VB_Name = "frmCustomersMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Reference to current object
Dim moCustomer As clsCustomer

'Function saved or cancelled
Dim bSaveRecord As Boolean

'Form state
Dim eFormChangeState As FormChangeState

Private Sub Form_Unload(Cancel As Integer)

   'Clean up any set reference
   Set moCustomer = Nothing

End Sub

Public Function Edit(oCustomer As clsCustomer) As Boolean

   'Do not save record by default
   bSaveRecord = False

   'Set global reference to object being edited
   Set moCustomer = oCustomer

   'Display values from object
   Call DisplayFields(oCustomer)

   'Form clean
   Call FormClean

   'Display form
   Me.Show vbModal

   'If saving record
   If bSaveRecord = True Then

      'Put fields in object
      Call SetFields(oCustomer)

      'Update database
      oCustomer.Update

   End If

   'Return; unload form
   Edit = bSaveRecord
   Unload Me

End Function

Public Function Add(oCustomer As clsCustomer) As Boolean

   'Do not save record by default
   bSaveRecord = False

   'Display default values for object if any set
   Call DisplayFields(oCustomer)

   'Form clean
   Call FormClean

   'Display form
   Me.Show vbModal

   'If saving record
   If bSaveRecord = True Then

      'Put fields in object
      Call SetFields(oCustomer)

      'Insert to database
      oCustomer.Insert

   End If

   'Return; unload form
   Add = bSaveRecord
   Unload Me

End Function

Private Sub cmdSave_Click()

   'Perform form level validations
   If FormValidations = False Then
      Exit Sub
   End If

   'Saving record
   bSaveRecord = True

   'Hide form
   Me.Hide

End Sub

Private Sub cmdCancel_Click()

   'Not saving record
   bSaveRecord = False

   'Hide form
   Me.Hide

End Sub

Public Sub DisplayFields(oCustomer As clsCustomer)

   'Display each property
   txtCustomerID = oCustomer.CustomerID
   txtCompanyName = oCustomer.CompanyName
   txtContactName = Format(oCustomer.ContactName)
   txtContactTitle = Format(oCustomer.ContactTitle)
   txtAddress = Format(oCustomer.Address)
   txtCity = Format(oCustomer.City)
   txtRegion = Format(oCustomer.Region)
   txtPostalCode = Format(oCustomer.PostalCode)
   txtCountry = Format(oCustomer.Country)
   txtPhone = Format(oCustomer.Phone)
   txtFax = Format(oCustomer.Fax)

End Sub

Public Sub SetFields(oCustomer As clsCustomer)

   'Set each property
   oCustomer.CustomerID = txtCustomerID
   oCustomer.CompanyName = txtCompanyName
   oCustomer.ContactName = Format(txtContactName)
   oCustomer.ContactTitle = Format(txtContactTitle)
   oCustomer.Address = Format(txtAddress)
   oCustomer.City = Format(txtCity)
   oCustomer.Region = Format(txtRegion)
   oCustomer.PostalCode = Format(txtPostalCode)
   oCustomer.Country = Format(txtCountry)
   oCustomer.Phone = Format(txtPhone)
   oCustomer.Fax = Format(txtFax)

End Sub

Private Function FormValidations() As Boolean

   'Check all required fields are filled out
   '..as based on your business rules.

   'Return true if no problems
   FormValidations = True

End Function

Private Sub FormClean()

   'Form clean; disable save option
   eFormChangeState = Clean
   cmdSave.Enabled = False

End Sub

Private Sub FormDirty()

   'Form changed; enable save option
   eFormChangeState = Dirty
   cmdSave.Enabled = True

End Sub

Private Sub txtCustomerID_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtCompanyName_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtContactName_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtContactTitle_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtRegion_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtPostalCode_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
   Call ValidateData(True, True, True, KeyAscii)
End Sub

Private Sub txtCustomerID_Change()
   Call FormDirty
End Sub

Private Sub txtCompanyName_Change()
   Call FormDirty
End Sub

Private Sub txtContactName_Change()
   Call FormDirty
End Sub

Private Sub txtContactTitle_Change()
   Call FormDirty
End Sub

Private Sub txtAddress_Change()
   Call FormDirty
End Sub

Private Sub txtCity_Change()
   Call FormDirty
End Sub

Private Sub txtRegion_Change()
   Call FormDirty
End Sub

Private Sub txtPostalCode_Change()
   Call FormDirty
End Sub

Private Sub txtCountry_Change()
   Call FormDirty
End Sub

Private Sub txtPhone_Change()
   Call FormDirty
End Sub

Private Sub txtFax_Change()
   Call FormDirty
End Sub

