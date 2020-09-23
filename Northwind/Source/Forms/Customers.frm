VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "frmCustomers"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   315
      Left            =   5100
      TabIndex        =   3
      Top             =   3930
      Width           =   1275
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   3780
      TabIndex        =   2
      Top             =   3930
      Width           =   1275
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   2460
      TabIndex        =   1
      Top             =   3930
      Width           =   1275
   End
   Begin MSComctlLib.ListView lvwCustomers 
      Height          =   3735
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6588
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Listview collection constants
Const C_CUSTOMERS_CUSTOMERID As Integer = 0
Const C_CUSTOMERS_COMPANYNAME As Integer = 1
Const C_CUSTOMERS_CONTACTNAME As Integer = 2
Const C_CUSTOMERS_CONTACTTITLE As Integer = 3
Const C_CUSTOMERS_ADDRESS As Integer = 4
Const C_CUSTOMERS_CITY As Integer = 5
Const C_CUSTOMERS_REGION As Integer = 6
Const C_CUSTOMERS_POSTALCODE As Integer = 7
Const C_CUSTOMERS_COUNTRY As Integer = 8
Const C_CUSTOMERS_PHONE As Integer = 9
Const C_CUSTOMERS_FAX As Integer = 10

'Global collection of objects
Dim moCustomers As New clsCustomers

Private Sub CollectionHeader(lvwListView As ListView)

   'Clear any existing list
   lvwListView.ColumnHeaders.Clear
   lvwListView.ListItems.Clear
   lvwListView.MultiSelect = False
   lvwListView.View = lvwReport

   'Add in headings for columns
   lvwListView.ColumnHeaders.Add , , "CustomerID", 1000
   lvwListView.ColumnHeaders.Add , , "Company Name", 1000
   lvwListView.ColumnHeaders.Add , , "Contact Name", 1000
   lvwListView.ColumnHeaders.Add , , "Contact Title", 1000
   lvwListView.ColumnHeaders.Add , , "Address", 1000
   lvwListView.ColumnHeaders.Add , , "City", 1000
   lvwListView.ColumnHeaders.Add , , "Region", 1000
   lvwListView.ColumnHeaders.Add , , "Postal Code", 1000
   lvwListView.ColumnHeaders.Add , , "Country", 1000
   lvwListView.ColumnHeaders.Add , , "Phone", 1000
   lvwListView.ColumnHeaders.Add , , "Fax", 1000
   
   'Default options for list
   lvwListView.Sorted = True
   lvwListView.SortKey = 0
   lvwListView.SortOrder = lvwAscending

End Sub

Private Sub AddItemToListView(oCustomer As clsCustomer, Optional bSelectItem As Boolean = False)

   Dim itmx As ListItem

   'Add listview item with (string) key as unique index
   Set itmx = lvwCustomers.ListItems.Add(, oCustomer.IndexKey & "_K", oCustomer.CustomerID)

   'Update subitems
   Call UpdateSubItems(itmx, oCustomer)

   'If selecting item
   If bSelectItem = True Then
      lvwCustomers.SelectedItem = itmx
   End If

End Sub

Private Sub UpdateSubItems(itmx As ListItem, oCustomer As clsCustomer)

   'Assign properties to subitems
   itmx.SubItems(C_CUSTOMERS_COMPANYNAME) = oCustomer.CompanyName
   itmx.SubItems(C_CUSTOMERS_CONTACTNAME) = Format(oCustomer.ContactName)
   itmx.SubItems(C_CUSTOMERS_CONTACTTITLE) = Format(oCustomer.ContactTitle)
   itmx.SubItems(C_CUSTOMERS_ADDRESS) = Format(oCustomer.Address)
   itmx.SubItems(C_CUSTOMERS_CITY) = Format(oCustomer.City)
   itmx.SubItems(C_CUSTOMERS_REGION) = Format(oCustomer.Region)
   itmx.SubItems(C_CUSTOMERS_POSTALCODE) = Format(oCustomer.PostalCode)
   itmx.SubItems(C_CUSTOMERS_COUNTRY) = Format(oCustomer.Country)
   itmx.SubItems(C_CUSTOMERS_PHONE) = Format(oCustomer.Phone)
   itmx.SubItems(C_CUSTOMERS_FAX) = Format(oCustomer.Fax)

   'Store key in tag
   itmx.Tag = oCustomer.IndexKey

End Sub

Private Sub lvwCustomers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call SortListView(lvwCustomers, ColumnHeader.Index - 1)
End Sub

Private Sub LoadCollectionListView()

   'Declare object
   Dim oCustomer As clsCustomer

   'Header
   Call CollectionHeader(lvwCustomers)

   'Load collection from db
   Call moCustomers.PopulateAll

   'For each individual item
   For Each oCustomer In moCustomers
      Call AddItemToListView(oCustomer)
   Next

   'Select first record in list
   If moCustomers.Count > 0 Then
      Set lvwCustomers.SelectedItem = lvwCustomers.ListItems(1)
      Call lvwCustomers_ItemClick(lvwCustomers.SelectedItem)
   End If

End Sub

Private Sub lvwCustomers_DblClick()

   'Edit currently selected record
   If lvwCustomers.SelectedItem Is Nothing Then
      Exit Sub
   End If

   'Call edit function
   Call cmdEdit_Click

End Sub

Private Sub lvwCustomers_ItemClick(ByVal Item As MSComctlLib.ListItem)

   'Display selected records information for editing
   'or show detail records in master-detail relationship

End Sub

Private Sub Form_Load()

   'Load collection
   Call LoadCollectionListView

End Sub

Private Sub Form_Unload(Cancel As Integer)

   'Clean up
   Set moCustomers = Nothing

End Sub

Private Sub cmdDelete_Click()

   'Remove from collection
   moCustomers.Delete lvwCustomers.SelectedItem.Tag

   'Remove from screen
   lvwCustomers.ListItems.Remove lvwCustomers.SelectedItem.Index

End Sub

Private Sub cmdEdit_Click()

   'Variables
   Dim oForm As frmCustomersMaint
   Dim oCustomer As clsCustomer

   'Create new form and object
   Set oForm = New frmCustomersMaint
   Set oCustomer = New clsCustomer

   'Locate record
   Set oCustomer = moCustomers(lvwCustomers.SelectedItem.Tag)

   'Edit record
   If oForm.Edit(oCustomer) = True Then

      'Update screen
      Call UpdateListView(oCustomer)

   End If

   'Clean up
   Set oForm = Nothing
   Set oCustomer = Nothing

End Sub

Private Sub cmdAdd_Click()

   'Variables
   Dim oForm As frmCustomersMaint
   Dim oCustomer As clsCustomer

   'Create new form and object
   Set oForm = New frmCustomersMaint
   Set oCustomer = New clsCustomer

   'Add record
   If oForm.Add(oCustomer) = True Then

      'Update collection (also sets index key)
      Call moCustomers.Add(oCustomer)

      'Update screen
      Call AddItemToListView(oCustomer)

   End If

   'Clean up
   Set oForm = Nothing
   Set oCustomer = Nothing

End Sub

Private Sub UpdateListView(oCustomer As clsCustomer)

   Dim itmx As ListItem

   'Locate listview item using (string) key
   Set itmx = lvwCustomers.ListItems.Item(oCustomer.IndexKey & "_K")

   'Adjust first column and subitems
   itmx.Text = oCustomer.CustomerID
   Call UpdateSubItems(itmx, oCustomer)

End Sub
