<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = True%>
<%
Sub DisplayRecords()

   Dim fldCustomerID, fldCompanyName, fldContactName, fldContactTitle, fldAddress, fldCity, fldRegion, fldPostalCode, fldCountry, fldPhone, fldFax
   Dim Count

   'Varibles
   Dim oCustomer
   Dim Rs

   'Error handling
   On Error Resume Next

   'Calls NW_SelAllCustomers
   Set oCustomer = Server.CreateObject("NorthwindData.clsDbCustomers")
   Set Rs = oCustomer.PopulateAll( )

   'Check for error
   If Err.Number <> 0 Then
      Response.Write ("Error - Your Query Could Not Be Run - Please Try Again.  The Error Message Was: " & Err.Description)
      Set oCustomer = Nothing
      Set Rs = Nothing
      Exit Sub
   End If

   If Not (Rs Is Nothing) Then
      If Not (Rs.EOF) Then
      
         'Assign fields
         Set fldCustomerID = Rs("CustomerID")
         Set fldCompanyName = Rs("CompanyName")
         Set fldContactName = Rs("ContactName")
         Set fldContactTitle = Rs("ContactTitle")
         Set fldAddress = Rs("Address")
         Set fldCity = Rs("City")
         Set fldRegion = Rs("Region")
         Set fldPostalCode = Rs("PostalCode")
         Set fldCountry = Rs("Country")
         Set fldPhone = Rs("Phone")
         Set fldFax = Rs("Fax")
         %>

         <table width="100%" border="1" cellspacing="0" cellpadding="2" bordercolordark="#000099" bordercolorlight="#000099" bgcolor="#FFFFFF">
         <tr>
         <th>Customerid</th>
         <th>Company Name</th>
         <th>Contact Name</th>
         <th>Contact Title</th>
         <th>Address</th>
         <th>City</th>
         <th>Region</th>
         <th>Postal Code</th>
         <th>Country</th>
         <th>Phone</th>
         <th>Fax</th>

         </tr>
         <%
         Do Until Rs.EOF%>
            <tr>
               <td><%=fldCustomerID.Value%></td>
               <td><%=fldCompanyName.Value%></td>
               <td><%=fldContactName.Value%></td>
               <td><%=fldContactTitle.Value%></td>
               <td><%=fldAddress.Value%></td>
               <td><%=fldCity.Value%></td>
               <td><%=fldRegion.Value%></td>
               <td><%=fldPostalCode.Value%></td>
               <td><%=fldCountry.Value%></td>
               <td><%=fldPhone.Value%></td>
               <td><%=fldFax.Value%></td>
            </tr><%
            Rs.MoveNext
         Loop
         Count = Rs.RecordCount
         Rs.Close
         Set Rs = Nothing
         Response.Write ("</table>")
         If Count > 0 Then
            Response.Write ("<p>" & Count & " Records.")
         End If
      Else
         Response.Write "<p> No Records Found Matching Your Search Criteria </p>"
      End If
   End If
End Sub%>

<html>
<head>
<link rel="stylesheet" href="StyleSheet.css">
<title>Page Name HereCustomer</title>
</head>
<body>
<h1>Customers</h1>
<h2>All Records Listed</h2>
<b>Sample Output Via Transaction Server</b>

<% Call DisplayRecords() %>

</body>
</html>
