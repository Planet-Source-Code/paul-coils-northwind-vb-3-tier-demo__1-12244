<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = True%>
<html>
<head>
<link rel="stylesheet" href="StyleSheet.css">
<title>Session Timed Out</title>
</head>
<body>
<h1>Your Session Has Timed Out</h1>
<h2>Please Login Again</h2>
<p>You would normally redirect users to this page if they have not logged into your site yet, or if there are security issues requiring a login before page access.  You could either created a link to your login page here, or #include the login page as a sub-page on this one.</p>
<p>This dummy page has changed your session to indicate it is logged in - press back on your browser and you will now be able to access other pages as if you were logged in.
<% Session("LoggedIn") = "Y" %>

