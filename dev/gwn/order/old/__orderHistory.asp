<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
'********************************************************************************
'*                                                                              *
'*   1.00.001 (June 15, 2006)													*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file="SFLib/ssOrderManager.asp"-->
<!--#include file="SFLib/myAccountSupportingFunctions.asp"-->
<%
'/////////////////////////////////////////////////
'/
'/  User Parameters
'/


'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Global variables
'**********************************************************

Dim mbytLoginDisplayType
Dim mbytLoginStatus
Dim mlngOrderID
Dim mrsOrderHistory
'Dim mstrAction
'Dim mstrEmail
'Dim mstrPassword
Dim mstrMessage

'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'*	Begin Page Code
'**********************************************************

mbytLoginDisplayType = 0
'Check for logged in, possible results
'
'Login Status saved to Session("ssLoginStatus")
'
'0) Not logged in
'1) Logged in with email/orderID - view order only
'2) Logged in with email/password - view order and order history
'2) Logged in using SF login:&nbsp;&nbsp; this condition is left as an excersize for the student :&nbsp;&nbsp;)

'mbytLoginStatus = Session("ssLoginStatus")

mlngOrderID = LoadRequestValue("OrderID")
mstrEmail = LoadRequestValue("Email")
mstrPassword = LoadRequestValue("Password")

'Only let in valid logins
If Not isLoggedIn Then
	If Len(Request.QueryString & Request.Form) > 0 Then	
		Call Login(mstrEmail,mstrPassword,mlngOrderID,mstrMessage)
	Else
		Call cleanup_dbconnopen	'This line needs to be included to close database connection
		Response.Redirect "myAccount.asp?PrevPage=" & Request.ServerVariables("SCRIPT_NAME") & Server.URLEncode("?" & Request.QueryString)
	End If
End If

If mblnShowOrderSummaries And mbytLoginStatus=0 Then mbytLoginDisplayType = 4

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Order History</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% Call preventPageCache %>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="keywords">
<meta name="description" content="description">
<meta name="Robot" content="Index,ALL">
<meta name="revisit-after" content="15 Days">
<meta name="Rating" content="General">
<meta name="Language" content="en">
<meta name="distribution" content="Global">
<meta name="Classification" content="classification">

<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">

<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/incae.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfEmailFriend.js" type="text/javascript"></script>
</head>
<body <%= mstrBodyStyle %>>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
	<tr>
		<td>
		<%
			Call ShowMyAccountBreadCrumbsTrail("", True)
			If isLoggedIn Then
				Call ShowOrderDetail(mlngOrderID)
				If LoadOrderHistory(visitorLoggedInCustomerID, mrsOrderHistory) Then
					Call ShowOrderHistory(mlngOrderID, True, mrsOrderHistory)
				End If
				mrsOrderHistory.Close
				Set mrsOrderHistory = Nothing
			Else
				Call ShowOrderDetail(mlngOrderID)
				Call ShowOrderHistory(mlngOrderID, False, mrsOrderHistory)
			End If
		%>
        </td>
	</tr>
</table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->
</body>
</html>
<%
Call cleanup_dbconnopen
%>