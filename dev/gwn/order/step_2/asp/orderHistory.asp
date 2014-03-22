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
  <link runat="server" rel="shortcut icon" type="../image/png" href="favicon.ico">
  <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Lato:100,400,900|Josefin+Sans:100,400,700,400italic,700italic">
  <link rel="stylesheet" href="../css/main.css">
<script language="javascript" src="../SFLib/jquery-1.10.2.min.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/incae.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfEmailFriend.js" type="text/javascript"></script>


<script>
$(document).ready(function() {
	$('.tdAltFont1 a > b').unwrap();
	$('.tdAltFont2 a > b').unwrap();
	$('#tblMainContent td').css('padding', '5px');
	$('#divCategoryTrail').hide();
})
</script>

<style>
body {
	text-align: center;
}
</style>
</head>

<body <%= mstrBodyStyle %>>

		<div id="header" style="margin-bottom: 2%;">
    <div id="gwn_logo">
      <a href="../index.html" title="Home"><image src="../images/gwn_logo.png" alt="GameWearNow Logo"></a>
    </div>
    <div id="heading">
      <span class="title_txt" id="title">CUSTOM JERSEYS FOR<br>YOUR SPORTS TEAM</span>
        <br>
      <span class="title_txt" id="sub_title">ORDER HISTORY</span>
    </div>
  </div>


<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent" style="background-color:#DEDEDE; margin: 0 auto 5%;">
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


  <div id="footer">
    <ul id="horizontal-nav">
      <li class="not_selected"><span><image src="images/shopping_cart.png" alt="Shopping Cart" id="shopping_cart">MY SHOPPING CART</span></a></li>
      <li class="pipe">|</li>
      <li id="current_page"><a href="myAccount.asp" title="My Account">MY ACCOUNT</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/faqs/faqs.html" title="FAQ's">FAQ'S</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/privacy_policy/privacy_policy.html" title="Privacy Policy">PRIVACY POLICY</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/contact_us/contact_us.html" title="Contact Us">CONTACT US <font>(877) 796-6639</font></a></li>
    </ul>
  </div>
</body>
</html>
<%
Call cleanup_dbconnopen
%>