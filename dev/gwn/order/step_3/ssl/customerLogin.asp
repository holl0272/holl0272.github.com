<%@ Language=VBScript %>
<% Option Explicit %>
<%
'********************************************************************************
'*
'*   process_order.asp
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins is process_order.asp APPVERSION: 50.4014.0.2
'*   from LaGarde, Incorporated. It has been heavily modified by Sandshot Software
'*   with permission from LaGarde, Incorporated who has NOT released any of the 
'*   original copyright protections. As such, this is a derived work covered by
'*   the copyright provisions of both companies.
'*
'*   LaGarde, Incorporated Copyright Statement
'*   The contents of this file is protected under the United States copyright
'*   laws and is confidential and proprietary to LaGarde, Incorporated.  Its 
'*   use ordisclosure in whole or in part without the expressed written 
'*   permission of LaGarde, Incorporated is expressly prohibited.
'*   (c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'*   
'*   Sandshot Software Copyright Statement
'*   The contents of this file are protected by United States copyright laws 
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************
%>
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="sfLib/incDesign.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLIB/incAE.asp"-->
<!--#include file="SFLib/incProcOrder.asp"-->
<!--#include file="SFLib/ssclsCustomer.asp"--> 
<!--#include file="SFLib/ssclsCustomerShipAddress.asp"--> 
<!--#include file="SFLib/ssclsLogin.asp"--> 
<%

Response.Buffer = True
On Error Goto 0

'**********************************************************
'	Developer notes
'**********************************************************

'this page can be access from:
'1) order.asp - standard checkout
'2) process_order.asp - logging in
'3) confirm.asp - rejected payment

'**********************************************************
'*	Page Level variables
'**********************************************************
Dim InboundVisitorID
Dim sEmail, sPassword

'**********************************************************
'*	Functions
'**********************************************************

Sub CleanupPageObjects

On Error Resume Next

	If Not isEmpty(mclsLogin) Then Set mclsLogin = Nothing
	Call cleanup_dbconnopen
	
	If Err.number <> 0 Then Err.Clear
	
End Sub	'CleanupPageObjects

'**********************************************************

Sub gotoProcessOrder

	Call CleanupPageObjects
	Server.Transfer "process_order.asp"
	
End Sub	'gotoProcessOrder

'**********************************************************
'*	Begin Page Code
'**********************************************************

If vDebug = 1 Then Call displayVisitorPreferences	'for debugging

'Check for login
sEmail			= Trim(Request.Form("Email"))
sPassword		= Trim(Request.Form("Passwd"))
If Len(sEmail) > 0 And Len(sPassword) > 0 Then
	Set mclsLogin = New clsLogin
	mstrLoginMessage = mclsLogin.ValidUserName(sEmail, sPassword)
	If len(sEmail) = 0 Then sEmail = Request.Cookies("Email")
	If mstrLoginMessage = "True" then	
		iCustID = mclsLogin.UserID
		Call setVisitorLoggedInCustomerID(iCustID)
		
		If Len(custID_cookie) > 0 AND iCustID <> custID_cookie  Then
			If CheckSavedCartCustomer(custID_cookie) Then
				' Delete SvdCartCustomer Row
				Call DeleteCustRow(custID_cookie)
				' See if saved cart has any remaining saved
				Call setUpdateSavedCartCustID(iCustID, custID_cookie)
			End If
		End If	
		
		Call gotoProcessOrder
	End If
	Set mclsLogin = Nothing
Else
	'This check added for a special case where ssl directory triggers new session
	InboundVisitorID = Request.Form("SessionID")
	If getvisitorPreference("visitorID") <> InboundVisitorID Then
		Set maryVisitorPreferences = Nothing
		Call loadVisitorPreferences(InboundVisitorID)
		
		Call setSessionID(InboundVisitorID)
		Call setCookie_SessionID(InboundVisitorID, DateAdd("d", 365, Now()))
		Call setCookie_visitorID(InboundVisitorID, DateAdd("d", 365, Now()))
	End If
End If	'Len(sEmail) > 0 And Len(sPassword) > 0

' Check if custID exists 
If Len(visitorLoggedInCustomerID) > 0 Then
	Set mclsCustomer = New clsCustomer
    If mclsCustomer.LoadCustomer(visitorLoggedInCustomerID) Then
		iCustID = visitorLoggedInCustomerID
		Call setCookie_custID(iCustID, Date() + 730)
		Call gotoProcessOrder
	Else
		Call setCookie_custID("", Now())
	End If
End If	

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Check Out - Customer Login</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Pragma" content="no-cache">
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
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" src="include_commonElements/css.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
				function isEmpty(theField,theMessage)
				{
				if (theField.value == "")
					{	
					alert(theMessage);
					theField.focus();
					theField.select();
					return(true);
					}
					{
					return(false);
					}
				}

				function ValidInput(theForm)
				{

				if (isEmpty(theForm.Email,"Please enter a Email.")) {return(false);}
				if (isEmpty(theForm.Passwd,"Please enter a password.")) {return(false);}
				  
				return(true);
				}

<% writeCurrencyConverterOpeningScript %>
</script>
</head>

<body <%= mstrBodyStyle %>>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr>
          <td align="center" class="tdMiddleTopBanner">Welcome to Checkout</td>
        </tr>
        <tr>
          <td class="tdBottomTopBanner" align="left">
This is the final step in completing your order. You are now connected using a secure (SSL) connection and all information is transmitted in an encrypted form. 
You should see a small key (Netscape) or lock (IE) indicating that your browser is communicating securely with our web store.
          </td>    
        </tr>	
        <tr>
          <td align="left" class="tdContent2"  valign="middle"><hr /><strong>Step 1: Customer Information</strong> | Step 2: Payment Information | Step 3: Complete Order<hr /></td>
        </tr>
        <tr>
          <td align="left" class="tdContent2"  valign="middle">
		        <table border="0" cellpadding="2" cellspacing="0">
		          <tr>
		            <td align="left" class="tdContent2" valign="top">
			<table class="Section" cellpadding="2" cellspacing="0" border="1">
				<tr>
				<td class="tdTopBanner">Returning Customers</td>
				</tr>
				<tr>
				<td nowrap align="left">
					  <form name="frmPassword" id="frmPassword" action="process_order.asp" method="post" onsubmit="return ValidInput(this);">
					  <table border="0" class="tdBottomTopBanner2" width="100%" cellpadding="3" cellspacing="1">
						<tr>
						  <td align="center" valign="middle" class="tdContent2">
					        <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
							  <tr>
							    <td width="100%" align="right" colspan="2" class="tdContent">
								<p align="left">Please use your e-mail address and password to log in and retrieve your customer information.</td>
							  </tr>
							  <tr>
							    <td width="15%" align="right">E-Mail:</td>
							    <td width="85%" align="left"><input type="text" size="25" name="Email" id="Email_frmPassword" title="E-Mail Address" class="formDesign"></td>
							  </tr>
							  <tr>
							    <td align="right">Password:</td>
							    <td align="left"><input type="password" size="20" name="Passwd" id="Passwd_frmPassword" title="Password" class="formDesign" maxlength="10"></td>
							  </tr>
							  <tr>
							    <td colspan="2" align="center">
							      <input Type="image" class="inputImage" src="<%= C_BTN16 %>" name="btnLogin" id="btnLogin">
							    </td>
							  </tr>
					        </table>					    
						  </td>
						</tr>
						<tr>
						  <td align="center" class="tdContent2"><a href="password.asp?status=fpwd">Forgot your password?</a></td>		        
						</tr>
					  </table>
					  </form>		
				   </td>
				   </tr>
				   </table>
				    </td>
				    <td>&nbsp;</td>
				    <td class="tdContent2" valign="top">
			<table class="Section" cellpadding="2" cellspacing="0" border="1">
				<tr>
				<td class="tdTopBanner">New Customers</td>
				</tr>
				<tr>
				<td align="left">
					  <table border="0" class="tdBottomTopBanner2" width="100%" cellpadding="3" cellspacing="1">
						<tr>
						  <td align="center" valign="middle" class="tdContent2">
					        <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
							  <tr>
							    <td width="100%" align="left" class="tdContent">
								You will have the opportunity during checkout to 
								register for online access to your order status.</td>
							  </tr>
							  <tr>
							    <td align="center">
				  <form action="process_order.asp" method="post" name="frmCheckout" ID="frmCheckout">
					<input type="hidden" name="SessionID" ID="SessionID" value="<%= SessionID %>">
					<input type="hidden" name="HideLogin" ID="HideLogin" value="True">
					<input type="image" class="inputImage" src="<%= C_BTN05 %>" name="checkout" ID="checkout" alt="New Customers click here!">
				  </form></td>
							  </tr>
					        </table>					    
						  </td>
						</tr>
						</table>
				   </td>
				   </tr>
				   </table>
				    </td>
		          </tr>
		        </table>	
          </td>
        </tr>
        </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->
</body>
</html>
<% Call CleanupPageObjects %>