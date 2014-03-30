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
<!--#include file="sfLib/ssclsCustomer.asp"-->
<!--#include file="sfLib/ssclsLogin.asp"-->
<!--#include file="SFLib/ssclsCustomerShipAddress.asp"-->
<!--#include file="SFLIB/processor_PayPalWebPayments.asp"-->
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
Dim sEmail, sPassword, sCondition
Dim sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2, sCustCity, sCustState, sCustZip, sCustCountry, sCustPhone, sCustFax, sCustEmail, sCustSubscribed
Dim sShipCustFirstName, sShipCustMiddleInitial,sShipCustLastName, sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustZip, sShipCustCountry, sShipCustPhone, sShipCustFax, sShipCustCompany, sShipCustEmail
Dim mstrAltState
Dim sSubmitAction
Dim paryPriorShippingAddresses
Dim pstrOptions_PriorShippingAddresses
Dim pstrOptionText_PriorShippingAddresses
Dim i,j
Dim sShipMethods
Dim mblnHideLogin
Dim InboundVisitorID

'**********************************************************
'*	Functions
'**********************************************************

Sub CleanupPageObjects

On Error Resume Next

	If Not isEmpty(mclsCustomer) Then Set mclsCustomer = Nothing
	If Not isEmpty(mclsCustomerShipAddress) Then Set mclsCustomerShipAddress = Nothing
	If Not isEmpty(mclsCartTotal) Then Set mclsCartTotal = Nothing
	Call cleanup_dbconnopen

	If Err.number <> 0 Then Err.Clear

End Sub	'CleanupPageObjects

'**********************************************************
'*	Begin Page Code
'**********************************************************

If vDebug = 1 Then Call displayVisitorPreferences	'for debugging

'Check for login
sEmail			= Trim(Request.Form("Email"))
sPassword		= Trim(Request.Form("Passwd"))
If Len(sEmail) = 0 And Len(sPassword) = 0 Then
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

If vDebug = 1 Then Call displayVisitorPreferences	'for debugging

Set mclsCustomer = New clsCustomer
If Len(visitorCity) > 0 Then mclsCustomer.custCity = visitorCity
If Len(visitorState) > 0 Then mclsCustomer.custState = visitorState
If Len(visitorZIP) > 0 Then mclsCustomer.custZip = visitorZIP
If Len(visitorCountry) > 0 Then mclsCustomer.custCountry = visitorCountry

Set mclsCartTotal = New clsCartTotal
With mclsCartTotal

	.City = visitorCity
	.State = visitorState
	.ZIP = visitorZIP
	.Country = visitorCountry
	.isCODOrder = False

	.ShipMethodCode = visitorPreferredShippingCode
	.LoadAllShippingMethods = True

	.LoadCartContents
	.checkInventoryLevels

	'.writeDebugCart

	'.displayOrder_CheckoutView

	If True Then	'True	False for testing
		If .isEmptyCart Then
			Session.Abandon
			Call CleanupPageObjects	'Clean up before the redirect
			Response.Redirect(adminDomainName & cstrTimeoutRedirectPage)
		ElseIf .isStockDepleted Then
			Call CleanupPageObjects	'Clean up before the redirect
			Response.Redirect(adminDomainName & "order.asp")
		End If
	Else
		.writeDebugCart
		Call displayVisitorPreferences
		If .isEmptyCart Then
			Response.Write "<h4>Empty Cart</h4>"
		ElseIf .isStockDepleted Then
			Response.Write "<h4>Stock Depleted</h4>"
		End If
	End If

	'Required for view rates pop-up
	Session("persistTotalPrice") = .SubTotal

End With	'mclsCartTotal

'Check for login
If Len(Trim(Request.Form("btnLogin.x"))) > 0 Then
' If login button is depressed
	iCustID = customerAuth(sEmail, sPassword, "loose")	'loose = email+password must match
	If iCustID > 0 Then
		If Len(custID_cookie) > 0 AND iCustID <> custID_cookie  Then
			If CheckSavedCartCustomer(custID_cookie) Then
				' Delete SvdCartCustomer Row
				Call DeleteCustRow(custID_cookie)
				' See if saved cart has any remaining saved
				Call setUpdateSavedCartCustID(iCustID, custID_cookie)
			End If
		End If

		Call SetSessionLoginParameters(iCustID, sEmail)
	Else
		iCustID = ""
		If customerAuth(sEmail,sPassword,"loosest") > 0 Then	'loosest = email already exist
			sCondition = "EmailMatch"
			Call expireCookie_sfCustomer
		Else
			sCondition = "WrongCombination"
			Call expireCookie_sfCustomer
		End If
	End If
End If

' Check if custID exists
If Not isValidRecordID(visitorLoggedInCustomerID) Then
    If mclsCustomer.LoadCustomer(visitorLoggedInCustomerID) Then
		iCustID = visitorLoggedInCustomerID
		Call setCookie_custID(iCustID, Date() + 730)
	Else
		Call setCookie_custID("", Now())
	End If
ElseIf mclsCustomer.custID <> visitorLoggedInCustomerID Then
'this condition occurs on login
    If mclsCustomer.LoadCustomer(visitorLoggedInCustomerID) Then
		iCustID = visitorLoggedInCustomerID
		Call setCookie_custID(iCustID, Date() + 730)
	End If
End If

'Set billing address
With mclsCustomer
	sCustFirstName		= .custFirstName
	sCustMiddleInitial	= .custMiddleInitial
	sCustLastName		= .custLastName
	sCustCompany		= .custCompany
	sCustAddress1		= .custAddr1
	sCustAddress2		= .custAddr2
	sCustCity			= .custCity
	sCustState			= .custState
	sCustZip			= .custZip
	sCustCountry		= .custCountry
	sCustPhone			= .custPhone
	sCustFax			= .custFax
	sCustEmail			= .custEmail
	sCustSubscribed		= .custIsSubscribed
End	With	'mclsCustomer

' Change display for saved cart customers
If instr(1, sCustFirstName, "Saved Cart Customer", 1) Then sCustFirstName = ""
If (sCustCountry = "US" Or sCustCountry = "CA") Then
	mstrAltState = ""
Else
	mstrAltState = sCustState
End If

Call PayPalExpressCheckout_process_order	'added for Sandshot Software's PayPal WebPayments Pro Integration

'Get prior shipping addresses, if any, to populate form
Set mclsCustomerShipAddress = New clsCustomerShipAddress
If mclsCustomerShipAddress.getPriorShippingAddresses(visitorShipAddressID, visitorLoggedInCustomerID, paryPriorShippingAddresses) Then
	With mclsCustomerShipAddress
		sShipCustFirstName		= .FirstName
		sShipCustMiddleInitial	= .MiddleInitial
		sShipCustLastName		= .LastName
		sShipCustCompany		= .Company
		sShipCustAddress1		= .Addr1
		sShipCustAddress2		= .Addr2
		sShipCustCity			= .City
		sShipCustState			= .State
		sShipCustZip			= .Zip
		sShipCustCountry		= .Country
		sShipCustPhone			= .Phone
		sShipCustFax			= .Fax
		sShipCustEmail			= .Email
	End	With	'mclsCustomerShipAddress
ElseIf False Then	'disabled to keep shipping address clear
	sShipCustFirstName		= sCustFirstName
	sShipCustMiddleInitial	= sCustMiddleInitial
	sShipCustLastName		= sCustLastName
	sShipCustCompany		= sCustCompany
	sShipCustAddress1		= sCustAddress1
	sShipCustAddress2		= sCustAddress2
	sShipCustCity			= sCustCity
	sShipCustState			= sCustState
	sShipCustZip			= sCustZip
	sShipCustCountry		= sCustCountry
	sShipCustPhone			= sCustPhone
	sShipCustFax			= sCustFax
	sShipCustEmail			= sCustEmail
End If

mblnHideLogin = Request.Form("HideLogin")
If Len(mblnHideLogin) = 0 Then mblnHideLogin = False

sSubmitAction = ""
If (NOT (isLoggedIn) OR Len(iCustID) = 0) And Not cblnDisableLogin Then sSubmitAction = "this.Password.password=true;this.Password.optional = true;this.Password2.optional = true;"
sSubmitAction = sSubmitAction & "this.Company.optional = true;this.Address2.optional = true;this.Fax.optional = true;this.Address2.optional = true;this.Instructions.optional = true;this.Email.eMail = true;this.Phone.phoneNumber = true;this.ShipState.optional = true;this.ShipCountry.optional = true;this.MiddleInitial.optional = true;"
If mclsCustomerShipAddress.NumAddresses > 0 Then sSubmitAction = "this.priorAddress.optional = true;" & sSubmitAction
If Not isLoggedIn And cblnIncludeEmailVerification Then
	sSubmitAction = sSubmitAction & "if (!validateEmails(this)){return false;} return validate_Me(this);"
Else
	sSubmitAction = sSubmitAction & "return validate_Me(this);"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Check Out/Second Step in Checkout</title>
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
  <link runat="server" rel="shortcut icon" type="../image/png" href="favicon.ico">
  <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Lato:100,400,900|Josefin+Sans:100,400,700,400italic,700italic">
  <link rel="stylesheet" href="../css/main.css">
<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">
<script language="javascript" src="../SFLib/jquery-1.10.2.min.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" src="ssShippingRates.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">

function validateEmails(theForm)
{
	if (theForm.Email.value == '')
	{
		alert("Please enter a valid email address.");
		theForm.Email.focus();
		return false;
	}

	if (!emailCheck(theForm.Email.value))
	{
		theForm.Email.focus();
		return false;
	}

	if (theForm.Email.value != theForm.EmailVerification.value)
	{
		alert("Your email verification does not match. Please check your typing.");
		theForm.EmailVerification.focus();
		return false;
	}

	return true;
}

function clearShipping(form)
{
	for (var i=0; i < form.length; i++)
	{
		e = form.elements[i];
		if (e.name.indexOf("Ship") == 0)
		{
			if (e.name != "Shipping"){e.value = "";}
		}
	}
}

function validate_Me(frm)
{

	frm.Zip.zipcode=true;
	frm.ShipZip.zipcode=true;

	if(frm.Country.options.selectedIndex!="-1")
	{
		if (frm.Country.options[frm.Country.options.selectedIndex].value=="CA" || frm.Country.options[frm.Country.options.selectedIndex].value == "US")
		{
			frm.Zip.optional=false;
		}else{
			frm.Zip.optional=true;
		}
	}else{
		frm.Zip.optional=true;
	}

	if(frm.ShipCountry.options.selectedIndex!="-1")
	{
		if (frm.ShipCountry.options[frm.ShipCountry.selectedIndex].value=="CA" || frm.ShipCountry.options[frm.ShipCountry.selectedIndex].value == "US")
		{
			frm.ShipZip.optional=false;
		}else{
			frm.ShipZip.optional=true;
		}
	}else{
		frm.ShipZip.optional=true;
	}

	var bshipping_is_good = isCompleteShippingAddress(frm);
	if(bshipping_is_good)
	{
		if(frm.ShipCountry.options.selectedIndex!="-1")
		{
			if (frm.ShipCountry.options[frm.ShipCountry.selectedIndex].value == "US")
			{
				frm.ShipZip.optional=false;
			}else{
				frm.ShipZip.optional=true;
			}
		}



		var bmain_is_good = sfCheck(frm);
		return bmain_is_good;
	}else{
		window.alert("You have only entered a partial shipping address. To ensure your order is sent to the proper address please fill in all required shipping fields or no shipping fields at all. The following information is required for a complete shipping address. \n\nFirst Name\nLast Name\nStreet Address\nCity\nState/Province\nZip/Postal Code\nCountry\nPhone\n\nIf you desire to use your billing information for your shipping information press the Clear Shipping Fields button.")
		return false;
	}

	if(bmain_is_good == true)
	{
		return bshipping_is_good;
	}else{
		return false;
	}
}

function isCompleteShippingAddress(frm)
{
	var blnAllEmpty = true;

	blnAllEmpty = (blnAllEmpty && (frm.ShipFirstName.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipMiddleInitial.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipLastName.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipCompany.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipAddress1.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipAddress2.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipCity.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipZip.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipStateAlt.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipPhone.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipFax.value.length == 0));
	blnAllEmpty = (blnAllEmpty && (frm.ShipEmail.value.length == 0));

	if (frm.ShipCountry.selectedIndex != -1){blnAllEmpty = (blnAllEmpty && (frm.ShipCountry.options[frm.ShipCountry.selectedIndex].value==""));}
	//if (frm.ShipState.selectedIndex != -1){blnAllEmpty = (blnAllEmpty && (frm.ShipState.options[frm.ShipState.selectedIndex].value==""));}
	//alert("blnAllEmpty: " + blnAllEmpty);

	if (!blnAllEmpty)
	{
		var blnPasses = true;

		blnPasses = (blnPasses && (frm.ShipZip.length == 0 && frm.ShipZip.optional==false));
		//alert("check: {" + frm.ShipZip.value.length + "} - " + (frm.ShipZip.value != ""));
		if ((frm.ShipZip.value == "" && frm.ShipZip.optional==false)
		     || frm.ShipFirstName.value == ""
		     || frm.ShipLastName.value == ""
		     || frm.ShipAddress1.value == ""
		     || frm.ShipCity.value == ""
		     || frm.ShipCountry.options[document.form1.ShipCountry.selectedIndex].text  == ""
		     || (frm.ShipPhone.value == "" && frm.ShipPhone.optional==false)
		     || (frm.ShipEmail.value == "" && frm.ShipEmail.optional==false)
		   )
		{
			return false;
		}
	}

	return true;
}

function checkCountryChange(theCountry, theState, theAltState)
{
	var billingCountry=getSelectValue(theCountry);
	var countryHasStates=((billingCountry=='US') || (billingCountry=='us') || (billingCountry=='CA') || (billingCountry=='ca'));

	if (countryHasStates)
	{
		theState.disabled=false;
		theState.optional=false;
		theAltState.disabled=true;
		theAltState.optional=true;
	}else{
		letSelectValue(theState,"NA");
		theState.disabled=true;
		theState.optional=true;
		theAltState.disabled=false;
		theAltState.optional=false;
	}
}

function minZipLength(theZip)
{
	var destCountry;

	if (theZip.name.indexOf("Ship") == -1)
	{
		destCountry = getSelectValue(theZip.form.Country);
	}else{
		destCountry = getSelectValue(theZip.form.ShipCountry);
	}

	if ((destCountry=="US")||(destCountry=="us")){return 5}
	if ((destCountry=="CA")||(destCountry=="ca")){return 6}
	if ((destCountry=="")||(destCountry=="ca")){return 0}

	return 4;	//default minimum postal code length
}

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

function setShippingSameAsBilling(theElement)
{
	var theForm = theElement.form;
	var autoFillShippingEnabled = false;

	switch (theElement.name)
	{
		case "matchBillingAddress":
			if (theForm.matchBillingAddress.checked)
			{
				theForm.ShipFirstName.value = theForm.FirstName.value;
				theForm.ShipMiddleInitial.value = theForm.MiddleInitial.value;
				theForm.ShipLastName.value = theForm.LastName.value;
				theForm.ShipCompany.value = theForm.Company.value;
				theForm.ShipAddress1.value = theForm.Address1.value;
				theForm.ShipAddress2.value = theForm.Address2.value;
				theForm.ShipCity.value = theForm.City.value;
				letSelectValue(theForm.ShipState,getSelectValue(theForm.State));
				theForm.ShipStateAlt.value = theForm.altState.value;
				theForm.ShipZip.value = theForm.Zip.value;
				letSelectValue(theForm.ShipCountry,getSelectValue(theForm.Country));
				theForm.ShipPhone.value = theForm.Phone.value;
				theForm.ShipFax.value = theForm.Fax.value;
				theForm.ShipEmail.value = theForm.Email.value;
			}
			break;
		case "FirstName":
			if (autoFillShippingEnabled) theForm.ShipFirstName.value = theForm.FirstName.value;
			break;
		case "MiddleInitial":
			if (autoFillShippingEnabled) theForm.ShipMiddleInitial.value = theForm.MiddleInitial.value;
			break;
		case "LastName":
			if (autoFillShippingEnabled) theForm.ShipLastName.value = theForm.LastName.value;
			break;
		case "Company":
			if (autoFillShippingEnabled) theForm.ShipCompany.value = theForm.Company.value;
			break;
		case "Address1":
			if (autoFillShippingEnabled) theForm.ShipAddress1.value = theForm.Address1.value;
			break;
		case "Address2":
			if (autoFillShippingEnabled) theForm.ShipAddress2.value = theForm.Address2.value;
			break;
		case "City":
			if (autoFillShippingEnabled) theForm.ShipCity.value = theForm.City.value;
			break;
		case "State":
			if (autoFillShippingEnabled) letSelectValue(theForm.ShipState,getSelectValue(theForm.State));
			break;
		case "altState":
			if (autoFillShippingEnabled) if (autoFillShippingEnabled) theForm.ShipStateAlt.value = theForm.altState.value;
			break;
		case "Zip":
			if (autoFillShippingEnabled) theForm.ShipZip.value = theForm.Zip.value;
			break;
		case "Country":
			if (autoFillShippingEnabled) letSelectValue(theForm.ShipCountry,getSelectValue(theForm.Country));
			break;
		case "Phone":
			if (autoFillShippingEnabled) theForm.ShipPhone.value = theForm.Phone.value;
			break;
		case "Fax":
			if (autoFillShippingEnabled) theForm.ShipFax.value = theForm.Fax.value;
			break;
		case "Email":
			if (autoFillShippingEnabled) theForm.ShipEmail.value = theForm.Email.value;
			break;
		default:
			if (autoFillShippingEnabled) alert(theElement.name);
			break;
	}

	/* This
	theForm.ShipFirstName.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipMiddleInitial.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipLastName.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipCompany.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipAddress1.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipAddress2.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipCity.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipState.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipStateAlt.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipZip.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipCountry.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipPhone.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipFax.disabled = (theForm.matchBillingAddress.checked);
	theForm.ShipEmail.disabled = (theForm.matchBillingAddress.checked);
	*/

}
</script>

<script>
$(document).ready(function() {
	$('.tdAltFont1 a > b').unwrap();
	$('.tdAltFont2 a > b').unwrap();
	$('#tblMainContent td').css('padding', '5px');
})
</script>

<% writeCurrencyConverterOpeningScript %>

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
      <span class="title_txt" id="sub_title">ORDER PROCESSING</span>
    </div>
  </div>

<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent" style="margin: 0 auto 5%;">
  <tr>
    <td>
      <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
          <td align="center" class="tdMiddleTopBanner">Customer Information</td>
        </tr>
        <tr>
          <td class="tdBottomTopBanner" align="left">This is the final step in completing your order. You are now connected using a secure (SSL) connection and all information is transmitted in an encrypted form.
You should see a small key (Netscape) or lock (IE) indicating that your browser is communicating securely with our web store.
          </td>
        </tr>
        <tr>
          <td align="left" class="tdContent2"  valign="middle"><hr /><strong>Step 1: Customer Information</strong> | Step 2: Payment Information | Step 3: Complete Order<hr /></td>
        </tr>
        <tr>
          <td align="left" class="tdContent2"  valign="middle">
            <% If Not isLoggedIn And Not cblnDisableLogin And Not mblnHideLogin Then %>
		        <table border="0" cellpadding="2" cellspacing="0">
		          <tr>
		            <td align="left" class="tdContent2" valign="top">
			<table class="Section" cellpadding="2" cellspacing="0" border="1">
				<tr>
				<td class="tdTopBanner">Returning Customer Login</td>
				</tr>
				<tr>
				<td nowrap align="left">
					  <form name="frmPassword" id="frmPassword" action="process_order.asp" method="post" onsubmit="return ValidInput(this);">
					  <% If PayPalExpressCheckoutEnabled Then	'added for Sandshot Software's PayPal WebPayments Pro Integration %>
					  <input type="hidden" name="token" id="token" value="<%= PayPalToken %>">
					  <input type="hidden" name="xpressCheckout" id="xpressCheckout" value="PayPal">
					  <% End If %>
					  <table border="0" class="tdBottomTopBanner2" width="100%" cellpadding="3" cellspacing="1">
						<tr>
						  <td align="center" valign="middle" class="tdContent2">
					        <table class="tdContent2" border="0" width="100%" cellpadding="2">
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
				      <center>
				      <font class="Error"><b>
				      <% If sCondition = "EmailMatch" or sCondition = "WrongCombination" Then %>
				      <font color="red">Login Failed</font>
				      <% Else %>
						Login Directions
				      <% End If %>
				      </b></font>
				      <hr width="90%" noshade size="1">
				      </center>
					  <% If sCondition = "EmailMatch" Then %>
						Your combination was wrong, but an e-mail match was found. Please login with the correct password or if you wish to open a new account, you must choose a new password.
					  <% ElseIf sCondition = "WrongCombination" Then %>
						Your e-mail and password combination is incorrect. Try again.
					  <% Else %>
						Please use your e-mail address and password to log in and retrieve your customer information.
					  <% End If %>
		            </td>
		          </tr>
		        </table>
	        <% End If %>
          </td>
        </tr>
        <tr>
          <td align="left" class="tdContent2" valign="middle"><%= displayValidationError %></td>
        </tr>
        <tr valign="middle">
          <td class="tdContent2" align="left" valign="middle">
		<form method="post" action="verify.asp" name="form1" id="form1" onSubmit="<%= sSubmitAction %>">
		<input type="hidden" name="pageSource" id="pageSource" value="process_order">
		<input type="hidden" name="projectedDeliveryDate" id="projectedDeliveryDate" value="<%= getShipmentDate %>">
		<table class="tdContent2" border="0" cellpadding="2" cellspacing="0">
			<tr>
			<td valign="top">
			<table class="Section" cellpadding="2" cellspacing="0" border="1" width="100%">
				<tr>
				<td class="tdTopBanner">Billing Information</td>
				</tr>
				<tr>
				<td nowrap align="left">
			<table class="tdContent2" cellpadding="2" cellspacing="0" border="0">
				<tr>
				<td nowrap align="right"><font color="#FF0000">*</font>First Name:</td>
				<td nowrap><input type="text" maxlength="50" name="FirstName" ID="FirstName" title="First Name" size="20" style="<%= C_FORMDESIGN%>" value="<%= sCustFirstName %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right">MI:</td>
				<td nowrap><input type="text" name="MiddleInitial" ID="MiddleInitial" size="1" style="<%= C_FORMDESIGN%>" value="<%= sCustMiddleInitial %>" maxlength="1" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right"><font color="#FF0000">*</font>Last Name:</td>
				<td nowrap><input type="text" maxlength="50" name="LastName" ID="LastName" title="Last Name" size="20" style="<%= C_FORMDESIGN%>" value="<%= sCustLastName %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right">Company:</td>
				<td nowrap><input type="text" style="<%= C_FORMDESIGN%>" maxlength="50" name="Company" ID="Company" title="Company" size="25" value="<%= sCustCompany %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right"><font color="#FF0000">*</font>Address 1</td>
				<td nowrap><input type="text" style="<%= C_FORMDESIGN%>" maxlength="50" name="Address1" ID="Address1" title="Street Address" size="25" value="<%= sCustAddress1 %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right" height="15">Address 2:</td>
				<td nowrap height="15"><input type="text" style="<%= C_FORMDESIGN%>" maxlength="50" name="Address2" ID="Address2" title="Address2" size="25" value="<%= sCustAddress2 %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right"><font color="#FF0000">*</font>City:</td>
				<td nowrap><input type="text" maxlength="50" name="City" ID="City" title="City" size="20" class="formDesign" value="<%= sCustCity %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td align="right"><font color="#FF0000">*</font>State&nbsp;or<br />Province:</td>
				<td nowrap>
					<select name="State" ID="State" title="State" style="<%= C_FORMDESIGN%>" onchange="setShippingSameAsBilling(this);"><%= getStateList(sCustState) %></select>
					<br /><font size="-1">For U.S./Canadian addresses</font>
				</td>
				</tr>
				<tr>
				<td nowrap align="right">&nbsp;</td>
				<td nowrap><input type="text" maxlength="25" size="25" name="altState" id="altState" title="Province for addresses outside the U.S and Canada" style="<%= C_FORMDESIGN%>" value="<%= mstrAltState %>" onchange="setShippingSameAsBilling(this);"><br /><font size="-1">For addresses outside the U.S./Canada</font></td>
				</tr>
				<tr>
				<td nowrap align="right"><font color="#FF0000">*</font>Postal Code:</td>
				<td nowrap><input type="text" maxlength="12" name="Zip" ID="Zip" title="Zip Code" size="6" style="<%= C_FORMDESIGN%>" value="<%= sCustZip %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right" height="15"><font color="#FF0000">*</font>Country:</td>
				<td nowrap height="15">
					<select name="Country" id="Country" title="Country" class="formDesign" onchange="setShippingSameAsBilling(this);checkCountryChange(this, this.form.State, this.form.altState);"><%= getCountryList(sCustCountry, adminOriginCountry) %></select></td>
				</tr>
				<tr>
				<td nowrap align="right"><font color="#FF0000">*</font>Phone:</td>
				<td nowrap><input type="text" name="Phone" ID="Phone" maxlength="20" title="Phone Number" size="20" style="<%= C_FORMDESIGN%>" value="<%= sCustPhone %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right">Fax:</td>
				<td nowrap><input type="text" name="Fax" ID="Fax" maxlength="20" title="Fax Number" size="20" style="<%= C_FORMDESIGN%>" value="<%= sCustFax %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<tr>
				<td nowrap align="right"><font color="#FF0000">*</font>E-Mail:</td>
				<td nowrap><input type="text" style="<%= C_FORMDESIGN%>" maxlength="100" name="Email" ID="Email" title="Email Address" size="25" value="<%= sCustEmail %>" onchange="setShippingSameAsBilling(this);"></td>
				</tr>
				<% If Not isLoggedIn And cblnIncludeEmailVerification Then %>
				<tr>
				<td nowrap align="right"><font color="#FF0000">*</font>Verify E-Mail:</td>
				<td nowrap><input type="text" style="<%= C_FORMDESIGN%>" maxlength="100" name="EmailVerification" ID="EmailVerification" title="Email Address" size="25" value=""></td>
				</tr>
				<% End If	'isLoggedIn %>
				<tr>
				<td nowrap align="right">&nbsp;</td>
				<td nowrap><% If adminEmailActive = "1" Then %><input type="checkbox" name="Subscribe" id="Subscribe" value="1" <%if trim(sCustSubscribed) = "1" or trim(sCustSubscribed)="" then Response.write "checked" %>>&nbsp;Add to mailing list<% End If %></td>
				</tr>
				<!--
				<tr>
				<td nowrap align="right">&nbsp;</td>
				<td nowrap><input type="checkbox" name="matchBillingAddress" id="matchBillingAddress" value="1" onclick="setShippingSameAsBilling(this);">&nbsp;Click here to use your billing address for your shipping address</td>
				</tr>
				-->
			</table>
				</td>
				</tr>
			</table>
			</td>
			<td>&nbsp;</td>
			<!-- Shipping Info -->
			<td valign="top">
			<table class="Section" cellpadding="2" cellspacing="0" border="1" width="100%">
				<tr>
				<td colspan="2" class="tdTopBanner">Shipping Information (If Different from Billing Information)</td>
				</tr>
				<tr>
				<td nowrap align="left">
			<table class="tdContent2" cellpadding="2" cellspacing="0" border="0" width="100%">
				<%
				If mclsCustomerShipAddress.NumAddresses > 0 Then

  					For i = 0 to mclsCustomerShipAddress.NumAddresses
						pstrOptionText_PriorShippingAddresses = paryPriorShippingAddresses(4, i) & ", " & paryPriorShippingAddresses(2, i) & " - " & paryPriorShippingAddresses(8, i)
						If paryPriorShippingAddresses(15, i) = 1 Then
							pstrOptions_PriorShippingAddresses = pstrOptions_PriorShippingAddresses & "<option selected>" & pstrOptionText_PriorShippingAddresses & "</option>"
						Else
							pstrOptions_PriorShippingAddresses = pstrOptions_PriorShippingAddresses & "<option>" & pstrOptionText_PriorShippingAddresses & "</option>"
						End If
  					Next
				%>
				<tr>
				<td colspan="2" align="right">
				<script language="javascript" type="text/javascript">

				function ChangeAddress(theSelect)
				{
				var theForm = theSelect.form;
				var theIndex = theSelect.selectedIndex;

				if (theIndex==0)
				{
				clearShipping(theForm);
				return false;
				}

				var parycshpaddrID = new Array;
				var parycshpaddrCustID = new Array;
				var parycshpaddrShipFirstName = new Array;
				var parycshpaddrShipMiddleInitial = new Array;
				var parycshpaddrShipLastName = new Array;
				var parycshpaddrShipCompany = new Array;
				var parycshpaddrShipAddr1 = new Array;
				var parycshpaddrShipAddr2 = new Array;
				var parycshpaddrShipCity = new Array;
				var parycshpaddrShipState = new Array;
				var parycshpaddrShipZip = new Array;
				var parycshpaddrShipCountry = new Array;
				var parycshpaddrShipPhone = new Array;
				var parycshpaddrShipFax = new Array;
				var parycshpaddrShipEmail = new Array;
				var parycshpaddrIsActive = new Array;

					parycshpaddrID[0] = "";
					parycshpaddrCustID[0] = "";
					parycshpaddrShipFirstName[0] = "";
					parycshpaddrShipMiddleInitial[0] = "";
					parycshpaddrShipLastName[0] = "";
					parycshpaddrShipCompany[0] = "";
					parycshpaddrShipAddr1[0] = "";
					parycshpaddrShipAddr2[0] = "";
					parycshpaddrShipCity[0] = "";
					parycshpaddrShipState[0] = "";
					parycshpaddrShipZip[0] = "";
					parycshpaddrShipCountry[0] = "";
					parycshpaddrShipPhone[0] = "";
					parycshpaddrShipFax[0] = "";
					parycshpaddrShipEmail[0] = "";
					parycshpaddrIsActive[0] = "";

				<% For i = 0 to mclsCustomerShipAddress.NumAddresses %>
					parycshpaddrID[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(0, i) & "") %>";
					parycshpaddrCustID[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(1, i) & "") %>";
					parycshpaddrShipFirstName[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(2, i) & "") %>";
					parycshpaddrShipMiddleInitial[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(3, i) & "") %>";
					parycshpaddrShipLastName[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(4, i) & "") %>";
					parycshpaddrShipCompany[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(5, i) & "") %>";
					parycshpaddrShipAddr1[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(6, i) & "") %>";
					parycshpaddrShipAddr2[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(7, i) & "") %>";
					parycshpaddrShipCity[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(8, i) & "") %>";
					parycshpaddrShipState[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(9, i) & "") %>";
					parycshpaddrShipZip[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(10, i) & "") %>";
					parycshpaddrShipCountry[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(11, i) & "") %>";
					parycshpaddrShipPhone[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(12, i) & "") %>";
					parycshpaddrShipFax[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(13, i) & "") %>";
					parycshpaddrShipEmail[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(14, i) & "") %>";
					parycshpaddrIsActive[<%= i + 1 %>] = "<%= Server.HTMLEncode(paryPriorShippingAddresses(15, i) & "") %>";
  				<% Next %>

					theForm.ShipFirstName.value = parycshpaddrShipFirstName[theIndex];
					theForm.ShipMiddleInitial.value = parycshpaddrShipMiddleInitial[theIndex];
					theForm.ShipLastName.value = parycshpaddrShipLastName[theIndex];
					theForm.ShipCompany.value = parycshpaddrShipCompany[theIndex];
					theForm.ShipAddress1.value = parycshpaddrShipAddr1[theIndex];
					theForm.ShipAddress2.value = parycshpaddrShipAddr2[theIndex];
					theForm.ShipCity.value = parycshpaddrShipCity[theIndex];
					for (var i = 0;  i < theForm.ShipState.options.length;  i++)
					{
						if (theForm.ShipState.options[i].value == parycshpaddrShipState[theIndex])	{theForm.ShipState.selectedIndex = i;}
					}
					theForm.ShipZip.value = parycshpaddrShipZip[theIndex];
					for (var i = 0;  i < theForm.ShipCountry.options.length;  i++)
					{
						if (theForm.ShipCountry.options[i].value == parycshpaddrShipCountry[theIndex])	{theForm.ShipCountry.selectedIndex = i;}
					}
					theForm.ShipPhone.value = parycshpaddrShipPhone[theIndex];
					theForm.ShipFax.value = parycshpaddrShipFax[theIndex];
					theForm.ShipEmail.value = parycshpaddrShipEmail[theIndex];

					checkCountryChange(theForm.ShipCountry, theForm.ShipState, theForm.ShipStateAlt);
				}

				</SCRIPT>
					Prior Shipping Address:&nbsp;&nbsp;
					<select name="priorAddress" id="priorAddress" onchange="ChangeAddress(this);">
					<option>Use a new shipping address</option>
					<%= pstrOptions_PriorShippingAddresses %>
					</select>
				</td>
				</tr>
		<%
				End If	'mclsCustomerShipAddress.NumAddresses > 1

				'added to set shipping = billing
				'If Len(sShipCustZip) = 0 Then sShipCustZip = sCustZip
				'If Len(sShipCustState) = 0 Then sShipCustState = sCustState
				'If Len(sShipCustCountry) = 0 Then sShipCustCountry = sCustCountry
		%>

				<tr>
				<td nowrap align="right">First Name:</td>
				<td nowrap><input type="text" maxlength="50" name="ShipFirstName" ID="ShipFirstName" size="20" style="<%= C_FORMDESIGN%>" value="<%= sShipCustFirstName %>"></td>
				</tr>
				<tr>
				<td nowrap align="right">MI:</td>
				<td nowrap><input type="text" name="ShipMiddleInitial" ID="ShipMiddleInitial" size="1" style="<%= C_FORMDESIGN%>" value="<%= sShipCustMiddleInitial %>" maxlength="1"></td>
				</tr>
				<tr>
				<td nowrap align="right">Last Name:</td>
				<td nowrap><input type="text" maxlength="50" name="ShipLastName" ID="ShipLastName" size="20" style="<%= C_FORMDESIGN%>" value="<%= sShipCustLastName %>"></td>
				</tr>
				<tr>
				<td nowrap align="right">Company:</td>
				<td nowrap><input type="text" style="<%= C_FORMDESIGN%>" maxlength="50" name="ShipCompany" ID="ShipCompany" size="25" value="<%= sShipCustCompany %>"></td>
				</tr>
				<tr>
				<td nowrap align="right">Address:</td>
				<td nowrap><input type="text" style="<%= C_FORMDESIGN%>" maxlength="50" name="ShipAddress1" ID="ShipAddress1" size="25" value="<%= sShipCustAddress1%>"></td>
				</tr>
				<tr>
				<td nowrap align="right">Address 2:</td>
				<td nowrap><input type="text" style="<%= C_FORMDESIGN%>" maxlength="50" name="ShipAddress2" ID="ShipAddress2" size="25" value="<%= sShipCustAddress2 %>"></td>
				</tr>
				<tr>
				<td nowrap align="right">City:</td>
				<td nowrap><input type="text" maxlength="50" name="ShipCity" ID="ShipCity" size="20" style="<%= C_FORMDESIGN%>" value="<%= sShipCustCity %>"></td>
				</tr>
				<tr>
				<td align="right">State&nbsp;or<br />Province:</td>
				<td nowrap><select size="1" name="ShipState" ID="ShipState" class="formDesign"><option value="">Select a State/Province</option><%= getStateList(sShipCustState) %></select><br /><font size="-1">For U.S./Canadian addresses</font></td>
				</tr>
				<%
				If (sShipCustCountry = "US" Or sShipCustCountry = "CA") Then
					mstrAltState = ""
				Else
					mstrAltState = sShipCustState
				End If
				%>
				<tr>
				<td nowrap align="right">&nbsp;</td>
				<td nowrap><input type="text" maxlength="25" size="25" name="ShipStateAlt" id="ShipStateAlt" title="Province for addresses outside the U.S and Canada" style="<%= C_FORMDESIGN%>" value="<%= mstrAltState %>"><br /><font size="-1">For addresses outside the U.S./Canada</font></td>
				</tr>
				<tr>
				<td nowrap align="right">Postal Code:</td>
				<td nowrap><input type="text" maxlength="12" name="ShipZip" ID="ShipZip" size="20" style="<%= C_FORMDESIGN%>" value="<%= sShipCustZip %>"></td>
				</tr>
				<tr>
				<td nowrap align="right" height="15">Country:</td>
				<td nowrap height="15"><select size="1" name="ShipCountry" ID="ShipCountry" style="<%= C_FORMDESIGN%>" onchange="checkCountryChange(this, this.form.ShipState, this.form.ShipStateAlt);"><option value="">Select a Country</option><%= getCountryList(sShipCustCountry, "") %></select></td>
				</tr>
				<tr>
				<td nowrap align="right">Phone:</td>
				<td nowrap><input type="text" maxlength="20" name="ShipPhone" ID="ShipPhone" size="20" style="<%= C_FORMDESIGN%>" value="<%= sShipCustPhone %>"></td>
				</tr>
				<tr>
				<td nowrap align="right">Fax:</td>
				<td nowrap><input type="text" maxlength="20" name="ShipFax" ID="ShipFax" size="20" style="<%= C_FORMDESIGN%>" value="<%= sShipCustFax %>"></td>
				</tr>
				<tr>
				<td nowrap align="right">E-Mail:</td>
				<td nowrap><input type="text" style="<%= C_FORMDESIGN%>" maxlength="100" name="ShipEmail" ID="ShipEmail" size="25" value="<%= sShipCustEmail %>"></td>
				</tr>
				<tr>
				<td>&nbsp;</td>
				<td><a href="javascript:clearShipping(window.document.form1);"><img border=0 src="<%= C_BTN23 %>" alt="Clear Shipping Information"></a></td>
				</tr>
			</table>
			</td>
			</tr>
			</table>
			</td>
			</tr>
			<tr>
			<td colspan="3" class="tdContent2" height="10"></td>
			</tr>
			<tr>
			<td valign="top" class="tdContent2">
			<%
				If mclsCartTotal.isOrderShipped Then
					If cblnssUsePostageRate AND CStr(adminShipType) = "2" Then
						'sShipMethods = getssShippingOptions
						If True Then
							sShipMethods = getssShippingOptions_new(visitorPreferredShippingCode)
							If Len(sShipMethods) > 0 Then sShipMethods = "<select name=Shipping ID=Shipping class=formDesign>" & sShipMethods & "</select>"
						Else
							sShipMethods = getssShippingOptions_Radio(visitorPreferredShippingCode)
						End If
					Else
						sShipMethods = getShippingList(visitorPreferredShippingCode, False)
						If Len(sShipMethods) > 0 Then sShipMethods = "<select name=Shipping ID=Shipping class=formDesign>" & sShipMethods & "</select>"
					End If
				End If
				If Len(sShipMethods) > 0 Then
			%>
			  <table class="Section" cellpadding="2" cellspacing="0" border="1" width="100%">
				<tr>
				<td class="tdTopBanner">Shipping Method</td>
				</tr>
				<tr>
				<td nowrap align="left">
				  <table class="tdContent2" cellpadding="2" cellspacing="0" border="0" width="100%">
					<tr>
						<td class="tdContent2" valign="top"><%= sShipMethods %></td>
						<td class="tdContent2" valign="top">
						<% Call DisplayShippingTimeMessage %>
						<% If CStr(adminShipType) = "2" Then %><a href="process_order.asp" onclick='ssGetRates(1); return false;'><img src='images/buttons/getrates.gif' border=0 alt="Get Shipping Rates"></a><% End If %>
						<a href="#" onClick="window.open('viewUPSMap.asp','IWIN', 'status=no,location=no,menu=no,scrollbars,width=550,height=600,');"><img border="0" src="images/upslink.gif" align="absbottom" alt="View UPS delivery times"></a>
						</td>
					</tr>
				  </table>
				</td>
				</tr>
			  </table>
			<% End If	'Len(sShipMethods) > 0 %>&nbsp;
			</td>
			<td>&nbsp;</td>
			<!-- Payment Method -->
			<td valign="top" class="tdContent2">
			<% If mclsCartTotal.AmountDue > 0 Or Not mclsCartTotal.CompleteCalculation Then	'check disabled for now since tax/shipping may not be calculated %>
			<table class="Section" cellpadding="2" cellspacing="0" border="1" width="100%">
				<tr>
				<td class="tdTopBanner">Payment Method</td>
				</tr>
				<tr>
				<td nowrap align="left">
			<table class="tdContent2" cellpadding="2" cellspacing="0" border="0" width="100%">
				<tr>
				<td valign="middle" nowrap align="left">
				  <% If PayPalExpressCheckoutEnabled Then	'added for Sandshot Software's PayPal WebPayments Pro Integration %>
					<div style="font-style:bold">PayPal Express</div>
					<input type="hidden" name="PayPalToken" id="PayPalToken" value="<%= PayPalToken %>">
					<input type="hidden" name="PayPalPayerID" id="PayPalPayerID" value="<%= PayPalPayerID %>">
					<input type="hidden" name="PaymentMethod" id="PaymentMethod" value="PayPal WebPayments">
				  <% Else %>
					<% If True Then %>
					<%= getPaymentList_radio("Credit Card") %>
					<% Else %>
					<select class="formDesign" name="PaymentMethod" ID="PaymentMethod"><%= getPaymentList("Credit Card") %></select>
					<% End If %>
				  <% End If	'PayPalExpressCheckoutEnabled %>
				</td>
				</tr>
			</table>
			</td>
			</tr>
			</table>
			<% End If	'mclsCartTotal.AmountDue > 0 %>
			</td>
			</tr>
			<tr>
			<td colspan="3" class="tdContent2" height="10"></td>
			</tr>
			<tr>
			<td class="tdContent2">
			<table class="Section" cellpadding="2" cellspacing="0" border="1" width="100%">
			  <tr>
				<td class="tdTopBanner">Special Instructions (Optional)</td>
			  </tr>
			  <tr>
			    <td nowrap align="left">
				<table class="tdContent2" cellpadding="2" cellspacing="0" border="0" width="100%">
					<tr>
					<td>
						<textarea rows="4" name="Instructions" ID="Instructions" cols="30" style="<%= C_FORMDESIGN%>"></textarea>
					</td>
					</tr>
				</table>
				</td>
			  </tr>
			</table>
			</td>
			<td>&nbsp;</td>
			<td valign="top">
			<table class="Section" cellpadding="2" cellspacing="0" border="1" width="100%">
			  <tr>
				<td class="tdTopBanner">Click &quot;Submit&quot; to enter payment
				information</td>
			  </tr>
			  <tr align="center">
			    <td><br /><input type="image" class="inputImage" src="<%= C_BTN20%>" name="Verify" ID="Verify"><br /></td>
			  </tr>
			</table>
			</td>
			</tr>
			<% If NOT (isLoggedIn) And Not cblnDisableLogin Then %>
			<tr>
			<td colspan="3" class="tdContent2" height="10"></td>
			</tr>
			<tr>
			  <td colspan="3" class="tdContent2">
				<table class="Section" cellpadding="2" cellspacing="0" border="1" width="100%">
			  	  <tr>
					<td colspan="2" class="tdTopBanner">New Customers: Choose Password</td>
			  	  </tr>
			  	  <tr>
			    	<td align="left">
					  <table class="tdContent2" cellpadding="2" cellspacing="0" border="0">
						<tr>
				  		  <td colspan="2">In order to serve you better, an account will
					be created for you as part of the checkout process. This will
					facilitate a speedier checkout for future orders. To specify a
					password, please enter it below. Otherwise, a password will be
					generated for you.
						  </td>
						</tr>
						<tr>
						<td align="right" width="15%">Password:</td>
						<td align="left" width="85%"><input type="password" name="Password" ID="Password1" maxlength="10" title="Password" style="<%= C_FORMDESIGN%>" size="20"></td>
						</tr>
						<tr>
						<td align="right" nowrap>Password Confirmation:</td>
						<td align="left"><input type="password" name="Password2" ID="Password2" maxlength="10" title="Password Confirmation" style="<%= C_FORMDESIGN%>" size="20"></td>
						</tr>
					  </table>
					</td>
			  	  </tr>
				</table>
			  </td>
			</tr>
			<% End If	'If NOT (isLoggedIn) And Not cblnDisableLogin %>
			</table>
		</form>
		<script language=javascript type="text/javascript">
			checkCountryChange(document.form1.Country, document.form1.State, document.form1.altState);
			//checkCountryChange(document.form1.ShipCountry, document.form1.ShipState, document.form1.ShipStateAlt);
		</script>
              </table>
            </td>
          </tr>
        </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
  <div id="footer">
    <ul id="horizontal-nav">
      <li id="current_page"><a title="Shopping Cart"><span><image src="../images/shopping_cart.png" alt="Shopping Cart" id="shopping_cart">MY SHOPPING CART</span></a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="../myAccount.asp" title="My Account">MY ACCOUNT</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="../footer/faqs/faqs.html" title="FAQ's">FAQ'S</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="../footer/privacy_policy/privacy_policy.html" title="Privacy Policy">PRIVACY POLICY</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="../footer/contact_us/contact_us.html" title="Contact Us">CONTACT US <font>(877) 796-6639</font></a></li>
    </ul>
  </div>
</body>
</html>
<% Call CleanupPageObjects %>