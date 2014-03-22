<%@ Language=VBScript %>
<% Option Explicit %>
<%
'********************************************************************************
'*
'*   verify.asp
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins is verify.asp APPVERSION: 50.4014.0.9
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
<!--#include file="SFLib/incAE.asp"-->
<!--#include file="SFLib/incConfirm.asp"-->
<!--#include file="SFLib/incVerify.asp"-->
<!--#include file="SFLib/incCC.asp"-->
<!--#include file="SFLib/processor_PayPalWebPayments.asp"-->
<!--#include file="SFLib/ssclsCustomer.asp"-->
<!--#include file="SFLib/ssclsCustomerShipAddress.asp"-->
<!--#include file="SFLib/ssincCustomFormValues.asp"-->
<%

'**********************************************************
'	Developer notes
'**********************************************************

'NOTES:

'page accepts customer billing information (required)
'may contain new user account password

'may accept shipping information
'should contain payment method and shipping method

'**********************************************************
'*	Page Level variables
'**********************************************************
Dim pblnMinimumRequiredShippingFields
Dim pblnNoShippingEntered
Dim sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2, sCustCity, sCustState, sCustZip, sCustCountry, sCustPhone, sCustFax, sCustEmail, sCustSubscribed
Dim sCustStateName, sCustCountryName, bCustSubscribed, sPassword
Dim sShipCustFirstName, sShipCustMiddleInitial,sShipCustLastName, sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustZip, sShipCustCountry, sShipCustPhone, sShipCustFax, sShipCustCompany, sShipCustEmail
Dim sShipCustStateName, sShipCustCountryName
Dim iShipMethod, sInstructions, sPaymentMethod, sTransMethod, sCCList
Dim i
Dim sSubmitActionAE, sSubmitAction, Path

'**********************************************************
'*	Functions
'**********************************************************

	Function StatesForCountry(byVal strCountry)

		Select Case UCase(strCountry)
			Case "CA", "US", "PR"
				StatesForCountry = True
			Case Else
				StatesForCountry = False
		End Select

	End Function	'CheckNoStateForCountry

	'**********************************************************

	Sub collectCustomerBillingDetail()

		sCustFirstName = Trim(Request.Form("FirstName"))
		sCustMiddleInitial = Trim(Request.Form("MiddleInitial"))
		sCustLastName = Trim(Request.Form("LastName"))
		sCustCompany = Trim(Request.Form("Company"))
		sCustAddress1 = Trim(Request.Form("Address1"))
		sCustAddress2 = Trim(Request.Form("Address2"))
		sCustCity = Trim(Request.Form("City"))
		sCustZip = Trim(Request.Form("Zip"))
		sCustCountry = Trim(Request.Form("Country"))
		If StatesForCountry(sCustCountry) Then
			sCustState = Trim(Request.Form("State"))
		Else
			sCustState = Trim(Request.Form("altState"))
		End If
		sCustPhone = Trim(Request.Form("Phone"))
		sCustFax = Trim(Request.Form("Fax"))
		sCustEmail = Trim(Request.Form("Email"))
		bCustSubscribed = Trim(Request.Form("Subscribe"))
		If Len(bCustSubscribed) = 0 Then bCustSubscribed = 0
		sPassword = Trim(Request.Form("Password"))

		'validate the billing information
		If Len(sCustFirstName) = 0 Then Call addValidationError("<em>First Name</em> is a required field")
		If Len(sCustLastName) = 0 Then Call addValidationError("<em>Last Name</em> is a required field")
		If Len(sCustAddress1) = 0 Then Call addValidationError("<em>Address</em> is a required field")
		If Len(sCustCity) = 0 Then Call addValidationError("<em>City</em> is a required field")
		If Len(sCustState) = 0 Then Call addValidationError("<em>State</em> is a required field")
		If Len(sCustCountry) = 0 Then Call addValidationError("<em>Country</em> is a required field")
		If Len(sCustPhone) = 0 Then Call addValidationError("<em>Phone</em> is a required field")
		If Len(sCustEmail) = 0 Then Call addValidationError("<em>Email</em> is a required field")
		If hasValidationError Then
			Call cleanupPageObjects
			Call returnValidationErrorToSender("process_order.asp")
		End If

	End Sub	'collectCustomerBillingDetail

	'**********************************************************

	Sub collectCustomerShippingDetail()

		sShipCustFirstName			= Trim(Request.Form("ShipFirstName"))
		sShipCustMiddleInitial		= Trim(Request.Form("ShipMiddleInitial"))
		sShipCustLastName			= Trim(Request.Form("ShipLastName"))
		sShipCustCompany			= Trim(Request.Form("ShipCompany"))
		sShipCustAddress1			= Trim(Request.Form("ShipAddress1"))
		sShipCustAddress2			= Trim(Request.Form("ShipAddress2"))
		sShipCustCity				= Trim(Request.Form("ShipCity"))
		sShipCustCountry			= Trim(Request.Form("ShipCountry"))
		If StatesForCountry(sShipCustCountry) Then
			sShipCustState			= Trim(Request.Form("ShipState"))
		Else
			sShipCustState			= Trim(Request.Form("ShipStateAlt"))
		End If
		sShipCustZip				= Trim(Request.Form("ShipZip"))
		sShipCustPhone				= Trim(Request.Form("ShipPhone"))
		sShipCustFax				= Trim(Request.Form("ShipFax"))
		sShipCustEmail				= Trim(Request.Form("ShipEmail"))

		pblnMinimumRequiredShippingFields = CBool(Len(sShipCustFirstName)>0 And _
												  Len(sShipCustLastName)>0 And _
												  (Len(sShipCustAddress1)>0 Or Len(sShipCustAddress2)>0)And _
												  Len(sShipCustCity)>0 And _
												  Len(sShipCustState)>0 And _
												  Len(sShipCustZip)>0 And _
												  Len(sShipCustCountry)>0)

		'this check is to suppress the error which results from the check above
		pblnNoShippingEntered = CBool((Len(sShipCustFirstName)=0) Or _
									  (Len(sShipCustMiddleInitial)=0) Or _
									  (Len(sShipCustLastName)=0) Or _
									  (Len(sShipCustCompany)=0) Or _
									  (Len(sShipCustAddress1)=0) Or _
									  (Len(sShipCustAddress1)=0) Or _
									  (Len(sShipCustCity)=0) Or _
									  (Len(sShipCustState)=0) Or _
									  (Len(sShipCustState)=0) Or _
									  (Len(sShipCustZip)=0) Or _
									  (Len(sShipCustCountry)=0) Or _
									  (Len(sShipCustCountry)=0) Or _
									  (Len(sShipCustPhone)=0) Or _
									  (Len(sShipCustFax)=0) Or _
									  (Len(sShipCustEmail)=0) _
									 )

		If Not pblnMinimumRequiredShippingFields Then
			With mclsCustomer
				Set .Connection = cnn
				sShipCustFirstName = .custFirstName
				sShipCustMiddleInitial = .custMiddleInitial
				sShipCustLastName = .custLastName
				sShipCustCompany = .custCompany
				sShipCustAddress1 = .custAddr1
				sShipCustAddress2 = .custAddr2
				sShipCustCity = .custCity
				sShipCustState = .custState
				sShipCustZip = .custZip
				sShipCustCountry = .CustCountry
				sShipCustPhone = .CustPhone
				sShipCustFax = .CustFax
				sShipCustEmail = .CustEmail
			End	With	'mclsCustomer

			sShipCustStateName			= sCustStateName
			sShipCustCountryName		= Trim(sCustCountryName)
		End If

	End Sub	'collectCustomerShippingDetail

	'**********************************************************

	Sub cleanupPageObjects

	On Error Resume Next

		Set mclsCustomer = Nothing
		Set mclsCustomerShipAddress = Nothing
		Set mclsCartTotal = Nothing
		Call cleanup_dbconnopen

		If Err.number <> 0 Then Err.Clear

	End Sub	'cleanupPageObjects

'**********************************************************
'*	Begin Page Code
'**********************************************************

'**********************************************************
'**********************************************************

'Disabled since some browser privacy programs block HTTP_REFERER
' #52 Start If InStr(1, Request.ServerVariables("HTTP_REFERER"), "process_order.asp",1) = 0 and Request.querystring("optionID")="" then	response.end

'Call displayVisitorPreferences	'for debugging

If Request.Form("pageSource") = "process_order" Then
	Call PayPalExpressCheckout_verify	'added for Sandshot Software's PayPal WebPayments Pro Integration
	Call collectCustomerBillingDetail

	Set mclsCustomer = New clsCustomer
	Set mclsCustomer.Connection = cnn
	' Check if custID exists

	If isValidRecordID(visitorLoggedInCustomerID) Then
		If mclsCustomer.LoadCustomer(visitorLoggedInCustomerID) Then
			iCustID = visitorLoggedInCustomerID
			Call setCookie_custID(iCustID, Date() + 730)
		Else
			Call setCookie_custID("", Now())
		End If
	ElseIf Len(sPassword) > 0 Then
	'check if this is a returning customer who simply didn't log in AND is using the same email/password as prior order
		If mclsCustomer.LoadCustomerByEmailPassword(sCustEmail, sPassword) Then
			iCustID = mclsCustomer.custID
			Call SetSessionLoginParameters(iCustID, sCustEmail)
		End If
	Else
	'shouldn't be here IF password is required so sent back to process order
	'	Call addValidationError("There was an error with the address information you provided. Please resubmit your information.")
	'	Call cleanupPageObjects
	'	Call returnValidationErrorToSender("process_order.asp")	'use adminSSLPath?
	End If	'Len(visitorLoggedInCustomerID) > 0

	With mclsCustomer
		.custFirstName = sCustFirstName
		.custMiddleInitial = sCustMiddleInitial
		.custLastName = sCustLastName
		.custCompany = sCustCompany
		.custAddr1 = sCustAddress1
		.custAddr2 = sCustAddress2
		.custCity = sCustCity
		.custState = sCustState
		.custZip = sCustZip
		.custCountry = sCustCountry
		.custPhone = sCustPhone
		.custFax = sCustFax
		.custEmail = sCustEmail
		.custIsSubscribed = sCustSubscribed
		.custLastAccess = Now()

		If iCustID = 0 Then
			If Len(sPassword) = 0 Then sPassword = generatePassword()
			.custPasswd = sPassword
			.custTimesAccessed = 1

			If .AddCustomer Then
				iCustID = .custID
				Call SetSessionLoginParameters(iCustID, sCustEmail)
			Else
				Call addValidationError(.Message)
			End If
		Else
			.custTimesAccessed = .custTimesAccessed + 1
			If Not .Update Then Call addValidationError(.Message)
		End If	'iCustID = 0

		sCustStateName = .stateName
		sCustCountryName = .countryName

	End With	'mclsCustomer
	If hasValidationError Then
		Call cleanupPageObjects
		Call returnValidationErrorToSender("process_order.asp")
	End If

	Call collectCustomerShippingDetail
	Set mclsCustomerShipAddress = New clsCustomerShipAddress
	With mclsCustomerShipAddress
		.Connection = cnn

		If CStr(VisitorShipAddressID) <> "0" And Len(CStr(VisitorShipAddressID)) > 0 Then Call .LoadAddress(VisitorShipAddressID)

		.CustID = mclsCustomer.custID
		.FirstName = sShipCustFirstName
		.MiddleInitial = sShipCustMiddleInitial
		.LastName = sShipCustLastName
		.Company = sShipCustCompany
		.Addr1 = sShipCustAddress1
		.Addr2 = sShipCustAddress2
		.City = sShipCustCity
		.State = sShipCustState
		.Zip = sShipCustZip
		.Country = sShipCustCountry
		.Phone = sShipCustPhone
		.Fax = sShipCustFax
		.Email = sShipCustEmail

		If .addressID = 0 Then
			If Not .addAddress Then Call addValidationError(.Message)
		Else
			If Not .Update Then Call addValidationError(.Message)
		End If

		sShipCustStateName = .stateName
		sShipCustCountryName = .countryName

	End	With	'mclsCustomerShipAddress
	If hasValidationError Then
		Call cleanupPageObjects
		Call returnValidationErrorToSender("process_order.asp")
	End If

	'Collect the form variables
	sPaymentMethod = Request.Form("paymentmethod")
	iShipMethod = Request.Form("Shipping")
	sInstructions = Trim(Request.Form("Instructions"))

	Call updateVisitorOrderVerification(mclsCustomer.custID, mclsCustomerShipAddress.addressID, sPaymentMethod, iShipMethod, sInstructions)

Else

	'kicked back from verify due to processor error
	Set mclsCustomer = New clsCustomer
	With mclsCustomer
		Set .Connection = cnn
		If .LoadCustomer(visitorLoggedInCustomerID) Then
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
			sCustStateName		= .stateName
			sCustCountryName	= .countryName
			sPassword		= .custPasswd
		Else
			'This can occur if the session times out. For security the logged in ID is cleared upon the session expiration
			Call addValidationError("<em>For your security you have been logged out due to an extended period of inactivity. You may continue checkout where you left off by logging back in.</em>")
			'Note: this check is inside the fatal error since you only get to this section IF there is already a validation error from confirm.asp
			If hasValidationError Then
				Call cleanupPageObjects
				Call returnValidationErrorToSender("process_order.asp")
			End If
		End If
	End	With	'mclsCustomer

	Set mclsCustomerShipAddress = New clsCustomerShipAddress
	With mclsCustomerShipAddress
		.Connection = cnn
		If .LoadAddress(VisitorShipAddressID) Then
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
			sShipCustStateName		= .stateName
			sShipCustCountryName	= .countryName
		Else
			'Houston, we have a problem
			Call addValidationError("<em>Internal application Error: Unable to locate customer billing record ID# " & VisitorShipAddressID & ". Please resubmit your information.</em>")
			'Note: this check is inside the fatal error since you only get to this section IF there is already a validation error from confirm.asp
			If hasValidationError Then
				Call cleanupPageObjects
				Call returnValidationErrorToSender("process_order.asp")
			End If
		End If
	End	With	'mclsCustomerShipAddress

	'Collect the form variables
	sPaymentMethod = visitorPaymentmethod
	iShipMethod = visitorPreferredShippingCode
	sInstructions = visitorInstructions
	pblnNoShippingEntered = True

	Call updateVisitorOrderVerification(mclsCustomer.custID, mclsCustomerShipAddress.addressID, sPaymentMethod, iShipMethod, sInstructions)

End If	'Request.Form("pageSource") <> "verify"

Set mclsCartTotal = New clsCartTotal
With mclsCartTotal
	.Connection = cnn

	.City = mclsCustomerShipAddress.City
	.State = mclsCustomerShipAddress.State
	.ZIP = mclsCustomerShipAddress.Zip
	.Country = mclsCustomerShipAddress.Country
	.isCODOrder = CBool(sPaymentMethod = cstrCODTerm)

	.ShipMethodCode = iShipMethod
	.LoadAllShippingMethods = False

	.LoadCartContents
	.checkInventoryLevels
	'.writeDebugCart

	'.displayOrder_CheckoutView

	If .isEmptyCart Then
		Session.Abandon
		Call cleanupPageObjects	'Clean up before the redirect
		Response.Redirect(adminDomainName & cstrTimeoutRedirectPage)
	ElseIf .isStockDepleted Then
		Call cleanupPageObjects	'Clean up before the redirect
		Response.Redirect(adminDomainName & "order.asp")
	End If

End With	'mclsCartTotal

sTransMethod = adminTransMethod
If sPaymentMethod = "Credit" Then sPaymentMethod = "Credit Card"

If mclsCartTotal.AmountDue > 0 Then
	If (sPaymentMethod = "Credit Card" AND (sTransMethod <> "15" AND sTransMethod <> "18" AND sPaymentMethod <> "PayPal")) Then
		sSubmitAction = "this.CardNumber.creditCardNumber = true;return sfCheck(this);"
		sSubmitAction = "this.CardExpiryMonth.special = true;this.CardExpiryYear.special = true;" & sSubmitAction
		'sSubmitAction = "this.CardStartMonth.optional = true;this.CardStartYear.optional = true;;this.CardIssueNumber.optional = true;" & sSubmitAction
		If cstrCCV_Optional And Len(cstrCCVFieldName) > 0 Then sSubmitAction = "this.payCardCCV.optional = true;" & sSubmitAction
	Elseif sPaymentMethod = cstrPhoneFaxTerm OR sPaymentMethod = cstrCODTerm Then
		sSubmitAction ="" ' "this.CardType.optional = true;this.CardName.special = true;this.CardNumber.special = true;this.CardExpiryMonth.special = true;this.CardExpiryYear.special = true;this.CheckNumber.special = true;this.BankName.special = true;this.RoutingNumber.special = true;this.CheckingAccountNumber.special = true;this.POName.special = true;this.PONumber.special = true;return sfCheck(this);"
	Elseif sPaymentMethod = cstrECheckTerm  Then
		'sSubmitAction = "return Check_EC_PO('CheckNumber', this);"
		sSubmitAction = "return ECheck(this);"
	Elseif sPaymentMethod = cstrPOTerm  Then
		sSubmitAction = "return POCheck(POName.value, PONumber.value);"
	Else
		sSubmitAction = "return sfCheck(this);"
	End If
	sSubmitAction = Replace(sSubmitAction, "return sfCheck(this);", "if (sfCheck(this)){return initiateConfirmation(this);}else{return false;}")
End If	'mclsCartTotal.AmountDue > 0

sCCList = getCreditCardList("")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Verification Page/Third Step in Checkout</title>
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
<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">
<script language="javascript" src="../SFLib/jquery-1.10.2.min.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
<!--

function ValidateMe(txtbox)
  {
     if(txtbox.value != "")
      {
       var sval;
       if(txtbox.name == 'CardExpiryMonth')
        {
         sval =txtbox.value
         if((sval < 1) || (sval > 12))
           {
            alert("Please enter a valid expiration month");
            txtbox.focus();
            return false;
           }
         if(sval < 10 && sval.length == 1 )
           {
             txtbox.value = "0" + sval;
           }
      }
     if(txtbox.name == 'CardExpiryYear')
      {
        var d = new Date();
        var yy = (d.getFullYear());
        var mm = (d.getMonth() +1);
        sval = txtbox.value
        if(document.frmVerify.CardExpiryMonth.value == "")
          {
           document.frmVerify.CardExpiryMonth.focus();
           return false;
          }
        if(txtbox.length < 4)
         {
          alert("Enter a valid 4 Digit Date ");
          txtbox.focus();
          return false;
         }
        if(sval < yy)
         {
          alert("Date is not Valid");
          txtbox.focus();
          return false;
         }
       if((sval == yy)&& (mm > document.frmVerify.CardExpiryMonth.value))
           alert("Date is not Valid");
           txtbox.focus();
           return false;
      }
   }
       return true;
}

var alreadyProcessed=false;
function initiateConfirmation(theForm)
{
	if (alreadyProcessed)
	{
		return false;
	}else{
		alreadyProcessed=true;
		return true;
	}
}
//-->
</SCRIPT>

<script>
$(document).ready(function() {
	$('.tdAltFont1 a > b').unwrap();
	$('.tdAltFont2 a > b').unwrap();
	$('#tblMainContent td').css('padding', '5px');
})
</script>

<% writeCurrencyConverterOpeningScript %>
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
<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<form method="post" name="frmVerify" ID="frmVerify" action="confirm.asp" onSubmit="<%= sSubmitAction %>">
<input type="hidden" name="pageSource" id="pageSource" value="verify">
<input type="hidden" name="PayPalToken" id="PayPalToken" value="<%= PayPalToken %>">
<input type="hidden" name="PayPalPayerID" id="PayPalPayerID" value="<%= PayPalPayerID %>">
<% Call WriteCustomHiddenFormFields %>

<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent" style="margin: 0 auto 5%;">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="center" class="tdMiddleTopBanner">Verify Your Order</td>
        </tr>
        <tr>
          <td class="tdBottomTopBanner">Please review all your information as listed below. Once the <b>Checkout</b> button is pressed, your transaction will be final.</td>
        </tr>
        <tr>
          <td width="100%" align="center" class="tdContent2" valign="middle"><hr />Step 1: Customer Information | <strong>Step 2: Payment Information</strong> | Step 3: Complete Order<hr /></td>
        </tr>
        <tr>
          <td class="tdContent2"><%= displayValidationError %></td>
        </tr>
        <tr>
          <td class="tdContent2" width="100%">
              <% mclsCartTotal.displayOrder_CheckoutView %>
          </td>
        </tr>
        <tr>
          <td class="tdContent2"></td>
        </tr>

        <!--Customer Information -->
        <tr>
          <td width="100%" class="tdContent2">
            <table border="0" width="100%" cellspacing="0" cellpadding="2">
              <tr>
		        <td colspan="2" width="100%" class="tdContentBar" align="left">Customer Information</td>
              </tr>
              <tr>
		        <td align="left"><strong>Billing:</strong></td>
		        <td align="left"><strong>Ship To:</strong></td>
              </tr>
              <tr>
		        <td align="left" valign="top">
				<% With mclsCustomer %>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr><td><%= .DisplayName %></td></tr>
					<% If Len(.custCompany) > 0 Then %><tr><td><%= .custCompany %></td></tr><% End If %>
					<tr><td><%= .custAddr1 %></td></tr>
					<% If Len(.custAddr2) > 0 Then %><tr><td><%= .custAddr2 %></td></tr><% End If %>
					<tr><td><%= .custCity  %>,&nbsp;<%= .stateName %>,&nbsp;<%= .custZip %></td></tr>
					<tr><td><%= .countryName %></td></tr>
					<% If Len(.custPhone) > 0 Then %><tr><td>Phone: <%= .custPhone %></td></tr><% End If %>
					<% If Len(.custFax) > 0 Then %><tr><td>Fax: <%= .custFax %></td></tr><% End If %>
					<tr><td>Email: <%= .custEmail %></td></tr>
				</table>
				<% End With	'mclsCustomer %>
		        </td>
		        <td align="left" valign="top">
				<% With mclsCustomerShipAddress %>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr><td><%= .DisplayName %></td></tr>
					<% If Len(.Company) > 0 Then %><tr><td><%= .Company %></td></tr><% End If %>
					<tr><td><%= .Addr1 %></td></tr>
					<% If Len(.Addr2) > 0 Then %><tr><td><%= .Addr2 %></td></tr><% End If %>
					<tr><td><%= .City  %>,&nbsp;<%= .stateName %>,&nbsp;<%= .Zip %></td></tr>
					<tr><td><%= .countryName %></td></tr>
					<% If Len(.Phone) > 0 Then %><tr><td>Phone: <%= .Phone %></td></tr><% End If %>
					<% If Len(.Fax) > 0 Then %><tr><td>Fax: <%= .Fax %></td></tr><% End If %>
					<tr><td>Email: <%= .Email %></td></tr>
				</table>
				<% End With	'mclsCustomer %>
		        </td>
              </tr>
            </table>
          </td>
        </tr>
        <% If Not pblnMinimumRequiredShippingFields And Not pblnNoShippingEntered Then %>
        <tr>
          <td width="100%" class="tdContent2">Error: The shipping information you entered is invalid. Your billing information has been used.</td>
        </tr>
        <% End If %>

        <!--Special Instructions-->
        <tr>
          <td width="100%" class="tdContent2">
			<table border="0" width="100%" cellspacing="0" cellpadding="2">
              <tr>
	            <td width="100%" colspan="2" class="tdContentBar" align="left">Special Instructions</td>
              </tr>
              <%If Len(sInstructions) = 0 Then %>
              <tr><td align="left" height="40" valign="middle"><i>None Specified</i></td></tr>
              <% Else %>
              <tr><td align="left" height="40" valign="middle"><font class="ECheck"><%= sInstructions %></font></td></tr>
              <% End If %>
			</table>
		  </td>
        </tr>

        <!-- Payment Selection -->
        <% If mclsCartTotal.AmountDue > 0 And ((sTransMethod <> "15" or sTransMethod <> "18") AND (sPaymentMethod <> "Credit Card" OR sPaymentMethod <> "PayPal")) Then %>
           <tr>
             <td width="100%" class="tdContent2">

           <% If (sPaymentMethod = "Credit Card" AND sTransMethod <> "15" AND sTransMethod <> "18") Then
				Call displayCreditCardPayment
			  ElseIf sPaymentMethod = "PayPal WebPayments" Then
		   %>
			      <table class="tdContent2" border="0" width="100%" cellpadding="0" cellspacing="0">
			        <tr>
			          <td width="100%" colspan="2">
			            <table border="0" width="100%" cellpadding="2" cellspacing="0" class="tdContentBar">
			              <tr>
			                <td width="100%" class="tdContentBar" align="left">Payment Method</td>
			              </tr>
			              <tr>
			              <td colspan="2" align="left" valign="middle" class="tdContent">
		        	        <div style="font-style:bold;margin-top:10pt;margin-bottom:10pt">PayPal Express</div>
			                </td>
			            </tr>
			            </table>
			          </td>
			        </tr>
			       </table>
		   <%
			  ElseIf sPaymentMethod = cstrPhoneFaxTerm Then
		   %>
			      <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
			        <tr>
			          <td width="100%" colspan="2">
			            <table border="0" width="100%" cellpadding="2" cellspacing="0" class="tdContentBar">
			              <tr>
			                <td width="100%" class="tdContentBar" align="left">Phone Fax Order Information</td>
			              </tr>
			              <tr>
			              <td colspan="2" align="center" valign="middle">
          			        <p style="margin-top:20pt">
		        	        Phone Fax method
			                <p style="margin-top:20pt">
			                </td>
			            </tr>
			            </table>
			          </td>
			        </tr>
			        <tr><td height="20">&nbsp;</td></tr>
			        <tr>
			          <td align="center">
				        <table width="80%">
				        <% if CheckPaymentMethod("Credit Card") = 1 then%>
				          <tr>
				            <td colspan="2" align="center"><b><font class="ECheck">Complete this section for Credit Card purchases</font></b><hr width="90%" size="1" noshade class="tdAltBG1"></td>
				          </tr>
				          <tr>
				            <td colspan="2" align="center"><% Call displayCreditCardPayment %></td>
				          </tr>
				        <% end if 'CheckPaymentMethod("Credit Card")=1 %>
				        <% if CheckPaymentMethod(cstrECheckTerm) = 1 then %>
				          <tr><td height="40">&nbsp;</td></tr>
				          <tr>
				            <td colspan="2" align="center"><b><font class="ECheck">Complete this section for eCheck purchases</font></b><hr width="90%" size="1" noshade class="tdAltBG1"></td>
				          </tr>
				          <tr>
				            <td colspan="2" align="center"><% Call displayECheckPayment %></td>
				          </tr>
				        <% end if 'CheckPaymentMethod(cstrECheckTerm)=1 %>
				        <% if CheckPaymentMethod(cstrPOTerm) = 1 then %>
				          <tr><td height="40">&nbsp;</td></tr>
				          <tr>
				            <td colspan="2" align="center"><b><font class="ECheck">Complete this section for Purchase Order purchases</font></b><hr width="90%" size="1" noshade class="tdAltBG1"></td>
				          </tr>
				          <tr>
				            <td colspan="2" align="center"><% Call displayPOPayment %></td>
				          </tr>
				        <% end if 'CheckPaymentMethod(cstrPOTerm)=1 %>
				        </table>
				      </td>
				    </tr>
			        <tr>
				      <td colspan="2"><p style="margin-top:10pt">&nbsp;</p></td>
			        </tr>
			      </table>
           <% ElseIf sPaymentMethod = cstrECheckTerm Then
				Call displayECheckPayment
			  ElseIf sPaymentMethod = cstrPOTerm Then
				Call displayPOPayment
			  ElseIf sPaymentMethod = cstrCODTerm Then
           %>
		      <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
		        <tr>
			      <td colspan="2" width="100%" class="tdContentBar">COD Payment Method</td>
		        </tr>
		        <tr>
			      <td colspan="2" align="center" valign="middle">
			        <p style="margin-top:20pt">COD payment method</p>
			        <p style="margin-top:20pt"></p>
			      </td>
		        </tr>
		      </table>
           <%End If%>
				</td></tr>
        <% End If 'mclsCartTotal.AmountDue > 0 And ((sTransMethod <> "15" or sTransMethod <> "18") AND (sPaymentMethod <> "Credit Card" OR sPaymentMethod <> "PayPal")) %>

                 <% Sub displayCreditCardPayment %>
				<table class="tdContent2" border="1" width="100%" cellpadding="0" cellspacing="0">
				  <tr><td align="left">
			<table class="Section" cellpadding="2" cellspacing="0" border="1">
				<tr>
				<td class="tdTopBanner">Credit Card Information</td>
				</tr>
				<tr>
				<td class="tdContent" nowrap align="left">
					  <table border="0" class="tdContent" width="100%" cellpadding="3" cellspacing="1">
			        <tr>
			          <td><b>Card Type</b><font color="#FF0000">*</font><b>:</b></td>
			          <td align="left">
			            <select name="CardType" title="Credit Card Type" class="formDesign"><%= sCCList %></select>
			          </td>
			        </tr>
			        <tr>
			          <td><b>Name on card<font color="#FF0000">*</font>:</b></td>
			          <td align="left"><input type="text" name="CardName" id="CardName" title="Name on Card" size="30" Style="<%= C_FORMDESIGN%>" value="<%= mclsCustomer.DisplayName %>"></td>
			        </tr>
			        <tr>
			          <td><b>Card Number<font color="#FF0000">*</font>:</b></td>
			          <td align="left"><input type="text" name="CardNumber" title="Credit Card Number" size="30" style="<%= C_FORMDESIGN%>"></td>
			        </tr>
			        <% If False Then %>
			        <tr>
			          <td align="right"><b>Start Date:</b></td>
			          <td align="left">
			          <select name="CardStartMonth" id="CardStartMonth" title="Start Month" style="<%= C_FORMDESIGN%>">
			            <option value=""></option>
			            <option value="01">01 - Jan</option>
			            <option value="02">02 - Feb</option>
			            <option value="03">03 - Mar</option>
			            <option value="04">04 - Apr</option>
			            <option value="05">05 - May</option>
			            <option value="06">06 - Jun</option>
			            <option value="07">07 - Jul</option>
			            <option value="08">08 - Aug</option>
			            <option value="09">09 - Sep</option>
			            <option value="10">10 - Oct</option>
			            <option value="11">11 - Nov</option>
			            <option value="12">12 - Dec</option>
			          </select>&nbsp;/&nbsp;
			          <select name="CardStartYear" id="CardStartYear" title="Start Year" style="<%= C_FORMDESIGN%>">
			            <option value=""></option>
			          <% For i = Year(Date) To Year(Date) + 10 %>
			            <option value="<%= i %>"><%= i %></option>
			          <% Next 'i %>
			          </select>
					  </td>
			        </tr>
			        <% End If %>
			        <tr>
			          <td><b>Expiration Date<font color="#FF0000">*</font>:</b></td>
			          <td align="left">
			          <select name="CardExpiryMonth" ID="CardExpiryMonth" title="Credit Card Month" style="<%= C_FORMDESIGN%>">
			            <option value="">Month</option>
			            <option value="01">01 - Jan</option>
			            <option value="02">02 - Feb</option>
			            <option value="03">03 - Mar</option>
			            <option value="04">04 - Apr</option>
			            <option value="05">05 - May</option>
			            <option value="06">06 - Jun</option>
			            <option value="07">07 - Jul</option>
			            <option value="08">08 - Aug</option>
			            <option value="09">09 - Sep</option>
			            <option value="10">10 - Oct</option>
			            <option value="11">11 - Nov</option>
			            <option value="12">12 - Dec</option>
			          </select>&nbsp;/&nbsp;
			          <select name="CardExpiryYear" ID="CardExpiryYear" title="Credit Card Year" style="<%= C_FORMDESIGN%>">
			            <option value="">Year</option>
			          <% For i = Year(Date) To Year(Date) + 10 %>
			            <option value="<%= i %>"><%= i %></option>
			          <% Next 'i %>
			          </select>
					  </td>
			        </tr>
			        <% If False Then %>
			        <tr>
					  <td align="right"><b>Issue Number:</b></td>
			          <td align="left"><input type="text" name="CardIssueNumber" title="Issue Number" size="4" style="<%= C_FORMDESIGN%>" id="CardIssueNumber">&nbsp;<font color="red">E.g. 01 (Required for Debit card Purchase)</font></td>
			        </tr>
			        <% End If %>
			        <% If Len(cstrCCVFieldName) > 0 Then %>
			        <tr>
			          <td valign="top"><b>Verification&nbsp;Number<font color="#FF0000">*</font>:</b></td>
			          <td align="left"><input type="text" name="payCardCCV" id="payCardCCV" size="5" title="Credit Card validation number">
			          &nbsp;<a href="ccvAbout/aboutCCV.htm" target="_blank">What's this?</a>
			          </td>
			        </tr>
			        <% End If	'Len(cstrCCVFieldName) > 0 %>
			      </table>
			      </td>
			      </tr>
			      </table>
			      </td></tr>
			      </table>
                 <% End Sub	'displayCreditCardPayment %>

                 <% Sub displayECheckPayment %>
				<table class="tdContent2" border="1" width="100%" cellpadding="0" cellspacing="0">
				  <tr><td>
					<table class="tdContent2" border="0" width="100%" cellpadding="4" cellspacing="0">
					  <tr>
						<td width="100%" colspan="2" class="tdContentBar">eCheck Information</td>
					  </tr>
					  <tr>
						<td align="left"><b><font class="ECheck">Check Number <font color="#FF0000">*</font>:</font></b></td>
						<td align="left"><font class="ECheck"><input type="Text" name="CheckNumber" title="Check Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>
					  </tr>
					  <tr>
						<td align="left"><font class="ECheck"><b>Checking Account Number <font color="#FF0000">*</font>:</b></font></td>
						<td align="left"><font class="ECheck"><input type="Text" name="CheckingAccountNumber" title="Checking Account Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>
					  </tr>
					  <tr>
						<td align="left"><b><font class="ECheck">Bank Name <font color="#FF0000">*</font>:</font></b></td>
						<td align="left"><font class="ECheck"><input type="Text" name="BankName" title="Bank Name" size="30" style="<%= C_FORMDESIGN%>"></font></td>
					  </tr>
					  <tr>
						<td align="left"><font class="ECheck"><b>Bank Routing Number <font color="#FF0000">*</font>:</b></font></td>
						<td align="left"><font class="ECheck"><input type="Text" name="RoutingNumber" title="Routing Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>
					  </tr>
					  <tr>
						<td colspan="2" height="20"></td>
					  </tr>
					</table>
				  </td></tr>
				</table>
                 <% End Sub	'displayECheckPayment %>

                 <% Sub displayPOPayment %>
				<table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
				<tr>
					<td colspan="2" width="100%" class="tdContentBar">Purchase Order Payment Information</td>
				</tr>
				<tr>
					<td align="left"><font class="ECheck"><b>Purchase Order Name <font color="#FF0000">*</font>:</b></font></td>
					<td align="left"><input type="text" size="25" name="POName" title="PO Name" class="formDesign" value="<%= mclsCustomer.DisplayName %>"></td>
				</tr>
				<tr>
					<td align="left"><font class="ECheck"><b>PO Purchase Number <font color="#FF0000">*</font>:</b></font></td>
					<td align="left"><input type="text" size="25" name="PONumber" title="PO Number" class="formDesign"></td>
				</tr>
				<tr>
					<td colspan="2"><p style="margin-top:10pt">&nbsp;</p></td>
				</tr>
				</table>
                 <% End Sub	'displayPOPayment %>

        <tr>
          <td width="100%" class="tdContent2" valign="top" align="center">
            <% If sPaymentMethod <> "InternetCash" OR  (sTransMethod <> "15" AND sPaymentMethod <> "Credit Card")Then %>
		    <input type="image" class="inputImage" src="<%= C_BTN05 %>" name="verify">
	        <% ElseIf sPaymentMethod = "InternetCash" Then %>
		    <font class="ECheck">You are using <b>InternetCash</b> to pay for your purchase, please enter payment information in the popup window and press continue.</font>
            <% End If %>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->
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
<% Call cleanupPageObjects %>