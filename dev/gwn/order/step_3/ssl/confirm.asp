<%@ Language=VBScript %>
<% Option Explicit %>
<% Server.ScriptTimeout = 300	'add some time to be sure the order completes
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
'*   This file's origins is confirm.asp
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
<!--#include file="SFLib/fraudRules.asp"-->
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLib/incAE.asp"-->
<!--#include file="SFLib/incConfirm.asp"-->
<!--#include file="SFLib/incCC.asp"-->
<!--#include file="SFLib/processor.asp"-->
<!--#include file="SFLib/ssclsPayPal.asp"-->
<!--#include file="SFLib/ssclsCustomer.asp"--> 
<!--#include file="SFLib/ssclsCustomerShipAddress.asp"--> 
<!--#include file="SFLib/ssincCustomFormValues.asp"-->
<!--#include file="SFLib/ssmodDownload.asp"-->
<!--#include file="confirm_Email.asp"-->
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
Dim mstrErrorMessage
Dim pblnMinimumRequiredShippingFields
Dim pblnNoShippingEntered
Dim sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2, sCustCity, sCustState, sCustZip, sCustCountry, sCustPhone, sCustFax, sCustEmail, sCustSubscribed
Dim sCustStateName, sCustCountryName, sMailPassword
Dim sShipCustFirstName, sShipCustMiddleInitial,sShipCustLastName, sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustZip, sShipCustCountry, sShipCustPhone, sShipCustFax, sShipCustCompany, sShipCustEmail
Dim sShipCustStateName, sShipCustCountryName
Dim sShipInstructions, sPaymentMethod, sCCList
Dim sTotalPrice, sTotalSTax, sTotalCTax, iCODAmount, sHandling, sShipping, sGrandTotal
Dim sProcErrMsg
Dim i
Dim sTransMethod, sPaymentServer, sLogin, sPassword, sMercType
Dim sCustCardType, sCustCardName, sCustCardNumber, sCustCardExpiryMonth, sCustCardExpiryYear, sCustCardExpiry, sCustCardTypeName, mstrPayCardCCV
Dim iRoutingNumber, sBankName, iCheckNumber, iCheckingAccountNumber
Dim sPOName, iPONumber
Dim sSubmitActionAE, sSubmitAction, Path
Dim aReferer(2)
Dim mblnPostBack	'is this being called from a non-integrated payment processor (ie. PayPal, WorldPay, etc.)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	The following constant, proc_live, controls the status of 
'	all payment processors supported:
'
		Const proc_live = 1 'is 'live' mode
'		Const proc_live = 0 'is 'test' mode
'
'	Before your store goes live, you are encouraged to run 
'	an order through with the constant set to 0 for testing.	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim sPhoneFaxPayType, iPhoneFaxType
Dim iPayID, iAddrID, sShipMethodName, iOrderID

'Dim iRow
'Dim sAuthResp,ProcResponse,ProcMessage,ProcCustNumber,ProcAddlData,ProcRefCode,ProcAuthCode,ProcMerchNumber,ProcActionCode,ProcErrMsg,ProcErrLoc,ProcErrCode,ProcAvsCode, iPremiumShipping
'Dim sGrandTotalOut , sTotalPriceOut 'SFUPDATE
'Dim dUnitPrice,bTaxShipIsActive,dTaxAble_Amount

'**********************************************************
'*	Functions
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

sTransMethod = adminTransMethod
sPaymentServer = adminPaymentServer
sLogin = adminLogin
sPassword = adminPassword
sMercType = adminMerchantType

'**********************************************************
'**********************************************************

iCustID = visitorLoggedInCustomerID
iAddrID = visitorShipAddressID
sPaymentMethod = visitorPaymentmethod

'Now for some checks; they shouldn't get to this point without these values
If Not isValidRecordID(iCustID) Or Not isValidRecordID(iAddrID) Then
	Call addValidationError("For your security you have been returned to this page. This may be due to an extended period of inactivity. Your order was <strong>NOT</strong> completed nor was your payment information collected. Please resubmit your order.")
	Call cleanupPageObjects
	Call returnValidationErrorToSender("process_order.asp")	'use adminSSLPath?
End If

sShipInstructions = visitorInstructions

aReferer(0) = visitor_REFERER
aReferer(1) = vistor_HTTP_REFERER
aReferer(2) = visitor_REMOTE_ADDR

'Set billing address information
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
		sMailPassword		= .custPasswd
	Else
		'Houston, we have a problem
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
	End If
End	With	'mclsCustomerShipAddress

Set mclsCartTotal = New clsCartTotal
With mclsCartTotal
	.Connection = cnn

	.City = mclsCustomerShipAddress.City
	.State = mclsCustomerShipAddress.State
	.ZIP = mclsCustomerShipAddress.Zip
	.Country = mclsCustomerShipAddress.Country
	.isCODOrder = CBool(sPaymentMethod = cstrCODTerm)

	.ShipMethodCode = visitorPreferredShippingCode
	.LoadAllShippingMethods = False
	
	.LoadCartContents
	.checkInventoryLevels
	sShipMethodName = .ShipMethodName
	'.writeDebugCart	
	
	If .isEmptyCart Then
		Session.Abandon
		Call cleanupPageObjects	'Clean up before the redirect
		Response.Redirect(adminDomainName & cstrTimeoutRedirectPage)
	ElseIf .isStockDepleted Then
		Call cleanupPageObjects	'Clean up before the redirect
		Response.Redirect(adminDomainName & "order.asp")
	End If
	
End With	'mclsCartTotal

'This section is disabled which means
'INTERNETCASH, PAYPAL, WORLDPAY, AND ???CSVP Disabled
If True Then
	mblnPostBack = False		
Else
	If Request.QueryString("message") <> "" Then
		Call InternetCashResp
		mblnPostBack = True		
	ElseIf Request.Form("custom") <> "" Then
		Call PayPalResp("1") ' #321
		mblnPostBack = True		
	ElseIf Request("CSVPOSRESPONSE") <> "" Then
		Call CSVPOSResp
		mblnPostBack = True		
	ElseIf Request.QueryString("wpresponse") <> "" Then 
		Call WorldPayResp("1") ' #321
		mblnPostBack = True		
	Else
		mblnPostBack = False		
	End If
	
	'This section is disabled which means
	'INTERNETCASH, PAYPAL, WORLDPAY, AND ???CSVP Disabled
	If False Then
		If Request.item("custom") <> "" Then
			Call PayPalResp("2")
		ElseIf Request.QueryString("wpresponse") <> "" Then 
			Call WorldPayResp("2")
		End If
	End If

End If
	
If Not mblnPostBack Then

	With Request
		'for credit cards
		sCustCardType = Trim(.Form("CardType"))
		sCustCardName = Trim(.Form("CardName"))
		sCustCardNumber = Trim(.Form("CardNumber"))
		sCustCardExpiryMonth = Trim(.Form("CardExpiryMonth"))
		If len(sCustCardExpiryMonth) = 1 Then sCustCardExpiryMonth = "0" & sCustCardExpiryMonth
		sCustCardExpiryYear = Trim(.Form("CardExpiryYear"))
		If Len(sCustCardType) > 0 Then sCustCardTypeName = getTransactionName(sCustCardType)
		sCustCardExpiry = sCustCardExpiryMonth & "/" & sCustCardExpiryYear
		mstrPayCardCCV = Trim(.Form("payCardCCV"))
		
		'for eCheck
		iRoutingNumber = Trim(.Form("RoutingNumber"))
		sBankName = Trim(.Form("BankName"))
		iCheckNumber = Trim(.Form("CheckNumber"))
		iCheckingAccountNumber = Trim(.Form("CheckingAccountNumber"))
		
		'for PO
		sPOName = Trim(.Form("POName"))
		iPONumber = Trim(.Form("PONumber"))
	End With	'Request
	
	iPayID = 0	'Set default
	If mclsCartTotal.AmountDue = 0 Then
		'pass payment verification
	ElseIf sPaymentMethod = "Credit Card" AND sTransMethod <> "15" AND sTransMethod <> "18" Then
		If Len(sCustCardType) = 0 Then Call addValidationError("<em>Card Type</em> is a required field")
		If Len(sCustCardName) = 0 Then Call addValidationError("<em>Name on card</em> is a required field")
		If Len(sCustCardNumber) = 0 Then Call addValidationError("<em>Card Number</em> is a required field")
		If Len(sCustCardExpiryMonth) = 0 Then Call addValidationError("<em>Expiration Date Month</em> is a required field")
		If Len(sCustCardExpiryYear) = 0 Then Call addValidationError("<em>Expiration Date Year</em> is a required field")
		If Len(mstrPayCardCCV) = 0 Then
			If Not cstrCCV_Optional Then Call addValidationError("<em>Card Validation Code</em> is a required field")
		ElseIf Not isNumeric(mstrPayCardCCV) Then
			Call addValidationError("<em>Card Validation Code</em> is in an improper format")
		End If
		If hasValidationError Then
			Call cleanupPageObjects
			Call returnValidationErrorToSender("verify.asp")
		End If
		
		iPayID = setPayments(iCustID, sCustCardType, sCustCardName, sCustCardNumber, sCustCardExpiryMonth, sCustCardExpiryYear, iCC)
	ElseIf sPaymentMethod	= cstrECheckTerm Then
		If Len(iRoutingNumber) = 0 Then Call addValidationError("<em>Bank Routing Number</em> is a required field")
		If Len(sBankName) = 0 Then Call addValidationError("<em>Bank Name</em> is a required field")
		If Len(iCheckNumber) = 0 Then Call addValidationError("<em>Check Number</em> is a required field")
		If Len(iCheckingAccountNumber) = 0 Then Call addValidationError("<em>Checking Account Number</em> is a required field")
	ElseIf sPaymentMethod = cstrPOTerm Then
		If Len(sPOName) = 0 Then Call addValidationError("<em>Purchase Order Name</em> is a required field")
		If Len(iPONumber) = 0 Then Call addValidationError("<em>PO Purchase Number</em> is a required field")
	ElseIf sPaymentMethod = cstrPhoneFaxTerm Then
		If Len(sCustCardNumber) > 0 Then
			sPhoneFaxPayType = "Credit Card"
		ElseIf Len(iRoutingNumber) > 0 Then
			sPhoneFaxPayType = cstrECheckTerm
		ElseIf Len(iPONumber) > 0 Then
			sPhoneFaxPayType = cstrPOTerm
		Else
			sPhoneFaxPayType = "Phone/Fax"
		End If			 
	End If

	If hasValidationError Then
		Call cleanupPageObjects
		Call returnValidationErrorToSender("verify.asp")
	End If
	
End If	'Not mblnPostBack

sTotalPrice = mclsCartTotal.SubTotalWithDiscount
sTotalSTax = mclsCartTotal.StateTax + mclsCartTotal.LocalTax
sTotalCTax = mclsCartTotal.CountryTax
iCODAmount = mclsCartTotal.COD
sHandling = mclsCartTotal.Handling
sShipping = mclsCartTotal.Shipping
sGrandTotal = mclsCartTotal.AmountDue
sShipInstructions = visitorInstructions

'If an orderID already exists then all that's necessary is to update the current order record
iOrderID = visitorOrderID
'Get the preliminary order number
If sPaymentMethod = "Credit Card" AND (sTransMethod <> "15" AND sTransMethod <> "18") Then
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,sPaymentMethod,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)		
Elseif sPaymentMethod = cstrECheckTerm Then
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,sPaymentMethod,iRoutingNumber,sBankName,iCheckNumber,iCheckingAccountNumber,"","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)
Elseif sPaymentMethod = cstrPOTerm Then
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,sPaymentMethod,"","","","",iPONumber,sPOName,sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)'changed iponame to sponame
Elseif sPaymentMethod = cstrPhoneFaxTerm Then
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,sPaymentMethod & "_" & sPhoneFaxPayType,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)
ElseIf sPaymentMethod = cstrCODTerm Then
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,sPaymentMethod,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,iCODAmount,aReferer)
ElseIf (sPaymentMethod = "PayPal Transaction" OR sPaymentMethod = "PayPal") Then
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,"PayPal Initial","","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)
ElseIf (sPaymentMethod = "PayPal WebPayments") Then
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,sPaymentMethod,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)
ElseIf (sPaymentMethod = "WorldPay") OR (sTransMethod = "15" OR sTransMethod = "18" OR sTransMethod = "5" OR sTransMethod = "6" OR sTransMethod = "12") Then
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,sPaymentMethod,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)
Else
	'Should never see this but go with the flow - figure it out in OM
	iOrderID = setOrderInitial(iOrderID, iCustID,iPayID,iAddrID,"UNKNOWN-REQUIRES INVESTIGATION",iRoutingNumber,sBankName,iCheckNumber,iCheckingAccountNumber,iPONumber,sPOName,sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,iCODAmount,aReferer)
End If
If Len(CStr(visitorOrderID)) = 0 Or CStr(visitorOrderID) = "0" Then Call setVisitorOrderID(iOrderID)

' Move from tmp order cart to orders, this way it is always recoverable even if the payment verification hangs
Call mclsCartTotal.saveOrderItemsToOrderDetails(iOrderID)

If cdbl(sGrandTotal) > 0 Then
	If sPaymentMethod = "PayPal WebPayments" Then
		sProcErrMsg =  PayPalExpressCheckout_confirm
	ElseIf sPaymentMethod = "Credit Card" Then
		Select Case CStr(sTransMethod)
			Case "2", "11", "13", "19":
				sProcErrMsg = AuthNet(proc_live, "1")
			Case "1":
				sProcErrMsg = CyberCash(proc_live)
			Case "16":
				sProcErrMsg = CyberCash(proc_live)
			Case "7":
				sProcErrMsg = LinkPoint(proc_live)
			Case "15", "PayPal":
				sProcErrMsg = PayPal()
			Case "8":
				sProcErrMsg = PSIGate(proc_live)
			Case "10":
				sProcErrMsg = SecurePay(proc_live)
			Case "3", "17":
				sProcErrMsg = SignioPayProFlow(proc_live)
			Case "4":
				sProcErrMsg = SurePay(proc_live)
			Case "18", "WorldPay":
				sProcErrMsg = WorldPay(proc_live)
			Case PayPalTransactionMethodID:	'added for Sandshot Software's PayPal WebPayments Pro Integration
				sProcErrMsg = PayPalWebPayments(proc_live)
			Case "21", "Orbital":	'Note: Orbital Is NOT installed by default
				sProcErrMsg = Orbital(proc_live)
			Case Else:
				sProcErrMsg = SimulatedProcessor(proc_live)
		End Select
	
		'For debugging
		'If sCustCardNumber = "4111111111111111" Then sProcErrMsg = ""
		'sProcErrMsg = "Invalid Address"
		If Len(Session("ssDebug_PreventOrderCompletion")) > 0 Then sProcErrMsg = "Order Processing Disabled for Debugging"
		
	ElseIf sPaymentMethod = "PayPal" Then
	
	Else
		'To do? add validation for other payment methods?
	End If  
End If

'For testing
'sProcErrMsg = "Your credit card was denied!"

If Not goodProcessorResponse(4) Then
	Call addValidationError("<b><font color='red'>An error occurred while processing your order</font></b><hr noshade width='100%' size='1'><b>Error Message: " & sProcErrMsg & "</b><br />")
	Call cleanupPageObjects
	Call returnValidationErrorToSender("verify.asp")
End If

If goodProcessorResponse(1) Then

	'Expire CustID before processing order
	Call expireCookie_sfCustomer       	

	Call mclsCartTotal.finalizeCart(iOrderID)

	If cblnSF5AE Then Confirm_SaveAmounts(iOrderID)

	If  iPhoneFaxType <> "1"  Then

     	' Set Order complete flag to 1
		If goodProcessorResponse(2) Then Call setOrderComplete(iOrderID)
		Call SaveBuyersClubOrder(cnn, iOrderID)
		If (sPaymentMethod = "PayPal Transaction" OR sPaymentMethod = "PayPal") Then Call saveFreeGift(iOrderID)

		' Begin email 
		If goodProcessorResponse(3) Then Call sendOrderConfirmationEmails(sCustEmail)
		
		Call setCookie_ReturningOrder(iOrderID)	'Set Cookie For NewOrder Page
	End If

End If	'Len(sProcErrMsg) = 0
	  
If iPhoneFaxType = "1" Then DeleteOrder iOrderId
   
'Time To Clean up
If goodProcessorResponse(4) Then
	Call setVisitorOrderID(0)
	
	'Keeps login status
	Call updateVisitorOrderVerification(visitorLoggedInCustomerID, 0, 0, visitorPreferredShippingCode, "")
	'Automatically resets login status
	'Call updateVisitorOrderVerification(0, 0, 0, visitorPreferredShippingCode, "")
	
	Call expireCookie_sfCustomer

End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Order Confirmation Page</title>
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
<% writeCurrencyConverterOpeningScript %>
</head>

<body onload="confirm_onload()" <%= mstrBodyStyle %>>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
	  <tr>
    	<td>
	      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        	<tr>
	          <td colspan="2" align="center" class="tdMiddleTopBanner">
			   </td>
	        </tr>
    	    <tr>
        	  <td align="left" colspan="2" class="tdBottomTopBanner">
            	Thank you for your Order. Following is a summary of the transaction. Please print it out to keep as a receipt of your transaction.
			  </td>
	        </tr>
    	    <tr>
        	  <td colspan="2" width="100%" align="center" class="tdContent" valign="middle"><hr />Step 1: Customer Information | Step 2: Payment Information | <strong>Step 3: Complete Order</strong><hr /></td>
	        </tr>
			<% If sPaymentMethod = cstrPhoneFaxTerm Or sPaymentMethod = cstrPOTerm Then %>
    	    <tr>
        	  <td colspan="2" width="100%" align="center" class="tdContent" valign="middle">
        	  <%= getPageFragmentByKey("POMailingInstructions") %>
			  </td>
	        </tr>
			<% End If 'sPaymentMethod = cstrPhoneFaxTerm Then %>
    	    <tr>
	    	  <td align="center" class="tdContent2" width="100%" colspan="2">        
            	<table border="0" width="100%" cellspacing="0" cellpadding="4">            
	              <% If iPhoneFaxType <> 1 Then %>
	              <tr align="center">
					  <td colspan="4" width="100%" align="center">				
						<table width="85%"  cellpadding="1" cellspacing="0" border="0" class="tdBottomTopBanner">   
						  <tr>
							<td align="center" width="100%">
							  <table width="100%"  cellpadding="8" cellspacing="0" class="tdContentBar">
								<tr>
								  <td align="center" width="100%" class="tdAltBG1"><font class="Content_Large">
								    <% If goodProcessorResponse(4) Then %>
								    <b>Your Order ID is <%= iOrderID %></b>
								    </font> 
								    <br />
								    <font size="-1"> <b>Please print out or write down this number for future reference. </b></font>
								    <% Else %>
							    	<b><font color="red">An error occurred while processing your order</font></b>
								    <hr noshade width="100%" size="1">
								    <b>Error Message: <%= sProcErrMsg %> </b><br />
								    <a href="javascript:window.history.go(-2)">Resubmit</a>
								    <% End If %>
							     </td>
							  </tr>
						    </table>
					      </td>
					    </tr>
					  </table>
				    </td>
				  </tr>            
	              <% End If	'iPhoneFaxType <> 1 %>
                  
                  <% If goodProcessorResponse(4) Then %>
				  <tr>
				    <td width="100%" colspan="4" class="tdContent2"><% mclsCartTotal.displayOrder_CheckoutView %></td>
				  </tr>
                  <%
                  If mclsCartTotal.hasDownloadableItems Then
					If mlngFraudScore <= clngFraudThreshhold_ImmediateDownload Then	
						Call authorizeDownloadsByOrder(iOrderID)
					%>
					<tr>
						<td width="100%" colspan="4" class="tdContent2"><a href="../priorOrders.asp">Your order is available for immediate download</a></td>
					</tr>
					<%
					Else
					%>
					<tr>
						<td width="100%" colspan="4" class="tdContent2">Your order is currently being processed. You will receive an email shortly with download instructions. If you have any questions please contact customer service with message code <em><%= mlngFraudScore %>.</em></td>
					</tr>
					<%
					End If
				  End If	'mclsCartTotal.hasDownloadableItems
				  %>
				  <% End If	'goodProcessorResponse(4) %>

            	</table>
	    	  </td>
    	    </tr>
		<%
		Call WritePayPalPaymentLine_SF5
		If sPaymentMethod = cstrECheckTerm Then
		%>
	    <tr>
	       <td colspan="2" class="tdContent2">
	              <table border="1" width="95%" class="tdAltBG2" cellpadding="1" cellspacing="0" align="center">
				      <tr><td>
					      <table border="0" cellpadding="4" cellspacing="0" width="100%" align="center">
					        <tr>
						      <td align="left"><b><font class="ECheck"><%= sCustFirstName & " " & sCustLastName %></font></b></td>
						      <td align="right"><font class="ECheck2"><b>Check Number: </b><%= iCheckNumber %></font></td>
					        </tr>
					        <tr>
						      <td align="left" colspan="2"><font class="ECheck"><%= sCustAddress1 %>
						        <br /><%= sCustAddress2 %></font>
						      </td>
					          <tr>
						        <td align="left" colspan="2"><font class="ECheck"><%= sCustCity & " "%> <%= sCustStateName %>, <%= sCustZip %></font></td>
					          </tr>
					          <tr>	
						        <td align="left" colspan="2"><font class="ECheck"><%= sCustCountryName %></font></td>						
					          </tr>
					          <tr>
						        <td align="left" colspan="2" height="10"></td>
					          </tr>
					          <tr>
					            <td align="center" colspan="2"><b><font class="ECheck">Pay the amount of : <%= FormatCurrency(mclsCartTotal.CartTotal)%></font></b>
					              <hr size="1" width="60%" color="#445566" align="center">
					            </td>
					          </tr> <%'added this close %>
					            <tr>
						          <td width="50%" height="20">&nbsp;</td>
						          <td align="center"><font class="ECheck2">Electronically Signed By: <b><%= sCustFirstName & " " & sCustLastName %></b></font>
						            <hr size="1" width="100%" class="tdAltBG1">
						          </td>
					              </tr>	
					              <tr>
						            <td align="left" colspan="2"><b><font class="ECheck"><%= sBankName %></font></b></td>
					              </tr>
					              <tr>
				      					<td colspan="2" align="center" class="tdAltBG1"><font class="ECheck2">Payment Authorized by Account Holder. Indemnification Agreement Provided by Depositor.</font></td>
				      				</tr>	
					                <tr>
						              <td colspan="2" align="center"><font class="ECheck2"><b><%= iRoutingNumber %>::<%= iCheckingAccountNumber %> </b></font></td>
						              </tr>
					</table>
			       </td>
				  </tr>
			     </table>
	           </td>
	    </tr>
		<% End If	'sPaymentMethod = cstrECheckTerm %>

		<% If sPaymentMethod = cstrPhoneFaxTerm Then %>
	    <tr>
	       <td colspan="2" class="tdContent2">
           		<table class="tdContent" border="0" width="100%" cellpadding="2" cellspacing="0">
		                      <tr><td width="100%" colspan="2">
			                      <table border="0" width="100%" cellspacing="0" cellpadding="3" class="tdContent2">
			                        <tr>
			                          <td width="100%" class="tdContentBar">Phone/Fax Printout</td>
			                        </tr>
			                      </table>
		                        </td></tr>	
	
	                          <!--Customer Information -->
	                          <tr><td width="100%" class="tdContent2" colspan="2">   
	                              <table border="0" width="100%" cellspacing="0" cellpadding="2">
	                                <tr>
		                              <td><b>Billing:</b></td>
		                              <td><b>Ship To:</b></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustFirstName %>&nbsp;&nbsp;<%= sCustMiddleInitial %>&nbsp;&nbsp;<%= sCustLastName %></td>
		                              <td><%= sShipCustFirstName %>&nbsp;&nbsp;<%= sShipCustMiddleInitial %>&nbsp;&nbsp;<%= sShipCustLastName %></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustCompany %></td>
		                              <td><%= sShipCustCompany%></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustAddress1 %></td>
		                              <td><%= sShipCustAddress1 %></td>
	                                </tr>       
                                    <%If sCustAddress2 <> "" Then%>
   	                                <tr>	
   		                              <td><%= sCustAddress2 %></td>
                                      <td><%= sShipCustAddress2%></td>
                                    </tr>    
                                    <%End If%>
	                                <tr>
		                              <td><%= sCustCity%>,&nbsp;<%= sCustStateName %>,&nbsp;<%= sCustZip%></td>
		                              <td><%= sShipCustCity%>,&nbsp;<%= sShipCustStateName %>,&nbsp;<%= sShipCustZip %></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustCountryName %></td>
		                              <td><%= sShipCustCountryName %></td>
	                                </tr>
	                                <tr>
		                              <td></td>
 		                              <td></td>
 	                                </tr>
	                                <tr>
		                              <td><%= sCustFax%></td>
		                              <td><%= sShipCustFax %></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustEmail%></td>
    	                              <td><%= sShipCustEmail %></td>
                                    </tr>
                                    <tr>
			                          <td width="100%" align="left" colspan="4">	
			                            <br /><b>Special Instructions:</b></td>
		                            </tr>
		                            <tr>
		                              <td><%= sShipInstructions%></td>
    	                            </tr>
                                  </table>

	                            </td>
	                          </tr>
	
	                          <tr><td width="100%" class="tdContent2" colspan="2">   
	                              <table border="0" width="100%" cellspacing="0" cellpadding="2">
	                                <% 	If sPhoneFaxPayType = "Credit Card" Then %>
		                            <tr>
			                          <td width="100%" align="left" colspan="4">	
			                            <br /><b>Credit Card information:</b></td>
		                            </tr>
		                            <tr>
			                          <td align="left">Card Type:</td><td align="left"><%= sCustCardTypeName %></td>
			                          <td align="left">Card Name:</td><td align="left"><%= sCustCardName %></td>
		                            </tr>
		                            <tr>
			                          <td align="left">Credit Card Number:</td><td align="left"><%= sCustCardNumber %></td>
			                          <td align="left">Credit Card Expiration Date:</td><td align="left"><%= sCustCardExpiry %></td>
		                            </tr>		
	                                <%	ElseIf sPhoneFaxPayType = cstrECheckTerm Then %>
		                            <tr>
			                          <td width="100%" align="left" colspan="4">	
			                            <br /><b>e-Check information:</b></td>
		                            </tr>
		                            <tr>
			                          <td>Account Number:</td> <td align="left"><%= iCheckingAccountNumber %></td>
			                          <td>Check Number:</td> <td align="left"><%= iCheckNumber %></td>
		                            </tr>
		                            <tr>
			                          <td>Bank Name:</td> <td align="left"><%= sBankName %></td>
			                          <td>Routing Number:</td> <td align="left"><%= iRoutingNumber %></td>
		                            </tr>	
		                            <tr>
			                          <td height="20" colspan="4">&nbsp;</td>
		                            </tr>
		                            <tr><td width="100%" colspan="4" align="center">	
			                            <table border="1" width="95%" class="tdAltBG2" cellpadding="1" cellspacing="0" align="center">
				                          <tr>
				                            <td align="center" class="tdAltBG1"><font class="ECheck2">Payment Authorized by Account Holder. Indemnification Agreement Provided by Depositor.</font></td>
				                            </tr>	
				                            <tr><td>
					                            <table border="0" cellpadding="4" cellspacing="0" width="100%" align="center">
					                              <tr>
						                            <td align="left"><b><font class="ECheck"><%= sCustFirstName & " " & sCustLastName %></font></b></td>
						                            <td align="right"><font class="ECheck2"><b>Check Number: </b><%= iCheckNumber %></font></td>
					                              </tr>
					                              <tr>
						                            <td align="left" colspan="2"><font class="ECheck"><%= sCustAddress1 %>
						                              <br /><%= sCustAddress2 %></font></td>
					
					                                <tr>
						                              <td align="left" colspan="2"><font class="ECheck"><%= sCustCity & " "%> <%= sCustStateName %>, <%= sCustZip %></font></td>
					                                </tr>
					                                <tr>	
						                              <td align="left" colspan="2"><font class="ECheck"><%= sCustCountryName %></font></td>						
					                                </tr>
					                                <tr>
						                              <td align="left" colspan="2" height="10"></td>
					                                </tr>
					                                <tr>
					                                  <td align="center" colspan="2"><b><font class="ECheck">Pay the amount of : <%= FormatCurrency(cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax))%></font></b>
					                                    <hr size="1" width="60%" color="#445566" align="center">
					                                  </td>
					                                  <tr>
						                                <td width="50%" height="20">&nbsp;</td>
						                                <td align="center"><font class="ECheck2">Electronically Signed By: <b><%= sCustFirstName & " " & sCustLastName %></b></font>
						                                  <hr size="1" width="100%" class="tdAltBG1"></td>
					                                    </tr>	
					                                    <tr>
						                                  <td align="left" colspan="2"><b><font class="ECheck"><%= sBankName %></font></b></td>
					                                    </tr>
					                                    <tr>
						                                  <td align="left"><font class="ECheck2"><b>Routing Number: </b> <%= iRoutingNumber %></font></td>
						                                  <td align="right"><font class="ECheck2"><b>Checking Account Number: </b> <%= iCheckingAccountNumber %> </font></td>
					                                    </tr>
					                                    <tr><td colspan="2" height="25"></td></tr>
					                                  </table>
			                                        </td></tr>
			                                    </table>
		                                      </td></tr></table>		
	                                    </td></tr>	
	                                    <% ElseIf sPhoneFaxPayType = cstrPOTerm Then %>
	                                    <tr>
		                                  <td width="100%" align="left" colspan="4">	
		                                    <br /><b>Purchase Order information:</b></td>
	                                    </tr>
	                                    <tr>
		                                  <td width="25%" align="left">Name:</td><td width="25%" align="left"><%=	sPOName %></td>
		                                  <td width="25%" align="left">Purchase Order Number:</td><td width="25%" align="left"><%= iPONumber %></td>
	                                    </tr>
	                                    <% End If	%>
	                                    <tr>
	                                      <td align="center" colspan="4" height="60" valign="middle">
											<a href="javascript:window.print();">Print this Page</a>
	                                      </td>
	                                    </tr>
	      </table>
		  </td>
		</tr>
		<% End If	'sPaymentMethod = cstrPhoneFaxTerm %>
              </table>
            </td>
          </tr>
        </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->

<script language="javascript" type="text/javascript">
<!--
function confirm_onload()
{
<% If Len(cstrGoogleAnalytics_uacct) > 0 Then Response.Write "__utmSetTrans();" %>
}
-->
</script>

<% If Len(cstrOvertureID) > 0 Then %>
<script language="javascript" type="text/javascript">
<!-- Overture Services Inc. 07/15/2003
var cc_tagVersion = "1.0";
var cc_accountID = "<%= cstrOvertureID %>";
var cc_marketID =  "0";
var cc_protocol="http";
var cc_subdomain = "convctr";
if(location.protocol == "https:")
{
	cc_protocol="https";
	cc_subdomain="convctrs";
}
var cc_queryStr = "?" + "ver=" + cc_tagVersion + "&aID=" + cc_accountID + "&mkt=" + cc_marketID +"&ref=" + escape(document.referrer);
var cc_imageUrl = cc_protocol + "://" + cc_subdomain + ".overture.com/images/cc/cc.gif" + cc_queryStr;
var cc_imageObject = new Image();
cc_imageObject.src = cc_imageUrl;
// -->
</script>
<% End If	'Len(cstrOvertureID) > 0 %>

<% If Len(cstrGoogleAdwords_conversion_id) > 0 Then %>
<!-- Google Conversion Code -->
<script language="javascript" type="text/javascript">
<!--
google_conversion_id = <%= cstrGoogleAdwords_conversion_id %>;
google_conversion_language = "en_US";
google_conversion_value = <%= mclsCartTotal.CartTotal %>;
google_conversion_label = "Purchase";
-->
</script>
<script language="JavaScript" src="https://www.googleadservices.com/pagead/conversion.js">
</script>
<noscript>
<a href="https://services.google.com/sitestats/en_US.html" target=_blank>
<img src="https://www.googleadservices.com/pagead/conversion/<%= cstrGoogleAdwords_conversion_id %>/?value=<%= mclsCartTotal.CartTotal %>&label=Purchase&hl=en">
</a>
</noscript>
<% End If	'Len(cstrGoogleAdwords_conversion_id) > 0 %>
<% If Len(cstrGoogleAnalytics_uacct) > 0 Then %>
<form style="display:none;" name="utmform" ID="utmform">
<textarea name="utmtrans" id="utmtrans">
<%
'Format for Google
'UTM:T|[order-id]|[affiliation]|[total]|[tax]|[shipping]|[city]|[state]|[country]

	Response.Write "UTM:T|" & mclsCartTotal.OrderID & "|" & "" & "|" & mclsCartTotal.CartTotal & "|" & CStr(mclsCartTotal.StateTax + mclsCartTotal.CountryTax) & "|" & mclsCartTotal.Shipping & "|"
	'For billing
	Response.Write mclsCustomer.custCity & "|" & mclsCustomer.custState  & "|" & mclsCustomer.countryName & vbcrlf
	'For shipping address
	'Response.Write mclsCustomerShipAddress.City & "|" & mclsCustomerShipAddress.State  & "|" & mclsCustomerShipAddress.countryName & vbcrlf
	Call mclsCartTotal.WriteGoogleAnalyticsEommerceTrackingItems
%>
</textarea>
</form>
<% End If	'Len(cstrGoogleAnalytics_uacct) > 0 %>
</body>

</html>
<% Call cleanupPageObjects %>