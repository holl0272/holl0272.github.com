<%
'********************************************************************************
'*   Sandshot Software PayPal WebPayment Pro									*
'*   Release Version:   1.00.001												*
'*   Release Date:      July 15, 2006											*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.				*
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	Const PayPalEnvironment = "sandbox"
	'Const PayPalEnvironment = "live"
	Const PayPalOrderDescription = "Order "

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim ActionCodeType
Dim PayPalToken
Dim PayPalPayerID
Dim PayPalExpressCheckoutEnabled:	PayPalExpressCheckoutEnabled = False
Dim PayPalTransactionMethodID:		PayPalTransactionMethodID = "20"

'***********************************************************************************************
'***********************************************************************************************

Class PayPalAPI

' Module variables
Private pobjPPcaller
Private pp_request
Private util
Private pstrAPIUsername
Private pstrAPIPassword
Private pstrAPISignature
Private pstrSubject
Private pstrEnvironment
Private pstrCurrency
Private pstrPaymentStatus
Private pstrToken
Private pblnIsSuccessful
Private pblnLocalDebug

	Private Sub Class_Initialize()
		pblnLocalDebug = True
		
		On Error Resume Next
		Set pobjPPcaller = CreateObject("com.paypal.sdk.COMNetInterop.COMAdapter2")
		If Err.number <> 0 Then
			Response.Write "<fieldset><legend>PayPal WebSite Payments Pro Error</legend>"
			Response.Write "<h3 style=""color:red"">It appears the PayPal SDK is not properly installed</h3>"
			Response.Write "Error " & err.number & ": " & err.Description
			Response.Write "</fieldset>"
			Err.Clear
		End If
	End Sub

	Private Sub SetAPIProfile()
		If IsEmpty(pobjPPcaller) Then Set pobjPPcaller = CreateObject("com.paypal.sdk.COMNetInterop.COMAdapter2")
		With pobjPPcaller
			.SetAPIUsername pstrAPIUsername
			.SetAPIPassword pstrAPIPassword
			.SetAPISignature pstrAPISignature
			.SetSubject pstrSubject
			.SetEnvironment pstrEnvironment
		End With
		
		If False Then
			Response.Write "<fieldset><legend>PayPal WebSite Payments Pro - API Profile</legend>"
			Response.Write "APIUsername: " & pstrAPIUsername & "<br />"
			Response.Write "APIPassword: " & pstrAPIPassword & "<br />"
			Response.Write "APISignature: " & pstrAPISignature & "<br />"
			Response.Write "Subject: " & pstrSubject & "<br />"
			Response.Write "Environment: " & pstrEnvironment & "<br />"
			Response.Write "</fieldset>"
		End If
		
	End Sub

	Private Sub class_Terminate()
		On Error Resume Next
		If Not isEmpty(util) Then Set util = Nothing
		If Not isEmpty(pobjPPcaller) Then Set pobjPPcaller = Nothing
		If Not isEmpty(pp_request) Then Set pp_request = Nothing
		If Err.number <> 0 Then Err.Clear
	End Sub

	Public Property Get pp_caller
		Set pp_caller = pobjPPcaller
	End Property

	Public Property Let DebugEnabled(byVal vntValue)
		pblnLocalDebug = vntValue
	End Property

	Public Property Let APIUsername(byVal vntValue)
		pstrAPIUsername = vntValue
	End Property

	Public Property Let APIPassword(byVal vntValue)
		pstrAPIPassword = vntValue
	End Property

	Public Property Let APISignature(byVal vntValue)
		pstrAPISignature = vntValue
	End Property

	Public Property Let Subject(byVal vntValue)
		pstrSubject = vntValue
	End Property

	Public Property Let Environment(byVal vntValue)
		pstrEnvironment = vntValue
	End Property

	'***********************************************************************************************

	Public Sub TransactionSearch(startDate, endDate)
		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
		' Create the request object
		Set pp_request = CreateObject("com.paypal.soap.api.TransactionSearchRequestType")
		
		'Convert time to GMT time
		pp_request.StartDate = GetGMTDate(CDate(startDate))
		pp_request.EndDate = GetGMTDate( DateAdd("d", 1, CDate(endDate)) ) 'end date inclusive
		pp_request.EndDateSpecified = true

		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pobjPPcaller.CallAPI "TransactionSearch" 
	End Sub

	'***********************************************************************************************

	Public Sub GetTransactionDetails(trxID)
		' Create the request object
		Set pp_request = CreateObject("com.paypal.soap.api.GetTransactionDetailsRequestType")

		pp_request.TransactionID = trxID

		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pobjPPcaller.CallAPI "GetTransactionDetails"
	End Sub

	'***********************************************************************************************

	Public Sub RefundTransaction(trxID, refundType, amount)
		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
		
		' Create the request object
		Set pp_request = CreateObject("com.paypal.soap.api.RefundTransactionRequestType")

		With pp_request
			.TransactionID = trxID
			Select Case refundType
				Case "Full"
					.RefundType = util.GetEnumValue("RefundPurposeTypeCodeType", "Full")
					.RefundTypeSpecified = true
				Case "Partial"
					.RefundType = util.GetEnumValue("RefundPurposeTypeCodeType", "Partial")
					.RefundTypeSpecified = true
					Set .Amount = CreateObject("com.paypal.soap.api.BasicAmountType")
					.Amount.currencyID = util.GetEnumValue("CurrencyCodeType", "USD")
					.Amount.Value = amount
			End Select
		End With

		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pobjPPcaller.CallAPI "RefundTransaction"
	End Sub

	'***********************************************************************************************

	Public Sub DoDirectPayment(byVal paymentAmount, byRef aryBillingAddress, byRef aryShippingAddress, _
							   byVal creditCardType, byVal creditCardNumber, byVal CVV2, byVal expMonth, byVal expYear, byVal actionCodeType, byVal strDescription)
		
		If Len(expMonth) = 1 Then expMonth = "0" & expMonth
		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
		Set pp_request = CreateObject("com.paypal.soap.api.DoDirectPaymentRequestType")

		With pp_request
			' Create the request details object
			.DoDirectPaymentRequestDetails = CreateObject("com.paypal.soap.api.DoDirectPaymentRequestDetailsType")

			.DoDirectPaymentRequestDetails.IPAddress = Request.ServerVariables("REMOTE_ADDR")
			.DoDirectPaymentRequestDetails.MerchantSessionId = Session.SessionID
			.DoDirectPaymentRequestDetails.PaymentAction = util.GetEnumValue("PaymentActionCodeType", actionCodeType)
				
			.DoDirectPaymentRequestDetails.CreditCard = CreateObject("com.paypal.soap.api.CreditCardDetailsType")
				
			.DoDirectPaymentRequestDetails.CreditCard.CreditCardNumber = creditCardNumber	
			Select Case creditCardType
				Case "Visa":		.DoDirectPaymentRequestDetails.CreditCard.CreditCardType = util.GetEnumValue("CreditCardTypeType", "Visa")
				Case "MasterCard":	.DoDirectPaymentRequestDetails.CreditCard.CreditCardType = util.GetEnumValue("CreditCardTypeType", "MasterCard")
				Case "Discover":	.DoDirectPaymentRequestDetails.CreditCard.CreditCardType = util.GetEnumValue("CreditCardTypeType", "Discover")
				Case "Amex":		.DoDirectPaymentRequestDetails.CreditCard.CreditCardType = util.GetEnumValue("CreditCardTypeType", "Amex")
			End Select
			.DoDirectPaymentRequestDetails.CreditCard.CVV2 = CVV2
			.DoDirectPaymentRequestDetails.CreditCard.ExpMonth = expMonth
			.DoDirectPaymentRequestDetails.CreditCard.ExpYear = expYear
				
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner = CreateObject("com.paypal.soap.api.PayerInfoType")
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Payer = ""
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.PayerID = ""
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.PayerStatus = util.GetEnumValue("PayPalUserStatusCodeType", "unverified")
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.PayerCountry = util.GetEnumValue("CountryCodeType", "US")

			'aryBillingAddress: Array(sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2, sCustCity, sCustState, sCustZip, sCustCountry, sCustPhone, sCustFax, sCustEmail), _
			'aryShippingAddress: Array(sShipCustFirstName, sShipCustMiddleInitial,sShipCustLastName, sShipCustCompany, sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustZip, sShipCustCountry, sShipCustPhone, sShipCustFax, sShipCustEmail), _
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address = CreateObject("com.paypal.soap.api.AddressType")
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.Street1 = aryBillingAddress(4)
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.Street2 = aryBillingAddress(5)
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.CityName = aryBillingAddress(6)
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.StateOrProvince= aryBillingAddress(7)
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.PostalCode = aryBillingAddress(8)
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.Country = util.GetEnumValue("CountryCodeType", aryBillingAddress(9))
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.CountrySpecified = true
			'.DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.CountryName = "USA"
			
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.PayerName = CreateObject("com.paypal.soap.api.PersonNameType")
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.PayerName.FirstName = aryBillingAddress(0)
			.DoDirectPaymentRequestDetails.CreditCard.CardOwner.PayerName.LastName = aryBillingAddress(2)
						
			.DoDirectPaymentRequestDetails.PaymentDetails = CreateObject("com.paypal.soap.api.PaymentDetailsType")

			.DoDirectPaymentRequestDetails.PaymentDetails.OrderTotal = CreateObject("com.paypal.soap.api.BasicAmountType")
			.DoDirectPaymentRequestDetails.PaymentDetails.OrderTotal.currencyID = util.GetEnumValue("CurrencyCodeType", "USD")
			.DoDirectPaymentRequestDetails.PaymentDetails.OrderTotal.Value = paymentAmount
			.DoDirectPaymentRequestDetails.PaymentDetails.OrderDescription = strDescription

			'Future Use
			'aryShippingAddress: Array(sShipCustFirstName, sShipCustMiddleInitial,sShipCustLastName, sShipCustCompany, sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustZip, sShipCustCountry, sShipCustPhone, sShipCustFax, sShipCustEmail), _
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress = CreateObject("com.paypal.soap.api.AddressType")
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.Name = aryShippingAddress(0) & " " & aryShippingAddress(2)
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.Street1 = aryShippingAddress(4)
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.Street2 = aryShippingAddress(5)
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.CityName = aryShippingAddress(6)
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.StateOrProvince= aryShippingAddress(7)
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.PostalCode = aryShippingAddress(8)
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.Country = util.GetEnumValue("CountryCodeType", aryShippingAddress(9))
			.DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.CountrySpecified = true

			If pblnLocalDebug Then
				Response.Write "<fieldset><legend>PayPal API - DoDirectPayment</legend>"
				Response.Write "FirstName: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.PayerName.FirstName & "<br />"
				Response.Write "LastName: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.PayerName.LastName & "<br />"
				Response.Write "CreditCardType: " & .DoDirectPaymentRequestDetails.CreditCard.CreditCardType & "<br />"
				Response.Write "creditCardNumber: " & .DoDirectPaymentRequestDetails.CreditCard.CreditCardNumber & "<br />"
				Response.Write "CVV2: " & .DoDirectPaymentRequestDetails.CreditCard.CVV2 & "<br />"
				Response.Write "ExpMonth: " & .DoDirectPaymentRequestDetails.CreditCard.ExpMonth & "<br />"
				Response.Write "ExpYear: " & .DoDirectPaymentRequestDetails.CreditCard.ExpYear & "<hr />"
				
				Response.Write "<strong>Billing Address</strong><br />"
				Response.Write "Street1: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.Street1 & "<br />"
				Response.Write "Street2: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.Street2 & "<br />"
				Response.Write "CityName: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.CityName & "<br />"
				Response.Write "StateOrProvince: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.StateOrProvince & "<br />"
				Response.Write "PostalCode: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.PostalCode & "<br />"
				Response.Write "CountryName: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.CountryName & "<br />"
				Response.Write "Country: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.Country & "<br />"
				Response.Write "CountrySpecified: " & .DoDirectPaymentRequestDetails.CreditCard.CardOwner.Address.CountrySpecified & "<hr />"

				'Response.Write "<strong>Shipping Address</strong><br />"
				'Response.Write "Street1: " & .DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.Street1 & "<br />"
				'Response.Write "Street2: " & .DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.Street2 & "<br />"
				'Response.Write "CityName: " & .DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.CityName & "<br />"
				'Response.Write "StateOrProvince: " & .DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.StateOrProvince & "<br />"
				'Response.Write "PostalCode: " & .DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.PostalCode & "<br />"
				'Response.Write "CountryName: " & .DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.CountryName & "<br />"
				'Response.Write "Country: " & .DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.Country & "<br />"
				'Response.Write "CountrySpecified: " & .DoDirectPaymentRequestDetails.PaymentDetails.ShipToAddress.CountrySpecified & "<hr />"

				Response.Write "OrderTotal: " & .DoDirectPaymentRequestDetails.PaymentDetails.OrderTotal.Value & "<br />"
				Response.Write "</fieldset>"
				Response.Flush
			End If

		End With
		
		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pobjPPcaller.CallAPI "DoDirectPayment"
	End Sub

	'***********************************************************************************************

	Public Sub SetExpressCheckout(paymentAmount, returnURL, cancelURL, paymentAction, curr)
		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")

		' Create the request object
		Set pp_request = CreateObject("com.paypal.soap.api.SetExpressCheckoutRequestType")

		With pp_request
			' Create the request details object
			.SetExpressCheckoutRequestDetails = CreateObject("com.paypal.soap.api.SetExpressCheckoutRequestDetailsType")

			.SetExpressCheckoutRequestDetails.PaymentAction = util.GetEnumValue("PaymentActionCodeType", paymentAction)
			.SetExpressCheckoutRequestDetails.PaymentActionSpecified = true
			.SetExpressCheckoutRequestDetails.OrderTotal = CreateObject("com.paypal.soap.api.BasicAmountType")
			.SetExpressCheckoutRequestDetails.OrderTotal.currencyID = util.GetEnumValue("CurrencyCodeType", curr)
			.SetExpressCheckoutRequestDetails.OrderTotal.Value = paymentAmount
			
			.SetExpressCheckoutRequestDetails.CancelURL = cancelURL
			.SetExpressCheckoutRequestDetails.ReturnURL = returnURL
		End With
		
		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pobjPPcaller.CallAPI "SetExpressCheckout"
	End Sub

	'***********************************************************************************************

	Public Sub GetExpressCheckoutDetails(token)
		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")

		' Create the request object
		Set pp_request = CreateObject("com.paypal.soap.api.GetExpressCheckoutDetailsRequestType")

		With pp_request
			.Token = token
		End With
		
		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pobjPPcaller.CallAPI "GetExpressCheckoutDetails"
	End Sub

	'***********************************************************************************************

	Public Sub DoExpressCheckoutPayment(token, payerID, paymentAmount, actionCodeType, currencyCodeType)
		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")

		' Create the request object
		Set pp_request = CreateObject("com.paypal.soap.api.DoExpressCheckoutPaymentRequestType")

		With pp_request
			' Create the request details object
			.DoExpressCheckoutPaymentRequestDetails = CreateObject("com.paypal.soap.api.DoExpressCheckoutPaymentRequestDetailsType")
			
			.DoExpressCheckoutPaymentRequestDetails.Token = token
			.DoExpressCheckoutPaymentRequestDetails.PayerID = payerID
			.DoExpressCheckoutPaymentRequestDetails.PaymentAction = util.GetEnumValue("PaymentActionCodeType", actionCodeType)
			
			.DoExpressCheckoutPaymentRequestDetails.PaymentDetails = CreateObject("com.paypal.soap.api.PaymentDetailsType")

			.DoExpressCheckoutPaymentRequestDetails.PaymentDetails.OrderTotal = CreateObject("com.paypal.soap.api.BasicAmountType")
			.DoExpressCheckoutPaymentRequestDetails.PaymentDetails.OrderTotal.currencyID = util.GetEnumValue("CurrencyCodeType", currencyCodeType)
			.DoExpressCheckoutPaymentRequestDetails.PaymentDetails.OrderTotal.Value = paymentAmount
			
		End With
		
		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pobjPPcaller.CallAPI "DoExpressCheckoutPayment"
	End Sub

	'***********************************************************************************************

	public sub DoCapture(authorizationId, note, value, currencyId, invoiceId)
		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
		set pp_request = CreateObject("com.paypal.soap.api.DoCaptureRequestType")
		pp_request.AuthorizationID = authorizationId
		pp_request.Note = note
		pp_request.Amount = CreateObject("com.paypal.soap.api.BasicAmountType")
		pp_request.Amount.Value = value
		pp_request.Amount.CurrencyID = util.GetEnumValue("CurrencyCodeType", currencyId)
		pp_request.InvoiceID = invoiceId
		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pobjPPcaller.CallAPI "DoCapture"
	end sub

	'***********************************************************************************************

	Public sub DoVoid(authorizationId, note)
		set pp_request = CreateObject("com.paypal.soap.api.DoVoidRequestType")
		pp_request.AuthorizationID = authorizationId
		pp_request.Note = note
		Call SetAPIProfile
		pobjPPcaller.Request = pp_request
		pp_caller.CallAPI "DoVoid"
	end sub

	'***********************************************************************************************

	Public Function GetCurrency(byVal currencyID)

		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
		
		Select Case currencyID
			Case util.GetEnumValue("CurrencyCodeType", "USD"):	pstrCurrency = "USD "
			Case util.GetEnumValue("CurrencyCodeType", "CAD"):	pstrCurrency = "CAD "
			Case util.GetEnumValue("CurrencyCodeType", "CNY"):	pstrCurrency = "CNY "
			Case util.GetEnumValue("CurrencyCodeType", "EUR"):	pstrCurrency = "EUD "
			Case util.GetEnumValue("CurrencyCodeType", "JPY"):	pstrCurrency = "JPY "
			Case util.GetEnumValue("CurrencyCodeType", "GBP"):	pstrCurrency = "GBP "
			Case Else:	pstrCurrency = ""
		End Select
		
		GetCurrency = pstrCurrency
		
	End Function	'GetCurrency

	'***********************************************************************************************

	Public Function GetPaymentStatus(byVal status)

		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
		
		Select Case status
			Case util.GetEnumValue("PaymentStatusCodeType", "None"):		pstrPaymentStatus = "None"
			Case util.GetEnumValue("PaymentStatusCodeType", "Completed"):	pstrPaymentStatus = "Completed"
			Case util.GetEnumValue("PaymentStatusCodeType", "Failed"):		pstrPaymentStatus = "Failed"
			Case util.GetEnumValue("PaymentStatusCodeType", "Pending"):		pstrPaymentStatus = "Pending"
			Case util.GetEnumValue("PaymentStatusCodeType", "Denied"):		pstrPaymentStatus = "Denied"
			Case util.GetEnumValue("PaymentStatusCodeType", "Refunded"):	pstrPaymentStatus = "Refunded"
			Case util.GetEnumValue("PaymentStatusCodeType", "Reversed"):	pstrPaymentStatus = "Reversed"
			Case util.GetEnumValue("PaymentStatusCodeType", "Processed"):	pstrPaymentStatus = "Processed"
			Case Else:	pstrPaymentStatus= ""
		End Select
		
		GetPaymentStatus = pstrPaymentStatus
		
	End Function	'GetPaymentStatus

	'***********************************************************************************************

	Public Function IsSuccessful(byVal ack)

		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
		
		Select Case ack
			Case util.GetEnumValue("AckCodeType", "Success"):				pblnIsSuccessful = true
			Case util.GetEnumValue("AckCodeType", "Failure"):				pblnIsSuccessful = false
			Case util.GetEnumValue("AckCodeType", "Warning"):				pblnIsSuccessful = false
			Case util.GetEnumValue("AckCodeType", "SuccessWithWarning"):	pblnIsSuccessful = true
			Case util.GetEnumValue("AckCodeType", "FailureWithWarning"):	pblnIsSuccessful = false
			Case Else:		pblnIsSuccessful = false
		End Select

		IsSuccessful = pblnIsSuccessful
		
	End Function	'IsSuccessful

	'***********************************************************************************************

	Public Function AckCode(byVal ack)
	
	Dim pstrAckCode

		If IsEmpty(util) Then Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
		
		Select Case ack
			Case util.GetEnumValue("AckCodeType", "Success"):				pstrAckCode = "Success"
			Case util.GetEnumValue("AckCodeType", "Failure"):				pstrAckCode = "Failure"
			Case util.GetEnumValue("AckCodeType", "Warning"):				pstrAckCode = "Warning"
			Case util.GetEnumValue("AckCodeType", "SuccessWithWarning"):	pstrAckCode = "SuccessWithWarning"
			Case util.GetEnumValue("AckCodeType", "FailureWithWarning"):	pstrAckCode = "FailureWithWarning"
			Case Else:		pblnIsSuccessful = false
		End Select

		AckCode = pstrAckCode
		
	End Function	'AckCode

	'***********************************************************************************************

	Public Function ErrorMessages(errors)

	Dim error
	Dim iterator
	Dim pstrErrorMessage
	
		Set iterator = CreateObject("com.paypal.sdk.COMNetInterop.COMIterator")
		iterator.csArray = errors
		Do While iterator.HasNext()
			Set error = iterator.Next()
			With error
				If Len(pstrErrorMessage) = 0 Then
					pstrErrorMessage = pstrErrorMessage & vbcrlf & .ErrorCode & " - " & .LongMessage
				Else
					pstrErrorMessage = pstrErrorMessage & vbcrlf & .ErrorCode & " - " & .LongMessage
				End If
 			End With
 		Loop
		Set error = Nothing
		Set iterator = Nothing
		
		ErrorMessages = pstrErrorMessage
	 	
	End Function

	'***********************************************************************************************

	Private Function GetGMTDate(d)
		set oShell = CreateObject("WScript.Shell")
		atb = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
		offsetMin = oShell.RegRead(atb)
		GetGMTDate = DateAdd("n", offsetMin, d)
	End Function

	Public Property Let Token(byVal strToken)
		pstrToken = strToken
	End Property
	
	Public Property Get Token
		Token = pstrToken
	End Property
	
	Public Property Get ExpressCheckoutURL()
		ExpressCheckoutURL = "https://www."  & pstrEnvironment & ".paypal.com/cgi-bin/webscr?cmd=_express-checkout&token=" & pstrToken
	End Property
	
End Class

'***********************************************************************************************
'***********************************************************************************************

Sub GetPayPalLogin(byRef login, byRef password)

Dim pstrSQL
Dim pobjRS

	login = adminLogin
	password = adminPassword
	If adminMerchantType = "authcapture" Then
		ActionCodeType = "Sale"
	Else
		ActionCodeType = "Authorization"
	End If
	
	If Len(login) = 0 Or Len(password) = 0 Then
		Response.Write "<fieldset><legend>PayPal WebSite Payments Pro Error</legend>"
		Response.Write "<font color=red>Either your username or login is empty</font>"
		Response.Write "</fieldset>"
		Response.Flush
	End If
	Exit Sub

	'Below is only necessary for standard installation
	pstrSQL = "Select trnsmthdID, trnsmthdLogin, trnsmthdPasswd From sfTransactionMethods Where trnsmthdName='PayPal WebPayments'"
	Set pobjRS = CreateObject("ADODB.RecordSet")
	With pobjRS
		.Open pstrSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
		If Not .EOF Then
			login = Trim(.Fields("trnsmthdLogin").Value & "")
			password = Trim(.Fields("trnsmthdPasswd").Value & "")
			PayPalTransactionMethodID = Trim(.Fields("trnsmthdID").Value & "")
		End If
		.Close
		
		pstrSQL = "Select adminMerchantType From sfAdmin"
		.Open pstrSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
		If Not .EOF Then
			If Trim(.Fields("adminMerchantType").Value & "") = "authcapture" Then
				ActionCodeType = "Sale"
			Else
				ActionCodeType = "Authorization"
			End If
		End If
		.Close
		
	End With
	Set pobjRS = Nothing
	
End Sub	'GetPayPalLogin

'***********************************************************************************************

Sub PayPalExpressCheckout_Initiate(byVal dblSubtotal)

Dim CancelURL
Dim paryCredential
Dim pblnLocalDebug
Dim pstrCustom
Dim pstrLogin
Dim pstrPassword
Dim pstrSQL
Dim pstrURLOut
Dim pobjPayPalWebPayment
Dim ReturnURL

	pblnLocalDebug = False	'True	False
	pstrURLOut = ""

	If Len(Request.Form("btn_xpressCheckout.x")) > 0 Then
		Call GetPayPalLogin(pstrLogin, pstrPassword)
		paryCredential = Split(pstrPassword, "|")
		Set pobjPayPalWebPayment = New PayPalAPI
		With pobjPayPalWebPayment
			.APIUsername = pstrLogin
			If UBound(paryCredential) >= 0 Then .APIPassword = paryCredential(0)
			If UBound(paryCredential) >= 1 Then .APISignature = paryCredential(1)
			.Subject = ""
			.Environment = PayPalEnvironment
			.DebugEnabled = pblnLocalDebug
			
			pstrCustom = "xpressCheckout=PayPal"
			ReturnURL = adminSSLPath & "?" & pstrCustom
			CancelURL = adminDomainName & "order.asp"

			If pblnLocalDebug Then
				Response.Write "<fieldset><legend>PayPal API Credentials</legend>"
				Response.Write "PayPalEnvironment: " & PayPalEnvironment & "<hr />"
				Response.Write "APIUsername: " & pstrLogin & "<br />"
				If UBound(paryCredential) >= 0 Then Response.Write "APIPassword: " & paryCredential(0) & "<br />"
				If UBound(paryCredential) >= 1 Then Response.Write "APISignature: " & paryCredential(1) & "<br />"
				Response.Write "ReturnURL: " & ReturnURL & "<br />"
				Response.Write "CancelURL: " & CancelURL & "<br />"
				Response.Write "</fieldset>"
				Response.Flush
			End If
			
			.SetExpressCheckout dblSubtotal, ReturnURL, CancelURL, ActionCodeType, "USD"
			
			If .IsSuccessful(.pp_caller.Response.Ack) Then
				.Token = .pp_caller.Response.Token
				pstrURLOut = .ExpressCheckoutURL

				pstrSQL = "Update sfTmpOrderDetails Set PayPalToken='" & .Token & "' Where odrdttmpSessionID=" & Session("SessionID")
				cnn.Execute pstrSQL,,128
				
				If pblnLocalDebug Then
					Response.Write "<fieldset><legend>SetExpressCheckout Result</legend>" _
								& "Token: " & .Token & "<br>" _
								& "ExpressCheckoutURL: " & .ExpressCheckoutURL & "<br>" _
								& "ReturnURL: " & ReturnURL & "<br>" _
								& "CancelURL: " & CancelURL & "<br>" _
								& "dblSubtotal: " & dblSubtotal & "<br>" _
								& "ErrorMessage: " & .ErrorMessages(.pp_caller.Response.Errors) & "<br>" _
								& "</fieldset>"
								
					Response.Write "<a href=""" & ReturnURL & """>Process Order (" & ReturnURL & ")</a>"
				Else
					Call CleanupPageObjects
					Response.Redirect pstrURLOut
				End If

			Else
				Response.Write .ErrorMessages(.pp_caller.Response.Errors)
			End If
		End With
		Set pobjPayPalWebPayment = Nothing
	End If	'Len(Request.Form("btn_xpressCheckout.x")) > 0

	
End Sub	'PayPalExpressCheckout_Initiate

'***********************************************************************************************

Sub PayPalExpressCheckout_process_order()

Dim paryCredential
Dim pblnLocalDebug
Dim pblnResult
Dim pobjPayPalWebPayment
Dim pstrLogin
Dim pstrPassword

	pblnLocalDebug = False	'True	False

	Call PayPalExpressCheckout_Initiate(mclsCartTotal.SubTotal)

	If LoadRequestValue("xpressCheckout") = "PayPal" Then	

		'Now catch the express checkout
		PayPalToken = LoadRequestValue("token")
		
		If Len(PayPalToken) > 0 Then
			Call GetPayPalLogin(pstrLogin, pstrPassword)
			paryCredential = Split(pstrPassword, "|")
			Set pobjPayPalWebPayment = New PayPalAPI
			With pobjPayPalWebPayment
				.APIUsername = pstrLogin
				If UBound(paryCredential) >= 0 Then .APIPassword = paryCredential(0)
				If UBound(paryCredential) >= 1 Then .APISignature = paryCredential(1)
				.Subject = ""
				.Environment = PayPalEnvironment
				.DebugEnabled = pblnLocalDebug
				
				.GetExpressCheckoutDetails PayPalToken 
				If .IsSuccessful(.pp_caller.Response.Ack) Then
					pblnResult = True
					With .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo
						PayPalPayerID = .PayerID
						
						sCustFirstName = .PayerName.FirstName
						sCustLastName = .PayerName.LastName
						sCustAddress1 = .Address.Street1
						sCustAddress2 = .Address.Street2
						sCustCity = .Address.CityName
						sCustState = .Address.StateOrProvince
						sCustZip = .Address.PostalCode
						sCustCountry = .Address.Country

						sShipCustFirstName		= sCustFirstName
						sShipCustLastName		= sCustLastName
						sShipCustAddress1		= sCustAddress1
						sShipCustAddress2		= sCustAddress2
						sShipCustCity			= sCustCity
						sShipCustState			= sCustState
						sShipCustZip			= sCustZip
						sShipCustCountry		= sCustCountry
						
						'sShipCustStateName		= sCustStateName
						'sShipCustCountryName	= sCustCountryName

					End With
					PayPalExpressCheckoutEnabled = True

					If pblnLocalDebug Then
						Response.Write "<fieldset><legend>GetExpressCheckoutDetails Result</legend>" _
									& "Token: " & PayPalToken & "<hr />" _
									& "PayerID: " & PayPalPayerID & "<br>" _
									& "PayerStatus: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.PayerStatus & "<br>" _
									& "FirstName: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.PayerName.FirstName & "<br>" _
									& "LastName: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.PayerName.LastName & "<br>" _
									& "Street1: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.Address.Street1 & "<br>" _
									& "Street2: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.Address.Street2 & "<br>" _
									& "CityName: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.Address.CityName & "<br>" _
									& "StateOrProvince: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.Address.StateOrProvince & "<br>" _
									& "PostalCode: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.Address.PostalCode & "<br>" _
									& "Country: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.Address.Country & "<br>" _
									& "CountryName: " & .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.Address.CountryName & "<br>" _
									& "ErrorMessage: " & .ErrorMessages(.pp_caller.Response.Errors) & "<br>" _
									& "</fieldset>"
					End If
				End If
			End With
			Set pobjPayPalWebPayment = Nothing

		End If	'Len(PayPalToken) > 0

	End If	'Request.QueryString("xpressCheckout") = "PayPal"

End Sub	'PayPalExpressCheckout_process_order

'***********************************************************************************************

Sub PayPalExpressCheckout_verify

	PayPalToken = Request.Form("PayPalToken")
	If Len(PayPalToken) > 0 Then
		PayPalPayerID = Request.Form("PayPalPayerID")
		PayPalExpressCheckoutEnabled = True
	End If

End Sub	'PayPalExpressCheckout_verify

'***********************************************************************************************

Function PayPalExpressCheckout_confirm

Dim iProcResponse
Dim paryCredential
Dim pblnLocalDebug
Dim pstrError
Dim pstrLogin
Dim pstrPassword
Dim pobjPayPalWebPayment
Dim pstrReferenceNumber
Dim pstrRespCode

	pblnLocalDebug = False

	PayPalToken = Request.Form("PayPalToken")
	If Len(PayPalToken) > 0 Then
		PayPalPayerID = Request.Form("PayPalPayerID")
		PayPalExpressCheckoutEnabled = True
		
		Call GetPayPalLogin(pstrLogin, pstrPassword)
		paryCredential = Split(pstrPassword, "|")
		Set pobjPayPalWebPayment = New PayPalAPI
		With pobjPayPalWebPayment
			.APIUsername = pstrLogin
			If UBound(paryCredential) >= 0 Then .APIPassword = paryCredential(0)
			If UBound(paryCredential) >= 1 Then .APISignature = paryCredential(1)
			.Subject = ""
			.Environment = PayPalEnvironment
			.DebugEnabled = pblnLocalDebug

			.DoExpressCheckoutPayment PayPalToken, PayPalPayerID, sGrandTotal, ActionCodeType, "USD"
			If .IsSuccessful(.pp_caller.Response.Ack) Then
				iProcResponse = 1
				pstrReferenceNumber = .pp_caller.Response.DoExpressCheckoutPaymentResponseDetails.PaymentInfo.TransactionID
				pstrRespCode = .GetPaymentStatus(.pp_caller.Response.DoExpressCheckoutPaymentResponseDetails.PaymentInfo.PaymentStatus)

				If pblnLocalDebug Then
					Response.Write "<fieldset><legend>GetExpressCheckoutDetails Result</legend>" _
								& "Token: " & PayPalToken & "<hr />" _
								& "TransactionID: " & pstrReferenceNumber & "<br>" _
								& "GrossAmount: " & .pp_caller.Response.DoExpressCheckoutPaymentResponseDetails.PaymentInfo.GrossAmount.Value & "<br>" _
								& "PaymentStatus: " & pstrRespCode & "(" & .pp_caller.Response.DoExpressCheckoutPaymentResponseDetails.PaymentInfo.PaymentStatus & ")<br>" _
								& "PendingReason: " & .pp_caller.Response.DoExpressCheckoutPaymentResponseDetails.PaymentInfo.PendingReason & "<br>" _
								& "ErrorMessage: " & .ErrorMessages(.pp_caller.Response.Errors) & "<br>" _
								& "</fieldset>"
				End If
			Else
				pstrError = .ErrorMessages(.pp_caller.Response.Errors)
				iProcResponse = 0
				If pblnLocalDebug Then
					Response.Write "<fieldset><legend>GetExpressCheckoutDetails Result</legend>" _
								& "Token: " & PayPalToken & "<hr />" _
								& "sGrandTotal: " & sGrandTotal & "<hr />" _
								& "ErrorMessage: " & pstrError & "<br>" _
								& "</fieldset>"
				End If
			End If
		End With
		Set pobjPayPalWebPayment = Nothing

		'Call setResponse("PayPal Express", iOrderID, pstrReferenceNumber, "", Array("", sGrandTotal, ""), "", pstrRespCode, "", "", pstrError, iProcResponse)
		Call setResponse("PayPal Express", iOrderID, pstrReferenceNumber, "", "", "", pstrRespCode, "", "", pstrError , iProcResponse)
	End If
	
	PayPalExpressCheckout_confirm = pstrError

End Function	'PayPalExpressCheckout_confirm

'***********************************************************************************************

Function PayPalWebPayments(proc_live)

Dim iProcResponse
Dim pblnLocalDebug
Dim ProcActionCode
Dim ProcAuthCode
Dim ProcAvsCode
Dim pstrRespCode
Dim pstrErrorMessage
Dim pstrCCVCode
Dim pstrLogin
Dim pstrPassword
Dim pstrReferenceNumber
Dim pobjPayPalWebPayment
Dim paryCredential

	'custom sections:
	'Session("debugCC") = "True"
	'Session("debugCC") = ""
	pblnLocalDebug = Len(Session("debugCC")) > 0
	'pblnLocalDebug = True
	
	Call GetPayPalLogin(pstrLogin, pstrPassword)
	paryCredential = Split(pstrPassword, "|")
	Set pobjPayPalWebPayment = New PayPalAPI
	With pobjPayPalWebPayment
		.APIUsername = pstrLogin
		If UBound(paryCredential) >= 0 Then .APIPassword = paryCredential(0)
		If UBound(paryCredential) >= 1 Then .APISignature = paryCredential(1)
		.Subject = ""
		.Environment = PayPalEnvironment
		.DebugEnabled = pblnLocalDebug

		If pblnLocalDebug Then
			Response.Write "<fieldset><legend>PayPal API Credentials</legend>"
			Response.Write "PayPalEnvironment: " & PayPalEnvironment & "<br />"
			Response.Write "Login: " & pstrLogin & "<br />"
			If UBound(paryCredential) >= 0 Then Response.Write "APIPassword: " & paryCredential(0) & "<br />"
			If UBound(paryCredential) >= 1 Then Response.Write "APISignature: " & paryCredential(1) & "<br />"
			Response.Write "</fieldset>"
			Response.Flush
		End If
		
		.DoDirectPayment _
			sGrandTotal, _
			Array(sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2, sCustCity, sCustState, sCustZip, sCustCountry, sCustPhone, sCustFax, sCustEmail), _
			Array(sShipCustFirstName, sShipCustMiddleInitial,sShipCustLastName, sShipCustCompany, sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustZip, sShipCustCountry, sShipCustPhone, sShipCustFax, sShipCustEmail), _
			sCustCardTypeName, _
			sCustCardNumber, _
			mstrPayCardCCV, _
			sCustCardExpiryMonth, _
			sCustCardExpiryYear, _
			ActionCodeType, _
			PayPalOrderDescription & iOrderID
			
		If .IsSuccessful(.pp_caller.Response.Ack) Then
			pstrReferenceNumber = .pp_caller.Response.TransactionID
			ProcAvsCode = .pp_caller.Response.AVSCode
			pstrCCVCode = .pp_caller.Response.CVV2Code
			pstrRespCode = .AckCode(.pp_caller.Response.Ack)
			'.pp_caller.Response.Amount.currencyID)
			'.pp_caller.Response.Amount.Value
			iProcResponse = 1
		Else
			pstrErrorMessage = .ErrorMessages(.pp_caller.Response.Errors)
			iProcResponse = 0
		End If
    End With	'pobjPayPalWebPayment
    Set pobjPayPalWebPayment = Nothing
    
	If pblnLocalDebug Then
		Response.Write "<fieldset><legend>PayPalWebPayments Result</legend>" _
					& "Approved: " & CBool(iProcResponse) & "<br>" _
					& "Approved: " & CBool(iProcResponse) & "<br>" _
					& "AVS: " & PayPalAvsCodeDefinition(ProcAvsCode) & " (" & ProcAvsCode & ")" & "<br>" _
					& "RespCode: " & pstrRespCode & "<br>" _
					& "CCV: " & PayPalCCVCodeDefinition(pstrCCVCode) & " (" & pstrCCVCode & ")" & "<br>" _
					& "OrderID: " & iOrderID & "<br>" _
					& "Reference Number: " & pstrReferenceNumber & "<br>" _
					& "Amount: " & sGrandTotal & "<br>" _
					& "ErrorMessage: " & pstrErrorMessage & "<br>" _
					& "</fieldset>"
	End If

	Call setResponse("PayPalWebPayment", iOrderID, pstrReferenceNumber, "", ProcAvsCode, PayPalAvsCodeDefinition(ProcAvsCode), pstrRespCode, "", "",pstrErrorMessage, iProcResponse)	
    'Call setResponse("PayPalWebPayment", iOrderID, pstrReferenceNumber, "", Array(PayPalCCVCodeDefinition(pstrCCVCode), sGrandTotal, pstrCCVCode), PayPalAvsCodeDefinition(ProcAvsCode), pstrRespCode, "", "", pstrErrorMessage, iProcResponse)
    
	PayPalWebPayments = pstrErrorMessage
	
End Function	'PayPalWebPayments

'************************************************************************************************************

Function PayPalAvsCodeDefinition(byVal strProcAvsCode)

Dim ProcAvsCode

	Select Case strProcAvsCode
		Case "A": ProcAvsCode = "Address only (no ZIP)"
		Case "B": ProcAvsCode = "Address only (no ZIP)"
		Case "C": ProcAvsCode = "None"
		Case "D": ProcAvsCode = "Address and Postal Code"
		Case "E": ProcAvsCode = "Not applicable"
		Case "F": ProcAvsCode = "Address and Postal Code"
		Case "G": ProcAvsCode = "Not applicable"
		Case "I": ProcAvsCode = "Not applicable"
		Case "N": ProcAvsCode = "None"
		Case "P": ProcAvsCode = "Postal Code only (no Address)"
		Case "R": ProcAvsCode = "Not applicable"
		Case "S": ProcAvsCode = "Not applicable"
		Case "U": ProcAvsCode = "Not applicable"
		Case "W": ProcAvsCode = "Nine-digit ZIP code (no Address)"
		Case "X": ProcAvsCode = "Address and nine-digit ZIP code"
		Case "Y": ProcAvsCode = "Address and five-digit ZIP"
		Case "Z": ProcAvsCode = "Five-digit ZIP code (no Address)"
		Case Else: ProcAvsCode = "Unknown Code"
	End Select
			
	PayPalAvsCodeDefinition = ProcAvsCode
	
End Function	'PayPalAvsCodeDefinition

'************************************************************************************************************

Function PayPalCCVCodeDefinition(byVal strCCVCode)

Dim ProcCCVCode

	Select Case strCCVCode
		Case "M": ProcCCVCode = "Match"
		Case "N": ProcCCVCode = "No Match"
		Case "P": ProcCCVCode = "Not Processed"
		Case "S": ProcCCVCode = "Service not Supported"
		Case "U": ProcCCVCode = "Unavailable"
		Case "X": ProcCCVCode = "No response"
		Case Else: ProcCCVCode = "Unknown Code"
	End Select
	
	PayPalCCVCodeDefinition = ProcCCVCode
	
End Function	'PayPalCCVCodeDefinition
%>