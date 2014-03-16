<%
'********************************************************************************
'*   PayPal Payments															*
'*   Release Version:   3.00.003												*
'*   Release Date:		August 10, 2003											*
'*   Revision Date:		September 24, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Release 3.00.004 (December 20, 2004)										*
'*	 - Enhancement - added ability to process Credit Cards via PayPal as well	*
'*	   as PayPal as a separate payment method									*
'*                                                                              *
'*   Release 3.00.003 (September 24, 2004)										*
'*	 - Bug Fix - Emails weren't sent for non-PayPal orders if not set to send Confirmation email
'*                                                                              *
'*   Release 3.00.002 (April 13, 2004)											*
'*	 - Enhancement - Streamlined installation									*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************


'////////////////////////////////////////////////////////////////////////////////
'//
'//		USER CONFIGURATION - StoreFront Specific

	'AE Settings
	Const cblnSendConfirmationEmail	= False	'Applies to AE only 

'//
'////////////////////////////////////////////////////////////////////////////////

'****************************************************************************************************************

Class clsPayPal

Private pobjConnection

Private pstrOrderID
Private pstrSessionID
Private pblnFoundOrder

'Information from PayPal pages
Private psngVeriedPaymentAmount
Private pblnCompleted

'Information from PayPal Form Collection
Private pstrPaymentType
Private pstrTransactionID
Private pdtTransactionDate
Private psngGrossAmount
Private psngFee
Private pstrFirstName
Private pstrLastName
Private pstrEmail
Private pstrAddress1
Private pstrAddress2
Private pstrCity
Private pstrState
Private pstrZIP
Private pstrCountry


'User Defined Constants
Private cstrSiteURL
Private cstrReturnPage
Private cstrPaymentLink
Private cstrImageURL
Private cstrOrderRef
Private cstrPayPalLogin
Private cstrPayPalPassword
Private cblnGetCustomerShipAddress
Private cbytno_shipping

Private pdblshipping
Private pdblshipping2
Private pstrnight_phone_a
Private pstrnight_phone_b
Private pstrnight_phone_c
Private pstrday_phone_a
Private pstrday_phone_b
Private pstrday_phone_c

Private pdblAmountDue
Private pstrCustom
Private cbytCurrencyType

'****************************************************************************************************************

Private Sub Class_Initialize

'////////////////////////////////////////////////////////////////////////////////
'//
'//		USER CONFIGURATION - PayPal Specific

	cstrPaymentLink = "here"			'what you want the customer to see for the payment link. This can be text or an image
	cstrImageURL = "http://www.YourWebSite.Com/ssl"	'path to your site logo; this should be located at a https address
	cstrReturnPage = ""					'the page you want the customer to return to. confirm.asp is the default if left empty
	'cstrReturnPage = "http://www.yoursite.com/OrderHistory.asp?OrderID={OrderID}&email={email}"
	cstrOrderRef = "Your Order"			'what you want the customer to see on the PayPal site
	cstrPayPalLogin = "Your Login"
	cbytno_shipping = 0				'do you want the customer to enter a shipping address 0-No, 1-Yes
	cbytCurrencyType = 0			'0 - U.S. Dollars ($) USD
									'1 - Canadian Dollars (C $) CAD
									'2 - Euros (€) EUR
									'3 - Pounds Sterling (£) GBP
									'4 -Yen (¥) JPY
 
	cstrReturnPage = "http://www.yoursite.com/OrderHistory.asp?OrderID={OrderID}&email={email}"
'//
'////////////////////////////////////////////////////////////////////////////////

End Sub

Private Sub Class_Terminate

End Sub

Public Function URL(lngOrderID,sngTotalDue)

Dim pstrURL

	pstrURL = "https://www.paypal.com/cgi-bin/webscr/?cmd=_xclick" _
			& "&business=" & Replace(cstrPayPalLogin,"@","%40") _
			& "&item_name=" & Replace(cstrOrderRef," ","+") _
			& "&item_number=" & pstrSessionID _
			& "&invoice=" & lngOrderID _
			& "&undefined_quantity=0" _
			& "&amount=" & formatnumber(sngTotalDue,2) _
			& "&currency_code=" & currencyCodeToString(pbytCurrencyType) _
			& "&no_shipping=" & cbytno_shipping _
			& "&return=" & cstrSiteURL & "%3FAction%3DSuccess%26OrderID%3D" & lngOrderID _
			& "&cancel_return=" & cstrSiteURL & "%3FAction%3DCancel%26OrderID%3D" & lngOrderID

	If len(cstrImageURL) > 0 Then pstrURL = pstrURL & "&image_url=" & cstrImageURL
	
	'Additional variables PayPal can accept but that aren't used for this application	
	'	& "&shipping=" & formatnumber(sngshipping,2) _
	'	& "&shipping2=" & formatnumber(sngshipping2,2) _
	'	& "&handling=" & formatnumber(snghandling,2) _
	'	& "&custom=" & "" _
	
	URL = pstrURL
	
End Function

Public Function RedirectToPayPal(lngOrderID,sngTotalDue)

	With Response
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>"
		.Write "<META HTTP-EQUIV='Refresh' CONTENT='5;URL=" & pstrURL & "'>"
		.Write "</head>"
		.Write "<body>"
		.Write "<h4>You will be redirected to the secure PayPal server to enter your payment.</h4>"
		.Write "<h4>If you are not automatically redirected follow this <a href='" & URL(lngOrderID,sngTotalDue) & "'>link</a>.</h4>"
		.Write "</body>"
		.Write "</html>"
	End With

End Function

Public Function RedirectToPayPalHeader(lngOrderID,sngTotalDue)
	RedirectToPayPalHeader = "<META HTTP-EQUIV='Refresh' CONTENT='5;URL=" & URL(lngOrderID,sngTotalDue) & "'>"
End Function

Public Function RedirectMessage(lngOrderID,sngTotalDue)
	RedirectMessage = "You will be redirected to the secure PayPal server to enter your payment.</h4>" _
				& "If you are not automatically redirected follow this <a href='" & URL(lngOrderID,sngTotalDue) & "'>link</a>."
End Function

'****************************************************************************************************************

Public Property Let Connection(objConnection)
	Set pobjConnection = objConnection
End Property

Public Property Get PaymentType
	PaymentType = pstrPaymentType
End Property

Public Property Get TransactionID
	TransactionID = pstrTransactionID
End Property

Public Property Get GrossAmount
	GrossAmount = psngGrossAmount
End Property

Public Property Get Fee
	Fee = psngFee
End Property

Public Property Get TransactionDate
	TransactionDate = pdtTransactionDate
End Property

Public Property Let LastName(strLastName)
	pstrLastName = strLastName
End Property
Public Property Get LastName
	LastName = pstrLastName
End Property

Public Property Let FirstName(strFirstName)
	pstrFirstName = strFirstName
End Property
Public Property Get FirstName
	FirstName = pstrFirstName
End Property

Public Property Let Email(strEmail)
	pstrEmail = strEmail
End Property
Public Property Get Email
	Email = pstrEmail
End Property

Public Property Let Address1(strAddress1)
	pstrAddress1 = strAddress1
End Property
Public Property Get Address1
	Address1 = pstrAddress1
End Property

Public Property Let Address2(strAddress2)
	pstrAddress2 = strAddress2
End Property
Public Property Get Address2
	Address2 = pstrAddress2
End Property

Public Property Let City(strCity)
	pstrCity = strCity
End Property
Public Property Get City
	City = pstrCity
End Property

Public Property Let State(strState)
	pstrState = strState
End Property
Public Property Get State
	State = pstrState
End Property

Public Property Let ZIP(strZIP)
	pstrZIP = strZIP
End Property
Public Property Get ZIP
	ZIP = pstrZIP
End Property

Public Property Get Country
	Country = pstrCountry
End Property

Public Property Get FoundOrder
	FoundOrder = pblnFoundOrder
End Property

Public Property Get Completed
	Completed = pblnCompleted
End Property

Public Property Let Custom(strCustom)
	pstrCustom = strCustom
End Property
Public Property Let ssPayPalReturnURL(strssPayPalReturnURL)

Dim pstrTemp

	If Len(cstrReturnPage) = 0 Then
		pstrTemp = strssPayPalReturnURL
	Else
		pstrTemp = cstrReturnPage
		pstrTemp = Replace(pstrTemp, "{OrderID}", pstrOrderID)
		pstrTemp = Replace(pstrTemp, "{email}", pstrEmail)
	End If
	
	cstrSiteURL = pstrTemp
	
End Property

Public Property Get VeriedPaymentAmount
	VeriedPaymentAmount = psngVeriedPaymentAmount
End Property

Public Property Let shipping(dblshipping)
	pdblshipping = dblshipping
End Property
Public Property Let shipping2(dblshipping2)
	pdblshipping2 = dblshipping2
End Property
Public Property Let night_phone_a(strnight_phone_a)
	pstrnight_phone_a = strnight_phone_a
End Property
Public Property Let night_phone_b(strnight_phone_b)
	pstrnight_phone_b = strnight_phone_b
End Property
Public Property Let night_phone_c(strnight_phone_c)
	pstrnight_phone_c = strnight_phone_c
End Property
Public Property Let day_phone_a(strday_phone_a)
	pstrday_phone_a = strday_phone_a
End Property
Public Property Let day_phone_b(strday_phone_b)
	pstrday_phone_b = strday_phone_b
End Property
Public Property Let day_phone_c(strday_phone_c)
	pstrday_phone_c = strday_phone_c
End Property
Public Property Let OrderID(strOrderID)
	pstrOrderID = strOrderID
End Property
Public Property Let SessionID(strSessionID)
	pstrSessionID = strSessionID
End Property
Public Property Let AmountDue(dblAmountDue)
	pdblAmountDue = dblAmountDue
End Property

Public Property Get PaymentLinkText
	PaymentLinkText = cstrPaymentLink
End Property

Public Sub Phone(strPhone, blnDay)

Dim p_strTemp
Dim p_lngLen
Dim p_strAreaCode, p_strPrefix, p_strSuffix

	'clean out the extra characters
	p_strTemp = Replace(strPhone," ","")
	p_strTemp = Replace(p_strTemp,"(","")
	p_strTemp = Replace(p_strTemp,")","")
	p_strTemp = Replace(p_strTemp,"-","")
	p_lngLen = Len(p_strTemp)
	
	If p_lngLen = 10 Then
		p_strAreaCode = Left(p_strTemp,3)
		p_strPrefix = Mid(p_strTemp,4,3)
		p_strSuffix = Right(p_strTemp,4)
		
		If blnDay Then
			pstrday_phone_a = p_strAreaCode
			pstrday_phone_b = p_strPrefix
			pstrday_phone_c = p_strSuffix
		Else
			pstrnight_phone_a = p_strAreaCode
			pstrnight_phone_b = p_strPrefix
			pstrnight_phone_c = p_strSuffix
		End If
	End If

End Sub

'****************************************************************************************************************

Private Function currencyCodeToString(bytCurrencyType)

	Select Case bytCurrencyType
		Case 0:		currencyCodeToString = "USD"	'U.S. Dollars ($) 
		Case 1:		currencyCodeToString = "CAD"	'Canadian Dollars (C $) 
		Case 2:		currencyCodeToString = "EUR"	'Euros (€) 
		Case 3:		currencyCodeToString = "GBP"	'Pounds Sterling (£) 
		Case 4:		currencyCodeToString = "JPY"	'Yen (¥) 
		Case Else:	currencyCodeToString = "USD"	'U.S. Dollars ($) 
	End Select

End Function	'currencyCodeToString

'****************************************************************************************************************

Public Function Save()

Dim sql
Dim sqlValues
Dim pobjCommand

On Error Resume Next

	set pobjCommand = CreateObject("ADODB.Command")
	with pobjCommand
					
		sql = "Insert Into PayPalTransactions (PayPalTransactionID,OrderID"
		sqlValues = ") Values (?,?"
		.CommandType = 1	'adCmdText
		.Parameters.Append .CreateParameter("PayPalTransactionID",200,1,len(pstrTransactionID),pstrTransactionID)
		.Parameters.Append .CreateParameter("OrderID",20,1,,pstrOrderID)
					
		if len(pstrPaymentType) > 0 then
			sql = sql & ",PaymentType"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("PaymentType",200,1,len(pstrPaymentType),pstrPaymentType)
		end if

		if len(pstrEmail) > 0 then
			sql = sql & ",Email"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("Email",200,1,len(pstrEmail),pstrEmail)
		end if

		if len(pstrFirstName) > 0 then
			sql = sql & ",FirstName"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("FirstName",200,1,len(pstrFirstName),pstrFirstName)
		end if

		if len(pstrLastName) > 0 then
			sql = sql & ",LastName"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("LastName",200,1,len(pstrLastName),pstrLastName)
		end if

		if len(pstrAddress1) > 0 then
			sql = sql & ",Address1"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("Address1",200,1,len(pstrAddress1),pstrAddress1)
		end if

		if len(pstrAddress2) > 0 then
			sql = sql & ",Address2"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("Address2",200,1,len(pstrAddress2),pstrAddress2)
		end if

		if len(pstrCity) > 0 then
			sql = sql & ",City"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("City",200,1,len(pstrCity),pstrCity)
		end if

		if len(pstrState) > 0 then
			sql = sql & ",State"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("State",200,1,len(pstrState),pstrState)
		end if

		if len(pstrZIP) > 0 then
			sql = sql & ",ZIP"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("ZIP",200,1,len(pstrZIP),pstrZIP)
		end if

		if len(pstrCountry) > 0 then
			sql = sql & ",Country"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("Country",200,1,len(pstrCountry),pstrCountry)
		end if

		if len(psngGrossAmount) > 0 then
			sql = sql & ",GrossAmount"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("GrossAmount",4,1,,psngGrossAmount)
		end if

		if len(psngFee) > 0 then
			sql = sql & ",Fee"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("Fee",4,1,,psngFee)
		end if

		if len(pdtTransactionDate) > 0 then
			sql = sql & ",TransactionDate"
			sqlValues = sqlValues & ",?"
			.Parameters.Append .CreateParameter("TransactionDate",7,1,,pdtTransactionDate)
		end if

'debugprint "sql",sql
		.CommandText = sql

		.ActiveConnection = pobjConnection
		.Execute
	end with
	set pobjCommand = Nothing
	
	If Err.number = "-2147467259" Then
		If Instr(1,Err.Description,"not successful because they would create duplicate values") > 0 Then
			Save = True
		Else
			Save = False
		End If
	Else
'debugprint "Err",Err.number & " - " & Err.Description
		Save = (Err.number = 0)
	End If
	If Err.number <> 0 Then Err.Clear

End Function	'Save

'****************************************************************************************************************

Private Sub LoadPayPalRequestForm

	pstrPaymentType = Request.Form("txn_type")				'web_accept
	pstrTransactionID = Request.Form("txn_id")				'Transaction ID - useful to reference item
	pdtTransactionDate = Request.Form("payment_date")

'Items available but not used in this application
'	pstrpayment_status = Request.Form("payment_status")		'either Completed or Pending
'	pstritem_number = Request.Form("item_number")
'	pstrnotify_version = Request.Form("notify_version")
'	pstrverify_sign = Request.Form("verify_sign")
'	pstritem_name = Request.Form("item_name")
'	pstrcustom = Request.Form("custom")
'	pstrreceiver_email = Request.Form("receiver_email")
'	pintquantity = Request.Form("quantity")
	
	psngGrossAmount = Request.Form("payment_gross")
	psngFee = Request.Form("payment_fee")					'only present if verified user

	pstrFirstName = Request.Form("first_name")
	pstrLastName = Request.Form("last_name")
	pstrEmail = Request.Form("payer_email")
	pstrAddress1 = Request.Form("address_street")			'only present if verified user
	pstrAddress2 = Request.Form("address_street2")			'only present if verified user
	pstrCity = Request.Form("address_city")					'only present if verified user
	pstrState = Request.Form("address_state")				'only present if verified user
	pstrZIP = Request.Form("address_zip")					'only present if verified user
	pstrCountry = Request.Form("address_country")			'only present if verified user

End Sub

'****************************************************************************************************************

	'***********************************************************************************************

	Function HTMLHiddenField(strFieldName, strFieldValue)

	HTMLHiddenField = "<input type=hidden id=" & chr(34) & Server.HTMLEncode(strFieldName) & chr(34) _
					& " name=" & chr(34) & Server.HTMLEncode(strFieldName) & chr(34) _
					& " value=" & chr(34) & Server.HTMLEncode(strFieldValue) & chr(34) & ">"
			 
	End Function	'HTMLHiddenField

	'***********************************************************************************************

	Public Function PaymentForm()

	If cblnDebugPayPalAddon Then pdblAmountDue = .01

	PaymentForm = "<form action='https://www.paypal.com/cgi-bin/webscr' method='post' id='frmSSPayment' name='frmSSPayment'>" & vbcrlf _
				& HTMLHiddenField("cmd", "_ext-enter") & vbcrlf _
				& HTMLHiddenField("redirect_cmd", "_xclick") & vbcrlf _
				& HTMLHiddenField("business", cstrPayPalLogin) & vbcrlf _
				& HTMLHiddenField("image_url", cstrImageURL) & vbcrlf _
				& HTMLHiddenField("return", cstrSiteURL) & vbcrlf _
				& HTMLHiddenField("cancel_return", cstrSiteURL & "?Action=Cancel&OrderID=" & pstrOrderID) & vbcrlf _
				& HTMLHiddenField("item_name", cstrOrderRef & " " & pstrOrderID) & vbcrlf _
				& HTMLHiddenField("item_number", pstrSessionID) & vbcrlf _
				& HTMLHiddenField("invoice", pstrOrderID) & vbcrlf _
				& HTMLHiddenField("amount", FormatNumber(pdblAmountDue,2)) & vbcrlf _
				& HTMLHiddenField("no_shipping", cbytno_shipping) & vbcrlf _
				& HTMLHiddenField("shipping", pdblshipping) & vbcrlf _
				& HTMLHiddenField("shipping2", pdblshipping2) & vbcrlf _
				& HTMLHiddenField("first_name", pstrFirstName) & vbcrlf _
				& HTMLHiddenField("last_name", pstrLastName) & vbcrlf _
				& HTMLHiddenField("address1", pstrAddress1) & vbcrlf _
				& HTMLHiddenField("address2", pstrAddress2) & vbcrlf _
				& HTMLHiddenField("city", pstrCity) & vbcrlf _
				& HTMLHiddenField("state", pstrState) & vbcrlf _
				& HTMLHiddenField("zip", pstrZIP) & vbcrlf _
				& HTMLHiddenField("night_phone_a", pstrnight_phone_a) & vbcrlf _
				& HTMLHiddenField("night_phone_b", pstrnight_phone_b) & vbcrlf _
				& HTMLHiddenField("night_phone_c", pstrnight_phone_c) & vbcrlf _
				& HTMLHiddenField("day_phone_a", pstrday_phone_a) & vbcrlf _
				& HTMLHiddenField("day_phone_b", pstrday_phone_b) & vbcrlf _
				& HTMLHiddenField("day_phone_c", pstrday_phone_c) & vbcrlf _
				& HTMLHiddenField("currency_code", currencyCodeToString(cbytCurrencyType)) & vbcrlf _
				& HTMLHiddenField("custom", pstrCustom) & vbcrlf _
				& "</form>" & vbcrlf

'this field is parked due to a bug on the PayPal site
'			 & HTMLHiddenField("undefined_quantity", "0") & vbcrlf _
	End Function	'PaymentForm
	
	'***********************************************************************************************

	Public Function PaymentLink(strPreMessage,strSubmitHTML,strPostMessage)

	Dim pstrForm

	PaymentLink = strPreMessage _
				& "&nbsp;<a href='' onclick='document.frmSSPayment.submit(); return false;'>" & strSubmitHTML & "</a>" _
				& "&nbsp;" & strPostMessage & vbcrlf

	End Function	'PaymentForm
	
	'***********************************************************************************************

	Public Function AutoSubmitScript(strMessage, delay)
		AutoSubmitScript = strMessage _
						 & "<script language='javascript'  type='text/javascript'>window.setTimeout ('document.frmSSPayment.submit();'," & delay * 1000 & ");</script>" & vbcrlf
	End Function	'AutoSubmitScript

	'***********************************************************************************************

End Class


'***********************************************************************************************
'***********************************************************************************************

'***********************************************************************************************

Sub ProcessPayPalReturn

Dim plngOrderID
Dim pobjrsOrder
Dim psngAmountPaid
Dim psngTotalDue
Dim pblnValidPayment
Dim pblnFoundOrder
Dim pblnPaymentMade
Dim pblnPaymentVerified
Dim pblnPaymentSufficient

	If Err.number <> 0 Then Err.Clear
	
	'this line will write to the transaction table
	Call PayPalResp("2")
	If Err.number <> 0 Then Err.Clear

	plngOrderID = Request.QueryString("OrderID")
	plngOrderID = iOrderID

	'Retrieve Order Information
	Set pobjrsOrder = CreateObject("ADODB.RECORDSET")
	sql = "Select orderGrandTotal from sfOrders where orderID = " & plngOrderID
	pobjrsOrder.Open sql,cnn,3,1
	If pobjrsOrder.EOF Then
		pblnFoundOrder = False 
		psngTotalDue = 0
	Else
		pblnFoundOrder = True 
		psngTotalDue = CSng(pobjrsOrder.Fields("orderGrandTotal").Value)
	End If
	pobjrsOrder.Close
	set pobjrsOrder = Nothing
	
	If pblnFoundOrder Then

		Set mclsPayPal = New clsPayPal
		With mclsPayPal
			Call .ValidatePayment(plngOrderID)
			pblnPaymentMade = .FoundOrder
			pblnPaymentVerified = .Completed
			psngAmountPaid = CSng(.VeriedPaymentAmount)
			pblnPaymentSufficient = (psngAmountPaid >= psngTotalDue)
			
			If False Then
				Response.Write "pblnPaymentMade: " & pblnPaymentMade & "<br />"
				Response.Write "pblnPaymentVerified: " & pblnPaymentVerified & "<br />"
				Response.Write "psngAmountPaid: " & psngAmountPaid & "<br />"
				Response.Write "pblnPaymentSufficient: " & pblnPaymentSufficient & "<br />"
			End If
			
			'User defined rules
			pblnValidPayment = pblnPaymentMade AND pblnPaymentSufficient 'AND pblnPaymentVerified
			If pblnFoundOrder Then
				If pblnPaymentMade Then
					If pblnPaymentSufficient Then
						If pblnPaymentVerified Then
							'sql = "Update sfOrders Set orderPaymentMethod = 'PayPal', orderIsComplete = 1 where orderID = " & plngOrderID
							sql = "Update sfOrders Set orderPaymentMethod = 'PayPal' where orderID = " & plngOrderID
						Else
							sql = "Update sfOrders Set orderPaymentMethod = 'PayPal - Unconfirmed' where orderID = " & plngOrderID
							'Dim iOrderID
							'iOrderID = plngOrderID
							'Call PayPal()
						End If
					Else
						sql = "Update sfOrders Set orderPaymentMethod = 'PayPal - Insufficient' where orderID = " & plngOrderID
					End If
				Else
					sql = "Update sfOrders Set orderPaymentMethod = 'PayPal - Unverified' where orderID = " & plngOrderID
				End If
				cnn.Execute sql,,128
			End If

			If pblnValidPayment Then
				.Connection = cnn
				.Save
			End If
		End With
		Set mclsPayPal = Nothing

	End If

	If pblnFoundOrder Then
		If pblnPaymentMade Then
			If pblnPaymentSufficient Then
				If pblnPaymentVerified Then
					mstrssPayPalMessage = "<h3>Thank You for your order. You will be receiving your order shortly.</h3>"
					
					'section to add custom code

				Else
					mstrssPayPalMessage = "<h3>Thank You for your order. Your order will be shipped as soon as Customer Service verifies payment.</h3>"
				End If
			Else
				mstrssPayPalMessage = "<h3>I was unable to verify the amount PayPal payment. Your order will be shipped as soon as Customer Service verifies payment.</h3>"
			End If
		Else
			mstrssPayPalMessage = "<h3>I was unable to verify your PayPal payment. Your order will be shipped as soon as Customer Service verifies payment.</h3>"
		End If
	Else
		mstrssPayPalMessage = "<h3>No Order Number</h3>"
		sProcErrMsg = "No Order Number"
	End If
	
End Sub	'ProcessPayPalReturn

	'***********************************************************************************************

Sub WritePayPalPaymentLine_SF5

	If isPaidByPayPal(sPaymentMethod, sTransMethod) And (CDbl(sGrandTotal) > 0) Then 
		Set mclsPayPal = New clsPayPal
		With mclsPayPal
			.OrderID = iOrderID
			.SessionID = Session("SessionID")
			.AmountDue = sGrandTotal
			.FirstName = sCustFirstName
			.LastName = sCustLastName
			.Address1 = sCustAddress1
			.Address2 = sCustAddress2
			.City = sCustCity
			.State = sCustState
			.ZIP = sCustZip
			.Phone sCustPhone, False	'Night Phone
			.Phone sCustPhone, True		'Day Phone
			.shipping = 0
			.shipping2 = 0
			.email = sCustEmail
			.Custom = mstrssPayPalCustomString
			
			If cblnDebugPayPalAddon Then
				Dim pstrCurrentLocation
				pstrCurrentLocation = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
				pstrCurrentLocation = Replace(pstrCurrentLocation, "confirm.asp", "ssPayPal_InstantNotification.asp")
				
				If LCase(Request.ServerVariables("HTTPS")) = "on" Then
					.ssPayPalReturnURL = "https://" & pstrCurrentLocation
				Else
					.ssPayPalReturnURL = "http://" & pstrCurrentLocation
				End If
			Else
				.ssPayPalReturnURL = mstrssPayPalReturnURL
			End If
			 
			Response.Write "				<tr><td colspan='2' align='center' class='tdContent2'>"
			Response.Write .PaymentForm
%>
<table width="90%" cellSpacing="0" cellPadding="4" border="1" style="border-collapse: collapse">
      <tr>
        <td>
          <% If cblnDebugPayPalAddon Then Response.Write "<h4>PayPal debugging Enabled: $0.01 payment being processed.</h4>" %>
          <p><b>Your order is not yet complete</b>, 
          <%= .AutoSubmitScript("You will be automatically redirected to the secure PayPal server to enter your payment information.",5) %>
          <%= .PaymentLink(" If you are not automatically redirected you may click","here","&nbsp;or the PayPal Check Out button below to finish your order.") %>
          </p>
        </td>
      </tr>
</table>
<%
			Response.Write "				</td></tr>"

			'Response.Write .AutoSubmitScript("You will be automatically redirected to the secure PayPal server to enter your payment information.",5)
			'Response.Write .PaymentLink("<br />If you are not automatically redirected you may click",.PaymentLinkText,".")
			'Response.Write "				</H3></td></tr>"
		End With
		Set mclsPayPal = Nothing
	End If

End Sub	'WritePayPalPaymentLine_SF5

'***********************************************************************************************
'
'	MAIN
'
'***********************************************************************************************

Const cstrConfirm_PayPalFailCode = "False"
Dim mclsPayPal
Dim mstrssPayPalCustomString
Dim mstrssPayPalReturnURL
Dim mstrssPayPalMessage

Dim cblnDebugPayPalAddon
cblnDebugPayPalAddon = Len(Session("ssDebug_PayPal")) > 0

'***********************************************************************************************

Function URLDecode(byVal strSource) 'Decodes an encoded QueryString (all URL Entity equivs. will be converted)

Dim aChars
Dim aANSICodes
Dim aLength
Dim sOut

	If Len(strSource) = 0 Then Exit Function

	aChars		= Array(" ",	"'",	"!",	"#",	"$",	"%",	"&",	"(",	")",	"/",	":",	";",	"[",	"\",	"]",	"^",	"`",	"{",	"|",	"}",	"+",	"<",	"=",	">")
	aANSICodes	= Array("+",	"%27",	"%21",	"%23",	"%24",	"%25",	"%26",	"%28",	"%29",	"%2F",	"%3A",	"%3B",	"%5B",	"%5C",	"%5D",	"%5E",	"%60",	"%7B",	"%7C",	"%7D",	"%2B",	"%3C",	"%3D",	"%3E")
	aLength = UBound(aChars)
	
	sOut = strSource
	For i = 0 To aLength
		sOut = (Replace(sOut, aANSICodes(i), aChars(i), 1, -1, vbTextCompare))
	Next
    
    URLDecode = sOut

End Function	'URLDecode

'***********************************************************************************************

Function checkForValidOrderID(byRef lngOrderID)

	If Len(Session("OrderID")) = 0 Then
		lngOrderID = Request.Form("item_number")
	Else
		lngOrderID = Session("OrderID")
	End If
	If Not isNumeric(lngOrderID) Then lngOrderID = ""	'check for malicious input
	  	    
    checkForValidOrderID = CBool(Len(iOrderID) = 0)

End Function	'checkForValidOrderID

'***********************************************************************************************

Function goodProcessorResponse(byVal bytLocation)
'all variables declared on confirm.asp

Dim pblnReturn
Dim pblnIsSF5AE

	pblnReturn = False
	pblnIsSF5AE = CBool(Application("AppName") = "StoreFrontAE")

	If bytLocation = 1 Then
	'process the order as if complete but do not set it complete
		If pblnIsSF5AE Then
			If isPaidByPayPal(sPaymentMethod, sTransMethod) Then
				pblnReturn = CBool(Len(Trim(sProcErrMsg)) = 0 Or sProcErrMsg = cstrConfirm_PayPalFailCode)
			Else
				pblnReturn = CBool(Len(Trim(sProcErrMsg)) = 0)
			End If
		Else
			'only displayed for SE sites since one cannot return to confirm.asp for AE sites
			pblnReturn = CBool(Len(Trim(sProcErrMsg)) = 0)
			If Len(mstrssPayPalMessage) > 0 Then Response.Write "<p>" & mstrssPayPalMessage & "</p>" & vbcrlf
		End If
	ElseIf bytLocation = 2 Then
	'set the order complete only if sProcErrMsg is empty
		If isPaidByPayPal(sPaymentMethod, sTransMethod) Then
			pblnReturn = False
		Else
			pblnReturn = CBool(Len(Trim(sProcErrMsg)) = 0)
		End If
	ElseIf bytLocation = 3 Then
	'send the email
		If pblnIsSF5AE Then
			If isPaidByPayPal(sPaymentMethod, sTransMethod) Then
				pblnReturn = cblnSendConfirmationEmail
			Else
				pblnReturn = True
			End If
		Else
			pblnReturn = True	'SE will only get to this point for valid emails
		End If
	ElseIf bytLocation = 4 Then
	'displays the order if sProcErrMsg is empty or the pseudo-PayPal fail code
		pblnReturn = CBool(Len(Trim(sProcErrMsg)) = 0 Or sProcErrMsg = cstrConfirm_PayPalFailCode)
	Else
		'should not see this
		pblnReturn = False
	End If

	goodProcessorResponse = pblnReturn
	  
End Function	'goodProcessorResponse

'***********************************************************************************************

Function isPaidByPayPal(byVal strPaymentMethod, byVal strTransactionMethod)
	isPaidByPayPal = CBool(strPaymentMethod = "PayPal Transaction" OR strPaymentMethod = "PayPal" OR (strTransactionMethod = "15" AND NOT strPaymentMethod = cstrPhoneFaxTerm))
End Function	'isPaidByPayPal

'***********************************************************************************************

Function modifyPayPayString(byVal strCustom, byVal strReturnURL)

	mstrssPayPalCustomString = strCustom
	mstrssPayPalReturnURL = strReturnURL

	modifyPayPayString = cstrConfirm_PayPalFailCode
	  
End Function	'modifyPayPayString

'***********************************************************************************************

%>

