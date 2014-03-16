<%@ LANGUAGE="VBScript" %>
<% 
Option Explicit 
'********************************************************************************
'*   PayPal Payments															*
'*   Release Version:   3.00.003												*
'*   Release Date:		March 17, 2003											*
'*   Revision Date:		April 25, 2005   										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   3.00.003 (April 25, 2005)		                                            *
'*   -- Added validation for non-order related payments                         *
'*                                                                              *
'*   3.00.002 (September 19, 2004)                                              *
'*   -- Added additional debugging code                                         *
'*                                                                              *
'*   3.01 (April 5, 2003)                                                       *
'*   -- Added additional debugging code                                         *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True
%>
<!--#include file="SFLib/ssPayPal_IPNCustomActions.asp"-->
<!--#include file="SFLib/ssPayPal_IPNDebugModule.asp"-->
<%
	
	'***********************************************************************************************

	Function RetrieveRemoteData(strURL,strFormData,blnPostData)
	
	Dim pobjXMLHTTP
	
	'set timeouts in milliseconds
	Const resolveTimeout = 1000
	Const connectTimeout = 1000
	Const sendTimeout = 1000
	Const receiveTimeout = 10000
	
	On Error Resume Next

		If Err.number <> 0 Then	Err.Clear
	
		'Use MSXML2 if possible - must have the Microsoft XML Parser v3 or later installed
		Set pobjXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		If Err.number <> 0 Then
			Err.Clear
			Set pobjXMLHTTP = CreateObject("Microsoft.XMLHTTP")
		End If
		
		With pobjXMLHTTP
			If blnPostData Then
				.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
				.open "POST", strURL, False
				.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
				.send strFormData
			Else
				.open "GET", strURL, False
				.send
			End If
			RetrieveRemoteData  = .responseText

		End With
		set pobjXMLHTTP = nothing

		If Err.number <> 0 Then
			Select Case Err.number
				Case -2147467259:	'Unspecified error 
					'This has only been seen when the server permissions are set incorrectly (access is denied msxml3.dll)
					Response.Write "<h3><font color=red>Server permissions error: <br /><i>msxml3.dll error '80070005'<br />Access is denied.</i><br />Please contact your server administrator</font></h3>"
				Case 438: 'Object doesn't support this property or method
					'This is from the set timeouts, no action required
				Case Else
					Response.Write "<h3><font color=red>Error " & Err.number & ": " & Err.Description & "</font></h3>"
			End Select
			Err.Clear
		End If

	End Function	'RetrieveRemoteData
		
	'**************************************************************************************

	Function VerifyPayPalIPN(strFormData)

	Dim pstrResult

'	On Error Resume Next
	
		' post back to PayPal system to validate
		Output "<h4>Verifying IPN</h4>" & vbcrlf

		pstrResult = RetrieveRemoteData("https://www.paypal.com/cgi-bin/webscr",strFormData & "&cmd=_notify-validate",True)
		Output "&nbsp;&nbsp;Result: " & pstrResult & "<br />" & vbcrlf
		
		' Check notification validation
		If (pstrResult = "VERIFIED") then
			VerifyPayPalIPN = True
		ElseIf (pstrResult = "INVALID") then
			VerifyPayPalIPN = False
		Else 
			VerifyPayPalIPN = False
		End If
		
	End Function	'VerifyPayPalIPN
	
	'**************************************************************************************

	Function orig_VerifyPayPalIPN(strFormData)

	Dim pobjHttp
	Dim pstrResult

'	On Error Resume Next
	
		' post back to PayPal system to validate
		Set pobjHttp = CreateObject("Msxml2.ServerXMLHTTP")
		pobjHttp.Open "POST", "https://www.paypal.com/cgi-bin/webscr" , false
		pobjHttp.Send strFormData & "&cmd=_notify-validate"

		' Check notification validation
		If (pobjHttp.status <> 200 ) then
		' HTTP error handling
			pstrResult = False
		ElseIf (pobjHttp.responseText = "VERIFIED") then
			pstrResult = True
		ElseIf (pobjHttp.responseText = "INVALID") then
			pstrResult = False
		Else 
			pstrResult = False
		End If
		Set pobjHttp = Nothing
		
		VerifyPayPalIPN = pstrResult

	End Function	'VerifyPayPalIPN

	'***********************************************************************************************

	Sub LoadFromRequest

		With Request.Form
			pstraddress_city = Trim(.Item("address_city"))
			pstraddress_country = Trim(.Item("address_country"))
			pstraddress_state = Trim(.Item("address_state"))
			pstraddress_status = Trim(.Item("address_status"))
			pstraddress_street = Trim(.Item("address_street"))
			pstraddress_zip = Trim(.Item("address_zip"))
			pstrcustom = Trim(.Item("custom"))
			pstrinvoice = Trim(.Item("invoice"))
			pstritem_name = Trim(.Item("item_name"))
			pstritem_number = Trim(.Item("item_number"))
			pstrfirst_name = Trim(.Item("first_name"))
			pstrlast_name = Trim(.Item("last_name"))
			pstrnotify_version = Trim(.Item("notify_version"))
			pstrpayer_email = Trim(.Item("payer_email"))
			pstrpayer_status = Trim(.Item("payer_status"))
			pstrpayment_date = Trim(.Item("payment_date"))
			pstrpayment_fee = Trim(.Item("payment_fee"))
			pstrpayment_gross = Trim(.Item("payment_gross"))
			pstrpayment_status = Trim(.Item("payment_status"))
			pstrpayment_type = Trim(.Item("payment_type"))
			pstrpending_reason = Trim(.Item("pending_reason"))
			pstrquantity = Trim(.Item("quantity"))
			pstrreceiver_email = Trim(.Item("receiver_email"))
			pstrtxn_id = Trim(.Item("txn_id"))
			pstrtxn_type = Trim(.Item("txn_type"))
			pstrverify_sign = Trim(.Item("verify_sign"))
			
			If Len(pstrInvoice) = 0 Then pstrInvoice = Trim(Replace(.Item("item_name"), "Your Order ", ""))
		End With

		pstrpayment_date = FixPayPalDate(pstrpayment_date)

		Output "<hr><h4>LoadFromRequest - Variables</h4>" & vbcrlf
		Output "pstraddress_city: " & pstraddress_city & "<br />" & vbcrlf
		Output "pstraddress_country: " & pstraddress_country & "<br />" & vbcrlf
		Output "pstraddress_state: " & pstraddress_state & "<br />" & vbcrlf
		Output "pstraddress_status: " & pstraddress_status & "<br />" & vbcrlf
		Output "pstraddress_street: " & pstraddress_street & "<br />" & vbcrlf
		Output "pstraddress_zip: " & pstraddress_zip & "<br />" & vbcrlf
		Output "pstrcustom: " & pstrcustom & "<br />" & vbcrlf
		Output "pstrinvoice: " & pstrinvoice & "<br />" & vbcrlf
		Output "pstritem_name: " & pstritem_name & "<br />" & vbcrlf
		Output "pstritem_number: " & pstritem_number & "<br />" & vbcrlf
		Output "pstrfirst_name: " & pstrfirst_name & "<br />" & vbcrlf
		Output "pstrlast_name: " & pstrlast_name & "<br />" & vbcrlf
		Output "pstrnotify_version: " & pstrnotify_version & "<br />" & vbcrlf
		Output "pstrpayer_email: " & pstrpayer_email & "<br />" & vbcrlf
		Output "pstrpayer_status: " & pstrpayer_status & "<br />" & vbcrlf
		Output "pstrpayment_date: " & pstrpayment_date & "<br />" & vbcrlf
		Output "pstrpayment_fee: " & pstrpayment_fee & "<br />" & vbcrlf
		Output "pstrpayment_gross: " & pstrpayment_gross & "<br />" & vbcrlf
		Output "pstrpayment_status: " & pstrpayment_status & "<br />" & vbcrlf
		Output "pstrpayment_type: " & pstrpayment_type & "<br />" & vbcrlf
		Output "pstrpending_reason: " & pstrpending_reason & "<br />" & vbcrlf
		Output "pstrquantity: " & pstrquantity & "<br />" & vbcrlf
		Output "pstrreceiver_email: " & pstrreceiver_email & "<br />" & vbcrlf
		Output "pstrtxn_id: " & pstrtxn_id & "<br />" & vbcrlf
		Output "pstrtxn_type: " & pstrtxn_type & "<br />" & vbcrlf
		Output "pstrverify_sign: " & pstrverify_sign & "<br />" & vbcrlf
		Output "<br />" & vbcrlf

	End Sub 'LoadFromRequest

	'***********************************************************************************************

	Function FixPayPalDate(dtOrig)

	Dim pdtTempDate
	'remove PST/PDT

		pdtTempDate = Replace(dtOrig, " PST", "")
		pdtTempDate = Replace(pdtTempDate, " PDT", "")

		If isDate(pdtTempDate) Then
			pdtTempDate = FormatDateTime(pdtTempDate)
		Else
			pdtTempDate = FormatDateTime(Now())
		End If
		
		FixPayPalDate = pdtTempDate
		
	End Function	'FixPayPalDate

	'***********************************************************************************************
	
	Sub SaveFieldValue(byRef objRS, byRef strFieldName, byRef vntValue)
	
	On Error Resume Next
	
		If Err.number <> 0 Then Err.Clear
		objRS(strFieldName).Value = vntValue
		If Err.number <> 0 Then
			Response.Write "Error " & Err.number & ": " & Err.Description & "<br />"
			Response.Write "Field length for " & strFieldName & ": " & objRS(strFieldName).ActualSize & "<br />"
		End If
	
	End Sub	'SaveFieldValue
	
	'***********************************************************************************************

	Function SaveIPN()

	Dim pstrSQL
	Dim pobjRS
	Dim pblnNew

	'On Error Resume Next

		Output "<h4>SaveIPN</h4>" & vbcrlf
		pstrSQL = "Select * from PayPalPayments where txn_id = '" & pstrtxn_id & "'"
		Output "SaveIPN - pstrSQL: " & pstrSQL & "<br />" & vbcrlf
		Set pobjRS = CreateObject("adodb.Recordset")
		pobjRS.open pstrSQL, cnn, adOpenKeyset, adLockOptimistic
		pblnNew = pobjRS.EOF
		Output "SaveIPN - pobjRS.EOF: " & pobjRS.EOF & "<br />" & vbcrlf
		If pblnNew Then
			pobjRS.AddNew
			
			pobjRS.Fields("receiver_email").Value = pstrreceiver_email
			pobjRS.Fields("item_name").Value = checkFieldLength(pstritem_name, 100, 0)
			pobjRS.Fields("item_number").Value = checkFieldLength(pstritem_number, 50, 0)
			If Len(pstrquantity) > 0 And isNumeric(pstrquantity) Then pobjRS.Fields("quantity").Value = pstrquantity
			pobjRS.Fields("invoice").Value = checkFieldLength(pstrinvoice, 50, 0)
			pobjRS.Fields("custom").Value = pstrcustom
			pobjRS.Fields("payment_status").Value = checkFieldLength(pstrpayment_status, 50, 0)
			pobjRS.Fields("pending_reason").Value = checkFieldLength(pstrpending_reason, 50, 0)
			If Len(pstrpayment_date) > 0 And isDate(pstrpayment_date) Then pobjRS.Fields("payment_date").Value = pstrpayment_date
			If Len(pstrpayment_gross) > 0 And isNumeric(pstrpayment_gross) Then pobjRS.Fields("payment_gross").Value = pstrpayment_gross
			If Len(pstrpayment_fee) > 0 And isNumeric(pstrpayment_fee) Then pobjRS.Fields("payment_fee").Value = pstrpayment_fee
			pobjRS.Fields("txn_type").Value = checkFieldLength(pstrtxn_type, 50, 0)
			pobjRS.Fields("first_name").Value = checkFieldLength(pstrfirst_name, 50, 0)
			pobjRS.Fields("last_name").Value = checkFieldLength(pstrlast_name, 50, 0)
			pobjRS.Fields("address_city").Value = checkFieldLength(pstraddress_city, 50, 0)
			pobjRS.Fields("address_street").Value = checkFieldLength(pstraddress_street, 50, 0)
			pobjRS.Fields("address_state").Value = checkFieldLength(pstraddress_state, 50, 0)
			pobjRS.Fields("address_zip").Value = checkFieldLength(pstraddress_zip, 50, 0)
			pobjRS.Fields("address_country").Value = checkFieldLength(pstraddress_country, 50, 0)
			pobjRS.Fields("address_status").Value = checkFieldLength(pstraddress_status, 50, 0)
			pobjRS.Fields("payer_email").Value = pstrpayer_email
			pobjRS.Fields("payer_status").Value = checkFieldLength(pstrpayer_status, 50, 0)
			pobjRS.Fields("payment_type").Value = checkFieldLength(pstrpayment_type, 50, 0)
			pobjRS.Fields("notify_version").Value = checkFieldLength(pstrnotify_version, 5, 0)
			pobjRS.Fields("verify_sign").Value = pstrverify_sign
			pobjRS.Fields("txn_id").Value = checkFieldLength(pstrtxn_id, 50, 0)

			'Fields left at defaults
			'pobjRS.Fields("Category").Value = 
			'pobjRS.Fields("ActionCompleted_Completed").Value = 
			'pobjRS.Fields("ActionCompleted_Pending").Value = 
			'pobjRS.Fields("ActionCompleted_Failed").Value = 
			'pobjRS.Fields("ActionCompleted_Denied").Value = 
			
			pobjRS.Update

		End If
		pobjRS.Close
		
		'Now since PayPal will submit multiple notifications, we'll only records those that materially change the status
		pstrSQL = "Select PayPalIPNID From PayPalIPNs Where (txn_id='" & pstrtxn_id & "') AND (payment_status='" & pstrpayment_status & "') AND (pending_reason='" & pstrpending_reason & "')"
		Output "SaveIPN - pstrSQL: " & pstrSQL & "<br />" & vbcrlf
		pobjRS.CursorLocation = adUseClient
		pobjRS.open pstrSQL, cnn, adOpenForwardOnly, adLockReadOnly
		Output "SaveIPN - pobjRS.EOF: " & pobjRS.EOF & "<br />" & vbcrlf
		If pobjRS.EOF Then
			'Now insert the Notification into the PayPalIPNs table
			pstrSQL = "Insert Into PayPalIPNs (txn_id, payment_status, pending_reason, DateIPNReceived) Values ('" & pstrtxn_id & "', '" & pstrpayment_status & "', '" & pstrpending_reason & "', " & sqlDateWrap(Now()) & ")"
			Output "SaveIPN - pstrSQL: " & pstrSQL & "<br />" & vbcrlf
			cnn.Execute pstrSQL,,adExecuteNoRecords
			
			'Now update the PayPalPayments table
			If Len(pstrpayment_fee) > 0 Then
				pstrSQL = "Update PayPalPayments Set payment_status='" & pstrpayment_status & "', pending_reason='" & pstrpending_reason & "', payment_fee=" & pstrpayment_fee & " where txn_id = '" & pstrtxn_id & "'"
			Else
				pstrSQL = "Update PayPalPayments Set payment_status='" & pstrpayment_status & "', pending_reason='" & pstrpending_reason & "' where txn_id = '" & pstrtxn_id & "'"
			End If
			Output "SaveIPN - pstrSQL: " & pstrSQL & "<br />" & vbcrlf
			cnn.Execute pstrSQL,,adExecuteNoRecords
			
			Call PerformCustomAction(pstrtxn_id,pstrpayment_status,pstrpending_reason,pstrpayer_status, pstraddress_status, False)
		Else
			Call PerformCustomAction(pstrtxn_id,pstrpayment_status,pstrpending_reason,pstrpayer_status, pstraddress_status, False)
		End If
		pobjRS.Close
		Set pobjRS = Nothing

	End Function    'Update

	'***********************************************************************************************

	Function InitializeConnection()
	'Initializes the connection to the database 

		Set cnn = CreateObject("ADODB.Connection")
		cnn.Open Application("DSN_Name")
		
		If Err.number <> 0 then
'			pstrError = Err.description
			InitializeConnection = False
			Response.Write "<h3>Your harddrive path is <i>" & ServerPath & "</i></h3>"
		Else
			InitializeConnection = (cnn.State = 1)
		End If
		
	End Function		' InitializeConnection
	
	'**************************************************************************************

	Function checkFieldLength(byVal strValue, byVal lngMaxLength, byVal bytTrim)
	'bytTrim can have the following values
	'0 - trim left, empty ok
	'1 - trim right, empty ok
	'2 - trim left, replace zero length string with null
	'3 - trim right, replace zero length string with null
	Dim pstrOut

		pstrOut = strValue
		If lngMaxLength > 0 Then
			If Len(pstrOut) > lngMaxLength Then
				Select Case bytTrim
					Case 0, 2: 'left portion
						pstrOut = Left(pstrOut, lngMaxLength)
					Case 1, 3: 'right portion
						pstrOut = Right(pstrOut, lngMaxLength)
					Case Else:
					
				End Select
			End If	'Len(pstrOut) > lngMaxLength
		End If	'lngMaxLength > 0
		
		'now check for null replacements
		If Len(pstrOut) = 0 Then
			If bytTrim = 2 Or bytTrim = 3 Then pstrOut = Null
		End If
		
		checkFieldLength = pstrOut
		
	End Function	'checkFieldLength
	
	'**************************************************************************************

	Function sqlSafe(strSQL)

		sqlSafe = Replace(strSQL,"'","''")

	End Function
	
	'**************************************************************************************

	Function sqlDateWrap(strValue)

	Dim pstrTempValue

		If Len(strValue & "") = 0 Then
			sqlDateWrap = "Null"
		Else
			If cblnSQLDatabase Then
				sqlDateWrap = "'" & strValue & "'"
			Else
				sqlDateWrap = "#" & strValue & "#"
			End If
		End If

	End Function

	'**************************************************************************************

	Function CheckOrderComplete(strOrderID)

	Dim pstrSQL
	Dim pobjRS

	'On Error Resume Next

		Output "<h4>CheckOrderComplete</h4>" & vbcrlf
		pstrSQL = "Select orderIsComplete from sfOrders where orderIsComplete = 1 AND orderID = " & strOrderID
		Output "CheckOrderComplete - pstrSQL: " & pstrSQL & "<br />" & vbcrlf

		'Check added for non-order related payments
		If Len(strOrderID) = 0 Or Not isNumeric(strOrderID) Then
			CheckOrderComplete = False
		Else

			Set pobjRS = CreateObject("adodb.Recordset")
			pobjRS.open pstrSQL, cnn, adOpenKeyset, adLockOptimistic
			Output "CheckOrderComplete - pobjRS.EOF: " & pobjRS.EOF & "<br />" & vbcrlf
			
			CheckOrderComplete = pobjRS.EOF
		
			pobjRS.Close
			Set pobjRS = Nothing
		End If
			
	End Function	'CheckOrderComplete

	'**************************************************************************************

'**************************************************************************************
'
'		BEGIN CODE EXECUTION
'
'**************************************************************************************

'---- CursorTypeEnum Values ----
'Const adOpenForwardOnly = 0
'Const adOpenKeyset = 1
'Const adOpenDynamic = 2
'Const adOpenStatic = 3

'---- LockTypeEnum Values ----
'Const adLockReadOnly = 1
'Const adLockPessimistic = 2
'Const adLockOptimistic = 3
'Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
'Const adUseServer = 2
'Const adUseClient = 3

'---- CommandTypeEnum Values ----
'Const adCmdUnknown = &H0008
'Const adCmdText = &H0001
'Const adCmdTable = &H0002
'Const adCmdStoredProc = &H0004

'Const adExecuteNoRecords = 128

Dim cnn
Dim cblnSQLDatabase
cblnSQLDatabase = CBool(Application("AppDatabase") <> "Access")

'cblnSQLDatabase = True				'Set this value to True for SQL Server databases, only need to set this manually for very early versions
'cblnSQLDatabase = False			'Set this value to False for Access databases, only need to set this manually for very early versions

Dim pstraddress_city, pstraddress_country, pstraddress_state, pstraddress_status, pstraddress_street, pstraddress_zip, pstrcustom, pstrinvoice, pstritem_name, pstritem_number, pstrfirst_name, pstrlast_name, pstrnotify_version, pstrpayer_email, pstrpayer_status, pstrpayment_date, pstrpayment_fee, pstrpayment_gross, pstrpayment_status, pstrpayment_type, pstrpending_reason, pstrquantity, pstrreceiver_email, pstrtxn_id, pstrtxn_type, pstrverify_sign

'---- Debugging Section ----
Dim vItem

	Output "-- File called on " & Now() & "<br />" & vbcrlf
	Output "<br />" & vbcrlf

	Output "<hr><h4>Server Variables</h4>" & vbcrlf
	For Each vItem in Request.ServerVariables
		Output vItem & ": " & Request.ServerVariables(vItem) & "<br />" & vbcrlf
	Next
	Output "<br />" & vbcrlf
	
	Output "<hr><h4>Form Variables</h4>" & vbcrlf
	For Each vItem in Request.Form
		Output vItem & ": " & Request.Form(vItem) & "<br />" & vbcrlf
	Next
	Output "<br />" & vbcrlf

	If Len(Request.Form) > 0 Then
		If VerifyPayPalIPN(Request.Form) Then
			Call LoadFromRequest
			If InitializeConnection Then Call SaveIPN
			
			'added to send email/complete cart for StoreFront 5
			'Transfer to confirm.asp to send email
			'Note IPN does not currently support session variables so the below will not work without rewriting SF5AE code - SE seems to function

			'If True Then
			If CheckOrderComplete(pstrInvoice) Then
				Call WriteToFile(True, mstrWriteToFile, False)
				Session("SessionID") = pstritem_number
				Session("OrderID") = pstrinvoice
			'	Session("ssDebug_DisableMail") = True
			'	Server.Transfer "confirm.asp"
			Else
				Call WriteToFile(True, mstrWriteToFile, False)
			End If

			On Error Resume Next
			cnn.Close
			Set cnn = Nothing
			On Error Goto 0
		Else
			Output "Invalid IPN<br />" & vbcrlf
			Call WriteToFile(True, mstrWriteToFile, False)
		End If	'VerifyPayPalIPN
	Else
		Output "No Form Present<br />" & vbcrlf
		Output Server.Mappath("../") & "<br />" & vbcrlf
		Call WriteToFile(True, mstrWriteToFile, False)
	End If	'Len(Request.Form) > 0

If Response.Buffer Then Response.Flush
%>
<html>
<head>
<title>Sandshot Software - PayPal IPN</title>
</head>
<body>OK</body>
</html>