<!--#include File = "modEncryption.asp"-->
<%
'********************************************************************************
'*
'*   incConfirm.asp
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins is incConfirm.asp
'*   from LaGarde, Incorporated. It has been heavily modified by Sandshot Software
'*   with permission from LaGarde, Incorporated who has NOT released any of the 
'*   original copyright protections. As such, this is a derived work covered by
'*   the copyright provisions of both companies.
'*
'*   LaGarde, Incorporated Copyright Statement                                                                           *
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

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	vDebug = 0	'overall application debugging - this can be set at the page level as well

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'The following functions were deleted as dead code
'Function CheckCardChange(sCardType,sCardName,sCardNumber,sCardExpiryMonth,sCardExpiryYear)
'Function CheckPhoneFax
'Function getShipping()

'**********************************************************
'*	Page Level variables
'**********************************************************


'**********************************************************
'*	Functions
'**********************************************************

'Sub setActive(sPrefix,iID)
'Function setOrderInitial(byVal lngOrderID, byVal lngCustID, byVal lngPayID, byVal lngAddrID, byVal strPaymentMethod, byVal strRoutingNumber, byVal strBankName, byVal strCheckNumber, byVal strCheckingAccountNumber, byVal strPONumber, byVal strPOName, byVal strShipMethodName, byVal dblTotalSTax, byVal dblTotalCTax, byVal dblHandling, byVal dblShipping, byVal dblSubTotal, byVal dblGrandTotal, byVal strShipInstructions, byVal dblCODAmount, byVal aryReferer)
'Function setPayments(byVal lngCustID, byVal sCardType, byVal sCardName, byVal sCardNumber, byVal sCardExpiryMonth, byVal sCardExpiryYear, byVal iCC)
'Sub setTransactionResponse(byVal lngOrderID, byVal ProcMessage, byVal ProcCustNumber, byVal ProcAddlData, byVal ProcRefCode, byVal ProcAuthCode, byVal ProcMerchNumber, byVal ProcActionCode, byVal ProcErrMsg, byVal ProcErrLoc, byVal ProcErrCode, byVal ProcAvsCode)

'**********************************************************
'*	Begin Page Code
'**********************************************************


'**********************************************************
'**********************************************************

Sub setActive(sPrefix,iID)

Dim pobjCmd
	
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		Set .ActiveConnection = cnn
		.Commandtext = pstrSQL
		.Parameters.Append .CreateParameter("id", adInteger, adParamInput, 4, custID_cookie)

		Select Case sPrefix 
			Case "cshpaddr"
				.Commandtext = "Update sfCShipAddresses Set cshpaddrIsActive=0 WHERE cshpaddrCustID=?"
				.Execute , , adExecuteNoRecords
				
				.Commandtext = "Update sfCShipAddresses Set cshpaddrIsActive=1 WHERE cshpaddrID=?"
				.Parameters("id").Value = iID
				.Execute , , adExecuteNoRecords
			Case "pay"
				.Commandtext = "Update sfCPayments Set payIsActive=0 WHERE payCustId=?"
				.Execute , , adExecuteNoRecords
				
				.Commandtext = "Update sfCPayments Set payIsActive=1 WHERE payID=?"
				.Parameters("id").Value = iID
				.Execute , , adExecuteNoRecords
		End Select
	End With	'pobjCmd
	closeobj(pobjCmd)
		
End Sub	'SetActive

'**********************************************************

Function setOrderInitial(byVal lngOrderID, byVal lngCustID, byVal lngPayID, byVal lngAddrID, byVal strPaymentMethod, byVal strRoutingNumber, byVal strBankName, byVal strCheckNumber, byVal strCheckingAccountNumber, byVal strPONumber, byVal strPOName, byVal strShipMethodName, byVal dblTotalSTax, byVal dblTotalCTax, byVal dblHandling, byVal dblShipping, byVal dblSubTotal, byVal dblGrandTotal, byVal strShipInstructions, byVal dblCODAmount, byVal aryReferer)

Dim pdblTotalHandling
Dim plngID
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim pstrTempTime

	pstrTempTime = CStr(Now())
	
	If dblCODAmount > "0" Then
		pdblTotalHandling = CDbl(dblHandling) + CDbl(dblCODAmount)
	Else
		pdblTotalHandling =  CDbl(dblHandling)
	End If
	
	'custID and order date can be indexed in sfOrders to speed retrieval
	pstrSQL = "Insert Into sfOrders (orderCustId, orderGrandTotal, orderDate) Values (?,?,?)"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		'.Commandtype = adCmdStoredProc
		Set .ActiveConnection = cnn
		
		If Len(CStr(lngOrderID)) = 0 Or CStr(lngOrderID)= "0" Then

			'If Len(pstrtmpAttrValue) = 0 Then pstrtmpAttrValue = Null
			'Note: orderGrandTotal parameter added since there was one case where orderID retrieval matched to an earlier ID
			.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, lngCustID)
			addParameter pobjCmd, "orderGrandTotal", adWChar, pstrTempTime, 50, 2

			.Parameters.Append .CreateParameter("orderDate", adDBTimeStamp, adParamInput, 16, pstrTempTime)
			.Execute , , adExecuteNoRecords

			pstrSQL = "Select orderID From sfOrders Where orderCustId=? And orderGrandTotal=? And orderDate=? Order By orderID Desc"
												  
			.Commandtext = pstrSQL
			Set pobjRS = .Execute
			If pobjRS.EOF Then
				plngID = -1
			Else
				plngID = pobjRS.Fields("orderID").Value
			End If
			closeobj(pobjRS)
			
			If vDebug = 1 Then Call WriteCommandParameters(pobjCmd, plngID, "setOrderInitial")
			
			'Remove two existing parameters
			.Parameters.Delete "orderCustId"	'technically you could reuse this one
			.Parameters.Delete "orderGrandTotal"
			.Parameters.Delete "orderDate"
		Else
			'debugprint "using existing order", lngOrderID
			plngID = lngOrderID
		End If	'Len(CStr(lngOrderID)) = 0 Or lngOrderID = 0
		
		pstrSQL = "Update sfOrders Set " _
				& "orderPayId=?, " _
				& "orderAddrId=?, " _
				& "orderAmount=?, " _
				& "orderComments=?, " _
				& "orderShipMethod=?, " _
				& "orderSTax=?, " _
				& "orderCTax=?, " _
				& "orderHandling=?, " _
				& "orderShippingAmount=?, " _
				& "orderGrandTotal=?, " _
				& "orderPaymentMethod=?, " _
				& "orderCheckAcctNumber=?, " _
				& "orderCheckNumber=?, " _
				& "orderBankName=?, " _
				& "orderRoutingNumber=?, " _
				& "orderPurchaseOrderName=?, " _
				& "orderPurchaseOrderNumber=?, " _
				& "orderRemoteAddress=?, " _
				& "orderTradingPartner=?, " _
				& "orderHttpReferrer=?" _
				& " Where orderID=?"
		.Commandtext = pstrSQL
				
		.Parameters.Append .CreateParameter("orderPayId", adInteger, adParamInput, 4, lngPayID)
		.Parameters.Append .CreateParameter("orderAddrId", adInteger, adParamInput, 4, lngAddrID)

		addParameter pobjCmd, "orderAmount", adWChar, dblSubTotal, 50, 2
		addParameter pobjCmd, "orderComments", adWChar, strShipInstructions, 2147483646, 2
		addParameter pobjCmd, "orderShipMethod", adWChar, strShipMethodName, 50, 2
		addParameter pobjCmd, "orderSTax", adWChar, dblTotalSTax, 50, 2
		addParameter pobjCmd, "orderCTax", adWChar, dblTotalCTax, 50, 2
		addParameter pobjCmd, "orderHandling", adWChar, pdblTotalHandling, 50, 2
		addParameter pobjCmd, "orderShippingAmount", adWChar, dblShipping, 50, 2
		addParameter pobjCmd, "orderGrandTotal", adWChar, dblGrandTotal, 50, 2
		addParameter pobjCmd, "orderPaymentMethod", adWChar, strPaymentMethod, 20, 2
		addParameter pobjCmd, "orderCheckAcctNumber", adWChar, strCheckingAccountNumber, 100, 2
		addParameter pobjCmd, "orderCheckNumber", adWChar, strCheckNumber, 100, 2
		addParameter pobjCmd, "orderBankName", adWChar, strBankName, 255, 2
		addParameter pobjCmd, "orderRoutingNumber", adWChar, strRoutingNumber, 255, 2
		addParameter pobjCmd, "orderPurchaseOrderName", adWChar, strPOName, 255, 2
		addParameter pobjCmd, "orderPurchaseOrderNumber", adWChar, strPONumber, 255, 2

		if isArray(aryReferer) then
   			addParameter pobjCmd, "orderRemoteAddress", adWChar, aryReferer(2), 255, 2
   			addParameter pobjCmd, "orderTradingPartner", adWChar, aryReferer(0), 100, 2
   			addParameter pobjCmd, "orderHttpReferrer", adWChar, aryReferer(1), 255, 2
		else
   			.Parameters.Append .CreateParameter("orderRemoteAddress", adWChar, adParamInput, 255, Null)
   			.Parameters.Append .CreateParameter("orderTradingPartner", adWChar, adParamInput, 100, Null)
   			.Parameters.Append .CreateParameter("orderHttpReferrer", adWChar, adParamInput, 255, Null)
		end if
		.Parameters.Append .CreateParameter("orderID", adInteger, adParamInput, 4, plngID)

		If vDebug = 1 Then Call WriteCommandParameters(pobjCmd, plngID, "setOrderInitial-update")
		.Execute , , adExecuteNoRecords
		
	End With	'pobjCmd
	closeobj(pobjCmd)
	setOrderInitial = plngID
	
	Call SaveCustomFormFields_SP(plngID)
	
End Function	'setOrderInitial

'**********************************************************

Function setPayments(byVal lngCustID, byVal sCardType, byVal sCardName, byVal sCardNumber, byVal sCardExpiryMonth, byVal sCardExpiryYear, byVal iCC)

Dim plngID
Dim pobjCmd
Dim pobjRS
Dim pstrPayCardNumber
Dim pstrSQL
Dim pstrCCV
	
	pstrPayCardNumber = Trim(sCardNumber)
	If CBool(adminEncodeCCIsActive = 1) Then pstrPayCardNumber = EnDeCrypt(Trim(pstrPayCardNumber), cstrRC4Key)
	
	Call DebugRecordSplitTime("setPayments . . .")
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		Set .ActiveConnection = cnn
		
		pstrSQL = "Update sfCPayments Set payIsActive=0 Where payCustId=?"
		.Commandtext = pstrSQL
		.Parameters.Append .CreateParameter("payCustId", adInteger, adParamInput, 4, lngCustID)
		.Execute , , adExecuteNoRecords

		If Len(cstrCCVFieldName) > 0 And cstrCCV_SaveToDB Then
			pstrSQL = "Insert Into sfCPayments (payCustId, payCardType, payCardName, payCardNumber, payCardExpires, " & cstrCCVFieldName & ", payIsActive) Values (?, ?, ?, ?, ?, ?, 1)"
			pstrCCV = Trim(Request.Form("payCardCCV"))
			If Len(pstrCCV) = 0 Then
				pstrCCV = Null
			ElseIf Len(pstrCCV) > 4 Then
				pstrCCV = Left(pstrCCV, 4)
			End If
		Else
			pstrSQL = "Insert Into sfCPayments (payCustId, payCardType, payCardName, payCardNumber, payCardExpires, payIsActive) Values (?, ?, ?, ?, ?, 1)"
		End If

		sCardType = trim(sCardType)
		sCardName = trim(sCardName)
		sCardExpiryMonth = trim(sCardExpiryMonth)
		sCardExpiryYear = trim(sCardExpiryYear)

		.Commandtext = pstrSQL
		addParameter pobjCmd, "payCardType", adWChar, sCardType, 50, 2
		addParameter pobjCmd, "payCardName", adWChar, sCardName, 50, 2
		addParameter pobjCmd, "payCardNumber", adWChar, pstrPayCardNumber, 50, 2
		addParameter pobjCmd, "payCardExpires", adWChar, sCardExpiryMonth & "/" & sCardExpiryYear, 50, 2

		If Len(cstrCCVFieldName) > 0 And cstrCCV_SaveToDB Then addParameter pobjCmd, "payCardCCV", adWChar, pstrCCV, 4, 2
		
		.Execute , , adExecuteNoRecords
		Call DebugRecordSplitTime("setPayments, payment set, getting id. . .")

		'Should be able to narrow down the select by payCustID and payIsActive
		'This failed if no CCV was entered so reverted to the only one customer's card is active at a time
		'If Len(cstrCCVFieldName) > 0 And cstrCCV_SaveToDB Then
		'	pstrSQL = "Select payID From sfCPayments Where payCustId=? And payCardType=? And payCardName=? And payCardNumber=? And payCardExpires=? And " & cstrCCVFieldName & "=? And payIsActive=1 Order By payID Desc"
		'Else
		'	pstrSQL = "Select payID From sfCPayments Where payCustId=? And payCardType=? And payCardName=? And payCardNumber=? And payCardExpires=? And payIsActive=1 Order By payID Desc"
		'End If
		
		pstrSQL = "Select payID From sfCPayments Where payCustId=? And payIsActive=1 Order By payID Desc"
		.Parameters.Delete "payCardType"
		.Parameters.Delete "payCardName"
		.Parameters.Delete "payCardNumber"
		.Parameters.Delete "payCardExpires"
		If Len(cstrCCVFieldName) > 0 And cstrCCV_SaveToDB Then .Parameters.Delete "payCardCCV"
		
		.Commandtext = pstrSQL
		Set pobjRS = .Execute
		
		If pobjRS.EOF Then
			plngID = -1
		Else
			plngID = pobjRS.Fields("payID").Value
		End If
		closeobj(pobjRS)

		'debugprint "setPayments - getIdentity", getIdentity
		If vDebug = 1 Then Call WriteCommandParameters(pobjCmd, plngID, "setPayments")

	End With	'pobjCmd
	closeobj(pobjCmd)
	Call DebugRecordSplitTime("setPayments Complete")

	setPayments = plngID
	
End Function	'setPayments

'**********************************************************

Sub setTransactionResponse(byVal lngOrderID, byVal ProcMessage, byVal ProcCustNumber, byVal ProcAddlData, byVal ProcRefCode, byVal ProcAuthCode, byVal ProcMerchNumber, byVal ProcActionCode, byVal ProcErrMsg, byVal ProcErrLoc, byVal ProcErrCode, byVal ProcAvsCode)

Dim pobjCmd
Dim pstrSQL
	
	Call DebugRecordSplitTime("setTransactionResponse . . .")
	pstrSQL = "Insert Into sfTransactionResponse (trnsrspOrderId, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode, trnsrspRetrievalCode, trnsrspAuthNo, trnsrspErrorMsg, trnsrspErrorLocation, trnsrspSuccess) Values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		'.Commandtype = adCmdStoredProc
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("trnsrspOrderId", adInteger, adParamInput, 4, lngOrderID)
		.Parameters.Append .CreateParameter("trnsrspCustTransNo", adWChar, adParamInput, parameterFieldLength(ProcCustNumber, 255), checkFieldLength(ProcCustNumber, 255, 2))
		.Parameters.Append .CreateParameter("trnsrspMerchTransNo", adWChar, adParamInput, parameterFieldLength(ProcMerchNumber, 255), checkFieldLength(ProcMerchNumber, 255, 2))
		.Parameters.Append .CreateParameter("trnsrspAVSCode", adWChar, adParamInput, parameterFieldLength(ProcAvsCode, 255), checkFieldLength(ProcAvsCode, 255, 2))
		.Parameters.Append .CreateParameter("trnsrspAUXMsg", adWChar, adParamInput, parameterFieldLength(ProcMessage, 255), checkFieldLength(ProcMessage, 255, 2))
		.Parameters.Append .CreateParameter("trnsrspActionCode", adWChar, adParamInput, parameterFieldLength(ProcActionCode, 255), checkFieldLength(ProcActionCode, 255, 2))
		.Parameters.Append .CreateParameter("trnsrspRetrievalCode", adWChar, adParamInput, parameterFieldLength(ProcRefCode, 255), checkFieldLength(ProcRefCode, 255, 2))
		.Parameters.Append .CreateParameter("trnsrspAuthNo", adWChar, adParamInput, parameterFieldLength(ProcAuthCode, 255), checkFieldLength(ProcAuthCode, 255, 2))
		.Parameters.Append .CreateParameter("trnsrspErrorMsg", adWChar, adParamInput, parameterFieldLength(ProcErrMsg, 255), checkFieldLength(ProcErrMsg, 255, 2))
		.Parameters.Append .CreateParameter("trnsrspErrorLocation", adWChar, adParamInput, parameterFieldLength(ProcErrLoc, 255), checkFieldLength(ProcErrLoc, 255, 2))
		If ProcResponse <> "failed" Then
			.Parameters.Append .CreateParameter("trnsrspSuccess", adWChar, adParamInput, parameterFieldLength(1, 255), checkFieldLength(1, 255, 2))
		Else
			.Parameters.Append .CreateParameter("trnsrspSuccess", adWChar, adParamInput, parameterFieldLength(0, 255), checkFieldLength(0, 255, 2))
		End If

		'.Parameters.Append .CreateParameter("CCV", adWChar, adParamInput, 4, checkFieldLength(ProcErrLoc, 4, 2))
		'.Parameters.Append .CreateParameter("trnsrspAuthorizationAmount", adWChar, adParamInput, 50, checkFieldLength(ProcErrLoc, 50, 2))

		'If True Then Call WriteCommandParameters(pobjCmd, "don't care", "setPayments")

	End With	'pobjCmd
	closeobj(pobjCmd)
	Call DebugRecordSplitTime("setTransactionResponse complete")

End Sub	'setTransactionResponse

'**********************************************************
'**********************************************************

'----------------------------------------------------------
' Sets the order complete flag to 1
'----------------------------------------------------------
Sub setOrderComplete(byVal lngOrderID)

Dim pobjCmd

	Call ssGiftCertificate_SaveRedemption(lngOrderID)
	Call SaveDiscounts(lngOrderID, sTotalPrice)
	Call setOrderManagerFraudPotential(lngOrderID)
	Call saveFreeGift(lngOrderID)
	
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "UPDATE sfOrders SET orderIsComplete = 1 WHERE orderID=?"
		'.Commandtype = adCmdStoredProc
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("trnsrspOrderId", adInteger, adParamInput, 4, lngOrderID)
		.Execute, ,adExecuteNoRecords
	End With	'pobjCmd
	closeobj(pobjCmd)
	
	Call updateVisitorSetOrderComplete

End Sub	'setOrderComplete

'***********************************************************************************************

Sub setOrderManagerFraudPotential(byVal iOrderID)

Dim pstrSQL
Dim pobjCommand
Dim plngOrderStatus
Dim plngInternalOrderStatus

	Call evaluateFraudPotentialScore(plngOrderStatus, plngInternalOrderStatus)
	pstrSQL = "Update ssOrderManager Set ssOrderStatus=?, ssInternalOrderStatus=? Where ssorderID=?"
			
	Set pobjCommand = CreateObject("ADODB.Command")
	With pobjCommand
		.CommandType = adCmdText
		.CommandText = pstrSQL
		.ActiveConnection = cnn

		.Parameters.Append .CreateParameter("ssOrderStatus", adInteger, adParamInput, 4, plngOrderStatus)
		.Parameters.Append .CreateParameter("ssInternalOrderStatus", adInteger, adParamInput, 4, plngInternalOrderStatus)
		.Parameters.Append .CreateParameter("ssorderID", adInteger, adParamInput, 4, iOrderID)

		On Error Resume Next	'added just in case
		.Execute ,,128	'adExecuteNoRecords
		
	End With	'pobjCommand
	Set pobjCommand = Nothing

End Sub	'setOrderManagerPayment


'Sandshot Software specific modifications

'--------------------------------------------------------
' Gets attribute number ' worst case scenario
'--------------------------------------------------------
Function getAttributeNumber
	Dim sLocalSQL, rsAttr, iCount
	
	sLocalSQL = "SELECT odrattrtmpID FROM sfTmpOrderAttributes INNER JOIN sfTmpOrderDetails ON sfTmpOrderAttributes.odrattrtmpOrderDetailId = sfTmpOrderDetails.odrdttmpID WHERE odrdttmpSessionID = " & SessionID
	Set rsAttr = CreateObject("ADODB.RecordSet")
	rsAttr.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	iCount = rsAttr.RecordCount
	closeobj(rsAttr)
	getAttributeNumber = iCount
End Function

%>