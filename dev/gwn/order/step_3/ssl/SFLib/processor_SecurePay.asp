<%
'--------------------------------------------------------------------------
' SecurePay Function
'--------------------------------------------------------------------------
Function SecurePay(proc_live)

Dim objOrder, sActionCode, iProcResponse, iTransactionID, sErrorMessage, iAVSCode, sAvsMsg, orderAVSREQ
Dim SPCOM_Response, sSP_Array, sReturn_Code, sApprov_Num, sCard_Response, sAVS_Response, sTr_Type
Dim pstrCVV
	
    If trim(sPaymentServer) = "" or isNull(sPaymentServer) Then sPaymentServer = "https://processing.securepay.net/secure1/index.asp"
    
	' Testing or live mode
	If proc_live = 1 Then
		orderAVSREQ = "1"
	ElseIf proc_live = 0 Then
		orderAVSREQ = "4"
		
		'Set the transaction type. Added multiple Tr_types 12/11/2002 TR
		if sMercType <> "" Then
			select case uCase(sMercType)
				case "AUTHONLY":	sTr_Type="PREAUTH"
				case "AUTHCREDIT":	sTr_Type="CREDIT"
				case else:			sTr_Type="SALE"
			end select
		else
			sTr_Type="SALE"
		end if
	End If	
	
	'create instance of component 
	If Err.number <> 0 Then Err.Clear
	On Error Resume Next
	
    Set objOrder = Server.CreateObject("SPCOM.clsSecureSend")
    If Err.number = 0 Then
		With objOrder
			.URL = sPaymentServer
			.Amount = FormatNumber(sGrandTotal,2,0,0,0)
			.AVSREQ = orderAVSREQ
			.Street = sCustAddress1 & " " & sCustAddress2
			.City = sCustCity
			.State = sCustState
			.Zip = sCustZip
			.NameOnCard	= sCustCardName
			.CreditCardNumber = sCustCardNumber 
			.Email = sCustEmail
			.Month = sCustCardExpiryMonth 
			.Year = right(sCustCardExpiryYear,2)
			.Merch_ID = sLogin 
			.tr_type = sTr_Type
			
			'For CCV enabled
			If Not cstrCCV_Optional Then
				'Non-U.S. issued cards may not support CCV
				Select Case sCustCountry
					Case "US", "CA"
						.CVV2 = mstrPayCardCCV
					Case Else
						.CVV2 = "PASS"
				End Select
			End If
			
			.Send
			SPCOM_Response = .ReturnCode  
		End With	'objOrder
		Set objOrder = Nothing
		
    Else
		SPCOM_Response = "SPCOM object not installed"
		Err.Clear
    End If	'Err.number = 0
    
	if instr(1,SPCOM_Response,",") then
		sSP_Array = Split(SPCOM_Response, ",")
		
		sActionCode = sSP_Array(0)		'ActionCode
		sApprov_Num = sSP_Array(1)		'ApprovalNumber
		sCard_Response = sSP_Array(2)	'VerboseResponse
		sAVS_Response = sSP_Array(3)	'AVScode
		Call scoreAVS(sAVS_Response)
		
		Dim paryVerbose
		paryVerbose = Split(sCard_Response, ";")
		If UBound(paryVerbose) > 0 Then
			pstrCVV = paryVerbose(1)	'Possible CVV2 responses "CVV2 MATCH", "CVV2 NOT AVAILABLE", "CVV2 NOMATCH"
			Call scoreCVV(pstrCVV)
		End If
	else
		sCard_Response = SPCOM_Response
	end if
	
    If sActionCode = "Y" Then
		iProcResponse = 1
	Else 
		iProcResponse = 0
		sErrorMessage = sCard_Response
	End If 		
	
	iTransactionID =  sApprov_Num    
    iAVSCode = sAVS_Response
 
    sAvsMsg = getAVSMessage("SecurePay", iAvsCode)
	iAVSCode = Array(sAVS_Response, GrandTotal, getCVVMessage("SecurePay", pstrCVV) & " (" & pstrCVV & ")")

    Call setResponse("SecurePay", iOrderID, iTransactionID, "", iAVSCode, sAvsMsg, sActionCode, "", "", sErrorMessage, iProcResponse)
    
	SecurePay = sErrorMessage
	
End Function	'SecurePay

%>