<%
'-----------------------------------------------------------------------------
' Verisign's Signio PayProFlow Subroutine
'-----------------------------------------------------------------------------
'https://payments.verisign.com/manager
Function SignioPayProFlow(proc_live)

Dim ccExp_Date, obj, ccAddress, Amt, TrxType, ObjResult, i, aPos, GrandTotal, iProcResponse, ProcRspMsg, ProcFailedMsg, strIn, aTemp,iTransactionID, iAuthCode, sAVSMsg
Dim sAVSADDR, sAVSZIP, sAVSMsg1, sAVSMsg2, ParmList, Ctx1, ProcMerchTransNum
Dim ProcResponse
Dim ProcAvsCode
Dim ProcErrMsg
Dim ProcActionCode
Dim pstrCVV
Dim pblnDebug_Local:	pblnDebug_Local = True	'True	False

	GrandTotal = FormatCurrency(sGrandTotal,2)
	ccExp_Date = sCustCardExpiryMonth & right(sCustCardExpiryYear,2) 'Issue #271

	ccAddress = sCustAddress1 & " " & sCustAddress2
	
	Amt = GrandTotal
	Amt = Mid(Amt,2)
		
	sMercType = UCase(sMercType)
	If sMercType = "AUTHCAPTURE" Then
		TrxType = "S"
	ElseIf sMercType = "AUTHONLY" Then
		TrxType = "A"
	End If

	If Len(Trim(sPaymentServer & "")) = 0 Then 
		If proc_live = 1 Then 
			sPaymentServer = "payflow.verisign.com"
		ElseIf proc_live =  0 Then 
			sPaymentServer = "test-payflow.verisign.com"
		End If
	End If

	'*********************************************
	'
	'  FOR TESTING
	'
	'  Note: PARTNER may need to be included in the login
	'        {username}&PARTNER={partner name}
	
	'sPaymentServer = "test-payflow.verisign.com"
	'sPaymentServer = "payflow.verisign.com"
	'sLogin = "{username}&PARTNER={partner name}"
	'sPassword = ""

	'
	'*********************************************

	If Len(mstrPayCardCCV) > 0 Then
		ParmList = "TRXTYPE=" & TrxType & "&TENDER=C&USER=" & sLogin & "&PWD=" & sPassword & "&ACCT=" & sCustCardNumber & "&EXPDATE=" & ccExp_Date & "&AMT=" & Amt & "&ZIP=" & sCustZip & "&City=" & sCustCity & "&State="& sCustState & "&STREET="& sCustAddress1 & "&PhoneNum="& sCustPhone & "&Country=" & sCustCountry & "&CUSTIP="& aReferer(2) & "&EMAIL=" & sCustEmail & "&FirstName=" & sCustFirstName & "&LastName="& sCustLastName & "&ShipFirstName=" & sShipCustFirstName & "&ShipLastName=" & sShipCustLastName & "&ShiptoCountry=" & sShipCustCountry & "&ShipToCity=" & sShipCustCity & "&ShiptoState=" & sShipCustState & "&ShipToZip=" & sShipCustZip & "&CVV2=" & mstrPayCardCCV
	Else
		ParmList = "TRXTYPE=" & TrxType & "&TENDER=C&USER=" & sLogin & "&PWD=" & sPassword & "&ACCT=" & sCustCardNumber & "&EXPDATE=" & ccExp_Date & "&AMT=" & Amt & "&ZIP=" & sCustZip & "&City=" & sCustCity & "&State="& sCustState & "&STREET="& sCustAddress1 & "&PhoneNum="& sCustPhone & "&Country=" & sCustCountry & "&CUSTIP="& aReferer(2) & "&EMAIL=" & sCustEmail & "&FirstName=" & sCustFirstName & "&LastName="& sCustLastName & "&ShipFirstName=" & sShipCustFirstName & "&ShipLastName=" & sShipCustLastName & "&ShiptoCountry=" & sShipCustCountry & "&ShipToCity=" & sShipCustCity & "&ShiptoState=" & sShipCustState & "&ShipToZip=" & sShipCustZip
	End If
	
	If sTransMethod = "17" Then	'PayFlowPro30
		On Error Resume Next 
		'Set obj = CreateObject("PFProSSControl.PFProSSControl2.1") 'No longer seems to work, assuming it ever did
		Set obj = CreateObject("PFProCOMControl.PFProCOMControl.1")
		If Err.number <> 0 Then
			Response.Write "<fieldset><legend>Error in PayFlow</legend>"
			Response.Write "<h1>Error " & err.number & ": " & err.Description & "</h1>"
			Response.Write "</fieldset>"
			Err.Clear
			Exit Function
		End If
		Ctx1 = obj.CreateContext(sPaymentServer, 443, 30, "", 0, "", "")
		strIn = obj.SubmitTransaction(Ctx1, ParmList, Len(ParmList))
		obj.DestroyContext (Ctx1)
			
	ElseIf sTransMethod = "3" Then	'Verisign PayFlow Pro
		Set obj = CreateObject("PFProSSControl.PFProSSControl.1")

		' Create object and process it
		obj.HostAddress = sPaymentServer
		obj.HostPort = 443
		obj.TimeOut = 30
		obj.DebugMode = 1
		obj.ParmList = ParmList
		
		obj.PNInit()
		obj.ProcessTransaction()
		strIn = obj.Response 	
		obj.PNCleanup 
	End If
	Set obj = Nothing

	If pblnDebug_Local Then Response.Write "<fieldset><legend>SignioPayProFlow (" & sTransMethod & ")</legend>"
	aPos = split(strIn, "&")
	For i = 0 to UBOUND(aPos)
		aTemp = split(aPos(i),"=")
		Select Case aTemp(0)
			Case "RESULT":		ProcResponse = Trim(aTemp(1))
			Case "RESPMSG":		ProcRspMsg = Trim(aTemp(1))
			Case "PNREF":		iTransactionID = Trim(aTemp(1))
			Case "AUTHCODE":	iAuthCode = Trim(aTemp(1))
			Case "ERRMSG":		ProcErrMsg = Trim(aTemp(1))
			Case "AVSADDR":	
								sAVSADDR = Trim(aTemp(1))
								sAVSMsg1 = getAVSMessage("", sAVSADDR)	
			Case "AVSZIP":	
								sAVSZIP = ProcAvsCode & Trim(aTemp(1))
								sAVSMsg2 = getAVSMessage("", sAVSZIP)	
			Case "ERRCODE":		ProcFailedMsg = Trim(aTemp(1))
			Case "FRAUDCODE":	ProcMerchTransNum = "<br />FraudCode: " & Trim(aTemp(1)) ' Item4
			Case "SCORE":		ProcMerchTransNum = ProcMerchTransNum & "<br />Score: " & Trim(aTemp(1)) 'Item4
			Case "FRAUDMSG":	ProcActionCode = "Fraud Msg: " & Trim(aTemp(1))
			Case "REASON1":		ProcActionCode = ProcActioncode & "<br />Reason 1: " & Trim(aTemp(1))
			Case "REASON2":		ProcActionCode = ProcActionCode & "<br />Reason 2: " & Trim(aTemp(1))
			Case "REASON3":		ProcActionCode = ProcActionCode & "<br />Reason 3: " & Trim(aTemp(1))
			Case "EXCEPTION1":	ProcActionCode = ProcActionCode & "<br />Exception 1: " & Trim(aTemp(1))
			Case "EXCEPTION2":	ProcActionCode = ProcActionCode & "<br />Exception 2: " & Trim(aTemp(1))
			Case "EXCEPTION3":	ProcActionCode = ProcActionCode & "<br />Exception 3: " & Trim(aTemp(1))
			Case "EXCEPTION4":	ProcActionCode = ProcActionCode & "<br />Exception 4: " & Trim(aTemp(1))
			Case "EXCEPTION5":	ProcActionCode = ProcActionCode & "<br />Exception 5: " & Trim(aTemp(1))
			Case "EXCEPTION6":	ProcActionCode = ProcActionCode & "<br />Exception 6: " & Trim(aTemp(1))
			Case "EXCEPTION7":	ProcActionCode = ProcActionCode & "<br />Exception 7: " & Trim(aTemp(1))
			Case "CVV2MATCH":	pstrCVV = Trim(aTemp(1))
		End Select
		If pblnDebug_Local Then
			If isArray(aTemp) Then
				If UBound(aTemp) >= 0 Then Response.Write aTemp(0)
				Response.Write ": "
				If UBound(aTemp) >= 1 Then Response.Write aTemp(1)
				Response.Write "<br />"
			Else
				Response.Write "aPos(" & i & "): " & aPos(i) & "<br />"
			End If
		End If
	Next	
	' The AVSCode is a combination of ADDR and ZIP AVS codes 
	ProcAvsCode = sAVSADDR & sAVSZIP
	sAVSMsg = sAVSMsg1 & ";" & sAVSMsg2

	' Failed or not  
	If ProcResponse = 0 Then
		iProcResponse = 1
	Else
		iProcResponse = 0
		Select Case ProcResponse
			Case "-1":	ProcFailedMsg = "Server Socket Unavailable"
			Case "-2":	ProcFailedMsg = "Hostname lookup failed"
			Case "-3":	ProcFailedMsg = "Server Timed Out"
			Case "-4":	ProcFailedMsg = "Socket Initialization Error"
			Case "-5":	ProcFailedMsg = "SSL Context Initialization Failed"
			Case "-6":	ProcFailedMsg = "SSL Verification Policy Failure"
			Case "-7":	ProcFailedMsg = "SSL Verify Location Failed"
			Case "-8":	ProcFailedMsg = "X509 Certification Verification Error"
			Case Else:	ProcFailedMsg = "General Processing Error - Please try to notify the merchant"
		End Select
		ProcFailedMsg = ProcFailedMsg & " " & ProcErrMsg & " " & ProcRspMsg
	End If

	If pblnDebug_Local Then
		Response.Write "ProcAvsCode: " & ProcAvsCode & "<br />"
		Response.Write "sAVSMsg: " & sAVSMsg & "<br />"
		Response.Write "CCV Result: " & " - Not Implemented - " & "<hr />"
		Response.Write "iProcResponse: " & iProcResponse & "<br />"
		Response.Write "ProcFailedMsg: " & ProcFailedMsg & "<br />"
		Response.Write "</fieldset>"
		
		'Now abort setting the response
		SignioPayProFlow = ProcFailedMsg 
		Exit Function
	End If
	
	' Write to response table		
	Call setResponse("Signio PFP",iOrderID,iTransactionID,ProcMerchTransNum,ProcAvsCode,sAVSMsg,ProcActionCode,iAuthCode,"",ProcFailedMsg ,iProcResponse)	
	SignioPayProFlow = ProcFailedMsg 
	
End Function	'SignioPayProFlow
%>








