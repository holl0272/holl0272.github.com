<%
'Reference http://uplinkearth.com/pdf/API_3.0_IntGuide.pdf

Function LinkPoint(byVal proc_live)

Dim host
Dim keyfile
Dim LPTxn
Dim order
Dim op
Dim outXml
Dim resp

	sLogin = "Your Login"
	
	If proc_live = 1 Then
		host = sPaymentServer
		If Len(host) = 0 Then host = "secure.linkpt.net"
	ElseIf proc_live = 0 Then
		host = "staging.linkpt.net"
	End If	

	keyfile = Replace(Server.MapPath(".") & "\Private\sf.pem", "\\", "\")
	keyfile = Replace(keyfile, "ssl\SFLib", "ssl")	'added since testing will add the SFLib directory

	'Prepare the card
	Dim pstrAuthorizationAmount
	Dim pstrCardNumber
	Dim pstrCardExpiration_Month
	Dim pstrCardExpiration_Year
	
	pstrAuthorizationAmount = FormatNumber(Replace(sGrandTotal, ",", ""), 2)
		
	' Payment variables
	pstrCardNumber = Replace(sCustCardNumber," ","")
	pstrCardNumber = Replace(pstrCardNumber,"-","")
	
	pstrCardExpiration_Month = sCustCardExpiryMonth 
	pstrCardExpiration_Year = Right(sCustCardExpiryYear, 2)

	' Create an empty order
	Set order = Server.CreateObject("LpiCom_6_0.LPOrderPart")
	order.setPartName("order")
	' Create an empty part
	Set op = Server.CreateObject("LpiCom_6_0.LPOrderPart")          

	' Build 'orderoptions'
	' For a test, set result to GOOD, DECLINE, or DUPLICATE
	'Call op.put("result", "GOOD")

	If UCase(sMercType) = "AUTHCAPTURE" Then
		Call op.put("ordertype", "SALE")
	ElseIf sMercType = "AUTHONLY" Then
		Call op.put("ordertype", "PREAUTH")
	End If
	' add 'orderoptions to order
	Call order.addPart("orderoptions", op)

	' Build 'merchantinfo'
	Call op.clear()
	Call op.put("configfile", sLogin)
	' add 'merchantinfo to order
	Call order.addPart("merchantinfo", op)

	' Build transactiondetails
	Call op.clear()
	Call op.put("oid", CStr(iOrderID))
	Call op.put("transactionorigin", "ECI")
	Call op.put("ip", Request.ServerVariables("REMOTE_ADDR"))
	' add transactiondetails to order
	Call order.addPart("transactiondetails", op)

	' Build 'creditcard'
	Call op.clear()
	Call op.put("cardnumber", sCustCardNumber)
	Call op.put("cardexpmonth", sCustCardExpiryMonth)
	Call op.put("cardexpyear", sCustCardExpiryYear)
	
	'For CCV enabled
	'Non-U.S. issued cards may not support CCV
	Select Case sCustCountry
		Case "US", "CA"
			If Len(mstrPayCardCCV) > 0 Then
				Call op.put("cvmvalue", mstrPayCardCCV)
				Call op.put("cvmindicator", "provided")
			Else
				Call op.put("cvmvalue", "")
				Call op.put("cvmindicator", "not_provided")
			End If
		Case Else
			Call op.put("cvmvalue", "")
			Call op.put("cvmindicator", "not_provided")
	End Select
	' add 'creditcard to order
	Call order.addPart("creditcard", op)

	' Build 'billing'
	Call op.clear()
	Call op.put("name", sCustCardName)
	Call op.put("company", sCustCompany)
	Call op.put("address1", sCustAddress1)
	Call op.put("address2", sCustAddress2)
	Call op.put("city", sCustCity)
	Call op.put("state", sCustState)
	Call op.put("zip", sCustZip)
	Dim plngPos
	plngPos = Instr(sCustAddress1, " ")
	If plngPos > 0 Then
		Call op.put("addrnum", Left(sCustAddress1, plngPos - 1))
	Else
		Call op.put("addrnum", "")
	End If
	Call op.put("country", sCustCountry)
	Call op.put("phone", sCustPhone)
	Call op.put("fax", sCustFax)
	Call op.put("email", sCustEmail)

	' add 'billing to order
	Call order.addPart("billing", op)

	' Build 'payment'
	Call op.clear()
	Call op.put("chargetotal", pstrAuthorizationAmount)
	' add 'payment to order
	Call order.addPart("payment", op)

	' create transaction object  
	Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")

	' get outgoing XML from 'order' object

	outXml = order.toXML()

	' Call LPTxn
	If True Then
		Response.Write "<fieldset><legend>LinkPoint Request</legend>"
		Response.Write "host: " & host & "<BR>"
		Response.Write "keyfile: " & keyfile & "<BR>"
		Response.Write "<hr><textarea rows=35 cols=60>" & outXml & "</textarea>"
		Response.Write "</fieldset>"
		Response.Flush
	End If

	resp = LPTxn.send(keyfile, host, 1129, outXml)

	Set LPTxn = Nothing
	Set order = Nothing
	Set op   = Nothing

	'Now process the response
	Dim R_Time
	Dim R_Ref
	Dim R_Approved
	Dim R_Code
	Dim R_Authresr
	Dim R_Error
	Dim R_OrderNum
	Dim R_Message
	Dim R_AVS
	Dim R_Score
	Dim R_TDate
	Dim R_Tax
	Dim R_Shipping
	Dim R_FraudCode
	Dim R_ESD

	R_Time = ParseTag("r_time", resp)
	R_Ref = ParseTag("r_ref", resp)
	R_Approved = ParseTag("r_approved", resp)
	R_Code = ParseTag("r_code", resp)
	R_Authresr = ParseTag("r_authresronse", resp)
	R_Error = ParseTag("r_error", resp)
	R_OrderNum = ParseTag("r_ordernum", resp)
	R_Message = ParseTag("r_message", resp)
	R_AVS = ParseTag("r_avs", resp)
	R_Score = ParseTag("r_score", resp)
	R_TDate = ParseTag("r_tdate", resp)
	R_Tax = ParseTag("r_tax", resp)
	R_Shipping = ParseTag("r_shipping", resp)
	R_FraudCode = ParseTag("r_fraudCode", resp)
	R_ESD = ParseTag("esd", resp)
	
	If True Then
		Response.Write "<fieldset><legend>LinkPoint Response</legend>"
		Response.Write "r_time: " & R_Time & "<BR>"
		Response.Write "r_ref: " & R_Ref & "<BR>"
		Response.Write "r_approved: " & R_Approved & "<BR>"
		Response.Write "r_code: " & R_Code & "<BR>"
		Response.Write "r_authresronse: " & R_Authresr & "<BR>"
		Response.Write "r_error: " & R_Error & "<BR>"
		Response.Write "r_ordernum: " & R_OrderNum & "<BR>"
		Response.Write "r_message: " & R_Message & "<BR>"
		Response.Write "r_avs: " & R_AVS & "<BR>"
		Response.Write "r_score: " & R_Score & "<BR>"
		Response.Write "r_tdate: " & R_TDate & "<BR>"
		Response.Write "r_avs: " & R_AVS & "<BR>"
		Response.Write "r_tax: " & R_Tax & "<BR>"
		Response.Write "r_shipping: " & R_Shipping & "<BR>"
		Response.Write "r_fraudCode: " & R_FraudCode & "<BR>"
		Response.Write "esd: " & R_ESD & "<BR>"
		Response.Write "<hr><textarea rows=35 cols=60>" & resp & "</textarea>"
		Response.Write "</fieldset>"
		Response.Flush
	End If
	
	Select Case R_Approved
		Case "APPROVED"
			ProcResponse = "approve"
			iProcResponse = 1
			ProcMessage = R_Message
			ProcActionCode = R_Ref
			ProcResponseCode = R_Approved
			ProcAuth = R_Code
			ProcAuthCode = Mid(ProcAuth, 1, 6)
			ProcRefCode = Mid(ProcAuth, 7, 10)
			
			ProcErrMessage = R_Error
			ProcCustNumber = R_OrderNum
			
			ProcAvsCode = R_AVS          'The Address Verification System (AVS) response for this transaction. The first character indicates whether the contents of the addrnum tag match the address number on file for the billing address. The second character indicates whether the billing zip code matches the billing records. The third character is the raw AVS response from the card-issuing bank. The last character indicates whether the cvmvalue was correct and may be "M" for Match, "N" for No Match, or "Z" if the match could not determined.
			If Len(ProcAvsCode) = 4 Then
				ProcAvsMsg = AVSMsg(CStr(Mid(ProcAvsCode, 3, 1)))
			Else
				ProcAvsMsg = AVSMsg(ProcAvsCode)
			End if
		Case "DUPLICATE"
			ProcResponse = R_Approved
			iProcResponse = 1
			ProcMessage = R_Message
			ProcActionCode = R_Ref
			ProcResponseCode = R_Approved
			ProcAuth = R_Code
			ProcAuthCode = Mid(ProcAuth, 1, 6)
			ProcRefCode = Mid(ProcAuth, 7, 10)
			
			ProcErrMessage = R_Error
			ProcCustNumber = R_OrderNum
		
			ProcAvsCode = R_AVS          'The Address Verification System (AVS) response for this transaction. The first character indicates whether the contents of the addrnum tag match the address number on file for the billing address. The second character indicates whether the billing zip code matches the billing records. The third character is the raw AVS response from the card-issuing bank. The last character indicates whether the cvmvalue was correct and may be "M" for Match, "N" for No Match, or "Z" if the match could not determined.
			If Len(ProcAvsCode) = 4 Then
				ProcAvsMsg = AVSMsg(CStr(Mid(ProcAvsCode, 3, 1)))
			Else
				ProcAvsMsg = AVSMsg(ProcAvsCode)
			End if
		Case "DECLINED"
			ProcResponse = R_Approved
			iProcResponse = 1
			ProcMessage = R_Message
			ProcActionCode = R_Ref
			ProcResponseCode = R_Approved
			ProcAuth = R_Code
			ProcAuthCode = Mid(ProcAuth, 1, 6)
			ProcRefCode = Mid(ProcAuth, 7, 10)
			
			ProcErrMessage = R_Error
			ProcCustNumber = R_OrderNum
		
			ProcAvsCode = R_AVS          'The Address Verification System (AVS) response for this transaction. The first character indicates whether the contents of the addrnum tag match the address number on file for the billing address. The second character indicates whether the billing zip code matches the billing records. The third character is the raw AVS response from the card-issuing bank. The last character indicates whether the cvmvalue was correct and may be "M" for Match, "N" for No Match, or "Z" if the match could not determined.
			If Len(ProcAvsCode) = 4 Then
				ProcAvsMsg = AVSMsg(CStr(Mid(ProcAvsCode, 3, 1)))
			Else
				ProcAvsMsg = AVSMsg(ProcAvsCode)
			End if
		Case "FRAUD"
			ProcResponse = R_Approved
			iProcResponse = 1
			ProcMessage = R_Message
			ProcActionCode = R_Ref
			ProcResponseCode = R_Approved
			ProcAuth = R_Code
			ProcAuthCode = Mid(ProcAuth, 1, 6)
			ProcRefCode = Mid(ProcAuth, 7, 10)
			
			ProcErrMessage = R_Error
			ProcCustNumber = R_OrderNum
		
			ProcAvsCode = R_AVS          'The Address Verification System (AVS) response for this transaction. The first character indicates whether the contents of the addrnum tag match the address number on file for the billing address. The second character indicates whether the billing zip code matches the billing records. The third character is the raw AVS response from the card-issuing bank. The last character indicates whether the cvmvalue was correct and may be "M" for Match, "N" for No Match, or "Z" if the match could not determined.
			If Len(ProcAvsCode) = 4 Then
				ProcAvsMsg = AVSMsg(CStr(Mid(ProcAvsCode, 3, 1)))
			Else
				ProcAvsMsg = AVSMsg(ProcAvsCode)
			End if
		Case Else
			ProcResponse = R_Approved
			iProcResponse = 1
			ProcResponse = "fail"
			ProcErrMessage = "Your transaction was NOT successful. Please verify your payment information and try again."
	End Select
	
	'Call setResponse("LinkPoint",iOrderID,ProcCustNumber,"",ProcAvsCode,ProcAvsMsg,ProcActionCode,ProcAuthCode,ProcResponseCode,ProcErrMessage,iProcResponse)
	LinkPoint = ProcErrMessage  
	
End Function

Function ParseTag(byVal tag , byVal rsp)

Dim sb
Dim idxSt, idxEnd 'As Integer

	sb = "<" & tag & ">"
	idxSt = -1
	idxEnd = -1
	idxSt = InStr(rsp,sb)
	
	If 0 = idxSt Then
		ParseTag = ""
		Exit Function
	End If
	
	idxSt = idxSt + Len(sb)
	sb = "</" & tag & ">"
	idxEnd = InStr(idxSt, rsp,sb)
	If 0 = idxEnd Then
		ParseTag = ""
		Exit Function
	End If
	
	ParseTag = Mid(rsp, idxSt, (idxEnd - idxSt))
	
End Function	'ParseTag
%>