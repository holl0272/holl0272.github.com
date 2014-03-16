<%
'************************************************************************************************************

Function OrbitalAvsCodeDefinition(byVal strProcAvsCode)

Dim ProcAvsCode

	Select Case strProcAvsCode
		Case "1": ProcAvsCode = "No address supplied"
		Case "2": ProcAvsCode = "Bill-to address did not pass Auth Host edit checks"
		Case "3": ProcAvsCode = "AVS not performed"
		Case "4", "R": ProcAvsCode = "Issuer does not participate in AVS"
		Case "5": ProcAvsCode = "Edit-error - AVS data is invalid"
		Case "6": ProcAvsCode = "System unavailable or time-out"
		Case "7": ProcAvsCode = "Address information unavailable"
		Case "8": ProcAvsCode = "Transaction Ineligible for AVS"
		Case "9": ProcAvsCode = "Zip Match / Zip4 Match / Locale match"
		Case "A": ProcAvsCode = "Zip Match / Zip 4 Match / Locale no match"
		Case "B": ProcAvsCode = "Zip Match / Zip 4 no Match / Locale match"
		Case "C": ProcAvsCode = "Zip Match / Zip 4 no Match / Locale no match"
		Case "D": ProcAvsCode = "Zip No Match / Zip 4 Match / Locale match"
		Case "E": ProcAvsCode = "Zip No Match / Zip 4 Match / Locale no match"
		Case "F": ProcAvsCode = "Zip No Match / Zip 4 No Match / Locale match"
		Case "G": ProcAvsCode = "No match at all"
		Case "H": ProcAvsCode = "Zip Match / Locale match"
		Case "J": ProcAvsCode = "Issuer does not participate in Global AVS"
		Case "JA": ProcAvsCode = "International street address and postal match"
		Case "JB": ProcAvsCode = "International street address match. Postal code not verified."
		Case "JC": ProcAvsCode = "International street address and postal code not verified."
		Case "JD": ProcAvsCode = "International postal code match. Street address not verified."
		Case "X": ProcAvsCode = "Zip Match / Zip 4 Match / Address Match"
		Case "Z": ProcAvsCode = "Zip Match / Locale no match"
		Case "": ProcAvsCode = "Not applicable (non-Visa)"
		Case Else: ProcAvsCode = "Unknown Code"
	End Select
	
	OrbitalAvsCodeDefinition = ProcAvsCode
	
End Function	'OrbitalAvsCodeDefinition

'************************************************************************************************************

Function OrbitalCCVCodeDefinition(byVal strCCVCode)

Dim ProcCCVCode

	Select Case strCCVCode
		Case "M": ProcCCVCode = "Match"
		Case "N": ProcCCVCode = "No Match"
		Case "P": ProcCCVCode = "Not Processed"
		Case "S": ProcCCVCode = "Should have been present"
		Case "U": ProcCCVCode = "Unsupported by issuer"
		Case "I", "Y" : ProcCCVCode = "Invalid"
		Case "": ProcCCVCode = "Not applicable (non-Visa)"
		Case Else: ProcCCVCode = "Unknown Code"
	End Select
	
	OrbitalCCVCodeDefinition = ProcCCVCode
	
End Function	'OrbitalCCVCodeDefinition

'************************************************************************************************************

Function Orbital(byVal proc_live)

Dim iProcResponse
Dim pblnLoadError
Dim plngAuthorizationAmount
Dim pobjNode
Dim pobjXMLDoc
Dim pobjXMLHTTP
Dim pstrCardExpiration
Dim pstrCountryCode
Dim pstrData
Dim pstrResult

Dim pstrAVSCode
Dim pstrCCVCode
Dim pstrResponseCode
Dim pstrReferenceNumber
Dim pblnSuccess
Dim pstrErrorMessage
Dim pstrRespCode
Dim orderAVSREQ
Dim sTr_Type

    If trim(sPaymentServer) = "" or isNull(sPaymentServer) Then sPaymentServer = "https://orbitalvar1.paymentech.net/"
    
	' Testing or live mode
	If proc_live = 1 Then
		orderAVSREQ = "1"
	ElseIf proc_live = 0 Then
		orderAVSREQ = "4"
	End If
	
	If UCase(sMercType) = "AUTHCAPTURE" Then
		sTr_Type = "AC"
	ElseIf sMercType = "AUTHONLY" Then
		sTr_Type = "A"
	End If
	sTr_Type = "A"
	
	'Amount is passed in cents
	plngAuthorizationAmount = FormatNumber(sGrandTotal * 100, 0, 0, 0, 0)
	
	If Len(sCustCardExpiryYear) > 2 Then
		pstrCardExpiration = sCustCardExpiryMonth & Right(sCustCardExpiryYear, 2)
	Else
		pstrCardExpiration = sCustCardExpiryMonth & sCustCardExpiryYear
	End If
	
	Select Case UCase(sCustCountry)
		Case "US", "CA", "GB", "UK"
			pstrCountryCode = UCase(sCustCountry)
		Case Else
			pstrCountryCode = ""
	End Select
	
	'For Testing
	sLogin = "700000001973"

	pstrData = "<?xml version=""1.0"" encoding=""UTF-8"" ?>" & vbcrlf _
				& "<Request>" & vbcrlf _
				& "	<AC>" & vbcrlf _
				& "		<CommonData>" & vbcrlf _
				& "			<CommonMandatory AuthOverrideInd=""N"" LangInd=""00"" CardHolderAttendanceInd=""01"" HcsTcsInd=""T"" TxCatg=""7"" MessageType=""" & sTr_Type & """ Version=""2"" TzCode=""705"">" & vbcrlf _
				& "				<AccountNum AccountTypeInd=""91"">" & sCustCardNumber & "</AccountNum>" & vbcrlf _
				& "				<POSDetails POSEntryMode=""01"" />" & vbcrlf _
				& "				<MerchantID>" & sLogin & "</MerchantID>" & vbcrlf _
				& "				<TerminalID TermEntCapInd=""05"" CATInfoInd=""06"" TermLocInd=""01"" CardPresentInd=""N"" POSConditionCode=""59"" AttendedTermDataInd=""01"">001</TerminalID>" & vbcrlf _
				& "				<BIN>000002</BIN>" & vbcrlf _
				& "				<OrderID>" & iOrderID & "</OrderID>" & vbcrlf _
				& "				<AmountDetails>" & vbcrlf _
				& "					<Amount>" & plngAuthorizationAmount & "</Amount>" & vbcrlf _
				& "				</AmountDetails>" & vbcrlf _
				& "				<TxTypeCommon TxTypeID=""G"" />" & vbcrlf _
				& "				<Currency CurrencyCode=""840"" CurrencyExponent=""2"" />" & vbcrlf _
				& "				<CardPresence>" & vbcrlf _
				& "					<CardNP>" & vbcrlf _
				& "						<Exp>" & pstrCardExpiration & "</Exp>" & vbcrlf _
				& "					</CardNP>" & vbcrlf _
				& "				</CardPresence>" & vbcrlf _
				& "				<TxDateTime />" & vbcrlf _
				& "			</CommonMandatory>" & vbcrlf _
				& "			<CommonOptional>" & vbcrlf _
				& "				<Comments />" & vbcrlf _
				& "				<ShippingRef />" & vbcrlf _
				& "				<CardSecVal CardSecInd=""1"">" & mstrPayCardCCV & "</CardSecVal>" & vbcrlf _
				& "				<ECommerceData ECSecurityInd=""07"">" & vbcrlf _
				& "					<ECOrderNum>" & iOrderID & "</ECOrderNum>" & vbcrlf _
				& "				</ECommerceData>" & vbcrlf _
				& "			</CommonOptional>" & vbcrlf _
				& "		</CommonData>" & vbcrlf _
				& "		<Auth>" & vbcrlf _
				& "			<AuthMandatory FormatInd=""H"" />" & vbcrlf _
				& "			<AuthOptional>" & vbcrlf _
				& "				<AVSextended>" & vbcrlf _
				& "					<AVSname>" & sCustCardName & "</AVSname>" & vbcrlf _
				& "					<AVSaddress1>" & sCustAddress1 & "</AVSaddress1>" & vbcrlf _
				& "					<AVSaddress2>" & sCustAddress2 & "</AVSaddress2>" & vbcrlf _
				& "					<AVScity>" & sCustCity & "</AVScity>" & vbcrlf _
				& "					<AVSstate>" & sCustState & "</AVSstate>" & vbcrlf _
				& "					<AVSzip>" & sCustZip & "</AVSzip>" & vbcrlf _
				& "					<AVScountryCode>" & pstrCountryCode & "</AVScountryCode>" & vbcrlf _
				& "				</AVSextended>" & vbcrlf _
				& "			</AuthOptional>" & vbcrlf _
				& "		</Auth>" & vbcrlf _
				& "		<Cap>" & vbcrlf _
				& "			<CapMandatory>" & vbcrlf _
				& "				<EntryDataSrc>02</EntryDataSrc>" & vbcrlf _
				& "			</CapMandatory>" & vbcrlf _
				& "			<CapOptional />" & vbcrlf _
				& "		</Cap>" & vbcrlf _
				& "	</AC>" & vbcrlf _
				& "</Request>"

    Set pobjXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    With pobjXMLHTTP
		.Open "post", sPaymentServer, False
		.setRequestHeader "MIME-Version", "1.0"
		.setRequestHeader "Content-Type", "application/PTI26"
		.setRequestHeader "Content-length", Len(pstrData)
		.setRequestHeader "Content-transfer-encoding", "text"
		.setRequestHeader "Request-number", "1"
		.setRequestHeader "Document-type", "Request"
		'.setRequestHeader "Interface-Version", "1"
		
		On Error Resume Next
		.send pstrData
		
		If Err.number <> 0 Then
			'Retry to alternate address
			Err.Clear
			.Open "post", "https://orbital2.paymentech.net/", False
			.send pstrData
		End If
		On Error Goto 0
		
		'Response.Write "<fieldset><legend>Orbital Error</legend>" & Server.HTMLEncode(.responseXML) & "</fieldset>"
		pstrResult = .responseText
		
    End With	'pobjXMLHTTP
    Set pobjXMLHTTP = Nothing

	'pstrResult = Replace(pstrResult, "Response_PTI26.dtd", "http://localhost/MasterTemplate/ssl/SFLib/Response_PTI26.dtd")
	pstrResult = Replace(pstrResult, "<!DOCTYPE Response SYSTEM ""Response_PTI26.dtd"">", "")
	
	Set pobjXMLDoc = CreateObject("MSXML2.DOMDocument.3.0")
	'Set pobjXMLDoc = CreateObject("MSXML.DOMDocument")
	With pobjXMLDoc
		.async = false
		.resolveExternals = False

		If .loadXML(pstrResult) Then

'On Error Resume Next
			pblnSuccess = True
			Set pobjNode = .SelectSingleNode("Response/QuickResponse/ProcStatus")
			If pobjNode is Nothing Then
				Set pobjNode = .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/ProcStatus")
				pstrResponseCode = CStr(pobjNode.Text)
				If pstrResponseCode <> "0" Then
					pblnSuccess = False
					pstrErrorMessage = .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/StatusMsg").Text
				End If
			Else
				'Error Case
				pstrResponseCode = CStr(pobjNode.Text)
				If pstrResponseCode <> "0" Then
					pblnSuccess = False
					pstrErrorMessage = .SelectSingleNode("Response/QuickResponse/StatusMsg").Text
				End If
			End If

If False Then
Response.Write "<fieldset><legend>Orbital Result</legend>" _
			 & "pstrResponseCode: " & pstrResponseCode & "<br />" _
			 & "pstrErrorMessage: " & pstrErrorMessage & "<br />" _
			 & "<br /><hr><br /><pre>" & Replace(Server.HTMLEncode(pstrData), "&gt;", "&gt;") & "</pre>" _
			 & "<br /><hr><br /><pre>" & Replace(Server.HTMLEncode(pstrResult), "&gt;", "&gt;") & "</pre>" _
			 & "</fieldset>"
			 '& "<br /><hr><br /><pre>" & Replace(pstrResult, vbcrlf, "<br />") & "</pre>" _
End If

			If pblnSuccess Then
				pblnSuccess = (CInt(.SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/ApprovalStatus").Text) = 1)
				pstrReferenceNumber = .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/TxRefNum").Text
				If Not .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/ResponseCodes/AVSRespCode") Is Nothing Then pstrAVSCode = .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/ResponseCodes/AVSRespCode").Text

				pstrRespCode = .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/ResponseCodes/RespCode").Text
				If Not .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/ResponseCodes/CVV2RespCode") Is Nothing Then pstrCCVCode = .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/ResponseCodes/CVV2RespCode").Text
				If Not pblnSuccess Then
					pstrErrorMessage = .SelectSingleNode("Response/ACResponse/CommonDataResponse/CommonMandatoryResponse/StatusMsg").Text
				End If
			End If	'pblnSuccess

			If pstrRespCode = "00" Then
				iProcResponse = 1
			Else 
				iProcResponse = 0
			End If 		

		Else
			Dim myErr
			Set myErr = .parseError
			Response.Write "<fieldset><legend>Orbital Error</legend>Error " & myErr.errorCode & ": " & myErr.reason & "</fieldset>"
			pblnLoadError = False
		End If
	End With	'pobjXMLDoc
    Set pobjXMLDoc = Nothing

	If False Then
		Response.Write "<fieldset><legend>Orbital Result</legend>" _
					& "Approved: " & CBool(iProcResponse) & "<br />" _
					& "Resp Code: " & pstrRespCode & "<br />" _
					& "AVS: " & OrbitalAvsCodeDefinition(pstrAVSCode) & " (" & pstrAVSCode & ")" & "<br />" _
					& "CCV: " & OrbitalCCVCodeDefinition(pstrCCVCode) & " (" & pstrCCVCode & ")" & "<br />" _
					& "CCV: " & pstrCCVCode & "<br />" _
					& "OrderID: " & iOrderID & "<br />" _
					& "Reference Number: " & pstrReferenceNumber & "<br />" _
					& "Amount: " & sGrandTotal & "<br />" _
					& "ErrorMessage: " & pstrErrorMessage & "<br />" _
					& "</fieldset>"
	End If

    Call setResponse("Orbital", iOrderID, pstrReferenceNumber, "", Array(OrbitalAvsCodeDefinition(pstrAVSCode), sGrandTotal, pstrCCVCode), OrbitalAvsCodeDefinition(pstrAVSCode), pstrRespCode, "", "", pstrErrorMessage, iProcResponse)
    
	Orbital = pstrErrorMessage
	
End Function	'Orbital

%>