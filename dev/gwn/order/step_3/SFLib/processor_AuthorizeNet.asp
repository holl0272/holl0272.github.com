<%
Function AuthNet(proc_live, iType)
	'AuthNet = AuthNet31(proc_live, iType)
	AuthNet = AuthNet_OLD(proc_live, iType)
End Function	'AuthNet

'************************************************************************************************************

Function AuthNetAvsCodeDefinition(byVal strProcAvsCode)

Dim ProcAvsCode

	Select Case strProcAvsCode
		Case "A": ProcAvsCode = "Address (Street) matches, ZIP does not"
		Case "B": ProcAvsCode = "Address information not provided for AVS check"
		Case "E": ProcAvsCode = "AVS error"
		Case "G": ProcAvsCode = "Non-U.S. Card Issuing Bank"
		Case "N": ProcAvsCode = "No Match on Address (Street) or ZIP"
		Case "P": ProcAvsCode = "AVS not applicable for this transaction"
		Case "R": ProcAvsCode = "Retry – System unavailable or timed out"
		Case "S": ProcAvsCode = "Service not supported by issuer"
		Case "U": ProcAvsCode = "Address information is unavailable"
		Case "W": ProcAvsCode = "9 digit ZIP matches, Address (Street) does not"
		Case "X": ProcAvsCode = "Address (Street) and 9 digit ZIP match"
		Case "Y": ProcAvsCode = "Address (Street) and 5 digit ZIP match"
		Case "Z": ProcAvsCode = "5 digit ZIP matches, Address (Street) does not"
		Case Else: ProcAvsCode = "Unknown Code"
	End Select
	
	AuthNetAvsCodeDefinition = ProcAvsCode
	
End Function	'AuthNetAvsCodeDefinition

'************************************************************************************************************

Function AuthNetCCVCodeDefinition(byVal strCCVCode)

Dim ProcCCVCode

	Select Case strCCVCode
		Case "M": ProcCCVCode = "Match"
		Case "N": ProcCCVCode = "No Match"
		Case "P": ProcCCVCode = "Not Processed"
		Case "S": ProcCCVCode = "Should have been present"
		Case "U": ProcCCVCode = "Issuer unable to process request"
		Case Else: ProcCCVCode = "Unknown Code"
	End Select
	
	AuthNetCCVCodeDefinition = ProcCCVCode
	
End Function	'AuthNetCCVCodeDefinition

'************************************************************************************************************

Function AuthNetResponseCodeDefinition(byVal strResponseCode)

Dim ResponseCode

	Select Case CStr(strResponseCode)
		Case "1": ResponseCode = "Approved"
		Case "2": ResponseCode = "Declined"
		Case "3": ResponseCode = "Error"
		Case Else: ResponseCode = "Unknown Code"
	End Select
	
	AuthNetResponseCodeDefinition = ResponseCode
	
End Function	'AuthNetResponseCodeDefinition

'************************************************************************************************************

Function AuthNet31(byVal proc_live, byVal iType)

Dim arResponse
Dim description
Dim i
Dim iProcResponse
Dim MType
Dim objHTTP
Dim pblnLocalDebug
Dim ProcActionCode
Dim ProcAddlData
Dim ProcAuthCode
Dim ProcAvsCode
Dim ProcErrCode
Dim ProcErrMsg
Dim ProcMerchNumber
Dim ProcMessage
Dim ProcRefCode
Dim pstrAuthorizationAmount
Dim pstrCardNumber
Dim pstrCardExpiration
Dim sCustAddress
Dim sCustName
Dim sErrorMessage
Dim sShipCustAddress
Dim sShipCustName
Dim sTstRqst
Dim strPost
Dim strReturn

	description = C_STORENAME & " Order " & iOrderID

	'custom sections:
	'Session("debugCC") = "True"
	'Session("debugCC") = ""
	pblnLocalDebug = Len(Session("debugCC")) > 0
	
	If sMercType = "authonly" Then
		MType = "AUTH_ONLY"
	ElseIf sMercType = "authcapture" Then
		MType = "AUTH_CAPTURE"
	End If
	
    'Response.Write "proc_live = "& proc_live 
	If proc_live = 1 Then
		sTstRqst = "false"	
	ElseIf proc_live = 0 Then
		sTstRqst = "true"	
	End If	
	
	' Customer Info
	sCustName		= sCustFirstName & " " & sCustLastName
	sCustAddress	= sCustAddress1 & ";" & sCustAddress2
	sShipCustName		= sShipCustFirstName & " " & sShipCustLastName
	sShipCustAddress	= sShipCustAddress1 & ";" & sShipCustAddress2
	
	' Order Variables
	pstrAuthorizationAmount			= FormatNumber(REPLACE(sGrandTotal,",",""),2)
		
	' Payment variables
	pstrCardNumber			= Replace(sCustCardNumber," ","")
	pstrCardNumber			= Replace(pstrCardNumber,"-","")
	pstrCardExpiration			= Replace(sCustCardExpiry,"/","")

'				& "&x_Password=" & sPassword _
	' post string
	If iType = 1 Then
		strPost = "x_login=" & sLogin _
				& "&x_tran_key=" & sPassword _
				& "&x_method=cc&x_type=" & mtype _
				& "&x_amount=" & pstrAuthorizationAmount _
				& "&x_invoice_num=" & iOrderID _
				& "&x_card_num=" & pstrCardNumber _
				& "&x_card_code=" & mstrPayCardCCV _
				& "&x_exp_date=" & pstrCardExpiration _
				& "&x_version=3.1" _
				& "&x_relay_response=FALSE" _
				& "&x_cust_id=" & visitorLoggedInCustomerID _
				& "&x_description=" & description _
				& "&x_first_name=" & sCustfirstname _
				& "&x_last_name=" & sCustlastname _
				& "&x_company=" & sCustcompany _
				& "&x_address=" & sCustAddress _
				& "&x_city=" & sCustcity _
				& "&x_state=" & sCuststate _
				& "&x_zip=" & sCustzip _
				& "&x_country=" & sCustcountry _
				& "&x_phone=" & sCustphone _
				& "&x_fax=" & sCustfax _
				& "&x_ship_to_first_name=" & sShipcustfirstname _
				& "&x_ship_to_last_name=" & sShipCustlastname _
				& "&x_ship_to_company=" & sShipCustcompany _
				& "&x_ship_to_address=" & sShipCustAddress _
				& "&x_ship_to_city=" & sShipCustcity _
				& "&x_ship_to_state=" & sShipCuststate _
				& "&x_ship_to_zip=" & sShipCustzip _
				& "&x_ship_to_country=" & sShipCustcountry _
				& "&x_email=" & sCustEmail _
				& "&x_delim_data=TRUE&x_delim_char=|"
	ElseIf iType = 2 Then
		strPost = "x_login=" & sLogin _
				& "&x_tran_key=" & sPassword _
				& "&x_method=ECHECK&x_type=AUTH_CAPTURE&x_amount=" & pstrAuthorizationAmount _
				& "&x_invoice_num=" & iOrderID _
				& "&x_bank_aba_code=" & iRoutingNumber _
				& "&x_bank_acct_num=" & iCheckingAccountNumber _
				& "&x_bank_acct_type=CHECKING&x_bank_name" & sBankname & "x_bank_acct_name" & sCustName & "&x_version=3.1&x_relay_response=FALSE" _
				& "&x_cust_id=" & visitorLoggedInCustomerID _
				& "&x_description=storefront web store order&&x_first_name=" & sCustfirstname _
				& "&x_last_name=" & sCustlastname _
				& "&x_company=" & sCustcompany _
				& "&x_address=" & sCustAddress _
				& "&x_city=" & sCustcity _
				& "&x_state=" & sCuststate _
				& "&x_zip=" & sCustzip _
				& "&x_country=" & sCustcountry _
				& "&x_phone=" & sCustphone _
				& "&x_fax=" & sCustfax _
				& "&x_ship_to_first_name=" & sShipcustfirstname _
				& "&x_ship_to_last_name=" & sShipCustlastname _
				& "&x_ship_to_company=" & sShipCustcompany _
				& "&x_ship_to_address=" & sShipCustAddress _
				& "&x_ship_to_city=" & sShipCustcity _
				& "&x_ship_to_state=" & sShipCuststate _
				& "&x_ship_to_zip=" & sShipCustzip _
				& "&x_ship_to_country=" & sShipCustcountry _
				& "&x_email=" & sCustEmail _
				& "&x_delim_data=TRUE&x_delim_char=|"
	End If
        
	If pblnLocalDebug Then Response.Write "strPost: " & strPost & "<br>"
	
    Set objHTTP = Server.CreateObject("MSXML2.XMLHTTP")
    With objHTTP
		.Open "post", "https://secure.authorize.net/gateway/transact.dll", False
		.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		.send strPost
		strReturn = .responseText
		ProcErrMsg = strReturn	'added by Sandshot Software to store retrieval
    End With	'objHTTP
    Set objHTTP = Nothing

    arResponse = Split(strReturn, "|")
    
	If pblnLocalDebug Then
		Response.Write "Response: " & ProcErrMsg & "<br>"
		for i=0 to UBound(arResponse)
			Response.write "i=" & i & "&nbsp;&nbsp;" & arResponse(i) & "<BR>"
		next
		Response.end
    End If
    
    If UBound(arResponse) > 0 Then
        
        ProcActionCode = arResponse(0)
        ProcAddlData = arResponse(1)
        ProcErrCode = arResponse(2)
        ProcMessage = arResponse(3)
        procAuthCode = arResponse(4)
        ProcAvsCode = arResponse(5)
        ProcRefCode = arResponse(6)
        ProcMerchNumber = arResponse(37)
        
        Select Case CStr(ProcActionCode)
			Case "1":
				iProcResponse = 1
			Case "2":
			    iProcResponse = 0
			    sErrorMessage = arResponse(3)
			Case "3":
			    iProcResponse = 0
			    sErrorMessage = arResponse(3)
			Case Else
                iProcResponse = 0
                sErrorMessage = "There was an error on the transaction processing network. Please check the account and resubmit the order.<BR> Additional Information from the Processor: " & ProcMessage 			
		End Select
        
 		ProcActionCode = AuthNetResponseCodeDefinition(ProcActionCode) & " (" & ProcActionCode & ")"
		ProcAvsCode = getAVSMessage("authorizenet", ProcAvsCode) & " (" & ProcAvsCode & ")"
		ProcAvsCode = ProcAvsCode & "|" & getCVVMessage("authorizenet", arResponse(38)) & " (" & arResponse(38) & ")"
		
	Else
		' no connection
        iProcResponse = 0
        sErrorMessage = "no connection error"
    End If
             
    'Save the response
	ProcAvsCode = Array(ProcAvsCode, pstrAuthorizationAmount)
	Call setResponse("authorizenet", iOrderID, "", ProcMerchNumber , ProcAvsCode, ProcMessage, ProcErrCode, procAuthCode, ProcRefCode, ProcErrMsg, iProcResponse)
	
    AuthNet31 = sErrorMessage
    
End Function	'AuthNet31

'************************************************************************************************************

'---------------------------------------------------------------------
'   AuthorizeNet Send Sub-Routine 
'	Com Object 1.0
'---------------------------------------------------------------------
Function AuthNet_OLD(proc_live,iType)

Dim AuthObj
Dim arResponse
Dim description
Dim i
Dim iProcResponse
Dim MType
Dim pblnLocalDebug
Dim ProcActionCode
Dim ProcAddlData
Dim ProcAuthCode
Dim ProcAvsCode
Dim ProcErrCode
Dim ProcErrMsg
Dim ProcMerchNumber
Dim ProcMessage
Dim ProcRefCode
Dim pstrAuthorizationAmount
Dim pstrCardNumber
Dim pstrCardExpiration
Dim pstrShipping
Dim pstrTax
Dim sCustAddress
Dim sCustName
Dim sErrorMessage
Dim sShipCustAddress
Dim sShipCustName
Dim strPost
Dim sTstRqst
Dim strReturn

	description = C_STORENAME & " Order " & iOrderID

	If sMercType = "authonly" Then MType = "AUTH_ONLY" 
	If sMercType = "authcapture" Then MType = "AUTH_CAPTURE" 
		
	sCustCardName = sCustCardName
	iOrderID =iOrderID
	sShipInstructions = sShipInstructions
    
    'Response.Write "proc_live = "& proc_live 
	If proc_live = 1 Then
		sTstRqst = "false"	
	ElseIf proc_live = 0 Then
		sTstRqst = "true"	
	End If	
	
	' Customer Info
	sCustName = sCustFirstName & " " & sCustLastName
	sCustAddress = sCustAddress1 & ";" & sCustAddress2
	sShipCustName = sShipCustFirstName & " " & sShipCustLastName
	sShipCustAddress = sShipCustAddress1 & ";" & sShipCustAddress2
	
	' Order Variables
	pstrShipping = (cDbl(sHandling) + cDbl(sShipping))
	pstrTax = (cDbl(sTotalSTax) + cDbl(sTotalCTax))
	pstrAuthorizationAmount = FormatNumber(REPLACE(sGrandTotal,",",""),2)
		
	' Payment variables
	pstrCardNumber			= Replace(sCustCardNumber," ","")
	pstrCardNumber			= Replace(pstrCardNumber,"-","")
	pstrCardExpiration			= Replace(sCustCardExpiry,"/","")

	' Post String
	If iType = 1 Then
		strPost = "x_Invoice_Num="&iOrderID&",x_Login="&sLogin&",x_Amount="&pstrAuthorizationAmount&",x_Freight="&pstrShipping&",x_Tax="&pstrTax&",x_Card_Num="&sCustCardNumber&",x_Exp_Date="&sCustCardExpiry&",x_Card_Code="&mstrPayCardCCV&",x_Password="&sPassword&",x_Method=CC,x_Type="&MType&",x_Cust_ID="&visitorLoggedInCustomerID&",x_Test_Request="&sTstRqst&",x_Cust_ID="&visitorLoggedInCustomerID&",x_Description=" & description & ",x_First_Name="&sCustFirstName&",x_Last_Name="&sCustLastName&",x_Company="&sCustCompany&",x_Address="&sCustAddress&",x_City="&sCustCity&",x_State="&sCustState&",x_Zip="&sCustZip&",x_Country="&sCustCountry&",x_Phone="&sCustPhone&",x_Fax="&sCustFax&",x_Ship_To_First_Name="&sShipCustFirstName&",x_Ship_To_Last_Name="&sShipCustLastName&",x_Ship_To_Company="&sShipCustCompany&",x_Ship_To_Address="&sShipCustAddress&",x_Ship_To_City="&sShipCustCity&",x_Ship_To_State="&sShipCustState&",x_Ship_To_Zip="&sShipCustZip&",x_Ship_To_Country="&sShipCustCountry&",x_Email="&sCustEmail
	ElseIf iType = 2 Then
		strPost = "x_Invoice_Num="&iOrderID&",x_Login="&sLogin&",x_Amount="&pstrAuthorizationAmount&",x_Freight="&pstrShipping&",x_Tax="&pstrTax&"x_Card_Num="&sCustCardNumber&",x_Exp_Date="&sCustCardExpiry&",x_Card_Code="&mstrPayCardCCV&",x_Password="&sPassword & ",x_Method=ECHECK,x_Type=" & sMercType & ",x_Bank_Name=" & sBankName & ",x_Bank_ABA_Code=" & iRoutingNumber & ",x_Bank_Acct_Num=" & iCheckingAccountNumber & ",x_Test_Request=" & sTstRqst & ",x_Cust_ID="&visitorLoggedInCustomerID&",x_Description=" & description & ",x_First_Name="&sCustFirstName&",x_Last_Name="&sCustLastName&",x_Company="&sCustCompany&",x_Address="&sCustAddress&",x_City="&sCustCity&",x_State="&sCustState&",x_Zip="&sCustZip&",x_Country="&sCustCountry&",x_Phone="&sCustPhone&",x_Fax="&sCustFax&",x_Ship_To_First_Name="&sShipCustFirstName&",x­_Ship_To_Last_Name="&sShipCustLastName&",x_Ship_To_Company="&sShipCustCompany&",x_Ship_To_Address="&sShipCustAddress&",x_Ship_To_City="&sShipCustCity&",x_Ship_To_State="&sShipCustState&",x_Ship_To_Zip="&sShipCustZip&",x_Ship_To_Country="&sShipCustCountry&",x_Email="&sCustEmail
	End If

	On Error Resume Next
	Set AuthObj = Server.CreateObject("AuthNetSSLConnect.SSLPost")
	If Err.number = 0 Then
		AuthObj.doSSLPost strPost
		If AuthObj.ErrorCode = 0 Then	
			If AuthObj.NumFields > 1 Then
					
				ProcActionCode = AuthObj.GetField(1)
				ProcAddlData = AuthObj.GetField(2)
				ProcErrCode = AuthObj.GetField(3)
				ProcMessage = AuthObj.GetField(4)
				ProcAuthCode = AuthObj.GetField(5)
				ProcAvsCode = AuthObj.GetField(6)
				ProcRefCode = AuthObj.GetField(7)

				Dim trnsrspAuthorizationAmount
				trnsrspAuthorizationAmount = AuthObj.GetField(10)
				
				If ProcActionCode = 1 Then
					iProcResponse = 1				
					
				ElseIf ProcActionCode = 2 Then
					iProcResponse = 0
					sErrorMessage = "The transaction has been declined.  Please check the account and resubmit, try another account or contact the card issuing bank.  Thank You.<br> Additional Information from the Processor: " & ProcMessage
					
				ElseIf ProcActionCode = 3 Then	 
					iProcResponse = 0
					sErrorMessage = "There was an error on the transaction processing network. Please check the account and resubmit the order.<br> Additional Information from the Processor: " & ProcMessage
					
				End If

				ProcAvsCode = getAVSMessage("authorizenet", ProcAvsCode) & " (" & ProcAvsCode & ")"
				
			Else	
				' No connection
				iProcResponse = 0
				sErrorMessage = "No Connection Error"	
			End If

		Else
			Select Case AuthObj.ErrorCode
			Case -1
				sErrorMessage = "A connection could not be established with the authorization network."
			Case -2
				sErrorMessage = "A connection could not be established with the authorization network."
			Case -3
				sErrorMessage = "A connection could not be established with the authorization."
			Case Else 
				sErrorMessage = "An error occured during processing."
			End Select
		End If					
	
	Else
		iProcResponse = 0
		sErrorMessage = "AuthNetSSLConnect does not appear to be installed correctly."	
	End If
	closeobj(AuthObj)
	
	On Error Goto 0
	
	'ProcAvsCode = Array(ProcAvsCode, trnsrspAuthorizationAmount)
	ProcAvsCode = Array(ProcAvsCode, pstrAuthorizationAmount)
	' Write to payment response table
	Call setResponse("AuthorizeNet",iOrderID,"","",ProcAvsCode,ProcMessage,ProcErrCode,ProcAuthCode,ProcRefCode,ProcErrMsg,iProcResponse)
	
	AuthNet_OLD = sErrorMessage
	
End Function	'AuthNet_OLD
%>