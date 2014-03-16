<%

'**************************************************
'
'	Variable Declarations
'

Const pbytDaysBackToCheckOrders = 180
Const cdblFraudThreshold = 7
Const cdblOrderThresholdRisk = 1500
Const cdblOrderThresholdAbsolute = 5000
Const ssInternalOrderStatus_ClearedToProcess = 6
Const ssInternalOrderStatus_ManuallyProcess = 7
Const ssInternalOrderStatus_OrderedWithVendor = 7
Const ssInternalOrderStatus_Unread = 0

Const cstrDelimiter = "|"
Const cstrPOPaymentMethodName = "PhoneFax_Phone/Fax"
Const pstrOrderDetailLink = "<a href='ssOrderAdmin.asp?OrderID={orderID}&Action=ViewOrder&optDisplay=0&optDate_Filter=0&optPayment_Filter=0&optShipment_Filter=0'>{orderID}</a>"

'Order Processing Rules Array
Const enProcessingRule_Index = 0
Const enProcessingRule_Title = 1
Const enProcessingRule_ProcessingAction = 2
Const enProcessingRule_FraudScore = 3

'Order Processing Actions Array
Const enProcessingAction_Index = 0
Const enProcessingAction_Title = 1
Const enProcessingAction_FraudThreshold = 2
Const enProcessingAction_CeaseFurtherProcessing = 3

'Order Processing Items Array
Const enProcessingItem_OrderNumber = 0
Const enProcessingItem_OrderItem = 1
Const enProcessingItem_OrderItemMfg = 2
Const enProcessingItem_OrderItemVend = 3
Const enProcessingItem_Rules = 4
Const enProcessingItem_Actions = 5
Const enProcessingItem_Messages = 6
Const enProcessingItem_FraudScore = 7
Const enProcessingItem_OrderDetailID = 8
Const enProcessingItem_InternalOrderStatus = 9
Const enProcessingItem_PaymentMethod = 10

Const cblnSMDorLocal = False	'True False
Const cDebugMode = 0			'0 - None, 1 - Minimal, 2 - Extensive

Dim mblnProcessOrdersTestMode

mblnProcessOrdersTestMode = CBool(Len(Request.QueryString("TestMode")) = 0)
'use to manually set test mode
'mblnProcessOrdersTestMode = True	'True	False


'**************************************************
'
'	Functions
'
'**********************************************************

Sub addProcessingRule(byRef aryProcessingItem, byVal aryProcessingRule, byVal strText)
'Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)

	If Len(aryProcessingItem(enProcessingItem_Rules)) > 0 Then
		aryProcessingItem(enProcessingItem_Rules) = aryProcessingItem(enProcessingItem_Rules) & cstrDelimiter & aryProcessingRule(enProcessingRule_Index)
	Else
		aryProcessingItem(enProcessingItem_Rules) = aryProcessingRule(enProcessingRule_Index)
	End If

	If Len(aryProcessingItem(enProcessingItem_Actions)) > 0 Then
		aryProcessingItem(enProcessingItem_Actions) = aryProcessingItem(enProcessingItem_Actions) & cstrDelimiter & aryProcessingRule(enProcessingRule_ProcessingAction)
	Else
		aryProcessingItem(enProcessingItem_Actions) = aryProcessingRule(enProcessingRule_ProcessingAction)
	End If
	
	aryProcessingItem(enProcessingItem_FraudScore) = aryProcessingItem(enProcessingItem_FraudScore) + aryProcessingRule(enProcessingRule_FraudScore)
	
	If Len(aryProcessingItem(enProcessingItem_Messages)) > 0 Then
		aryProcessingItem(enProcessingItem_Messages) = aryProcessingItem(enProcessingItem_Messages) & cstrDelimiter & strText
	Else
		aryProcessingItem(enProcessingItem_Messages) = strText
	End If

	If cDebugMode = 2 Then
		Response.Write "<fieldset><legend>Processing Rules</legend>"
		Response.Write "Rules: " & aryProcessingItem(enProcessingItem_Rules) & "<br />"
		Response.Write "Actions: " & aryProcessingItem(enProcessingItem_Actions) & "<br />"
		Response.Write "FraudScore: " & aryProcessingItem(enProcessingItem_FraudScore) & "<br />"
		Response.Write "Messages: " & aryProcessingItem(enProcessingItem_Messages) & "<br />"
		Response.Write "</fieldset>"
	End If
	
End Sub	'addProcessingRule

'**********************************************************

Sub writeRules()

Dim i

	Response.Write "<fieldset><legend>Processing Rules</legend>"
	For i = 0 To UBound(maryProcessingRules)
		Response.Write i & ": " & maryProcessingRules(i)(enProcessingRule_Title) & "<br />"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Action: " & maryProcessingActions(maryProcessingRules(i)(enProcessingRule_ProcessingAction))(enProcessingAction_Title) & "<br />"
	Next 'i
	Response.Write "</fieldset>"
	
	Response.Write "<fieldset><legend>Processing Actions</legend>"
	For i = 0 To UBound(maryProcessingActions)
		Response.Write i & ": " & maryProcessingActions(i)(enProcessingAction_Title) & "<br />"
	Next 'i
	Response.Write "</fieldset>"
	
End Sub	'writeRules

'**********************************************************

Function incrementRule()
	mlngRuleCounter = mlngRuleCounter + 1
	incrementRule = mlngRuleCounter
End Function

'**********************************************************

Function updateProcessingArraySize(byRef ary, byVal index)
	If isArray(ary) Then
		If UBound(ary) < index Then
			ReDim Preserve ary(index)
		End If
	Else
		ReDim ary(index)
	End If
	updateProcessingArraySize = index
End Function

'**********************************************************

Function isBillingShippingAddressSame(byRef objRS)

Dim pstrBillingAddress
Dim pstrShippingAddress

	With objRS
		pstrBillingAddress = Trim(.Fields("custCompany").Value & "") _
						   & Trim(.Fields("custFirstName").Value & "") _
						   & Trim(.Fields("custMiddleInitial").Value & "") _
						   & Trim(.Fields("custLastName").Value & "") _
						   & Trim(.Fields("custAddr1").Value & "") _
						   & Trim(.Fields("custAddr2").Value & "") _
						   & Trim(.Fields("custCity").Value & "") _
						   & Trim(.Fields("custState").Value & "") _
						   & Trim(.Fields("custZip").Value & "") _
						   & Trim(.Fields("custCountry").Value & "")

		pstrShippingAddress= Trim(.Fields("cshpaddrShipCompany").Value & "") _
						   & Trim(.Fields("cshpaddrShipFirstName").Value & "") _
						   & Trim(.Fields("cshpaddrShipMiddleInitial").Value & "") _
						   & Trim(.Fields("cshpaddrShipLastName").Value & "") _
						   & Trim(.Fields("cshpaddrShipAddr1").Value & "") _
						   & Trim(.Fields("cshpaddrShipAddr2").Value & "") _
						   & Trim(.Fields("cshpaddrShipCity").Value & "") _
						   & Trim(.Fields("cshpaddrShipState").Value & "") _
						   & Trim(.Fields("cshpaddrShipZip").Value & "") _
						   & Trim(.Fields("cshpaddrShipCountry").Value & "")
						   
	End With
	
	isBillingShippingAddressSame = CBool(LCase(pstrBillingAddress) = LCase(pstrBillingAddress))

End Function	'isBillingShippingAddressSame

'**********************************************************

Function isPriorCustomer(byVal lngCustomerID)
'Prior customers are defined as a custor who place a previous order which has been shipped

Dim pblnResult
Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	pblnResult = False
	'look for repeat orders over 48 hours which have shipped
	pstrSQL = "SELECT sfOrders.orderDate" _
			& " FROM ssOrderManager INNER JOIN sfOrders ON ssOrderManager.ssorderID = sfOrders.orderID" _
			& " WHERE ((ssOrderManager.ssDateOrderShipped Is Not Null) AND (sfOrders.orderCustId=?) AND (sfOrders.orderDate<?))"

	Set pobjCmd  = Server.CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, lngCustomerID)
		.Parameters.Append .CreateParameter("orderDate", adDBTimeStamp, adParamInput, 16, DateAdd("h", -48, Now()))
		Set pobjRS = .Execute
		pblnResult = Not pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

	isPriorCustomer = pblnResult

End Function	'isPriorCustomer

'**********************************************************

Function AVS_Code(byRef strField)

Dim paryCodes

	paryCodes = Split(strField & "", "|")
	If UBound(paryCodes) >= 0 Then AVS_Code = paryCodes(0)
			
End Function	'AVS_Code

'**********************************************************

Function CVV_Code(byRef strField)

Dim paryCodes

	paryCodes = Split(strField & "", "|")
	If UBound(paryCodes) >= 1 Then CVV_Code = paryCodes(1)
			
End Function	'CVV_Code

'**********************************************************

Function scoreAVS(byVal strAVS)

Dim plngScore
	
	'AVS may include actual message
	Dim plngPos
	Dim pstrTemp
	plngPos = InStrRev(strAVS, "(")
	If plngPos > 0 Then
		pstrTemp = Right(strAVS, Len(strAVS) - plngPos)
		strAVS = Left(pstrTemp, 1)
	End If

	Select Case strAVS
		Case "A"	'"Address (Street)matches, ZIP does not. (Code A)"
			plngScore = 1
		Case "D"	'"Street address and Postal Code match (International Issuer). (Code D)"
			plngScore = 1
		Case "E"	'"Ineligible transaction. (Code E)"
			plngScore = 1
		Case "G"	'"Service not supported by issuer (International). (Code G)"
			plngScore = 1
		Case "N"	'"Neither address nor ZIP matches. (Code N)"
			plngScore = 1
		Case "R"	'"Retry (system unavailable or timed out). (Code R)"
			plngScore = 1
		Case "S"	'"Card type not supported. (Code S)"
			plngScore = 1
		Case "U"	'"Address information unavailable. (Code U)"
			plngScore = 1
		Case "W"	'"9 digit zip match, address does not. (Code W)"
			plngScore = 1
		Case "X"	'"Exact match (9 digit zip and address). (Code X)"
			plngScore = 0
		Case "Y"	'"Address and 5 digit zip match. (Code Y)"
			plngScore = 0
		Case "Z"	'"5 digit zip matches, address does not. (Code Z)"
			plngScore = 1
		Case ""		'covers situation with no code for non-CC payments
			plngScore = 0
		Case Else	'"Unknown AVS Code."
			Response.Write("AVS code <strong>" & strAVS & "</strong> has no match. Please contact the developer.<br />")
			plngScore = 1
	End Select	
	
	scoreAVS = plngScore
	
End Function	'scoreAVS

'**********************************************************

Function scoreCVV(byVal strCVV)

Dim plngScore

	Select Case strCVV
		Case "CVV2 MATCH", "Match (M)"			'SecurePay Specific?
			plngScore = 0
		Case "CVV2 NOT AVAILABLE", "Not Processed (P)"	'SecurePay Specific?
			plngScore = 1
		Case "CVV2 NOMATCH", "No Match (N)"			'SecurePay Specific?
			plngScore = 1
		Case ""		'covers situation with no code for non-CC payments
			plngScore = 0
		Case Else	'"Unknown"
			Response.Write("CVV code <strong>" & strCVV & "</strong> is unknown. Please contact the developer.<br />")
			plngScore = 1
	End Select
	
	scoreCVV = plngScore
	
End Function	'scoreCVV

'***********************************************************************************************

Private Function getMatchingOrderIDs(byVal bytDaysBackToCheckOrders)

Dim pobjRS
Dim pstrSQL
Dim pstrOut

	pstrSQL = "SELECT sfOrders.orderID" _
			& " FROM ssOrderManager RIGHT JOIN sfOrders ON ssOrderManager.ssorderID = sfOrders.orderID" _
			& " WHERE (" _
			& "       (sfOrders.orderIsComplete=1)" _
			& "       AND (ssOrderManager.ssInternalOrderStatus Is Null Or ssOrderManager.ssInternalOrderStatus=" & ssInternalOrderStatus_Unread & " Or ssOrderManager.ssInternalOrderStatus=" & ssInternalOrderStatus_ClearedToProcess & ")" _
			& "	      AND (sfOrders.orderDate >= " & wrapSQLValue(DateAdd("d", -1 * bytDaysBackToCheckOrders, Date()) & " 12:00:00 AM", True, enDatatype_date) & ")" _
			& "       )"

	Set pobjRS = GetRS(pstrSQL)
	If Not pobjRS.EOF Then
		pstrOut = pobjRS.getString(,,,",")
		If Right(pstrOut, 1) = "," Then pstrOut = Left(pstrOut, Len(pstrOut)-1)
	End If
	Call ReleaseObject(pobjRS)
	If mblnProcessOrdersTestMode Then Response.Write "<fieldset><legend>getMatchingOrderIDs</legend>SQL: " & pstrSQL & "<hr />Order IDs: " & pstrOut & "</fieldset>"

	getMatchingOrderIDs = pstrOut
	
End Function	'getMatchingOrderIDs

'**************************************************

Function loadOrdersToProcess(byVal bytDaysBackToCheckOrders, byRef referenceMessage)

	Dim pstrSQL
	Dim pobjRS
	Dim pstrMessage
	Dim pstrVendorName
	
	pstrSQL = getMatchingOrderIDs(bytDaysBackToCheckOrders)
	If Len(pstrSQL) = 0 Then
		loadOrdersToProcess = False
		Exit Function
	End If
	
	pstrSQL = "SELECT sfOrders.*, ssOrderManager.*, sfOrderDetails.*, sfProducts.prodName, sfManufacturers.mfgName, sfVendors.vendName, sfTransactionResponse.trnsrspAVSCode, sfCustomers.*, sfCShipAddresses.*" _
			& " FROM sfCShipAddresses INNER JOIN (sfCustomers INNER JOIN (sfTransactionResponse RIGHT JOIN ((sfOrderDetails RIGHT JOIN (ssOrderManager RIGHT JOIN sfOrders ON ssOrderManager.ssorderID = sfOrders.orderID) ON sfOrderDetails.odrdtOrderId = sfOrders.orderID) LEFT JOIN ((sfProducts LEFT JOIN sfManufacturers ON sfProducts.prodManufacturerId = sfManufacturers.mfgID) LEFT JOIN sfVendors ON sfProducts.prodVendorId = sfVendors.vendID) ON sfOrderDetails.odrdtProductID = sfProducts.prodID) ON sfTransactionResponse.trnsrspOrderId = sfOrders.orderID) ON sfCustomers.custID = sfOrders.orderCustId) ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId" _
			& " WHERE sfOrders.orderID In (" & pstrSQL & ")" _
			& " Order By sfOrders.OrderID"
	referenceMessage = "Only completed orders placed after " & FormatDateTime(DateAdd("d", -1 * bytDaysBackToCheckOrders, Date())) & " with an internal status of none, unread, or ready for shipping are being examined."

	Set pobjRS = GetRS(pstrSQL)
	'If mblnProcessOrdersTestMode Then Response.Write "<fieldset><legend>getMatchingOrderIDs</legend>SQL: " & pstrSQL & "<hr />RecordCount: " & pobjRS.RecordCount & "</fieldset>"
	
	If Not pobjRS.EOF Then
		'Call DebugPrintRecordset("Order Processing", pobjRS)
		ReDim maryOrders(pobjRS.RecordCount - 1)
		For RecordCounter = 0 To pobjRS.RecordCount - 1
			maryOrders(RecordCounter) = Array( _
											  pobjRS.Fields("orderID").Value, _
											  Trim(pobjRS.Fields("prodName").Value & ""), _
											  Trim(pobjRS.Fields("mfgName").Value & ""), _
											  Trim(pobjRS.Fields("vendName").Value & ""), _
											  "", _
											  "", _
											  "", _
											  0, _
											  pobjRS.Fields("odrdtID").Value, _
											  pobjRS.Fields("ssInternalOrderStatus").Value, _
											  pobjRS.Fields("orderPaymentMethod").Value _
											  )

			'Response.Write maryOrders(RecordCounter)(enProcessingItem_OrderNumber) & "<br />"
			For RuleCounter = 0 To UBound(maryProcessingRules)
				pstrMessage = ""
				Select Case RuleCounter
					Case enProcessingRule_OldOrder:	'Old Orders
						If isDate(pobjRS.Fields("orderDate").Value) Then
							If CDate(pobjRS.Fields("orderDate").Value) < DateAdd("d", -60, Date()) Then
								If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
								Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
							End If
						End If
					Case enProcessingRule_NoOrderDetails:	'Orders with no order details
						If isNull(pobjRS.Fields("odrdtID").Value) Then
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_InternationalShipping:	'Orders with International Shipping Addresses
						If UCase(pobjRS.Fields("cshpaddrShipCountry").Value <> "US") Then
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_InternationalBilling:	'Orders with International Billing Addresses
						If UCase(pobjRS.Fields("custCountry").Value <> "US") Then
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_AddressesDoNotMatch:	'Orders where Shipping/Billing Addresses do not match
						If Not isBillingShippingAddressSame(pobjRS) Then
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_HighDollarOrder:	'Order subtotal over limit
						If isNumeric(pobjRS.Fields("orderGrandTotal").Value) Then
							If CBool(CDbl(pobjRS.Fields("orderGrandTotal").Value) > cdblOrderThresholdRisk) Then
								If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
								Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
							End If
						End If
					Case enProcessingRule_VeryHighDollarOrder:	'Order subtotal over limit
						If isNumeric(pobjRS.Fields("orderGrandTotal").Value) Then
							If CBool(CDbl(pobjRS.Fields("orderGrandTotal").Value) > cdblOrderThresholdAbsolute) Then
								If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
								Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
							Else
								'Response.Write("<h4>Not a high dollar order</h4>")
							End If
						End If
					Case enProcessingRule_PaidByPayPal:
						Select Case Trim(pobjRS.Fields("orderPaymentMethod").Value & "")
							Case "PayPal Initial"
								If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
								Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
							Case "PayPal"
								'check to make sure IPN returned
								If Not isPayPalPaymentVerified(pobjRS.Fields("orderID").Value) Then
									If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
									Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
								End If
						End Select
					Case enProcessingRule_PaidByPO:
						If Trim(pobjRS.Fields("orderPaymentMethod").Value & "") = cstrPOPaymentMethodName Then
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_Mfg_Spotlight:
						If Trim(pobjRS.Fields("mfgName").Value & "") = "Spotlight" Then
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_Vend_GenericVendor:
						pstrVendorName = Trim(pobjRS.Fields("vendName").Value & "")
						Select Case pstrVendorName
							Case "DS Hull Company", _
								 "International Dock Products", _
								 "Mermaid Marine Air", _
								 "Megafend", _
								 "Navimo", _
								 "Power Bright", _
								 "Power House", _
								 "Quick USA", _
								 "Repair Tech", _
								 "Revere", _
								 "Star Marine Depot", _
								 "Land N Sea", _
								 "North State Nursery"
								If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
								Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
							Case Else
								'do nothing
						End Select
					Case enProcessingRule_Vend_CWR
						pstrVendorName = Trim(pobjRS.Fields("vendName").Value & "")
						If pstrVendorName = "CWR Electronics" Then
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)			
						End If
					Case enProcessingRule_ScorePriorCustomer
						If isPriorCustomer(pobjRS.Fields("orderCustID").Value) Then
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_ScoreAVS
						If scoreAVS(AVS_Code(pobjRS.Fields("trnsrspAVSCode").Value)) <> 0 Then
							pstrMessage = "AVS code (" & AVS_Code(pobjRS.Fields("trnsrspAVSCode").Value) & ") resulted in a fraud score of <em>" & scoreAVS(AVS_Code(pobjRS.Fields("trnsrspAVSCode").Value)) & "</em>"
							maryOrders(RecordCounter)(enProcessingItem_FraudScore) = maryOrders(RecordCounter)(enProcessingItem_FraudScore) + scoreAVS(AVS_Code(pobjRS.Fields("trnsrspAVSCode").Value))
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_ScoreCVV
						If scoreCVV(CVV_Code(pobjRS.Fields("trnsrspAVSCode").Value)) <> 0 Then
							pstrMessage = "CVV code (" & CVV_Code(pobjRS.Fields("trnsrspAVSCode").Value) & ") resulted in a fraud score of <em>" & scoreCVV(CVV_Code(pobjRS.Fields("trnsrspAVSCode").Value)) & "</em>"
							maryOrders(RecordCounter)(enProcessingItem_FraudScore) = maryOrders(RecordCounter)(enProcessingItem_FraudScore) + scoreCVV(CVV_Code(pobjRS.Fields("trnsrspAVSCode").Value))
							If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
							Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
						End If
					Case enProcessingRule_CheckFraudScore
						If maryOrders(RecordCounter)(enProcessingItem_InternalOrderStatus) = ssInternalOrderStatus_ClearedToProcess Then
							'no processing rule required as it has been manually cleared
						Else
							If maryOrders(RecordCounter)(enProcessingItem_FraudScore) > maryProcessingRules(RuleCounter)(enProcessingRule_FraudScore) Then
								If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
								Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
							End If
						End If
					Case Else:
						pstrMessage = "<h4 style=""color:red"">No test for rule <em>" & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & "</em> has been defined</h4>"
						If cDebugMode = 2 Then Response.Write "Add Rule " & maryProcessingRules(RuleCounter)(enProcessingRule_Title) & " (" & RuleCounter & ")<br />"
						Call addProcessingRule(maryOrders(RecordCounter), maryProcessingRules(RuleCounter), pstrMessage)
				End Select	'RuleCounter
			Next 'RuleCounter
			pobjRS.MoveNext
		Next 'RecordCounter
		loadOrdersToProcess = True
	Else
		loadOrdersToProcess = False
	End If	'Not pobjRS.EOF
	pobjRS.Close
	Set pobjRS = Nothing
	
End Function	'loadOrdersToProcess

'**************************************************

Function ProcessOrderActions()

Dim paryRule
Dim paryAction
Dim plngOrderID
Dim plngOrderDetailID
Dim pstrActionMessage
Dim pstrLocalError
Dim pstrRuleMessage
Dim pstrSQL
Dim pobjRS
Dim pdicVendors
Dim orderItemCounter
Dim pstrAttributes
Dim FieldCounter
Dim pstrKey

'Call writeOrderProcessingLogEntry(Array(1, 1, "Test Message", 1))
'Call writeOrderProcessingLogEntry(Array(1, "", "Test Message", 1))
'Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage, 0))

	If isArray(maryOrders) Then
		Set pdicVendors = CreateObject("Scripting.Dictionary")
		For RecordCounter = 0 To UBound(maryOrders)
			If Len(maryOrders(RecordCounter)(enProcessingItem_Rules)) > 0 Then
				paryRule = Split(maryOrders(RecordCounter)(enProcessingItem_Rules), cstrDelimiter)
				paryAction = Split(maryOrders(RecordCounter)(enProcessingItem_Actions), cstrDelimiter)
				plngOrderID = maryOrders(RecordCounter)(enProcessingItem_OrderNumber)
				plngOrderDetailID = maryOrders(RecordCounter)(enProcessingItem_OrderDetailID)
				Call checkssOrderManagerRecordExists(plngOrderID)
				
				For i = 0 To UBound(paryAction)
					pstrActionMessage = maryProcessingActions(paryAction(i))(enProcessingAction_Title)
					If UBound(paryRule) >= i Then
						pstrRuleMessage = maryProcessingRules(paryRule(i))(enProcessingRule_Title)
					Else
						pstrRuleMessage = "Unknown"
					End If

					Select Case CLng(paryAction(i))
						Case enProcessingAction_None:	'Do nothing
						Case enProcessingAction_DeleteOrder:
							'add code to delete order; this is left for future use
							Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage & ";Order Deletion is currently not enabled", 0))
						Case enProcessingAction_SetOrderToFlagged:
							If mblnProcessOrdersTestMode Then
								Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage & "|" & pstrActionMessage & ";Action Not Performed due to test mode.", 0))
							Else
								If SetOrderToFlagged(plngOrderID) Then
									Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage & "|" & pstrActionMessage, 1))
								Else
									Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage & "|" & pstrActionMessage & ";" & pstrLocalError, 0))
								End If
							End If
						Case enProcessingAction_ChangeInternalStatusToFlagForManualProcessing:
							If mblnProcessOrdersTestMode Then
								Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage & "|" & pstrActionMessage & ";Action Not Performed due to test mode.", 0))
							Else
								If SetInternalOrderStatusToFlaggedForManual(plngOrderID) Then
									Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage & "|" & pstrActionMessage, 1))
								Else
									Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage & "|" & pstrActionMessage & ";" & pstrLocalError, 0))
								End If
							End If
						Case enProcessingAction_Email_PO_GenericVendor
							'Just add to the group processing
							pstrKey = maryOrders(RecordCounter)(enProcessingItem_OrderItemVend)
							If pdicVendors.Exists(pstrKey) Then
								'pdicVendors(pstrKey) = pdicVendors(pstrKey) & "," & RecordCounter
								pdicVendors(pstrKey) = pdicVendors(pstrKey) & "," & plngOrderDetailID
								
							Else
								'pdicVendors.Add pstrKey, RecordCounter
								pdicVendors.Add pstrKey, plngOrderDetailID
							End If
						Case enProcessingAction_EDI_CWR
							'Just add to the group processing
							pstrKey = "CWR Electronics"
							If pdicVendors.Exists(pstrKey) Then
								'pdicVendors(pstrKey) = pdicVendors(pstrKey) & "," & RecordCounter
								pdicVendors(pstrKey) = pdicVendors(pstrKey) & "," & plngOrderDetailID
								
							Else
								'pdicVendors.Add pstrKey, RecordCounter
								pdicVendors.Add pstrKey, plngOrderDetailID
							End If
						Case Else
							Response.Write "<h4 style=""color:red"">No action for <em>" & pstrRuleMessage & "</em> has been defined</h4>"
							Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, pstrRuleMessage & "|" & pstrActionMessage & ";No Action Defined.", 0))
					End Select
					
					'check if selected rule aborts
					'Response.Write "Abort? " & maryProcessingActions(paryAction(i))(enProcessingAction_CeaseFurtherProcessing) & "<BR>"
					If maryProcessingActions(paryAction(i))(enProcessingAction_CeaseFurtherProcessing) Then
					    'Response.Write "Order Action processing aborted by rule " & maryProcessingActions(paryAction(i))(enProcessingAction_Title) & "<BR>"
					    Exit For
					End If
				Next 'i

			End If
		Next 'RecordCounter
		
		'Now process the grouped actions
		Dim vItem
		Dim pclsEmail
		Dim paryEmails
		Dim pstrEmailTemplate
				
		For Each vItem in pdicVendors
			Select Case vItem
				Case "DS Hull Company", _
					 "International Dock Products", _
					 "Mermaid Marine Air", _
					 "Megafend", _
					 "Navimo", _
					 "Power Bright", _
					 "Power House", _
					 "Quick USA", _
					 "Repair Tech", _
					 "Revere", _
					 "Star Marine Depot", _
					 "Land N Sea", _
					 "North State Nursery"

					Response.Write "Items for " & vItem & ": " & pdicVendors(vItem) & "<br />"
					If vItem = "DS Hull Company" Then
					    pstrEmailTemplate = "DS_Hull_CompanyVendorEmail.txt"
					ElseIf vItem = "Land N Sea" Then
					    pstrEmailTemplate = "LandNSea_CompanyVendorEmail.txt"
					Else
					    pstrEmailTemplate = "NorthStateVendorEmail.txt"
					End If
					pstrEmailTemplate = "Generic_VendorEmail.txt"
					Set pclsEmail = New clsEmail
					With pclsEmail
						Call .setReplacementValue("OrderDate", Date())
						If .LoadEmailTemplates(ssAdminPath & emailTemplateDirectory, pstrEmailTemplate, paryEmails) Then
							.MailMethod = adminMailMethod
							.MailServer = adminMailServer
							.ShowFailures = True
							.From = EmailFromAddress(adminPrimaryEmail)
							.CC = "purchasing@starmarinedepot.com"  'EmailFromAddress(adminPrimaryEmail)
							.BCC = ""
							
							If cblnSMDorLocal Then
								'works on SMD
								pstrSQL = "SELECT sfCShipAddresses.*, sfOrders.orderID, sfOrders.orderShipMethod, sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtManufacturer, sfOrderDetails.odrdtVendor, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, sfProducts.*, sfVendors.*" _
										& " FROM sfVendors INNER JOIN sfOrderDetails INNER JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID ON sfVendors.vendID = sfProducts.prodVendorId LEFT OUTER JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId RIGHT OUTER JOIN sfCShipAddresses INNER JOIN sfOrders ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId ON sfOrderDetails.odrdtOrderId = sfOrders.orderID" _
										& " WHERE sfOrderDetails.odrdtID In (" & pdicVendors(vItem) & ")" _
										& " ORDER BY sfOrders.orderID, sfOrderDetails.odrdtID"
							Else
								'works in Access
								pstrSQL = "SELECT sfCShipAddresses.*, sfOrders.orderID, sfOrders.orderDate, sfOrders.orderShipMethod, sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtManufacturer, sfOrderDetails.odrdtVendor, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, sfProducts.*, sfVendors.*" _
										& " FROM ((sfCShipAddresses INNER JOIN sfOrders ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId) LEFT JOIN (sfOrderDetails LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN (sfVendors RIGHT JOIN sfProducts ON sfVendors.vendID = sfProducts.prodVendorId) ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
										& " WHERE sfOrderDetails.odrdtID In (" & pdicVendors(vItem) & ")" _
										& " ORDER BY sfOrders.orderID, sfOrderDetails.odrdtID"
							End If

					        'Response.Write "pstrSQL: " & pstrSQL & "<br />"
					        'Response.Flush
							Set pobjRS = GetRS(pstrSQL)
							If Not pobjRS.EOF Then
								plngOrderID = ""
								.To = Trim(pobjRS.Fields("vendEmail").Value)
								.RepeatingItemTag = "orderItems"
								Do While Not pobjRS.EOF
									If plngOrderID <> pobjRS.Fields("orderID").Value Then
										If Len(plngOrderID) <> 0 Then
											Call .setReplacementValue_Repeating("attributes", pstrAttributes)
											Call .SetRepeatingItemReplacementText
											'Send email
											If mblnProcessOrdersTestMode Then
												Response.Write .mailAsHTML
											Else
												.Send
												Call SetInternalOrderStatusToOrderedWithVendor(plngOrderID)
											End If
											.ClearRepeatingItemReplacementText
										End If
										plngOrderID = pobjRS.Fields("orderID").Value
										plngOrderDetailID = 0
										orderItemCounter = 0
										Call setOrderLevelReplacements(pclsEmail, pobjRS)
										Call .setReplacementValue("TemplateVendorName", vItem)
									End If

									If plngOrderDetailID <> pobjRS.Fields("odrdtID").Value Then
										If Not mblnProcessOrdersTestMode Then Call addInternalOrderStatusMessage(plngOrderID, Trim(pobjRS.Fields("odrdtProductID").Value) & " sent to vendor.")
										plngOrderDetailID = pobjRS.Fields("odrdtID").Value
										Call .setReplacementValue_Repeating("attributes", pstrAttributes)
										If orderItemCounter > 0 Then Call .SetRepeatingItemReplacementText
										orderItemCounter = orderItemCounter + 1
										pstrAttributes = ""
										Call setOrderDetailReplacements(pclsEmail, pobjRS)
										Call .setReplacementValue_Repeating("itemCounter", orderItemCounter)
									End If
									pstrAttributes = pstrAttributes & getOrderAttributeText(pobjRS)
									
									pobjRS.MoveNext
								Loop

								Call .setReplacementValue_Repeating("attributes", pstrAttributes)
								Call .SetRepeatingItemReplacementText	
							End If
							Call ReleaseObject(pobjRS)

							If mblnProcessOrdersTestMode Then
								Response.Write .mailAsHTML
							Else
								.Send
								Call SetInternalOrderStatusToOrderedWithVendor(plngOrderID)
							End If
						Else
							Response.Write "<h4 style=""color:red"">Unable to load email template</h4>"
							Response.Write .writeErrorMessages
						End If	'.LoadEmailTemplates
						
					End With	'pclsEmail
					
					Set pclsEmail = Nothing
			
				Case "CWR Electronics"
					Response.Write "Items for " & vItem & ": " & pdicVendors(vItem) & "<br />"
					Set pclsEmail = New clsEmail
					With pclsEmail
						Call .setReplacementValue("OrderDate", Date())
						If .LoadEmailTemplates(ssAdminPath & emailTemplateDirectory, "CWR_EDI.txt", paryEmails) Then
							.MailMethod = adminMailMethod
							.MailServer = adminMailServer
							.ShowFailures = True
							.From = EmailFromAddress(adminPrimaryEmail)
							.CC = "purchasing@starmarinedepot.com"  'EmailFromAddress(adminPrimaryEmail)
							.BCC = ""
							
							If cblnSMDorLocal Then
								'works on SMD
								pstrSQL = "SELECT sfCShipAddresses.*, sfOrders.orderID, sfOrders.orderDate, sfOrders.orderShipMethod, sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtManufacturer, sfOrderDetails.odrdtVendor, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, sfProducts.*, sfVendors.*" _
										& " FROM sfVendors INNER JOIN sfOrderDetails INNER JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID ON sfVendors.vendID = sfProducts.prodVendorId LEFT OUTER JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId RIGHT OUTER JOIN sfCShipAddresses INNER JOIN sfOrders ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId ON sfOrderDetails.odrdtOrderId = sfOrders.orderID" _
										& " WHERE sfOrderDetails.odrdtID In (" & pdicVendors(vItem) & ")" _
										& " ORDER BY sfOrders.orderID, sfOrderDetails.odrdtID"
	
								pstrSQL = "SELECT sfCShipAddresses.*, sfOrders.orderID, sfOrders.orderDate, sfOrders.orderShipMethod, sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtManufacturer, sfOrderDetails.odrdtVendor, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, sfProducts.*, sfVendors.*" _
										& " FROM sfVendors INNER JOIN sfOrderDetails INNER JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID ON sfVendors.vendID = sfProducts.prodVendorId LEFT OUTER JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId RIGHT OUTER JOIN sfCShipAddresses INNER JOIN sfOrders ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId ON sfOrderDetails.odrdtOrderId = sfOrders.orderID" _
										& " WHERE sfOrderDetails.odrdtID In (" & pdicVendors(vItem) & ")" _
										& " ORDER BY sfOrders.orderID, sfOrderDetails.odrdtID"
							Else
								'works in Access
								pstrSQL = "SELECT sfCShipAddresses.*, sfOrders.orderID, sfOrders.orderDate, sfOrders.orderShipMethod, sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtManufacturer, sfOrderDetails.odrdtVendor, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, sfProducts.*, sfVendors.*" _
										& " FROM ((sfCShipAddresses INNER JOIN sfOrders ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId) LEFT JOIN (sfOrderDetails LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN (sfVendors RIGHT JOIN sfProducts ON sfVendors.vendID = sfProducts.prodVendorId) ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
										& " WHERE sfOrderDetails.odrdtID In (" & pdicVendors(vItem) & ")" _
										& " ORDER BY sfOrders.orderID, sfOrderDetails.odrdtID"
							End If


					        'Response.Write "pstrSQL: " & pstrSQL & "<br />"
							Set pobjRS = GetRS(pstrSQL)
							If Not pobjRS.EOF Then
								plngOrderID = ""
								.To = Trim(pobjRS.Fields("vendEmail").Value)
								.RepeatingItemTag = "orderItems"
								
								Do While Not pobjRS.EOF
									If plngOrderID <> pobjRS.Fields("orderID").Value Then
										If Len(plngOrderID) <> 0 Then
											Call .setReplacementValue_Repeating("attributes", pstrAttributes)
											Call .SetRepeatingItemReplacementText
										End If
										plngOrderID = pobjRS.Fields("orderID").Value
										plngOrderDetailID = 0
										orderItemCounter = 0
										Call setOrderLevelReplacements(pclsEmail, pobjRS)
			                            Call .setReplacementValue("customDate", customFormatDate(Trim(pobjRS.Fields("orderDate").Value & ""), "mmddyyyy"))
										
			                            'Call .setReplacementValue("Address1", Trim(pobjRS.Fields("cshpaddrShipCompany").Value & ""))
										'If Len(.getReplacementValue("Address1")) = 0 Then  Call .setReplacementValue("Address1", .RecipientName)
			                            Call .setReplacementValue("Address1", Trim(pobjRS.Fields("cshpaddrShipAddr1").Value & ""))
			                            Call .setReplacementValue("Address2", Trim(pobjRS.Fields("cshpaddrShipAddr2").Value & ""))
										If Len(Trim(pobjRS.Fields("cshpaddrShipCompany").Value & "")) = 0 Then
											Call .setReplacementValue("CWR_Name", .RecipientName)
											Call .setReplacementValue("CWR_Attn", "")
											Call .setReplacementValue("Address3", "")
										Else
											Call .setReplacementValue("CWR_Name", Trim(pobjRS.Fields("cshpaddrShipCompany").Value & ""))
											Call .setReplacementValue("CWR_Attn", .RecipientName)
											
											Call .setReplacementValue("Address2", Trim(pobjRS.Fields("cshpaddrShipCompany").Value & ""))
											Call .setReplacementValue("Address3", Trim(pobjRS.Fields("cshpaddrShipAddr2").Value & ""))
										End If
			                            If Not mblnProcessOrdersTestMode Then Call SetInternalOrderStatusToOrderedWithVendor(plngOrderID)
									End If
									
									If plngOrderDetailID <> pobjRS.Fields("odrdtID").Value Then
										If Not mblnProcessOrdersTestMode Then Call addInternalOrderStatusMessage(plngOrderID, Trim(pobjRS.Fields("odrdtProductID").Value) & " added to EDI file.")
										plngOrderDetailID = pobjRS.Fields("odrdtID").Value
										Call .setReplacementValue_Repeating("attributes", pstrAttributes)
										If orderItemCounter > 0 Then
										    Call .SetRepeatingItemReplacementText
                                        End If
										orderItemCounter = orderItemCounter + 1
										pstrAttributes = ""
										Call setOrderDetailReplacements(pclsEmail, pobjRS)
										Call .setReplacementValue_Repeating("itemCounter", orderItemCounter)
										
										'temporary rule
										Call .setReplacementValue_Repeating("shipCode", "DHFP")
										If False Then
										Select Case UCase(Trim(pobjRS.Fields("orderShipMethod").Value & ""))
										    Case "DHL PROMOTION":               Call .setReplacementValue_Repeating("shipCode", "DHFP")
										    Case "CUSTOMER PICK UP":            Call .setReplacementValue_Repeating("shipCode", "PU")
										    Case "OUR DELIVERY":                Call .setReplacementValue_Repeating("shipCode", "OD")
										    Case "DHL 2 DAY AIR":               Call .setReplacementValue_Repeating("shipCode", "DHL2")
										    Case "DHL GROUND":                  Call .setReplacementValue_Repeating("shipCode", "DHLG")
										    Case "DHL OVERNIGHT":               Call .setReplacementValue_Repeating("shipCode", "DHLO")
										    Case "FEDEX 2 DAY":                 Call .setReplacementValue_Repeating("shipCode", "FED2")
										    Case "FEDEX EXPRESS SAVER":         Call .setReplacementValue_Repeating("shipCode", "FEDXP")
										    Case "FEDEX PRIORITY OVERNIGHT":    Call .setReplacementValue_Repeating("shipCode", "FEDP")
										    Case "FEDEX SATURDAY DELIVERY":     Call .setReplacementValue_Repeating("shipCode", "FEDPS")
										    Case "FEDEX STANDARD OVERNIGHT":    Call .setReplacementValue_Repeating("shipCode", "FEDS")
										    Case "TRUCK FREIGHT":               Call .setReplacementValue_Repeating("shipCode", "TF")
										    Case "UPS 2ND DAY":                 Call .setReplacementValue_Repeating("shipCode", "UPS2D")
										    Case "UPS 2ND DAY AM":              Call .setReplacementValue_Repeating("shipCode", "UPS2DAM")
										    Case "UPS 3 DAY":                   Call .setReplacementValue_Repeating("shipCode", "UPS3D")
										    Case "UPS GROUND", "UPS STANDARD GROUND":                  Call .setReplacementValue_Repeating("shipCode", "UPSG")
										    Case "UPS NEXT DAY":                Call .setReplacementValue_Repeating("shipCode", "UPS1D")
										    Case "UPS NEXT DAY EARLY":          Call .setReplacementValue_Repeating("shipCode", "UPS1DAM")
										    Case "UPS NEXT DAY SAVER":          Call .setReplacementValue_Repeating("shipCode", "UPS1DS")
										    Case "UPS SATURDAY DELIVERY":       Call .setReplacementValue_Repeating("shipCode", "UPSSAT")
										    Case "USPS EXPRESS MAIL":           Call .setReplacementValue_Repeating("shipCode", "USPE")
										    Case "USPS PRIORITY MAIL":          Call .setReplacementValue_Repeating("shipCode", "USPP")
										    Case "DHL INTERNATIONAL":           Call .setReplacementValue_Repeating("shipCode", "DHLI")
										    Case "FEDEX INTERNATIONAL ECONOMY": Call .setReplacementValue_Repeating("shipCode", "FEDI")
										    Case "FEDEX INTERNATIONAL PRIORITY":Call .setReplacementValue_Repeating("shipCode", "FEDIP")
										    Case "UPS CANADA EXPEDITED":        Call .setReplacementValue_Repeating("shipCode", "UPSCP")
										    Case "UPS CANADA EXPRESS":          Call .setReplacementValue_Repeating("shipCode", "UPSCE")
										    Case "UPS CANADA STANDARD":         Call .setReplacementValue_Repeating("shipCode", "UPSCS")
										    Case "UPS WORLDWIDE EXPEDITED":     Call .setReplacementValue_Repeating("shipCode", "UPSWE")
										    Case "UPS WORLDWIDE EXPRESS":       Call .setReplacementValue_Repeating("shipCode", "UPSI")
										    Case "USPS EXPRESS INTERNATIONAL":  Call .setReplacementValue_Repeating("shipCode", "USPIE")
										    Case "HOLIDAY - SPECIAL OVERNIGHT DELIVERY":  Call .setReplacementValue_Repeating("shipCode", "DHFP")
										    Case Else:                          Call .setReplacementValue_Repeating("shipCode", "Unknown Ship Method (" & Trim(pobjRS.Fields("orderShipMethod").Value & "") & ")")
										End Select
										End If
                                    Else
                                        'same order detail, different attribute; the attribute is handled below
									End If  'plngOrderDetailID <> pobjRS.Fields("odrdtID").Value
									pstrAttributes = pstrAttributes & getOrderAttributeText(pobjRS)
									
									pobjRS.MoveNext
								Loop

								Call .setReplacementValue_Repeating("attributes", pstrAttributes)
								Call .SetRepeatingItemReplacementText	
							End If
							Call ReleaseObject(pobjRS)

							Dim pstrFilePath
							pstrFilePath = rootPath & "fpdb/myCWR.txt"
							Call writeToFile(pstrFilePath, .bodyWithReplacements)
							Response.Write "<div style=""background-color:yellow""><a href=""" & mstrBaseHRef & "fpdb/myCWR.txt" & """ target=""_blank"">Download CWR File</a></div>"
							'Call SendFileToCWR(pstrFilePath)
							If mblnProcessOrdersTestMode Then
								'Response.Write .mailAsHTML
								Response.Write "<textarea rows=15 cols=900>" & .bodyWithReplacements & "</textarea><br />"
							Else
								'Call SendFileToCWR(pstrFilePath)
							End If
						Else
							Response.Write "<h4 style=""color:red"">Unable to load email template</h4>"
							Response.Write .writeErrorMessages
						End If	'.LoadEmailTemplates
						
					End With	'pclsEmail
					
					Set pclsEmail = Nothing
			
				Case "General"
Response.Write "Items: " & pdicVendors(vItem) & "<br />"
					Set pclsEmail = New clsEmail
					With pclsEmail
						If .LoadEmailTemplates(ssAdminPath & emailTemplateDirectory, "NorthStateVendorEmail.txt", paryEmails) Then
							.MailMethod = adminMailMethod
							.MailServer = adminMailServer
							.ShowFailures = True
							
							pstrSQL = "SELECT sfCShipAddresses.*, sfOrders.orderID, sfOrders.orderShipMethod, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtQuantity, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName" _
									& " FROM (sfCShipAddresses INNER JOIN sfOrders ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId) LEFT JOIN (sfOrderDetails LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
									& " WHERE sfOrderDetails.odrdtID In (" & pdicVendors(vItem) & ")" _
									& " ORDER BY sfOrders.orderID, sfOrderDetails.odrdtID"
							Set pobjRS = GetRS(pstrSQL)
							If Not pobjRS.EOF Then
								Call setOrderLevelReplacements(pclsEmail, pobjRS)
							End If
							Call ReleaseObject(pobjRS)

							.From = EmailFromAddress(adminPrimaryEmail)
							.To = ""
							.CC = EmailFromAddress(adminPrimaryEmail)
							.BCC = ""
							'.Send
							Response.Write .mailAsHTML
						Else
							Response.Write "<h4 style=""color:red"">Unable to load email template</h4>"
							Response.Write .writeErrorMessages
						End If	'.LoadEmailTemplates
						
					End With	'pclsEmail
					
					Set pclsEmail = Nothing
			
				Case Else
					Response.Write "<h4 style=""color:red"">No grouped action for <em>" & vItem & "</em> has been defined</h4>"
					Call writeOrderProcessingLogEntry(Array(plngOrderID, plngOrderDetailID, "No Grouped Action Defined for " & vItem & ".", 0))
					Response.Write vItem & ": " & pdicVendors(vItem) & "<br />"
			
			End Select
		Next
		Set pdicVendors = Nothing
	End If

End Function	'ProcessOrderActions

'**************************************************

Function checkssOrderManagerRecordExists(byVal lngOrderID)

Dim pstrSQL
Dim pobjRS
Dim pblnExists

	pstrSQL = "Select ssOrderID from ssOrderManager Where ssOrderID = " & wrapSQLValue(lngOrderID, False, enDatatype_number)
	Set pobjRS = GetRS(pstrSQL)
	If pobjRS.EOF Then
		pstrSQL = "Insert Into ssOrderManager (ssOrderID) Values (" & wrapSQLValue(lngOrderID, False, enDatatype_number) & ")"
		cnn.Execute pstrSQL,,128
		pblnExists = True
	Else
		pblnExists = True
	End If
	Call ReleaseObject(pobjRS)

	checkssOrderManagerRecordExists = pblnExists

End Function	'checkssOrderManagerRecordExists

'**************************************************

Function MassUpdateOrders_SetInternalOrderStatusToFlaggedForManual(byVal strOrderIDs)

Dim i
Dim paryOrderIDs

	If Len(strOrderIDs) > 0 Then
		paryOrderIDs = Split(strOrderIDs, ", ")
		For i = 0 To UBound(paryOrderIDs)
			If SetInternalOrderStatusToFlaggedForManual(Trim(paryOrderIDs(i))) Then
				Response.Write "Order " & paryOrderIDs(i) & " set to require manual processing.<br />"
			Else
			End If
		Next 'i
	End If
	
End Function

Function SetOrderToFlagged(byVal lngOrderID)

Dim pstrSQL
Dim pstrLocalError

	pstrSQL = "Update ssOrderManager Set ssOrderFlagged=1, ssInternalOrderStatus=3 Where ssorderID=" & lngOrderID
	SetOrderToFlagged = Execute_NoReturn(pstrSQL, pstrLocalError)

End Function

Function SetInternalOrderStatusToFlaggedForManual(byVal lngOrderID)
	SetInternalOrderStatusToFlaggedForManual = SetInternalOrderStatus(lngOrderID, ssInternalOrderStatus_ManuallyProcess)
End Function

Function SetInternalOrderStatusToOrderedWithVendor(byVal lngOrderID)
	SetInternalOrderStatusToOrderedWithVendor = SetInternalOrderStatus(lngOrderID, ssInternalOrderStatus_OrderedWithVendor)
End Function

Function SetInternalOrderStatus(byVal lngOrderID, byVal bytStatus)

Dim pstrSQL
Dim pstrLocalError

	pstrSQL = "Update ssOrderManager Set ssOrderFlagged=1, ssInternalOrderStatus=" & bytStatus & " Where ssorderID=" & lngOrderID
	SetInternalOrderStatus = Execute_NoReturn(pstrSQL, pstrLocalError)
	
End Function	'SetInternalOrderStatus

'**************************************************

Sub addInternalOrderStatusMessage(byVal lngOrderID, byVal strMessage)

Dim pobjRS
Dim pstrMessage
Dim pstrSQL

	pstrSQL = "Select ssInternalNotes from ssOrderManager Where ssOrderID = " & wrapSQLValue(lngOrderID, False, enDatatype_number)
	Set pobjRS = GetRS(pstrSQL)
	If pobjRS.EOF Then
		pstrSQL = "Insert Into ssOrderManager (ssOrderID, ssInternalNotes) Values (" & wrapSQLValue(lngOrderID, False, enDatatype_number) & ", " & wrapSQLValue(strMessage, True, enDatatype_string) & ")"
		cnn.Execute pstrSQL,,128
	Else
		pstrMessage = Trim(pobjRS.Fields("ssInternalNotes").Value & "")
		If Len(pstrMessage) > 0 Then
			'make sure it isn't a duplicate message
			If InStr(1, pstrMessage, strMessage) = 0 Then pstrMessage = pstrMessage & vbcrlf & strMessage
		Else
			pstrMessage = strMessage
		End If
		pstrSQL = "Update ssOrderManager Set ssInternalNotes=" & wrapSQLValue(pstrMessage, True, enDatatype_string) & " Where ssOrderID=" & wrapSQLValue(lngOrderID, False, enDatatype_number)
		cnn.Execute pstrSQL,,128
	End If
	Call ReleaseObject(pobjRS)

End Sub	'addInternalOrderStatusMessage

'**************************************************

Function isPayPalPaymentVerified(byVal lngOrderID)

Const cstrIPNPaymentNote = "Payment Recorded via PayPal IPN"

Dim pblnResult
Dim pobjRS
Dim pstrMessage
Dim pstrSQL

	pblnResult = False
	
	pstrSQL = "Select ssOrderID from ssOrderManager Where ssInternalNotes= " & wrapSQLValue(cstrIPNPaymentNote, False, enDatatype_string) & " And ssOrderID = " & wrapSQLValue(lngOrderID, False, enDatatype_number)
	Set pobjRS = GetRS(pstrSQL)
	pblnResult = Not pobjRS.EOF
	Call ReleaseObject(pobjRS)
	
	isPayPalPaymentVerified = pblnResult

End Function	'isPayPalPaymentVerified

'**************************************************

Function writeOrderProcessingLogEntry(byVal aryResult)

Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	pstrSQL = "Insert Into ssOrderProcessingLog (orderProcessingOrderID, orderProcessingOrderDetailID, orderProcessingAction, orderProcessingSuccess, orderProcessingDate) Values (?, ?, ?, ?, ?)"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		
		.Parameters.Append .CreateParameter("orderProcessingOrderID", adInteger, adParamInput, 4, aryResult(0))
		
		If Len(aryResult(1)) > 0 Then
			.Parameters.Append .CreateParameter("orderProcessingOrderDetailID", adInteger, adParamInput, 4, aryResult(1))
		Else
			.Parameters.Append .CreateParameter("orderProcessingOrderDetailID", adInteger, adParamInput, 4, Null)
		End If
		
		If Len(aryResult(2)) > 255 Then
			.Parameters.Append .CreateParameter("orderProcessingAction", adVarChar, adParamInput, 255, Left(aryResult(2), 255))
		ElseIf Len(aryResult(2)) > 0 Then
			.Parameters.Append .CreateParameter("orderProcessingAction", adVarChar, adParamInput, Len(aryResult(2)), aryResult(2))
		Else
			.Parameters.Append .CreateParameter("orderProcessingAction", adVarChar, adParamInput, 1, Null)
		End If

		If aryResult(3) = 1 Then
			.Parameters.Append .CreateParameter("orderProcessingSuccess", adInteger, adParamInput, 1, 1)
		Else
			.Parameters.Append .CreateParameter("orderProcessingSuccess", adInteger, adParamInput, 1, 0)
		End If
		.Parameters.Append .CreateParameter("orderProcessingDate", adDBTimeStamp, adParamInput, 16, Now())
		.Execute , , adExecuteNoRecords
		
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Function	'writeOrderProcessingLogEntry

'**************************************************

Function cleanOrderProcessingLog(byVal bytDaysToKeep)

Dim pobjCmd
Dim pstrSQL

	pstrSQL = "Delete From ssOrderProcessingLog Where orderProcessingDate<?"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		
		.Parameters.Append .CreateParameter("orderProcessingDate", adDBTimeStamp, adParamInput, 16, DateAdd("d", -1 * bytDaysToKeep, Now()))
		.Execute , , adExecuteNoRecords
		
	End With	'pobjCmd
	Set pobjCmd = Nothing
	Call writeOrderProcessingLogEntry(Array(0, "", "Log entries prior to " &  DateAdd("d", -1 * bytDaysToKeep, Now()) & " deleted.", 1))

End Function	'cleanOrderProcessingLog

'**************************************************

Sub setOrderLevelReplacements(ByRef objclsEmail, ByRef objRS)

Dim OrderNodes
Dim FieldCounter

	With objclsEmail
	
		'Call .setReplacementValue("storeName", StoreName)
		Call .setReplacementValue("orderNumber", objRS.Fields("orderID").Value)
		
		Call .setReplacementValue("recipientFirstName", Trim(objRS.Fields("cshpaddrShipFirstName").Value & ""))
		Call .setReplacementValue("recipientLastName", Trim(objRS.Fields("cshpaddrShipLastName").Value & ""))
		Call .setReplacementValue("recipientMI", Trim(objRS.Fields("cshpaddrShipMiddleInitial").Value & ""))
		Call .setReplacementValue("recipientAddress1", Trim(objRS.Fields("cshpaddrShipAddr1").Value & ""))
		'Call .setReplacementValue("recipientCompany", Trim(objRS.Fields("cshpaddrShipCompany").Value & ""))
		If Len(Trim(objRS.Fields("cshpaddrShipCompany").Value & "")) > 0 Then
			Call .setReplacementValue("recipientCompany", vbcrlf & Trim(objRS.Fields("cshpaddrShipCompany").Value & ""))
		Else
			Call .setReplacementValue("recipientCompany", "")
		End If
		If Len(Trim(objRS.Fields("cshpaddrShipAddr2").Value & "")) > 0 Then
			Call .setReplacementValue("recipientAddress2", vbcrlf & Trim(objRS.Fields("cshpaddrShipAddr2").Value & ""))
		Else
			Call .setReplacementValue("recipientAddress2", "")
		End If
		Call .setReplacementValue("recipientCity", Trim(objRS.Fields("cshpaddrShipCity").Value & ""))
		Call .setReplacementValue("recipientState", Trim(objRS.Fields("cshpaddrShipState").Value & ""))
		Call .setReplacementValue("recipientZip", Trim(objRS.Fields("cshpaddrShipZip").Value & ""))
		Call .setReplacementValue("recipientCountry", Trim(objRS.Fields("cshpaddrShipCountry").Value & ""))
		Call .setReplacementValue("recipientPhone", Trim(objRS.Fields("cshpaddrShipPhone").Value & ""))
		Call .setReplacementValue("recipientFax", Trim(objRS.Fields("cshpaddrShipFax").Value & ""))
		Call .setReplacementValue("recipientEmail", Trim(objRS.Fields("cshpaddrShipEmail").Value & ""))
		Call .setReplacementValue("shipMethod", Trim(objRS.Fields("orderShipMethod").Value & ""))

		For FieldCounter = 0 To objRS.Fields.Count - 1
			If objRS.Fields(FieldCounter).Name <> "upsize_ts" Then Call .setReplacementValue(objRS.Fields(FieldCounter).Name, Trim(objRS.Fields(FieldCounter).Value & ""))
		Next 'FieldCounter

	End With	'objclsEmail
	
End Sub	'setOrderLevelReplacements

'**************************************************

Sub setOrderDetailReplacements(ByRef objclsEmail, ByRef objRS)

Dim FieldCounter
Dim pblnLocalDebug
Dim pstrTemp

	pblnLocalDebug = True	'True False
	With objclsEmail
		Call .clearReplacementValues_Repeating
		
		For FieldCounter = 0 To objRS.Fields.Count - 1
            'debugprint FieldCounter, objRS.Fields(FieldCounter).Name
			If objRS.Fields(FieldCounter).Name <> "upsize_ts" Then Call .setReplacementValue_Repeating(objRS.Fields(FieldCounter).Name, Trim(objRS.Fields(FieldCounter).Value & ""))
		Next 'FieldCounter
		
		Call .setReplacementValue_Repeating("productID", Trim(objRS.Fields("odrdtProductID").Value & ""))
		Call .setReplacementValue_Repeating("productName", Trim(objRS.Fields("odrdtProductName").Value & ""))
		Call .setReplacementValue_Repeating("quantity", Trim(objRS.Fields("odrdtQuantity").Value & ""))
		
		'Now Set the mfg/location codes
		Dim pstrAttributeName
		Dim pstrAttributeDetailName
	    ReDim paryAttributeDetails(1)
	    pstrAttributeName = Trim(objRS.Fields("odrattrName").Value & " ")
		
		If pblnLocalDebug Then
			Response.Write "<fieldset><legend>Order: " & objRS.Fields("orderID").Value & "</legend>"
			Response.Write "Product: " & .getReplacementValue_Repeating("productID") & " - " & .getReplacementValue_Repeating("productName") & "<BR>"
			On Error Resume Next
			Response.Write "Product Level: VenderNumber: " & objRS.Fields("prodVenderNumber").Value & "<BR>"
			Response.Write "Product Level: MFGNumber: " & objRS.Fields("prodMFGNumber").Value & "<BR>"
			On Error Goto 0
		End If	'pblnLocalDebug
	    If Len(pstrAttributeName) > 0 Then
		    pstrAttributeDetailName = Trim(objRS.Fields("odrattrAttribute").Value)

		    If InStr(1, pstrAttributeName, Left(pstrAttributeDetailName, Len(pstrAttributeName))) > 0 Then pstrAttributeDetailName = Replace(pstrAttributeDetailName, pstrAttributeName, "", 1, 1)
			If pblnLocalDebug Then
				Response.Write "Attribute: " & pstrAttributeName & "<BR>"
				Response.Write "Attribute Detail: " & pstrAttributeDetailName & "<BR>"
			End If	'pblnLocalDebug
    	
		    Call getProductAttributeDetails(objRS.Fields("odrdtProductID").Value, pstrAttributeName, pstrAttributeDetailName, Array("attrdtSKU", "attrdtExtra"), paryAttributeDetails)
			
			'Reverse attributes per Greg's direction
			pstrTemp = paryAttributeDetails(0)
			paryAttributeDetails(0) = paryAttributeDetails(1)
			paryAttributeDetails(1) = pstrTemp
			 
		    'If Len(paryAttributeDetails(0)) = 0 Then paryAttributeDetails(0) = getNameFromID("sfManufacturers", "mfgName", "mfgID", True, objRS.Fields("odrdtManufacturer").Value)
		    'If Len(paryAttributeDetails(1)) = 0 Then paryAttributeDetails(1) = getNameFromID("sfVendors", "vendName", "vendID", True, objRS.Fields("odrdtVendor").Value)
			If pblnLocalDebug Then
				Response.Write "Attribute Level: MFGNumber: " & paryAttributeDetails(0) & "<BR>"
				Response.Write "Attribute Level: VenderNumber: " & paryAttributeDetails(1) & "<BR>"
			End If	'pblnLocalDebug
		Else
			On Error Resume Next
		    paryAttributeDetails(0) = objRS.Fields("prodVenderNumber").Value
		    paryAttributeDetails(1) = objRS.Fields("prodMFGNumber").Value
		    If Err.number <> 0 Then Err.Clear
	    End If
		Call .setReplacementValue_Repeating("ManufactuerCode", paryAttributeDetails(0))
		Call .setReplacementValue_Repeating("LocationCode", paryAttributeDetails(1))

		If pblnLocalDebug Then
			Response.Write "<hr />Manufactuer Code: " & .getReplacementValue_Repeating("ManufactuerCode") & "<BR>"
			Response.Write "Location Code: " & .getReplacementValue_Repeating("LocationCode") & "<BR>"
			Response.Write "</fieldset>"
		End If	'pblnLocalDebug
	End With	'objclsEmail
	
End Sub	'setOrderDetailReplacements

'**************************************************

Function getOrderAttributeText(ByRef objRS)

Dim pstrAttribute
Dim pstrAttributeName
Dim pstrAttributeDetailName

	pstrAttributeName = Trim(objRS.Fields("odrattrName").Value & " ")
	If Len(pstrAttributeName) > 0 Then
		pstrAttributeDetailName = Trim(objRS.Fields("odrattrAttribute").Value)

		If InStr(1, pstrAttributeName, Left(pstrAttributeDetailName, Len(pstrAttributeName))) > 0 Then pstrAttributeDetailName = Replace(pstrAttributeDetailName, pstrAttributeName, "", 1, 1)

		'Remove the trailing semi-colon, if present
		If Right(pstrAttributeName, 1) = ":" Then pstrAttributeName = Left(pstrAttributeName, Len(pstrAttributeName) - 1)
		
		pstrAttribute = pstrAttributeName & ": " & pstrAttributeDetailName & vbcrlf
    End If
    
    getOrderAttributeText = pstrAttribute
	
End Function	'getOrderAttributeText

'***********************************************************************************************

Function customFormatDate(byVal dtDate, byVal strFormat)

Dim pstrDateOut
Dim plngDay
Dim plngMonth
Dim plngYear
Dim plngDay_TwoDigit
Dim plngMonth_TwoDigit
Dim plngYear_TwoDigit

	If isDate(dtDate) Then
		plngDay = Day(dtDate)
		plngMonth = Month(dtDate)
		plngDay_TwoDigit = plngDay
		plngMonth_TwoDigit = plngMonth
		plngYear = Year(dtDate)
		
		If plngDay_TwoDigit < 10 Then plngDay_TwoDigit = "0" & CStr(plngDay_TwoDigit)
		If plngMonth_TwoDigit < 10 Then plngMonth_TwoDigit = "0" & CStr(plngMonth_TwoDigit)
		plngYear_TwoDigit = Right(plngYear, 2)
		
		pstrDateOut = Replace(strFormat, "yyyy", plngYear)
		pstrDateOut = Replace(pstrDateOut, "yy", plngYear_TwoDigit)
		pstrDateOut = Replace(pstrDateOut, "mm", plngMonth_TwoDigit)
		pstrDateOut = Replace(pstrDateOut, "m", plngMonth)
		pstrDateOut = Replace(pstrDateOut, "dd", plngDay_TwoDigit)
		pstrDateOut = Replace(pstrDateOut, "d", plngDay)
	Else
		pstrDateOut = ""
	End If

	customFormatDate = pstrDateOut
	
End Function	'customFormatDate

'***********************************************************************************************


Sub SendFileToCWR(byVal strFilePath)

Dim frmPost
Dim pstrFormData
Dim pstrResult
Dim pstrURL

	Set frmPost = New MultiPartFormPost
	
	pstrURL = "https://shop.cwrelectronics.com/login.php"
	frmPost.http.open "GET", pstrURL, False
	frmPost.http.send
	'Response.Write "<div style=""border:solid 1pt black"">" & frmPost.http.responseText & "</div>"

	'Set the login
	pstrURL = "https://shop.cwrelectronics.com/index.php"
	pstrFormData = "account=248406" _
					& "&user=248406" _
					& "&password=greg406" _
					& "&Login=Login"
	frmPost.http.open "POST", pstrURL, False
	frmPost.http.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
	frmPost.http.send pstrFormData
	'Response.Write "<div style=""border:solid 1pt black"">" & frmPost.http.responseText & "</div>"

	'Send the file
	pstrURL = "https://shop.cwrelectronics.com/edi.php"
	frmPost.ClearData
	frmPost.AddFile strFilePath, "file[]"
	frmPost.AddField "submitted", "TRUE"
	frmPost.AddField "MAX_FILE_SIZE", "1048576"
	pstrResult = frmPost.Send(pstrURL)
	Response.Write "<div style=""border:solid 1pt black"">" & pstrResult & "</div>"


	'Add a file (as if you clicked the browse button on a form)
	'Add a Form Field
	'frmPost.AddField ("FieldName","FieldValue")
	'you can call the above steps as many times as nessesary
	'next, we send the output to the form
Set frmPost = Nothing

End Sub

Sub SendFileToCWR2()

	
Dim pobjXMLHTTP

'set timeouts in milliseconds
Const resolveTimeout = 10000
Const connectTimeout = 10000
Const sendTimeout = 10000
Const receiveTimeout = 100000
Dim plngCounter
Dim pstrResult
Dim pstrURL
Dim pstrFormData

'On Error Resume Next

	If Err.number <> 0 Then	Err.Clear
	'Use MSXML2 if possible - must have the Microsoft XML Parser v3 or later installed
	Set pobjXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	For plngCounter = 0 To 1	'added because of unexplained error on the first call
		With pobjXMLHTTP
			pstrURL = "https://shop.cwrelectronics.com/login.php"
			.open "GET", pstrURL, False
			.send
'Response.Write "<div style=""border:solid 1pt black"">" & .responseText & "</div>"

			'Set the login
			pstrURL = "https://shop.cwrelectronics.com/index.php"
			pstrFormData = "account=248406" _
						 & "&user=248406" _
						 & "&password=greg406" _
						 & "&Login=Login"
			.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
			.open "POST", pstrURL, False
			.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
			.send pstrFormData
'Response.Write "<div style=""border:solid 1pt black"">" & .responseText & "</div>"
						 
sHeader = "Content-Type: multipart/form-data; boundary=913114112" & vbCrLf

lpszPostData = "--913114112" & vbNewLine & "Content-Disposition:" _
			 & "multipart/form-data; name=""filePath""" & vbNewLine & vbNewLine _
			 & "This is a test of the input file post" & vbNewLine & "--913114112--"
			'Post the form
			pstrURL = "https://shop.cwrelectronics.com/edi.php"
			pstrFormData = "submitted=TRUE" _
						 & "&MAX_FILE_SIZE=1048576" _
						 & "&file[]="
			.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
			.open "POST", pstrURL, False
			.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
			.send pstrFormData
Response.Write "<div style=""border:solid 1pt black"">" & .responseText & "</div>"
			'If blnPostData Then
			'Else
			'	.open "GET", strURL, False
			'	.send
			'End If
			pstrResult  = .responseText
			'RetrieveRemoteData  = pstrResult
			'Response.Write "() Error:" & Err.number & " - " & Err.Description & " (" & Err.Source & ")" & "<br />"	
			'If pblnDebug Then Response.Write "responseText =" & .responseText & "<br />" & vbcrlf
		End With
		If Err.number <> -2147467259 Then Exit For
	Next 'plngCounter
	set pobjXMLHTTP = nothing

End Sub
























'################################################# ################################
'## Created in 2004 Robert Collyer (WWW.WEBFORUMZ.COM)
'## and Nick Jones (WWW.CACTUSOFT.COM)
'##
'## This notice
'## must remain intact in the scripts
'##
'## This code is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 
'##
'## Portions of this code are based upon work carried out by Antonin Foller, PSTRUH Software
'## and credits are duly given to him for his work.
'##
'## Support can be obtained from the ASP support forums at:
'## http://www.webforumz.com
'################################################# ################################

Class MultiPartFormPost

Public Boundary

Public http 
Private NewData 
Private PreviousData 
Private ItemString 
Private blnLastItem
Private pstrResponseText

Private Sub Class_Initialize

Const resolveTimeout = 10000
Const connectTimeout = 10000
Const sendTimeout = 10000
Const receiveTimeout = 100000

	Set http = CreateObject("MSXML2.ServerXMLHTTP")
	http.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout

End Sub

Private Sub class_Terminate()
	On Error Resume Next
	set http = nothing
End Sub

Public Property Get ResponseText
	ResponseText = pstrResponseText
End Property

Public Sub ClearData()
	ItemString = ""
End Sub

'The next two functions build up a string of form elements and files to add
'to the form post.
Public Sub AddFile(LocalPath,FieldName)
	ItemString = ItemString & "FILE^^" & LocalPath & "^^" & FieldName & "||"
End Sub

Public Sub AddField(FieldName,FieldValue)
	ItemString = ItemString & "FIELD^^" & FieldName & "^^" & FieldValue & "||"
End Sub

Public Function Send(URL)

Dim count
Dim Items
Dim ItemPart

	If Boundary = "" then Boundary = "webforumz.com-MultipartFormPost" 'Default Boundary
	'Remove the last divider element
	If Len(ItemString) > 2 Then ItemString = left(itemstring,len(itemstring) - 2)
	'Create an array of the various form elements to post
	Items = Split(ItemString,"||")
	For count = 0 to ubound(Items)
	'Preserve the Current Binary Data
	PreviousData = NewData
	'Grab the data needed to post each form element 
	ItemPart = Split(Items(Count),"^^")
	'Are we dealing with the last element?
	If count = UBound(Items) Then blnLastItem = True
	If ItemPart(0) = "FILE" Then 
		AddItem 0,ItemPart(1),ItemPart(2) 'Add File
	else 
		AddItem 1,ItemPart(1),ItemPart(2) 'Add Field
	end if
	Next 
	
	'Create HTTP object to Post the data
	http.Open "POST", URL, False
	http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" + Boundary
	'Send the Data, and grab the response.
	http.send NewData : pstrResponseText = http.responseText
	
	Send = pstrResponseText

End Function 

Private Sub AddItem(FType,arg1, arg2)

Dim objFSO
Dim Stream
Dim FileContents
Dim NewData

	If FType = 1 then 'Add field
		NewData = BuildFormData(arg1, arg2,"")
	else 'Add file
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		'Does the file exist?
		if Not objFSO.FileExists(arg1) then Exit Sub
		'Grab the file data as binary.
		Set Stream = CreateObject("ADODB.Stream")
		Stream.Type = 1 : Stream.Open : Stream.LoadFromFile arg1
		FileContents = stream.read
		'Build this elements RAW HTTP Data
		NewData = BuildFormData(FileContents, arg1, arg2)
		'Clear up!
		stream.close : set stream = nothing : Set objFSO = Nothing
	end if
End Sub

Private Function BuildFormData(arg1, arg2, arg3)

Dim pre
Dim Po
Dim RS
Dim LenData
Dim FormData

	'Have any items been added yet?
	If lenb(PreviousData) > 0 then 
		pre = vbCrLf 
	else 
		pre = "" 
	end if

	'Arg3 will be blank if dealing with a Field Element
	If arg3 <> "" then 'File Element
		'Set the Element's preceding HTTP String
		Pre = Pre & Boundary & vbCrLf & "Content-Disposition: form-data; " & _
		"name=""" & arg3 & """; filename=""" & arg2 & """" & vbCrLf & _
		"Content-Type: application/upload" & vbCrLf & vbCrLf
	else 'Field Element
		'Set the Element's preceding HTTP String
		Pre = Pre & Boundary & vbCrLf & "Content-Disposition: form-data; " & _
		"name=""" & arg1 & """" & vbcrlf & vbcrlf
	end if
	
	'Are we dealing with the last element?
	If blnLastItem then
		'Set the last element's finishing HTTP String
		Po = vbcrlf & Boundary + "--" + vbCrLf 
	else 
		Po = "" 
	end if
	
	'Create a recordset instance so we can manipulate binary data
	Set RS = CreateObject("ADODB.Recordset") 
	If arg3 <> "" then 'File Element
		RS.Fields.Append "b", 205, Len(Pre) + LenB(arg1) + Len(Po)
	else 'Field Element
		RS.Fields.Append "b", 205, Len(Pre) + Len(arg2) + Len(Po)
	end if
	RS.Open : RS.AddNew 'Create a record within the recordset object
	LenData = Len(Pre)
	'Convert the preceeding HTTP String to binary
	RS("b").AppendChunk (StringToMB(Pre) & ChrB(0))
	Pre = RS("b").GetChunk(LenData) : RS("b") = ""
	If blnLastItem then ' Last element?
		'Convert the last element's finishing HTTP String to binary
		LenData = Len(Po) : RS("b").AppendChunk (StringToMB(Po) & ChrB(0))
		Po = RS("b").GetChunk(LenData) : RS("b") = ""
	end if
	if arg3 = "" then 'Convert Field's Value to binary.
		LenData = Len(arg2) : RS("b").AppendChunk (StringToMB(arg2) & ChrB(0))
		arg2 = RS("b").GetChunk(LenData) : RS("b") = ""
	end if
	'If there was already Binary Data (form elements), then we add this in front
	'of this element's binary data
	If LenB(PreviousData) > 0 then RS("b").AppendChunk (PreviousData)
		'Add the preceding HTTP String Binary
		RS("b").AppendChunk (Pre)
		if arg3 <> "" then 'Add the Binary File Data
		RS("b").AppendChunk (Arg1) 
	else 'Add the Binary Field Value
		RS("b").AppendChunk (arg2)
	end if
	'If we are on the last element, add the finishing binary data
	If blnLastItem then RS("b").AppendChunk (Po)
	'Return the Binary to calling function
	RS.Update : FormData = RS("b") : BuildFormData = FormData
	'Clear up.
	RS.Close : Set RS = Nothing
End Function

Private Function StringToMB(S)

Dim B
Dim I

	'This function converts a string, to a binary string
	B = ""
	For I = 1 To Len(S)
	B = B & ChrB(Asc(Mid(S, I, 1)))
	Next
	StringToMB = B
End Function
End Class

'Example Usage:

'	Set frmPost = New MultiPartFormPost
'	'Add a file (as if you clicked the browse button on a form)
'	frmPost.AddFile Server.MapPath("/images/myimage.jpg","FileInputFieldName")
'	'Add a Form Field
'	frmPost.AddField ("FieldName","FieldValue")
'	'you can call the above steps as many times as nessesary
'	'next, we send the output to the form
'	frmPost.Send("http://www.myurl.com/myFormHandler.asp")
'Set frmPost = Nothing


%>


