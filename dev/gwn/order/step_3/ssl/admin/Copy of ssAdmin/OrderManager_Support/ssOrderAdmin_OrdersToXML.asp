<%
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   2.00.003												*
'*   Release Date:		December 26, 2003										*
'*   Revision Date:		November 21, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Release 2.00.003 (November 21, 2004)								        *
'*	   - Enhancement: General clean-up											*
'*																				*
'*   Release 2.00.002 (April 12, 2004)									        *
'*	   - Enhancement: support for taxable shipping added						*
'*																				*
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

'	NONE

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Global variables
'**********************************************************

Const cblnOrdersToXMLDebug = False	'True	False
Dim mstrLastOrderExported:	mstrLastOrderExported = getLastInvoiceExported

'**********************************************************
'*	Functions
'**********************************************************

'Sub addItemToArray(byRef ary, byVal vntItem)
'Sub addNode(byRef objXMLDoc, byRef objXMLNode, byVal strNodeName, byVal strText)
'Sub addNodes(byRef objXMLDoc, byRef objXMLNode, byVal objRS, byVal ary)
'Function customFormatDate(byVal dtDate)
'Function CustomTaxEntity(byRef objRS, byVal blnInitial, byVal dblTaxRate, byVal strExisting)
'Function exportOrders(byRef strOrderIDs, byVal strXSLFilePath)
'Function FormatTaxRate(byVal dblTaxCharged, byVal blnFormatTaxAsPercent)
'Function getLastInvoiceExported()
'Function getStartingInvoiceNumber(byRef objRS)
'function gettaxrate(byval dbltaxcharged, byval dblshipping, byval blnshippingtaxable, byval curdbsubtotal)
'Function LimitPossibleTaxRate(byVal dblTaxRate)
'Sub Initialize_XMLFields(byRef paryXMLOrderFields, byRef paryXMLOrderDetailFields, byRef paryXMLShippingFields, byRef paryXMLBillingFields)
'Function isShippingTaxable()
'Function LoadOrdersXML(byVal strOrderIDs, byRef objXML)
'Sub saveStartingInvoiceNumber(byVal strInvoiceNumber)
'Function ssCreateOrderXML(byRef objRS)
'Function stepCounter(lngCounter)
'Sub TestWriteXML(byRef objXMLDoc)
'Sub WriteRSData(byRef objRS)
'Function WriteXSL(byRef objXML, byVal strXSLFilePath)

'***********************************************************************************************
'***********************************************************************************************

Sub addItemToArray(byRef ary, byVal vntItem)

Dim plngNumItems

	If Err.number <> 0 Then Err.Clear
	On Error Resume Next
	plngNumItems = UBound(ary)
	If Err.number <> 0 Then
		plngNumItems= -1
		Err.Clear
	End If
	On Error Goto 0
	
	plngNumItems = plngNumItems + 1
	ReDim Preserve ary(plngNumItems)
	ary(plngNumItems) = vntItem

End Sub	'addItemToArray

'***************************************************************************************************************************************************************

Function StripHTML2(byRef asHTML)

Dim loRegExp	' Regular Expression Object
    
    Set loRegExp = New RegExp
    loRegExp.Pattern = "<[^>]*>"
    
    StripHTML = loRegExp.Replace(asHTML, "")
    
    Set loRegExp = Nothing
    
End Function	'StripHTML

Function StripHTML(byRef asHTML)

Dim paryTemp
Dim pstrOut
Dim plngPos
Dim i

	If Len(asHTML) > 0 Then
		paryTemp = Split(asHTML, "<")
		For i = 0 To UBound(paryTemp)
			plngPos = InStr(1, paryTemp(i), ">")
			If plngPos > 0 Then pstrOut = pstrOut & Right(paryTemp(i), Len(paryTemp(i)) - plngPos)
		Next 'i
	End If	'Len(asHTML) > 0
	
	If Len(pstrOut) = 0 Then pstrOut = asHTML

    StripHTML = pstrOut
    
End Function	'StripHTML

'***************************************************************************************************************************************************************

Sub addNode(byRef objXMLDoc, byRef objXMLNode, byVal strNodeName, byVal strText)

Dim pobjxmlElement

	If strNodeName = "upsize_ts" Then Exit Sub

	Set pobjxmlElement = objXMLDoc.createElement(strNodeName)
	pobjxmlElement.Text = Trim(strText & "")
	objXMLNode.appendChild pobjxmlElement

End Sub	'addNode

'***************************************************************************************************************************************************************

Sub addCDATA(byRef objXMLDoc, byRef objXMLNode, byVal strNodeName, byVal strText)

Dim pobjxmlCDATASection
Dim pobjxmlElement

	If strNodeName = "upsize_ts" Then Exit Sub

	Set pobjxmlElement = objXMLDoc.createElement(strNodeName)
	Set pobjxmlCDATASection = objXMLDoc.createCDATASection(stripIllegalCharacters(strText))
	pobjxmlElement.appendChild pobjxmlCDATASection
	objXMLNode.appendChild pobjxmlElement

End Sub	'addCDATA

'***************************************************************************************************************************************************************

Function stripIllegalCharacters(byVal strText)

Dim char
Dim i
Dim length
Dim pstrOut
Dim temp

	If isNull(strText) Then
		temp = ""
	Else
		temp = strText
	End If
	
	length = Len(temp)
	For i = 1 To length
		char = Mid(temp, i, 1)
		Select Case Asc(char)
			Case 7
				'do nothing
			Case Else
				pstrOut = pstrOut & char
		End Select
	Next
	
	stripIllegalCharacters = pstrOut

End Function	'stripIllegalCharacters

'***************************************************************************************************************************************************************

Function XMLEncode(byVal strText)

Dim pstrOut

	pstrOut = Trim(strText & "")
	
	If False Then
		pstrOut = Replace(pstrOut, "&", "&#38;#38;")
		pstrOut = Replace(pstrOut, "<", "&#38;#60;")
		pstrOut = Replace(pstrOut, ">", "&#62;")

'		pstrOut = Replace(pstrOut, Chr(34), "&#39;")
'		pstrOut = Replace(pstrOut, "'", "&#34;")
	Else
'		pstrOut = Replace(pstrOut, "&", "&amp;")
		pstrOut = Replace(pstrOut, "<", "&lt;")
		pstrOut = Replace(pstrOut, ">", "&gt;")

'		pstrOut = Replace(pstrOut, Chr(34), "&quot;")
'		pstrOut = Replace(pstrOut, "'", "&apos;")
	End If
	
	XMLEncode = pstrOut

End Function	'XMLEncode

'***************************************************************************************************************************************************************

Sub addNodes(byRef objXMLDoc, byRef objXMLNode, byVal objRS, byVal ary)

Dim i
Dim pstrNodeName
Dim pstrFieldName
Dim pstrValue

	For i = 0 To UBound(ary)
		pstrFieldName = ary(i)(0)
		pstrNodeName = ary(i)(1)
		If Len(pstrNodeName) = 0 Then pstrNodeName = pstrFieldName
		
'		On Error Resume Next
		Select Case pstrFieldName
			Case "orderAmount", "orderGrandTotal":
					pstrValue = Trim(objRS.Fields(pstrFieldName).Value & "")
					pstrValue = FormatNumber(Round(pstrValue, 2), 2, False, False, False)
			Case "TodaysDate":
					pstrValue = customFormatDate(Date())
			Case "ssDatePaymentReceived", "ssDateOrderShipped":
					pstrValue = Trim(objRS.Fields(pstrFieldName).Value & "")
					If Len(pstrValue) = 0 Then pstrValue = CStr(Date())
					pstrValue = customFormatDate(pstrValue)
			'Case "orderShipMethod":
			'		pstrValue = "UPS"
			Case "CreditCardNumberLastFour":
					pstrValue = DecryptCardNumber(objRS.Fields("payCardNumber").Value, True)
					If Len(pstrValue) > 4 Then pstrValue = Right(pstrValue, 4)
			Case "CreditCardNumber":
					pstrValue = DecryptCardNumber(objRS.Fields("payCardNumber").Value, True)
			Case "payCardNumber":
					pstrValue = DecryptCardNumber(objRS.Fields("payCardNumber").Value, False)
			Case "odrdtPrice":
					If objRS.Fields("odrdtQuantity").Value <> 0 Then
						pstrValue = objRS.Fields("odrdtSubTotal").Value / objRS.Fields("odrdtQuantity").Value
					Else
						pstrValue = objRS.Fields("odrdtSubTotal").Value
					End If
			Case "payCardType":
					pstrValue = Trim(objRS.Fields(pstrFieldName).Value & "")
					If Len(pstrValue) > 0 Then pstrValue = getNameFromID("sfTransactionTypes", "transID", "transName", False, pstrValue)
			Case "prodStateTaxIsActive":
					pstrValue = Trim(objRS.Fields("orderSTax").Value & "")
					If Len(pstrValue) = 0 Then
						pstrValue = 0
					Else
						pstrValue = CDbl(pstrValue)
					End If
					
					If pstrValue > 0 Then '
						pstrValue = Trim(objRS.Fields(pstrFieldName).Value & "")
						If Len(pstrValue) = 0 Then pstrValue = 0
						pstrValue = ConvertToBoolean(pstrValue, False)
						If pstrValue Then
							pstrValue = "Y"
						Else
							pstrValue = "N"
						End If
					Else
						pstrValue = "N"
					End If
			Case Else:
					'Need to escape returns
					pstrValue = Trim(objRS.Fields(pstrFieldName).Value & "")
					pstrValue = Replace(pstrValue, vbcrlf, "|")
		End Select
		If cblnOrdersToXMLDebug Then debugprint pstrFieldName, pstrValue
		If Err.number <> 0 Then
			Response.Write "<font color=red>Error in addNodes: Error " & Err.number & ": " & Err.Description & " - Node being added was <i>" & pstrFieldName & "</i></font><BR>" & vbcrlf
			Err.Clear
		Else
			Call addNode(objXMLDoc, objXMLNode, pstrNodeName, pstrValue)
		End If
		On Error Goto 0
	Next 'i

End Sub	'addNodes

'***********************************************************************************************

Function customFormatDate(byVal dtDate)

Dim pstrDateOut
Dim plngDay
Dim plngMonth
Dim plngYear

	If isDate(dtDate) Then
		plngDay = Day(dtDate)
		plngMonth = Month(dtDate)
		plngYear = Year(dtDate)
		
		If plngDay < 10 Then plngDay = "0" & CStr(plngDay)
		If plngMonth < 10 Then plngMonth = "0" & CStr(plngMonth)
		If True Then plngYear = Right(plngYear, 2)
		
		pstrDateOut = plngMonth & "/" & plngDay & "/" & plngYear
	Else
		pstrDateOut = ""
	End If

	customFormatDate = pstrDateOut
	
End Function	'customFormatDate

'***************************************************************************************************************************************************************

Function CustomTaxEntity(byRef objRS, byVal blnInitial, byVal dblTaxRate, byVal strExisting)

Dim pstrTaxEntity
Dim pobjRSTax
Dim pstrSQL

'Available Replacements
'
' {taxRate}
' {taxDecimal}
' {stateAbbr}
' {countyName}

	If blnInitial Then
		pstrTaxEntity = cstrTaxEntity
		If Not objRS.EOF Then
			If Len(Trim(objRS.Fields("orderSTax").Value & "")) > 0 Then
				If CDbl(objRS.Fields("orderSTax").Value) > 0 Then
					If cblnAddon_TaxRateMgr Then
						pstrSQL = "Select County From ssTaxTable Where PostalCode='" & Trim(objRS.Fields("cshpaddrShipZIP").Value & "") & "'"
						Set pobjRSTax = GetRS(pstrSQL)
						If pobjRSTax.EOF Then
							pstrTaxEntity = Replace(pstrTaxEntity, "{countyName}", "NO COUNTY")
						Else
							pstrTaxEntity = Replace(pstrTaxEntity, "{countyName}", Trim(pobjRSTax.Fields("County").Value & ""))
						End If
						Call ReleaseObject(pobjRSTax)
					End If
					pstrTaxEntity = Replace(pstrTaxEntity, "{stateAbbr}", Trim(objRS.Fields("cshpaddrShipState").Value & ""))
				End If
			End If
		End If
	Else
		pstrTaxEntity = Trim(strExisting & "")	'protect against nulls
		If dblTaxRate > 0 Then
			pstrTaxEntity = Replace(pstrTaxEntity, "{taxRate}", dblTaxRate * 100 & "%")
			pstrTaxEntity = Replace(pstrTaxEntity, "{taxDecimal}", dblTaxRate)
		Else
			pstrTaxEntity = cstrTaxEntity_NoTax
		End If
	End If

	CustomTaxEntity = pstrTaxEntity

End Function	'CustomTaxEntity

'***************************************************************************************************************************************************************

Function exportOrders(byRef strOrderIDs, byVal strXSLFilePath)

Dim objXML
Dim pstrBody

'	On Error Resume Next

	If Len(strOrderIDs) = 0 Then Exit Function
		
	If LoadOrdersXML(strOrderIDs, objXML) Then
		pstrBody = WriteXSL(objXML, strXSLFilePath)
		pstrBody = Replace(pstrBody,"<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-16" & Chr(34) & "?>", "")

		Set objXML = Nothing
	Else
		pstrBody = "LoadOrdersXML Failed"
	End If

	'Response.Write pstrBody
	exportOrders = pstrBody

End Function	'exportOrders

'***************************************************************************************************************************************************************

Sub SendOrdersXMLToResponse(byRef strOrderIDs)

Dim objXML

'	On Error Resume Next

	If Len(strOrderIDs) = 0 Then Exit Sub
	'If Right(strOrderIDs, 1) = "," Then strOrderIDs = Left(strOrderIDs, Len(strOrderIDs) - 1)	'check added because sometimes a stray comma appears
		
	If LoadOrdersXML(strOrderIDs, objXML) Then
		Response.ContentType = "text/xml"
		objXML.Save Response
		'Call debug.SaveToDisk("db\text.xml", objXML.XML, True)
		Set objXML = Nothing
	End If

End Sub	'SendOrdersXMLToResponse

'***************************************************************************************************************************************************************

Function FormatTaxRate(byVal dblTaxCharged, byVal blnFormatTaxAsPercent)

	If blnFormatTaxAsPercent Then
		FormatTaxRate = FormatNumber(CStr(dblTaxCharged * 100), 2, True) & "%"
	Else
		FormatTaxRate = dblTaxCharged
	End If
	
End Function	'FormatTaxRate

'***************************************************************************************************************************************************************

Function getLastInvoiceExported()

Dim pstrLastInvoice
Dim pobjXMLDoc

	Set pobjXMLDoc = server.CreateObject("MSXML2.DOMDocument")
	pobjXMLDoc.async = false
	If pobjXMLDoc.load(ssAdminPath & "ssOrderAdmin.xml") Then
		pstrLastInvoice = pobjXMLDoc.documentelement.selectsinglenode("LastInvoice").text
	End If
	Set pobjXMLDoc = Nothing
	
	getLastInvoiceExported = pstrLastInvoice

End Function	'getLastInvoiceExported

'***************************************************************************************************************************************************************

Function getStartingInvoiceNumber(byRef objRS)

Dim pstrLastInvoice

	If cblnUseCustomInvoiceNumber Then
		If Not objRS.EOF Then
			If Len(CStr(mstrLastOrderExported)) = 0 Then
				'Need to subtract one since it will be added back
				pstrLastInvoice = objRS.Fields("orderID").Value - 1
			Else
				pstrLastInvoice = mstrLastOrderExported
			End If
		Else
			pstrLastInvoice = mstrLastOrderExported
		End If
	End If

	getStartingInvoiceNumber = pstrLastInvoice

End Function	'getStartingInvoiceNumber

'***************************************************************************************************************************************************************

Function GetTaxRate(byVal dblTaxCharged, byVal dblShipping, byVal blnShippingTaxable, byVal curDBSubTotal)

Dim pdblTaxRate

	pdblTaxRate = 0
	If blnShippingTaxable Then
		If (curDBSubTotal + dblShipping) > 0 Then pdblTaxRate = dblTaxCharged / (curDBSubTotal + dblShipping)
	Else
		If curDBSubTotal > 0 Then pdblTaxRate = dblTaxCharged / curDBSubTotal
	End If
	
	pdblTaxRate = LimitPossibleTaxRate(pdblTaxRate)
	
	GetTaxRate = pdblTaxRate

End Function	'GetTaxRate

'***************************************************************************************************************************************************************

Function LimitPossibleTaxRate(byVal dblTaxRate)
'This function limits returned tax rates to specific values if you need precise decimals

Dim pdblTaxRateToReturn
Dim pintOrdinal
Dim pintDecimal

	If cbytTaxFraction >= 0 Then
		pintOrdinal = Int(Round(dblTaxRate * 100, 3))
		pintDecimal = dblTaxRate * 100 - pintOrdinal
		
		pintDecimal = Round(pintDecimal * cbytTaxFraction, 0)/cbytTaxFraction
		pdblTaxRateToReturn = (pintOrdinal + pintDecimal)/100
	Else
		pdblTaxRateToReturn = dblTaxRate
	End If
	
	LimitPossibleTaxRate = pdblTaxRateToReturn

End Function	'LimitPossibleTaxRate

'***********************************************************************************************

Sub Initialize_XMLFields(byRef paryXMLOrderFields, byRef paryXMLOrderDetailFields, byRef paryXMLShippingFields, byRef paryXMLBillingFields)

	'Note: array of the format: Array(dbFieldName, Node Name (defaults to dbFieldName if left blank))

	'Set the Order Fields to use
	Call addItemToArray(paryXMLOrderFields, Array("orderID", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderDate", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderAmount", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderShippingAmount", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderSTax", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderCTax", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderHandling", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderGrandTotal", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderComments", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderShipMethod", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderTradingPartner", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderHttpReferrer", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderRemoteAddress", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderPaymentMethod", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderCheckAcctNumber", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderCheckNumber", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderBankName", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderRoutingNumber", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderPurchaseOrderName", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderPurchaseOrderNumber", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderStatus", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderProcessed", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderTracking", ""))
	Call addItemToArray(paryXMLOrderFields, Array("orderIsComplete", ""))
	Call addItemToArray(paryXMLOrderFields, Array("payCardType", ""))
	Call addItemToArray(paryXMLOrderFields, Array("payCardName", ""))
	Call addItemToArray(paryXMLOrderFields, Array("payCardNumber", ""))
	Call addItemToArray(paryXMLOrderFields, Array("payCardExpires", ""))
	Call addItemToArray(paryXMLOrderFields, Array(cstrCCV, ""))
	
	'Derived fields handled individually
	Call addItemToArray(paryXMLOrderFields, Array("CreditCardNumber", ""))
	Call addItemToArray(paryXMLOrderFields, Array("CreditCardNumberLastFour", ""))
	Call addItemToArray(paryXMLOrderFields, Array("TodaysDate", ""))
	Call addItemToArray(paryXMLOrderFields, Array("ssDateOrderShipped", ""))
	
	'AE Specific Fields
	If cblnSF5AE Then
		Call addItemToArray(paryXMLOrderFields, Array("orderCouponCode", ""))
		Call addItemToArray(paryXMLOrderFields, Array("orderCouponDiscount", ""))
		Call addItemToArray(paryXMLOrderFields, Array("orderBillAmount", ""))
		Call addItemToArray(paryXMLOrderFields, Array("orderBackOrderAmount", ""))
	End If
	
	'Set the Billing Fields to use
	Call addItemToArray(paryXMLBillingFields, Array("custID", "custID"))
	Call addItemToArray(paryXMLBillingFields, Array("custFirstName", "FirstName"))
	Call addItemToArray(paryXMLBillingFields, Array("custMiddleInitial", "MiddleInitial"))
	Call addItemToArray(paryXMLBillingFields, Array("custLastName", "LastName"))
	Call addItemToArray(paryXMLBillingFields, Array("custCompany", "Company"))
	Call addItemToArray(paryXMLBillingFields, Array("custAddr1", "Addr1"))
	Call addItemToArray(paryXMLBillingFields, Array("custAddr2", "Addr2"))
	Call addItemToArray(paryXMLBillingFields, Array("custCity", "City"))
	Call addItemToArray(paryXMLBillingFields, Array("custState", "State"))
	Call addItemToArray(paryXMLBillingFields, Array("custZip", "Zip"))
	Call addItemToArray(paryXMLBillingFields, Array("custCountry", "Country"))
	Call addItemToArray(paryXMLBillingFields, Array("billToCountryName", "CountryName"))
	Call addItemToArray(paryXMLBillingFields, Array("custPhone", "Phone"))
	Call addItemToArray(paryXMLBillingFields, Array("custFax", "Fax"))
	Call addItemToArray(paryXMLBillingFields, Array("custEmail", "Email"))
	Call addItemToArray(paryXMLBillingFields, Array("custPasswd", "Password"))
	Call addItemToArray(paryXMLBillingFields, Array("custIsSubscribed", "IsSubscribed"))

	'Set the Shipping Fields to use
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipFirstName", "FirstName"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipMiddleInitial", "MiddleInitial"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipLastName", "LastName"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipCompany", "Company"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipAddr1", "Addr1"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipAddr2", "Addr2"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipCity", "City"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipState", "State"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipZip", "Zip"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipCountry", "Country"))
	Call addItemToArray(paryXMLShippingFields, Array("shipToCountryName", "CountryName"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipPhone", "Phone"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipFax", "Fax"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrShipEmail", "Email"))
	Call addItemToArray(paryXMLShippingFields, Array("cshpaddrIsSubscribed", "IsSubscribed"))

	'Set the Order Detail Fields to use
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtID", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtProductID", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtProductName", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtPrice", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtQuantity", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtSubTotal", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("prodWeight", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("prodFileName", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("prodStateTaxIsActive", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtCategory", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtManufacturer", ""))
	Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtVendor", ""))

	'Note the below fields are NOT included because they are individually addressed in the code
	If False Then
		Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtGiftWrapUnitPrice", ""))
		Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtGiftWrapQTY", ""))
		Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtGiftWrapPrice", ""))
		Call addItemToArray(paryXMLOrderDetailFields, Array("odrdtBackOrderQTY", ""))
	End If

End Sub	'Initialize_XMLFields

'***************************************************************************************************************************************************************

Function isShippingTaxable()

Dim pobjRS
Dim pstrSQL
Dim pblnShippingTaxable

	pblnShippingTaxable = False	'default value
	
	pstrSQL = "Select adminTaxShipIsActive From sfAdmin Where adminID=1"
	If Len(mlngStoreID) > 0 Then pstrSQL = "Select adminTaxShipIsActive From sfAdmin Where adminID=" & mlngStoreID
	Set pobjRS = GetRS(pstrSQL)
	If isObject(pobjRS) Then
		If Not pobjRS.EOF Then
			pblnShippingTaxable = CBool(pobjRS.Fields("adminTaxShipIsActive").Value = "1")
		End If
	End If
	Call ReleaseObject(pobjRS)
	
	isShippingTaxable = pblnShippingTaxable

End Function	'isShippingTaxable

'***************************************************************************************************************************************************************

Function LoadOrdersXML(byVal strOrderIDs, byRef objXML)

Dim pstrSQL
Dim pstrSQL_Where
Dim pobjRS

	If Len(strOrderIDs) = 0 Then Exit Function
		
	pstrSQL_Where = Replace(strOrderIDs, ", ", " Or orderID=")
	pstrSQL_Where = Replace(strOrderIDs, ",", " Or orderID=")

	If cblnSF5AE Then
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderAmount, sfOrders.orderComments, sfOrders.orderShipMethod, sfOrders.orderTradingPartner, sfOrders.orderHttpReferrer, sfOrders.orderCTax, sfOrders.orderSTax, sfOrders.orderShippingAmount, sfOrders.orderHandling, sfOrders.orderGrandTotal, sfOrders.orderPaymentMethod, sfOrders.orderRemoteAddress, sfOrders.orderCheckAcctNumber, sfOrders.orderCheckNumber, sfOrders.orderBankName, sfOrders.orderRoutingNumber, sfOrders.orderPurchaseOrderName, sfOrders.orderPurchaseOrderNumber, sfOrders.orderStatus, sfOrders.orderProcessed, sfOrders.orderTracking, sfOrders.orderIsComplete, " _
				& "		  ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssTrackingNumber, " _
				& "		  sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtPrice, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, sfOrderDetails.odrdtCategory, sfOrderDetails.odrdtManufacturer, sfOrderDetails.odrdtVendor," _
				& "		  sfOrderAttributes.odrattrID, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, " _
				& "		  sfCPayments.payCardType, sfCPayments.PayCardName, sfCPayments.payCardNumber, sfCPayments.payCardExpires, sfCPayments." & cstrCCV & "," _
				& "		  sfCustomers.custID, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custLastName, sfCustomers.custCompany, sfCustomers.custAddr1, sfCustomers.custAddr2, sfCustomers.custCity, sfCustomers.custState, sfCustomers.custZip, sfCustomers.custCountry, sfCustomers.custPhone, sfCustomers.custFax, sfCustomers.custEmail, sfCustomers.custPasswd, sfCustomers.custIsSubscribed, " _
				& "		  sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipMiddleInitial, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipPhone, sfCShipAddresses.cshpaddrShipFax, sfCShipAddresses.cshpaddrShipEmail, sfLocalesCountry1.loclctryName AS shipToCountryName, sfLocalesCountry.loclctryName AS billToCountryName, '' AS cshpaddrIsSubscribed, " _
				& "		  sfOrdersAE.orderCouponCode, sfOrdersAE.orderBillAmount, sfOrdersAE.orderBackOrderAmount, sfOrdersAE.orderCouponDiscount, " _
				& "		  sfOrderDetailsAE.odrdtGiftWrapPrice, sfOrderDetailsAE.odrdtGiftWrapQTY, sfOrderDetailsAE.odrdtBackOrderQTY, sfOrderDetailsAE.odrdtAttDetailID, " _
				& "		  sfProducts.prodWeight, sfProducts.prodFileName, sfProducts.prodStateTaxIsActive" _
				& " FROM ssOrderManager RIGHT JOIN ((((sfLocalesCountry AS sfLocalesCountry1 RIGHT JOIN (sfLocalesCountry RIGHT JOIN (((((sfCustomers RIGHT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) LEFT JOIN sfCShipAddresses ON sfOrders.orderAddrId = sfCShipAddresses.cshpaddrID) LEFT JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) ON sfLocalesCountry.loclctryAbbreviation = sfCustomers.custCountry) ON sfLocalesCountry1.loclctryAbbreviation = sfCShipAddresses.cshpaddrShipCountry) LEFT JOIN sfOrdersAE ON sfOrders.orderID = sfOrdersAE.orderAEID) LEFT JOIN sfOrderDetailsAE ON sfOrderDetails.odrdtID = sfOrderDetailsAE.odrdtAEID) LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID) ON ssOrderManager.ssorderID = sfOrders.orderID" _
				& " WHERE sfOrders.orderID=" & pstrSQL_Where
	Else
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderAmount, sfOrders.orderComments, sfOrders.orderShipMethod, sfOrders.orderCTax, sfOrders.orderSTax, sfOrders.orderShippingAmount, sfOrders.orderHandling, sfOrders.orderGrandTotal, sfOrders.orderPaymentMethod, sfOrders.orderCheckAcctNumber, sfOrders.orderCheckNumber, sfOrders.orderBankName, sfOrders.orderRoutingNumber, sfOrders.orderPurchaseOrderName, sfOrders.orderPurchaseOrderNumber, sfOrders.orderStatus, sfOrders.orderProcessed, sfOrders.orderTracking, sfOrders.orderIsComplete, sfOrders.orderTradingPartner, sfOrders.orderHttpReferrer, sfOrders.orderRemoteAddress, " _
				& "		  ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssTrackingNumber, " _
				& "		  sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtPrice, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, sfOrderDetails.odrdtCategory, sfOrderDetails.odrdtManufacturer, sfOrderDetails.odrdtVendor," _
				& "		  sfOrderAttributes.odrattrID, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, " _
				& "		  sfCPayments.payCardType, sfCPayments.payCardName, sfCPayments.payCardNumber, sfCPayments.payCardExpires, " _
				& "		  sfCustomers.custID, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custLastName, sfCustomers.custCompany, sfCustomers.custAddr1, sfCustomers.custAddr2, sfCustomers.custCity, sfCustomers.custState, sfCustomers.custZip, sfCustomers.custCountry, sfCustomers.custPhone, sfCustomers.custFax, sfCustomers.custEmail, sfCustomers.custPasswd, sfCustomers.custIsSubscribed, " _
				& "		  sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipMiddleInitial, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipPhone, sfCShipAddresses.cshpaddrShipFax, sfCShipAddresses.cshpaddrShipEmail, sfLocalesCountry1.loclctryName AS shipToCountryName, sfLocalesCountry.loclctryName AS billToCountryName, 0 As cshpaddrIsSubscribed, " _
				& "		  sfProducts.prodWeight, sfProducts.prodFileName, sfProducts.prodStateTaxIsActive, " _
				& "		  '' As orderCouponCode, 0 As orderBillAmount, 0 As orderBackOrderAmount, 0 As orderCouponDiscount, " _
				& "		  '' As odrdtGiftWrapPrice, 0 As odrdtGiftWrapQTY, 0 As odrdtBackOrderQTY, '' As odrdtAttDetailID, " _
				& "		  sfTransactionResponse.trnsrspSuccess, sfTransactionResponse.trnsrspErrorLocation, sfTransactionResponse.trnsrspErrorMsg, sfTransactionResponse.trnsrspAuthNo, sfTransactionResponse.trnsrspRetrievalCode, sfTransactionResponse.trnsrspAVSCode, sfTransactionResponse.trnsrspMerchTransNo, sfTransactionResponse.trnsrspCustTransNo" _
				& " FROM ssOrderManager RIGHT JOIN (((sfLocalesCountry AS sfLocalesCountry1 RIGHT JOIN (sfLocalesCountry RIGHT JOIN (((((sfCustomers RIGHT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) LEFT JOIN sfCShipAddresses ON sfOrders.orderAddrId = sfCShipAddresses.cshpaddrID) LEFT JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) ON sfLocalesCountry.loclctryAbbreviation = sfCustomers.custCountry) ON sfLocalesCountry1.loclctryAbbreviation = sfCShipAddresses.cshpaddrShipCountry) LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID) LEFT JOIN sfTransactionResponse ON sfOrders.orderID = sfTransactionResponse.trnsrspOrderId) ON ssOrderManager.ssorderID = sfOrders.orderID" _
				& " WHERE sfOrders.orderID=" & pstrSQL_Where
	End If
	pstrSQL = pstrSQL & " Order By sfOrders.orderID Asc, sfOrderDetails.odrdtID Asc, sfOrderAttributes.odrattrID Asc"
	
	'Custom insertion
	'pstrSQL = Replace(pstrSQL, "sfOrders.orderID,", "sfOrders.orderID, sfProducts.UPC, sfProducts.Brand, sfProducts.QtyinPallet, sfProducts.GWeight, sfProducts.prodSize,")
	'Response.Write("LoadOrdersXML: pstrSQL = " & pstrSQL & "<BR>")

	If cblnOrdersToXMLDebug Then
		debugprint "cblnSF5AE", 	cblnSF5AE
		debugprint "pstrSQL", 	pstrSQL
	End If
	
	Set pobjRS = server.CreateObject("adodb.Recordset")
	pobjRS.CursorLocation = cExport_CursorLocation

	On Error Resume Next
	pobjRS.open pstrSQL, cnn, adOpenForwardOnly, adLockReadOnly
	If Err.number <> 0 Then
		Response.Write "<font color=red>Error in LoadOrdersXML: Error " & Err.number & ": " & Err.Description & "</font><BR>" & vbcrlf
		Response.Write "<font color=red>cblnSF5AE = " & cblnSF5AE & "</font><BR>" & vbcrlf
		Response.Write "<font color=red>pstrSQL = " & pstrSQL & "</font><BR>" & vbcrlf
		Response.Flush
		LoadOrdersXML = False
	Else
		On Error Goto 0
		If cblnOrdersToXMLDebug Then Call WriteRSData(pobjRS)
		
		'Load the XML object with XML from the database
		Set objXML = Server.CreateObject("MSXML2.DOMDocument")
		objXML.async = false
		If objXML.LoadXML(ssCreateOrderXML(pobjRS)) Then
			LoadOrdersXML = True
			If cblnOrdersToXMLDebug Then
				Response.Write "<fieldset><legend>LoadOrdersXML - XML</legend>"
				Call TestWriteXML(objXML)
				Response.Write "</fieldset>"
			End If
		Else
			LoadOrdersXML = False
		End If
		pobjRS.Close
	End If
	Set pobjRS = Nothing

End Function	'LoadOrdersXML

'***************************************************************************************************************************************************************

Sub saveStartingInvoiceNumber(byVal strInvoiceNumber)

Dim pobjElement

	Set pobjXMLDoc = server.CreateObject("MSXML2.DOMDocument")
	pobjXMLDoc.async = false
	If pobjXMLDoc.load(ssAdminPath & "ssOrderAdmin.xml") Then
		Set pobjElement = pobjXMLDoc.documentelement.selectsinglenode("LastInvoice")
		pobjElement.text = strInvoiceNumber
		
		On Error Resume Next
		pobjXMLDoc.Save(ssAdminPath & "ssOrderAdmin.xml")
		If Err.number <> 0 Then
			Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><BR>" & vbcrlf
			Response.Write "Error saving last invoice number to disk. <BR>" & vbcrlf
			Response.Write "File: <i>" & ssAdminPath & "ssOrderAdmin.xml" & "</i><BR>" & vbcrlf
			Response.Write "This is likely due to MSXML not having write permissions. Please contact your host.<BR>" & vbcrlf
			Err.Clear
		End If
	End If
	Set pobjXMLDoc = Nothing
	
End Sub	'saveStartingInvoiceNumber

'***********************************************************************************************

Function ssCreateOrderXML(byRef objRS)
'This function creates the following:
'orders
'- order (xmlOrderDetail)
'  - billingAddress (xmlAddress)
'  - shippingAddress (xmlAddress - reused)
'  - orderDetail (xmlProduct)

Dim xmlDoc
Dim xmlRoot
Dim xmlNode
Dim xmlOrderDetail
Dim xmlProduct
Dim xmlProductAttr
Dim xmlAddress
Dim xmlElement
Dim plngCurrentOrderID: plngCurrentOrderID = 0
Dim pstrSQL

Dim pstrBackOrderQty
Dim pstrGWQty
Dim pstrGWPrice

Dim plngCounter
Dim pdblOrderWeight
Dim plngodrdtID
Dim plngAttdtID
Dim pstrProdID

Dim pcurGrandTotal
Dim pcurAccessories
Dim pcurRealSubTotal
Dim pcurShipping
Dim pcurDBSubTotal
Dim pcurcalcDiscount
Dim pcurCouponDiscount
Dim pdblTaxCharged
Dim pdblTaxRate
Dim pblnisShippingTaxable
Dim pstrInvoiceNumber
Dim pstrAttrCategory, pstrAttrName

Dim pstrXMLOut

Dim paryXMLOrderFields()
Dim paryXMLOrderDetailFields()
Dim paryXMLShippingFields()
Dim paryXMLBillingFields()
Dim pstrShortOrderDate
Dim pstrTaxEntity
Dim pblnWriteAdditionalNodes

Dim plngItemOnBackOrder:	plngItemOnBackOrder = 0

	Call Initialize_XMLFields(paryXMLOrderFields, paryXMLOrderDetailFields, paryXMLShippingFields, paryXMLBillingFields)

	set xmlDoc = server.CreateObject("MSXML2.DOMDocument")
	' Create processing instruction and document root
    Set xmlNode = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    Set xmlNode = xmlDoc.insertBefore(xmlNode, xmlDoc.childNodes.Item(0))
    
	' Create document root
    Set xmlRoot = xmlDoc.createElement("orders")
    Set xmlDoc.documentElement = xmlRoot
    xmlRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"

	With objRS
		pstrInvoiceNumber = getStartingInvoiceNumber(objRS)
		Do While Not .EOF
		
			If plngCurrentOrderID <> .Fields("orderID").Value Then
				plngCurrentOrderID = .Fields("orderID").Value
				If cblnUseCustomInvoiceNumber Then
					pstrInvoiceNumber = pstrInvoiceNumber + 1
				Else
					pstrInvoiceNumber = plngCurrentOrderID
				End If

				If Not .EOF Then
					pstrShortOrderDate = customFormatDate(.Fields("orderDate").Value)
					pstrShortOrderDate = Replace(pstrShortOrderDate, "200", "0")
					pcurGrandTotal = .Fields("orderGrandTotal").Value
					pcurGrandTotal = .Fields("orderGrandTotal").Value
					pstrTaxEntity = CustomTaxEntity(objRS, True, 0, "")
					
					pcurAccessories = 0
					pcurShipping = Trim(.Fields("orderShippingAmount").Value & "")
					If Len(pcurShipping) > 0 Then
						pcurShipping = CDbl(pcurShipping)
					Else
						pcurShipping = 0
					End If
					pcurAccessories = pcurAccessories + pcurShipping
					If Len(Trim(.Fields("orderSTax").Value & "")) > 0 Then
						pdblTaxCharged = CDbl(.Fields("orderSTax").Value)
						pcurAccessories = pcurAccessories + pdblTaxCharged
					End If
					If Len(Trim(.Fields("orderCTax").Value & "")) > 0 Then pcurAccessories = pcurAccessories + CDbl(.Fields("orderCTax").Value)
					If Len(Trim(.Fields("orderHandling").Value & "")) > 0 Then pcurAccessories = pcurAccessories + CDbl(.Fields("orderHandling").Value)

					pcurCouponDiscount = 0
					If cblnSF5AE Then
						If Len(Trim(.Fields("orderCouponDiscount").Value & "")) > 0 Then pcurCouponDiscount = CDbl(.Fields("orderCouponDiscount").Value)
					End If

					pcurDBSubTotal = Round(CDbl(.Fields("orderAmount").Value), 2)
				End If
			
				'Create Order Node
				Set xmlOrderDetail = xmlDoc.createElement("order")
				xmlRoot.appendChild xmlOrderDetail
	    
				xmlOrderDetail.setAttribute "id", .Fields("orderID").Value
				'Set xmlElement = xmlDoc.createElement("emptySpace")
				'xmlElement.Text = "&nbsp;"
				'Set xmlElement = xmlDoc.createCDATASection("&nbsp;")
				'xmlOrderDetail.appendChild xmlElement
		
				Call addNodes(xmlDoc, xmlOrderDetail, objRS, paryXMLOrderFields)
				
				'Create Billing Address
				Set xmlAddress = xmlDoc.createElement("billingAddress")
				xmlOrderDetail.appendChild xmlAddress
				Call addNodes(xmlDoc, xmlAddress, objRS, paryXMLBillingFields)
			    
				'Shipping Address Details
	'			pstrNodeName = "": Call addNode(xmlDoc, xmlAddress, pstrNodeName, .Fields(pstrNodeName).Value)
				Set xmlAddress = xmlDoc.createElement("shippingAddress")
				xmlOrderDetail.appendChild xmlAddress
				Call addNodes(xmlDoc, xmlAddress, objRS, paryXMLShippingFields)
			    
				'Order Payment Details
				pcurRealSubTotal = 0
				pdblOrderWeight = 0
				plngCounter = 0


			End If	'plngCurrentOrderID <> .Fields("orderID").Value
			
			If plngodrdtID <> Trim(.Fields("odrdtID").Value) Then
				plngodrdtID = Trim(.Fields("odrdtID").Value)
				pstrProdID = Trim(.Fields("odrdtProductID").Value)

				Set xmlProduct = xmlDoc.createElement("orderDetail")
				xmlOrderDetail.appendChild xmlProduct
				xmlProduct.setAttribute "id", .Fields("odrdtID").Value
				
				Call addNodes(xmlDoc, xmlProduct, objRS, paryXMLOrderDetailFields)

				pcurRealSubTotal = pcurRealSubTotal + CDbl(.Fields("odrdtSubTotal").Value)
				If isNumeric(.Fields("prodWeight").Value) Then pdblOrderWeight = pdblOrderWeight + CDbl(.Fields("prodWeight").Value)

				pstrGWQty = Trim(.Fields("odrdtGiftWrapQTY").Value & "")
				If Len(pstrGWQty) > 0 Then
					pstrGWPrice = objRS.Fields("odrdtGiftWrapPrice").Value
					If Not isNumeric(pstrGWPrice) Then pstrGWPrice = 0
					pcurRealSubTotal = pcurRealSubTotal + CDbl(pstrGWPrice)
					If CLng(pstrGWQty) > 0 Then pstrGWPrice = pstrGWPrice / pstrGWQty
				Else
					pstrGWQty = 0
					pstrGWPrice = 0
				End If

				Call addNode(xmlDoc, xmlProduct, "odrdtGiftWrapQTY", pstrGWQty)
				Call addNode(xmlDoc, xmlProduct, "odrdtGiftWrapPrice", pstrGWPrice * pstrGWQty)
				Call addNode(xmlDoc, xmlProduct, "odrdtGiftWrapUnitPrice", Round(pstrGWPrice,2))

				'Check for Back Orders
				pstrBackOrderQty = Trim(.Fields("odrdtBackOrderQTY").Value & "")
				If Len(pstrBackOrderQty) > 0 Then plngItemOnBackOrder = plngItemOnBackOrder + CLng(pstrBackOrderQty)
				Call addNode(xmlDoc, xmlProduct, "odrdtBackOrderQTY", objRS.Fields("odrdtBackOrderQTY").Value)

			End If	'plngodrdtID <> Trim(.Fields("odrdtID").Value)
				
			If (plngAttdtID <> Trim(.Fields("odrattrID").Value)) And Not isNull(.Fields("odrattrID").Value) Then
				plngAttdtID = Trim(.Fields("odrattrID").Value)
				pstrAttrCategory = Trim(.Fields("odrattrName").Value & "")
				pstrAttrName = Trim(.Fields("odrattrAttribute").Value & "")
				If Len(pstrAttrCategory & pstrAttrName) > 0 Then
					Set xmlProductAttr = xmlDoc.createElement("odrdtAttDetailID")
					xmlProduct.appendChild xmlProductAttr

					Set xmlElement = xmlDoc.createElement("odrattrName")
					xmlElement.Text = pstrAttrCategory
					xmlProductAttr.appendChild xmlElement
					
					Set xmlElement = xmlDoc.createElement("odrattrAttribute")
					xmlElement.Text = pstrAttrName
					xmlProductAttr.appendChild xmlElement
				End If
			End If	'plngAttdtID <> Trim(.Fields("odrattrID").Value)

			.MoveNext

			'check to see if new order
			pblnWriteAdditionalNodes = .EOF
			If Not pblnWriteAdditionalNodes Then pblnWriteAdditionalNodes = CBool(plngCurrentOrderID <> .Fields("orderID").Value)
			If pblnWriteAdditionalNodes Then

				'add the ProcessorResponse
				Dim pobjRSResponse
				Dim pobjXMLElement
				Dim fieldCounter
				pstrSQL = "Select * From sfTransactionResponse Where trnsrspOrderId=" & plngCurrentOrderID
				Set pobjRSResponse = GetRS(pstrSQL)
				Do While Not pobjRSResponse.EOF
					Set pobjXMLElement = xmlDoc.createElement("ProcessorResponse")
					xmlOrderDetail.appendChild pobjXMLElement
					For fieldCounter = 1 To pobjRSResponse.Fields.Count
						Select Case pobjRSResponse.Fields(fieldCounter-1).Name
							Case "FullResponse", "trnsrspErrorMsg"
								Call addCDATA(xmlDoc, pobjXMLElement, pobjRSResponse.Fields(fieldCounter-1).Name, pobjRSResponse.Fields(fieldCounter-1).Value)
							Case Else
								Call addNode(xmlDoc, pobjXMLElement, pobjRSResponse.Fields(fieldCounter-1).Name, pobjRSResponse.Fields(fieldCounter-1).Value)
						End Select
					Next 'fieldCounter
					pobjRSResponse.MoveNext
				Loop
				Call ReleaseObject(pobjRSResponse)

				pcurcalcDiscount = pcurGrandTotal - pcurRealSubTotal - pcurCouponDiscount - pcurAccessories
				pcurcalcDiscount = Abs(Round(pcurcalcDiscount, 2))

				pblnisShippingTaxable = isShippingTaxable
				pdblTaxRate = GetTaxRate(pdblTaxCharged, pcurShipping, pblnisShippingTaxable, CDbl(pcurDBSubTotal-pcurcalcDiscount))
				pstrTaxEntity = CustomTaxEntity(objRS, False, pdblTaxRate, pstrTaxEntity)

				If False Then
					Response.Write "<fieldset><legend>Calc Discount</legend>"
					Response.Write "pcurRealSubTotal: " & pcurRealSubTotal & "<BR>"
					Response.Write "pcurcalcDiscount: " & pcurcalcDiscount & "<BR>"
					Response.Write "pcurDBSubTotal: " & pcurDBSubTotal & "<BR>"
					Response.Write "pcurCouponDiscount: " & pcurCouponDiscount & "<BR>"
					Response.Write "pcurAccessories: " & pcurAccessories & "<BR>"
					Response.Write "pcurGrandTotal: " & pcurGrandTotal & "<BR>"
					Response.Write "</fieldset>"

					Response.Write "<table border=1 cellspacing=0 cellpadding=1>"
					Response.Write "<tr><td>Calc SubTotal: </td><td>" & pcurRealSubTotal & "</td>"
					Response.Write "<tr><td>DB SubTotal: </td><td>" & pcurDBSubTotal & "</td>"
					Response.Write "<tr><td>Calc Discount: </td><td>" & pcurcalcDiscount & "</td>"
					Response.Write "<tr><td>Accessories: </td><td>" & pcurAccessories & "</td>"
					Response.Write "<tr><td>DB GrandTotal: </td><td>" & pcurGrandTotal & "</td>"
					Response.Write "<tr><td>calc GrandTotal: </td><td>" & CDbl(pcurDBSubTotal) + CDbl(pcurAccessories) & "</td>"
					Response.Write "</table>"
				End If

				If pdblTaxRate > 0 Then
					Call addNode(xmlDoc, xmlOrderDetail, "Taxable", "Y")
					Call addNode(xmlDoc, xmlOrderDetail, "TaxCalcMeth", "AUTOSTAX")
				Else
					Call addNode(xmlDoc, xmlOrderDetail, "Taxable", "N")
					Call addNode(xmlDoc, xmlOrderDetail, "TaxCalcMeth", "AUTOSTAX")
				End If
				Call addNode(xmlDoc, xmlOrderDetail, "TaxEntity", pstrTaxEntity)
				Call addNode(xmlDoc, xmlOrderDetail, "TaxRate", FormatTaxRate(pdblTaxRate, True))
				Call addNode(xmlDoc, xmlOrderDetail, "TaxDecimal", FormatTaxRate(pdblTaxRate, False))
				Call addNode(xmlDoc, xmlOrderDetail, "ShippingIsTaxable", pblnisShippingTaxable)
				Call addNode(xmlDoc, xmlOrderDetail, "OrderWeight", pdblOrderWeight)
				Call addNode(xmlDoc, xmlOrderDetail, "orderDiscount", pcurcalcDiscount)
				Call addNode(xmlDoc, xmlOrderDetail, "orderSubTotal", pcurRealSubTotal)
				Call addNode(xmlDoc, xmlOrderDetail, "ItemOnBackOrder", plngItemOnBackOrder)
				Call addNode(xmlDoc, xmlOrderDetail, "shortOrderDate", pstrShortOrderDate)
				Call addNode(xmlDoc, xmlOrderDetail, "InvoiceNumber", cstrInvoiceOrderPrefix & pstrInvoiceNumber)
				
				If cblnAddon_GCMgr Then 
					If GC_LoadByOrder(plngCurrentOrderID, pcurGrandTotal) Then
						Dim xmlGiftCertficate
						Set xmlGiftCertficate = xmlDoc.createElement("ssGiftCertificate")
						xmlOrderDetail.appendChild xmlGiftCertficate

						Set xmlElement = xmlDoc.createElement("CertificateNumber")
						xmlElement.Text = mstrCertificate
						xmlGiftCertficate.appendChild xmlElement

						Set xmlElement = xmlDoc.createElement("RedemptionAmount")
						xmlElement.Text = mdblssCertificateAmount
						xmlGiftCertficate.appendChild xmlElement
						
						Set xmlElement = xmlDoc.createElement("ssGCNewTotalDue")
						xmlElement.Text = mdblssGCNewTotalDue
						xmlGiftCertficate.appendChild xmlElement
						
					End If	'GC_LoadByOrder 
				End If	'cblnAddon_GCMgr
				
			End If	'pblnWriteAdditionalNodes
		Loop

	End With

	pstrXMLOut = xmlDoc.xml

	Set xmlElement = Nothing
	Set xmlDoc = Nothing
	
	mstrLastOrderExported = pstrInvoiceNumber

	ssCreateOrderXML = pstrXMLOut
	
	'Response.Write "<hr>" & vbcrlf & vbcrlf & pstrXMLOut & vbcrlf  & vbcrlf & "<hr>"

End Function	'ssCreateOrderXML

'***********************************************************************************************

Function stepCounter(lngCounter)

	lngCounter = lngCounter + 1
	stepCounter = lngCounter
End Function	'stepCounter

'***********************************************************************************************

Sub TestWriteXML(byRef objXMLDoc)
	If cblnOrdersToXMLDebug Then	
		Response.Write objXMLDoc.xml
	End If
End Sub	'TestWriteXML

'***********************************************************************************************

Sub WriteRSData(byRef objRS)

Dim i
Dim j

	Response.Write "<table border=1 cellpadding=0 cellspacing=0>"
	Response.Write "<tr>"
	Response.Write "<td>&nbsp;</td>"
	For i = 1 To objRS.Fields.Count
		Response.Write "<th>" & objRS.Fields(i-1).Name & "</th>"
	Next 'i
	Response.Write "</tr>"
	
	j = 0
	Do While Not objRS.EOF
		j = j + 1
		Response.Write "<tr>"
		Response.Write "<td>" & j & "</td>"
		For i = 1 To objRS.Fields.Count
			Response.Write "<th>" & objRS.Fields(i-1).Value & "&nbsp;</th>"
		Next 'i
		Response.Write "</tr>"
		objRS.Movenext
	Loop
	Response.Write "</table>"
	objRS.MoveFirst

End Sub	'WriteRSData

'***************************************************************************************************************************************************************

Function WriteXSL(byRef objXML, byVal strXSLFilePath)

Dim objXSL
Dim strOutput

	' Load the XSL from the XSL file
	set objXSL = Server.CreateObject("MSXML2.DOMDocument")
	objXSL.async = false
	'objXSL.preserveWhiteSpace = True
	'debugprint "strXSLFilePath", strXSLFilePath
	If objXSL.Load(strXSLFilePath) Then
		If cblnOrdersToXMLDebug Then
			Response.Write "<fieldset><legend>WriteXSL - XSL</legend>"
			Call TestWriteXML(objXSL)
			Response.Write "</fieldset>"
		End If
		strOutput = objXML.transformNode(objXSL)
	Else
		strOutput = "Error Loading XSL document " & strXSLFilePath & "."
	End If
	Set objXML = Nothing
	
	strOutput = Replace(strOutput,"&amp;nbsp;","&nbsp;")
	WriteXSL = strOutput

End Function	'WriteXSL

'***************************************************************************************************************************************************************

%>
