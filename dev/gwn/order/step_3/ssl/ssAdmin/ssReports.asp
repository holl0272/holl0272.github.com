<% 
'********************************************************************************
'*   Sales Report							                                    *
'*   Release Version:   1.00.002												*
'*   Release Date:		November 15, 2003										*
'*   Revision Date:		October 18, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Release 1.00.002 (October 18, 2004)										*
'*	   - Added report of sales by category										*
'*                                                                              *
'*   Release 1.00.001 (November 15, 2003)										*
'*	   - Initial release														*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Const enOrderSummary_SalesTotal = 0
Const enOrderSummary_SalesCount = 1
Const enOrderSummary_AverageOrder = 2
Const enOrderSummary_TopSellers = 3

Const enOrderDetail_orderID = 0
Const enOrderDetail_custName = 1
Const enOrderDetail_orderPaymentMethod = 2
Const enOrderDetail_orderGrandTotal = 3
Const enOrderDetail_orderDate = 4
Const enOrderDetail_orderPaymentDate = 5
Const enOrderDetail_orderCertificateRedemptions = 6

Const enCriteria_DoNotInclude = 0
Const enCriteria_True = 1
Const enCriteria_False = 2

'***********************************************************************************************

Function convertDateArray(byVal ary, byVal strDefaultTime)

	If isArray(ary) Then
		convertDateArray = ary(0) & " " & ary(1)
	ElseIf Len(ary) > 0 Then
		convertDateArray = ary & " " & strDefaultTime
	End If

End Function	'convertDateArray

'***********************************************************************************************

Function getOrderDetails(byRef ary, byVal dtStartDate, byVal dtEndDate, byVal bytPaid, byVal bytShipped, byVal bytComplete)

Dim pstrSQL
Dim pstrSQLWhere
Dim pobjRSOrders
Dim i
Dim pstrStartDate
Dim pstrEndDate

	pstrStartDate = convertDateArray(dtStartDate, "12:00:00 AM")
	pstrEndDate = convertDateArray(dtEndDate, "11:59:59 PM")

	Select Case bytComplete
		Case enCriteria_True:	pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And (sfOrders.orderIsComplete=1)"
		Case enCriteria_False:	pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And ((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete is Null))"
		Case enCriteria_DoNotInclude:	pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And ((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete=1) OR (sfOrders.orderIsComplete is Null))"
	End Select
	
	If Len(pstrStartDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " AND (orderDate >= " & wrapSQLValue(pstrStartDate, True, enDatatype_date) & ")"
	If Len(pstrEndDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate <= " & wrapSQLValue(pstrEndDate, True, enDatatype_date) & ")"

	Select Case bytPaid
		Case enCriteria_True:	pstrsqlWhere = pstrsqlWhere & " AND (ssDatePaymentReceived is not null)"
		Case enCriteria_False:	pstrsqlWhere = pstrsqlWhere & " AND (ssDatePaymentReceived is null)"
	End Select
	
	Select Case bytShipped
		Case enCriteria_True:	pstrsqlWhere = pstrsqlWhere & " AND (ssDateOrderShipped is not null)"
		Case enCriteria_False:	pstrsqlWhere = pstrsqlWhere & " AND (ssDateOrderShipped is null)"
	End Select
	
	pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custLastName + ', ' + sfCustomers.custFirstName As custName, sfOrders.orderPaymentMethod, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived" _
			& " FROM sfCustomers INNER JOIN (ssOrderManager RIGHT JOIN sfOrders ON ssOrderManager.ssorderID = sfOrders.orderID) ON sfCustomers.custID = sfOrders.orderCustId" _
			& pstrSQLWhere

	'added support for gift certificates
	pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custLastName+', '+sfCustomers.custFirstName AS custName, sfOrders.orderPaymentMethod, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, qryGiftCertificateRedemptionsByOrder.SumOfssGCRedemptionAmount" _
			& " FROM qryGiftCertificateRedemptionsByOrder RIGHT JOIN (sfCustomers INNER JOIN (ssOrderManager RIGHT JOIN sfOrders ON ssOrderManager.ssorderID = sfOrders.orderID) ON sfCustomers.custID = sfOrders.orderCustId) ON qryGiftCertificateRedemptionsByOrder.ssGCRedemptionOrderID = sfOrders.orderID" _
			& pstrSQLWhere

	pstrSQL = pstrSQL & " ORDER BY sfOrders.orderID"
	
	'debugprint "pstrSQL",pstrSQL
	'Response.Flush

	Set	pobjRSOrders = GetRS(pstrSQL)
	If pobjRSOrders.State <> 1 Then
		If createGCView Then Set pobjRSOrders = GetRS(pstrSQL)
		Err.Clear
	End If
	If pobjRSOrders.EOF Then
		ary = ""
		getOrderDetails = False
	Else
		ReDim ary(pobjRSOrders.RecordCount - 1)
		For i = 0 To pobjRSOrders.RecordCount - 1
			ary(i) = Array("", "", "", 0, "", "", 0)
			ary(i)(enOrderDetail_orderID) = pobjRSOrders.Fields(enOrderDetail_orderID).Value
			ary(i)(enOrderDetail_custName) = pobjRSOrders.Fields(enOrderDetail_custName).Value
			ary(i)(enOrderDetail_orderPaymentMethod) = pobjRSOrders.Fields(enOrderDetail_orderPaymentMethod).Value
			ary(i)(enOrderDetail_orderGrandTotal) = pobjRSOrders.Fields(enOrderDetail_orderGrandTotal).Value
			ary(i)(enOrderDetail_orderDate) = pobjRSOrders.Fields(enOrderDetail_orderDate).Value
			ary(i)(enOrderDetail_orderPaymentDate) = pobjRSOrders.Fields(enOrderDetail_orderPaymentDate).Value
			ary(i)(enOrderDetail_orderCertificateRedemptions) = pobjRSOrders.Fields(enOrderDetail_orderCertificateRedemptions).Value
			If Len(ary(i)(enOrderDetail_orderCertificateRedemptions) & "") = 0 Then ary(i)(enOrderDetail_orderCertificateRedemptions) = 0
			pobjRSOrders.MoveNext
		Next 'i
		getOrderDetails = True
	End If

	Call ReleaseObject(pobjRSOrders)

End Function	'getOrderDetails

'******************************************************************************************************************************************

Function GetOrderSummaries(byVal dtStartDate, byVal dtEndDate, byVal blnShowIncomplete, byVal bytNumProductsToReturn)

Dim pstrSQL
Dim pstrSQLWhere
Dim pobjRSOrders
Dim paryOrders
Dim pblnLocalDebug
Dim i
Dim pstrStartDate
Dim pstrEndDate

	pstrStartDate = convertDateArray(dtStartDate, "12:00:00 AM")
	pstrEndDate = convertDateArray(dtEndDate, "11:59:59 PM")

	paryOrders = Array(0, 0, 0, "")
	pblnLocalDebug = False

	If blnShowIncomplete Then
		pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And ((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete=1) OR (sfOrders.orderIsComplete is Null))"
	Else
		pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And (sfOrders.orderIsComplete=1)"
	End If

	If Len(pstrStartDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate >= " & wrapSQLValue(pstrStartDate, True, enDatatype_date) & ")"
	If Len(pstrEndDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate <= " & wrapSQLValue(pstrEndDate, True, enDatatype_date) & ")"
	
	If cblnSQLDatabase Then
		pstrSQL = "SELECT Sum(convert(money,orderGrandTotal)) AS SumOforderGrandTotal" _
				& " FROM sfOrders" _
				& pstrSQLWhere
		Set	pobjRSOrders = GetRS(pstrSQL)
		With pobjRSOrders
			If Not .EOF Then
				paryOrders(enOrderSummary_SalesTotal) = .Fields("SumOforderGrandTotal").Value
			End If
		End With
		Call ReleaseObject(pobjRSOrders)
		
		pstrSQL = "SELECT Count(orderGrandTotal) AS CountOforderGrandTotal" _
				& " FROM sfOrders" _
				& pstrSQLWhere
		Set	pobjRSOrders = GetRS(pstrSQL)
		With pobjRSOrders
			If Not .EOF Then
				paryOrders(enOrderSummary_SalesCount) = .Fields("CountOforderGrandTotal").Value
			End If
		End With
		Call ReleaseObject(pobjRSOrders)
	Else
		pstrSQL = "SELECT Sum(sfOrders.orderGrandTotal) AS SumOforderGrandTotal, Count(sfOrders.orderGrandTotal) AS CountOforderGrandTotal" _
				& " FROM sfOrders" _
				& pstrSQLWhere
		Set	pobjRSOrders = GetRS(pstrSQL)
		With pobjRSOrders
			If Not .EOF Then
				paryOrders(enOrderSummary_SalesTotal) = .Fields("SumOforderGrandTotal").Value
				paryOrders(enOrderSummary_SalesCount) = .Fields("CountOforderGrandTotal").Value
			End If
		End With
		Call ReleaseObject(pobjRSOrders)
	End If
	
	If bytNumProductsToReturn > -1 Then
		Dim paryBestSellers
		If bytNumProductsToReturn > 0 Then
			bytNumProductsToReturn = "Top " & bytNumProductsToReturn
		Else
			bytNumProductsToReturn = ""
		End If

		If cblnSQLDatabase Then
			pstrSQL = "SELECT " & bytNumProductsToReturn & " Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName,  Sum(convert(money,sfOrderDetails.odrdtSubTotal)) AS SumOfodrdtSubTotal " _
					& "  FROM sfOrders RIGHT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId " _
					& pstrSQLWhere _
					& "  GROUP BY sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrders.orderIsComplete " _
					& "  ORDER BY Sum(sfOrderDetails.odrdtQuantity) DESC "
		Else
			pstrSQL = "SELECT " & bytNumProductsToReturn & " Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName,  Sum(sfOrderDetails.odrdtSubTotal) AS SumOfodrdtSubTotal " _
					& "  FROM sfOrders RIGHT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId " _
					& pstrSQLWhere _
					& "  GROUP BY sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrders.orderIsComplete " _
					& "  ORDER BY Sum(sfOrderDetails.odrdtQuantity) DESC "
		End If
		
		Set	pobjRSOrders = GetRS(pstrSQL)
		With pobjRSOrders
			If Not .EOF Then
				ReDim paryBestSellers(.RecordCount - 1)
				For i = 0 To UBound(paryBestSellers)
					paryBestSellers(i) = Array(.Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value)
					.MoveNext
				Next 'i
				paryOrders(enOrderSummary_TopSellers) = paryBestSellers
			End If
		End With
		Call ReleaseObject(pobjRSOrders)
	End If

	If Not isNumeric(paryOrders(enOrderSummary_SalesTotal)) Then paryOrders(enOrderSummary_SalesTotal) = 0
	If Not isNumeric(paryOrders(enOrderSummary_SalesCount)) Then paryOrders(enOrderSummary_SalesCount) = 0
	If paryOrders(enOrderSummary_SalesCount) > 0 Then paryOrders(enOrderSummary_AverageOrder) = FormatNumber(paryOrders(enOrderSummary_SalesTotal) / paryOrders(enOrderSummary_SalesCount), 2)
	
	GetOrderSummaries = paryOrders

End Function	'GetOrderSummaries

'******************************************************************************************************************************************

Function GetDetailedOrderSummaries(byVal dtStartDate, byVal dtEndDate, byVal blnShowIncomplete, byVal bytNumProductsToReturn, byVal strSortField)

Dim paryBestSellers
Dim pstrSQL
Dim pstrSQLWhere
Dim pobjRSOrders
Dim paryOrders
Dim pblnLocalDebug
Dim i
Dim pstrStartDate
Dim pstrEndDate
Dim pstrProductName
Dim pstrAttribute
Dim paryAttributes
Dim plngPrevID
Dim plngUniqueOrderCounter
Dim pdicUniqueProducts
Dim pstrKey
Dim paryValue
Dim pblnAddItem
Dim plngQty
Dim pstrProductID
Dim pdblSubTotal
Dim vItem

	pstrStartDate = convertDateArray(dtStartDate, "12:00:00 AM")
	pstrEndDate = convertDateArray(dtEndDate, "11:59:59 PM")

	paryOrders = Array(0, 0, 0, "")
	pblnLocalDebug = False

	If blnShowIncomplete Then
		pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And ((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete=1) OR (sfOrders.orderIsComplete is Null))"
	Else
		pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And (sfOrders.orderIsComplete=1)"
	End If

	If Len(pstrStartDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate >= " & wrapSQLValue(pstrStartDate, True, enDatatype_date) & ")"
	If Len(pstrEndDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate <= " & wrapSQLValue(pstrEndDate, True, enDatatype_date) & ")"
	
	If cblnSQLDatabase Then
		pstrSQL = "SELECT Sum(convert(money,orderGrandTotal)) AS SumOforderGrandTotal" _
				& " FROM sfOrders" _
				& pstrSQLWhere
		Set	pobjRSOrders = GetRS(pstrSQL)
		With pobjRSOrders
			If Not .EOF Then
				paryOrders(enOrderSummary_SalesTotal) = .Fields("SumOforderGrandTotal").Value
			End If
		End With
		Call ReleaseObject(pobjRSOrders)
		
		pstrSQL = "SELECT Count(orderGrandTotal) AS CountOforderGrandTotal" _
				& " FROM sfOrders" _
				& pstrSQLWhere
		Set	pobjRSOrders = GetRS(pstrSQL)
		With pobjRSOrders
			If Not .EOF Then
				paryOrders(enOrderSummary_SalesCount) = .Fields("CountOforderGrandTotal").Value
			End If
		End With
		Call ReleaseObject(pobjRSOrders)
	Else
		pstrSQL = "SELECT Sum(sfOrders.orderGrandTotal) AS SumOforderGrandTotal, Count(sfOrders.orderGrandTotal) AS CountOforderGrandTotal" _
				& " FROM sfOrders" _
				& pstrSQLWhere
		Set	pobjRSOrders = GetRS(pstrSQL)
		With pobjRSOrders
			If Not .EOF Then
				paryOrders(enOrderSummary_SalesTotal) = .Fields("SumOforderGrandTotal").Value
				paryOrders(enOrderSummary_SalesCount) = .Fields("CountOforderGrandTotal").Value
			End If
		End With
		Call ReleaseObject(pobjRSOrders)
	End If
	
	If bytNumProductsToReturn > -1 Then

		If bytNumProductsToReturn > 0 Then
			bytNumProductsToReturn = "Top " & bytNumProductsToReturn
		Else
			bytNumProductsToReturn = ""
		End If

		pstrSQL = "SELECT sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtSubTotal, sfOrderAttributes.odrattrName, sfOrderAttributes.odrattrAttribute, sfOrderDetails.odrdtCategory, sfOrderDetails.odrdtManufacturer, sfOrderDetails.odrdtVendor" _
				& "  FROM (sfOrders RIGHT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId" _
				& pstrSQLWhere _
				& "  ORDER BY sfOrderDetails.odrdtID"
				'& "  ORDER BY sfOrderDetails.odrdtID, sfOrderAttributes.odrattrName, sfOrderAttributes.odrattrAttribute"

		plngPrevID = -1
		Set pdicUniqueProducts = Server.CreateObject("Scripting.Dictionary")
		
		Set	pobjRSOrders = GetRS(pstrSQL)
		With pobjRSOrders
			If Not .EOF Then
				ReDim paryBestSellers(.RecordCount - 1)
				For i = 0 To UBound(paryBestSellers)
					If plngPrevID <> .Fields(2).Value Then
						plngPrevID = .Fields(2).Value
						
						pstrProductID = Trim(.Fields(1).Value)
						pstrProductName = .Fields(3).Value
						plngQty = .Fields(0).Value
						If isNumeric(.Fields(4).Value) And Len(.Fields(4).Value & "") > 0 Then
							pdblSubTotal = CDbl(.Fields(4).Value)
						Else
							pdblSubTotal = 0
						End If
						pstrAttribute = ""
					End If
					
					If Len(Trim(.Fields(5).Value & .Fields(6).Value)) > 0 Then
						If Len(pstrAttribute) > 0 Then pstrAttribute = pstrAttribute & "|attribute|"
						pstrAttribute = pstrAttribute & Trim(.Fields(5).Value) & "|" & Trim(.Fields(6).Value)
					End If
					
					paryBestSellers(i) = Array(plngQty, pstrProductID, pstrProductName, pstrProductName)
					.MoveNext
					
					If .EOF Then
						pblnAddItem= True
					ElseIf plngPrevID <> .Fields(2).Value Then
						pblnAddItem = True
					Else
						pblnAddItem = False
					End If
					
					If pblnAddItem Then
						pstrKey = pstrProductID 
						If Len(pstrAttribute) > 0 Then pstrKey = pstrKey & "|attributes|" & pstrAttribute
						
						If pdicUniqueProducts.Exists(pstrKey) Then
							paryValue = pdicUniqueProducts(pstrKey)
							paryValue(0) = paryValue(0) + plngQty
							paryValue(3) = paryValue(3) + pdblSubTotal
							pdicUniqueProducts(pstrKey) = paryValue
						Else
							paryValue = Array(plngQty, pstrProductID, pstrProductName, pdblSubTotal)
							pdicUniqueProducts.Add pstrKey, paryValue
						End If
					End If
						
				Next 'i
			End If
		End With

		ReDim paryBestSellers(pdicUniqueProducts.Count - 1)
		'Response.Write "<hr>"
		plngUniqueOrderCounter = -1
		For Each vItem in pdicUniqueProducts
			plngUniqueOrderCounter = plngUniqueOrderCounter + 1
			paryValue = pdicUniqueProducts(vItem)
			paryValue(2) = Replace(vItem, paryValue(1), paryValue(2))
			paryBestSellers(plngUniqueOrderCounter) = paryValue
		'	Response.Write vItem & ": " & paryValue(0) & "<br />"
		Next 'vItem
		'Response.Write "<hr>"
		Call SortOrderItems(paryBestSellers, 0, True, False)
		paryOrders(enOrderSummary_TopSellers) = paryBestSellers

		Set pdicUniqueProducts = Nothing
		Call ReleaseObject(pobjRSOrders)
	End If

	If Not isNumeric(paryOrders(enOrderSummary_SalesTotal)) Then paryOrders(enOrderSummary_SalesTotal) = 0
	If Not isNumeric(paryOrders(enOrderSummary_SalesCount)) Then paryOrders(enOrderSummary_SalesCount) = 0
	If paryOrders(enOrderSummary_SalesCount) > 0 Then paryOrders(enOrderSummary_AverageOrder) = FormatNumber(paryOrders(enOrderSummary_SalesTotal) / paryOrders(enOrderSummary_SalesCount), 2)
	
	GetDetailedOrderSummaries = paryOrders

End Function	'GetDetailedOrderSummaries

'******************************************************************************************************************************************

Sub SortOrderItems(ByRef aryItems, byVal lngSortIndex, byVal blnIsNumeric, byVal blnAscending)

Dim i
Dim j
Dim pblnSwapItems
Dim paryTemp
Dim plngNumItems

	plngNumItems = UBound(aryItems)
	For i = 0 To plngNumItems
		For j = 0 To plngNumItems-1
			If blnAscending Then
				If blnIsNumeric Then
					pblnSwapItems = CDbl(aryItems(j)(lngSortIndex)) > CDbl(aryItems(j+1)(lngSortIndex))
				Else
					pblnSwapItems = aryItems(j)(lngSortIndex) > aryItems(j+1)(lngSortIndex)
				End If
			Else
				If blnIsNumeric Then
					pblnSwapItems = CDbl(aryItems(j)(lngSortIndex)) < CDbl(aryItems(j+1)(lngSortIndex))
				Else
					pblnSwapItems = aryItems(j)(lngSortIndex) < aryItems(j+1)(lngSortIndex)
				End If
			End If
			
			If pblnSwapItems Then
				paryTemp = aryItems(j+1)
				aryItems(j+1) = aryItems(j)
				aryItems(j) = paryTemp
			End If
		Next 'j
	Next 'i
	
End Sub	'SortOrderItems

'******************************************************************************************************************************************

Function writeProductName(byVal strProductName)

Dim paryProduct
Dim paryAttribute
Dim paryAttributes
Dim pstrOut
Dim i

	If InStr(1, strProductName, "|attributes|") > 0 Then
		paryProduct = Split(strProductName, "|attributes|")
		pstrOut = paryProduct(0)
		
		paryAttributes = Split(paryProduct(1), "|attribute|")
		For i = 0 To UBound(paryAttributes)
			paryAttribute  = Split(paryAttributes(i), "|")
			If Len(paryAttribute(0) & paryAttribute(1)) > 0 Then
				pstrOut = pstrOut _
						& "<br />-- " & paryAttribute(0) & ": " & paryAttribute(1)
			End If
		Next 'i
	Else
		pstrOut = strProductName
	End If
	
	writeProductName = pstrOut

End Function	'writeProductName

'******************************************************************************************************************************************

Function GetOrderReport(byVal dtStartDate, byVal dtEndDate, byVal blnShowIncomplete)

Dim pstrSQL
Dim pstrSQLWhere
Dim pobjRSOrders
Dim paryOrders
Dim pblnLocalDebug
Dim i
Dim pstrStartDate
Dim pstrEndDate

	pstrStartDate = convertDateArray(dtStartDate, "12:00:00 AM")
	pstrEndDate = convertDateArray(dtEndDate, "11:59:59 PM")

	paryOrders = Array(0, 0, 0, "")
	pblnLocalDebug = False

	If blnShowIncomplete Then
		pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And ((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete=1) OR (sfOrders.orderIsComplete is Null))"
	Else
		pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And (sfOrders.orderIsComplete=1)"
	End If

	If Len(pstrStartDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate >= " & wrapSQLValue(pstrStartDate, True, enDatatype_date) & ")"
	If Len(pstrEndDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate <= " & wrapSQLValue(pstrEndDate, True, enDatatype_date) & ")"
	
	pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrderDetails.odrdtCategory, sfOrders.orderGrandTotal, sfOrders.orderGrandTotal, sfOrders.orderAmount, sfOrders.orderHandling, sfOrders.orderShippingAmount, sfOrders.orderShipMethod, sfOrders.orderSTax, sfOrders.orderPaymentMethod, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, qryGiftCertificateRedemptionsByOrder.SumOfssGCRedemptionAmount" _
			& " FROM qryGiftCertificateRedemptionsByOrder RIGHT JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON qryGiftCertificateRedemptionsByOrder.ssGCRedemptionOrderID = sfOrders.orderID" _
			& pstrSQLWhere _
			& " ORDER BY sfOrders.orderDate, sfOrders.orderPaymentMethod, sfOrderDetails.odrdtCategory"

	Set	pobjRSOrders = GetRS(pstrSQL)
	If pobjRSOrders.State <> 1 Then
		If createGCView Then Set pobjRSOrders = GetRS(pstrSQL)
		Err.Clear
	End If

	'debugprint "pstrSQL", pstrSQL
	'debugprint "pobjRSOrders.RecordCount", pobjRSOrders.RecordCount
	
	Set GetOrderReport = pobjRSOrders

End Function	'GetOrderReport

'******************************************************************************************************************************************

Function createGCView()

Dim pstrSQL
Dim pstrLocalError

	pstrSQL = "CREATE VIEW qryGiftCertificateRedemptionsByOrder AS " _
			& "SELECT Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount) AS SumOfssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionOrderID" _
			& " FROM ssGiftCertificateRedemptions" _
			& " GROUP BY ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
			& " HAVING ((ssGiftCertificateRedemptions.ssGCRedemptionType=1) AND (ssGiftCertificateRedemptions.ssGCRedemptionActive<>0))"
	If Execute_NoReturn(pstrSQL, pstrLocalError) Then
		createGCView = True
	Else
		createGCView = False
		Response.Write pstrLocalError
	End If
	
End Function	'createGCView

'******************************************************************************************************************************************

Function GetOrderReportWithCCDetail(byVal dtStartDate, byVal dtEndDate, byVal blnShowIncomplete)

Dim pstrSQL
Dim pstrSQLWhere
Dim pobjRSOrders
Dim paryOrders
Dim pblnLocalDebug
Dim i
Dim pstrStartDate
Dim pstrEndDate

	pstrStartDate = convertDateArray(dtStartDate, "12:00:00 AM")
	pstrEndDate = convertDateArray(dtEndDate, "11:59:59 PM")

	paryOrders = Array(0, 0, 0, "")
	pblnLocalDebug = False

	If blnShowIncomplete Then
		pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And ((sfOrders.orderIsComplete=0) OR (sfOrders.orderIsComplete=1) OR (sfOrders.orderIsComplete is Null))"
	Else
		pstrSQLWhere = " Where (orderVoided=0 Or orderVoided is Null) And (sfOrders.orderIsComplete=1)"
	End If

	If Len(pstrStartDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate >= " & wrapSQLValue(pstrStartDate, True, enDatatype_date) & ")"
	If Len(pstrEndDate) > 0 Then pstrsqlWhere = pstrsqlWhere & " and (orderDate <= " & wrapSQLValue(pstrEndDate, True, enDatatype_date) & ")"
	
	If cblnSQLDatabase Then
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderGrandTotal, sfOrders.orderHandling, sfOrders.orderShippingAmount, sfOrders.orderShipMethod, sfOrders.orderSTax, sfOrders.orderPaymentMethod, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, sfTransactionTypes.transName" _
				& " FROM (sfCPayments INNER JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCPayments.payID = sfOrders.orderPayId) LEFT JOIN sfTransactionTypes ON sfCPayments.payCardType = convert(varchar,sfTransactionTypes.transID)" _
				& pstrSQLWhere _
				& " ORDER BY sfOrders.orderPaymentMethod, sfTransactionTypes.transName, sfOrders.orderDate"

		'Added GC redemption support
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderGrandTotal, sfOrders.orderHandling, sfOrders.orderShippingAmount, sfOrders.orderShipMethod, sfOrders.orderSTax, sfOrders.orderPaymentMethod, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, sfTransactionTypes.transName, qryGiftCertificateRedemptionsByOrder.SumOfssGCRedemptionAmount" _
				& " FROM qryGiftCertificateRedemptionsByOrder RIGHT JOIN ((sfCPayments INNER JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCPayments.payID = sfOrders.orderPayId) LEFT JOIN sfTransactionTypes ON sfCPayments.payCardType = convert(varchar,sfTransactionTypes.transID)) ON qryGiftCertificateRedemptionsByOrder.ssGCRedemptionOrderID = sfOrders.orderID" _
				& pstrSQLWhere _
				& " ORDER BY sfOrders.orderPaymentMethod, sfTransactionTypes.transName, sfOrders.orderDate"
	Else
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderGrandTotal, sfOrders.orderHandling, sfOrders.orderShippingAmount, sfOrders.orderShipMethod, sfOrders.orderSTax, sfOrders.orderPaymentMethod, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, sfTransactionTypes.transName" _
				& " FROM (sfCPayments INNER JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCPayments.payID = sfOrders.orderPayId) LEFT JOIN sfTransactionTypes ON sfCPayments.payCardType = CStr(sfTransactionTypes.transID)" _
				& pstrSQLWhere _
				& " ORDER BY sfOrders.orderPaymentMethod, sfTransactionTypes.transName, sfOrders.orderDate"

		'Added GC redemption support
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderGrandTotal, sfOrders.orderHandling, sfOrders.orderShippingAmount, sfOrders.orderShipMethod, sfOrders.orderSTax, sfOrders.orderPaymentMethod, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, sfTransactionTypes.transName, qryGiftCertificateRedemptionsByOrder.SumOfssGCRedemptionAmount" _
				& " FROM qryGiftCertificateRedemptionsByOrder RIGHT JOIN ((sfCPayments INNER JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCPayments.payID = sfOrders.orderPayId) LEFT JOIN sfTransactionTypes ON sfCPayments.payCardType = CStr(sfTransactionTypes.transID)) ON qryGiftCertificateRedemptionsByOrder.ssGCRedemptionOrderID = sfOrders.orderID" _
				& pstrSQLWhere _
				& " ORDER BY sfOrders.orderPaymentMethod, sfTransactionTypes.transName, sfOrders.orderDate"
	End If
	Set	pobjRSOrders = GetRS(pstrSQL)
	'debugprint "pstrSQL", pstrSQL
	'debugprint "pobjRSOrders.RecordCount", pobjRSOrders.RecordCount
	
	Set GetOrderReportWithCCDetail = pobjRSOrders

End Function	'GetOrderReportWithCCDetail

'******************************************************************************************************************************************

Function writeTopSellers(byVal aryBestSellers, byVal strItemTemplate, byVal strOpenTag, byVal strCloseTag)

Dim i
Dim pstrOut
Dim pstrItem

	If isArray(aryBestSellers) Then
		pstrOut = strOpenTag
		For i = 0 To UBound(aryBestSellers)
			pstrItem = Trim(strItemTemplate & "")
			pstrItem = Replace(pstrItem, "<salesQty>", aryBestSellers(i)(0) & "")
			pstrItem = Replace(pstrItem, "<prodID>", aryBestSellers(i)(1) & "")
			pstrItem = Replace(pstrItem, "<prodName>", aryBestSellers(i)(2) & "")
			pstrItem = Replace(pstrItem, "<salesTotal>", aryBestSellers(i)(3) & "")
			pstrOut = pstrOut & pstrItem
		Next 'i
		pstrOut = pstrOut & strCloseTag
	End If
	
	writeTopSellers = pstrOut

End Function	'writeTopSellers

'******************************************************************************************************************************************

Function createProductSalesReport(byVal aryOrderSummary)

Dim i, j, k
Dim paryTopSellers
Dim pstrProductID
Dim plngNumYears
Dim plngArrayLength
Dim paryProduct
Dim paryProductsOut
Dim pdicProducts

		'Now build the summary
		plngNumYears = UBound(aryOrderSummary) + 1
		plngArrayLength = plngNumYears * 2 + 3

		ReDim paryProduct(plngArrayLength)	'Product ID, Product Name, salesQty (year), salesTotal (year), salesQty (total), salesTotal (total)
		
		Set pdicProducts = CreateObject("Scripting.Dictionary")
		
		For i = 0 To UBound(aryOrderSummary)
			If isArray(aryOrderSummary(i)) Then
				paryTopSellers = aryOrderSummary(i)(enOrderSummary_TopSellers)
				If isArray(paryTopSellers) Then
					For j = 0 To UBound(paryTopSellers)
						pstrProductID = paryTopSellers(j)(1)
						If pdicProducts.Exists(pstrProductID) Then
							paryProduct = pdicProducts(pstrProductID)
						Else
							paryProduct(0) = paryTopSellers(j)(1)
							paryProduct(1) = paryTopSellers(j)(2)
							For k = 2 To UBound(paryProduct)
								paryProduct(k) = 0
							Next 'k
							pdicProducts.Add pstrProductID, paryProduct
						End If
						
						'Set the totals
						paryProduct(i * 2 + 2) = paryTopSellers(j)(0)
						paryProduct(i * 2 + 3) = paryTopSellers(j)(3)
						paryProduct(plngArrayLength - 1) = paryProduct(plngArrayLength - 1) + paryProduct(i * 2 + 2)
						paryProduct(plngArrayLength) = paryProduct(plngArrayLength) + paryProduct(i * 2 + 3)
						pdicProducts.Item(pstrProductID) = paryProduct
					Next 'j
				End If	'isArray(aryOrderSummary(i))
			End If	'isArray(aryOrderSummary(i))
		Next 'i
		
		'Now transfer the dictionary to an array
		ReDim paryProductsOut(pdicProducts.Count - 1)
		i = 0
		For Each pstrProductID in pdicProducts
			paryProductsOut(i) = pdicProducts(pstrProductID)
			i = i + 1
		Next 'vItem
		
		Set pdicProducts = Nothing
		
		'now for the results
		If False Then
		Response.Write "<table border=1>"
		For i = 0 To UBound(paryProductsOut)
			paryProduct = paryProductsOut(i)
			Response.Write "<tr>"
			For j = 0 To UBound(paryProduct)
				Response.Write "<td>" & paryProduct(j) & "</td>"
			Next
			Response.Write "</tr>"
		Next
		Response.Write "</table>"
		End If
		
		createProductSalesReport = paryProductsOut

End Function	'createProductSalesReport

%>
