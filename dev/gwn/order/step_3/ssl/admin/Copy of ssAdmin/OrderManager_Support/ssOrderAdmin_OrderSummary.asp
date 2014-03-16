<%
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   1.01.004												*
'*   Release Date:		November 15, 2003										*
'*   Revision Date:		October 14, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Dim pXML_OrderSummaries

	'***********************************************************************************************

	Public Function LoadOrderSummaryXML(byVal blnCheckAll)

	Dim xmlDoc
	Dim xmlRoot
	Dim xmlNode
	Dim xmlOrderDetail
	Dim xmlOrderItemsAttributes
	Dim xmlOrderAttributeDetail
	Dim xmlOrderItem
	Dim xmlOrderStatusOptions
	dim plnguBound, plnglbound, pstrDisplay
	Dim plngCurrentOrderID: plngCurrentOrderID = 0
	Dim plngCurrentOrderNumber
	Dim pstrFieldName
	Dim pstrFieldValue
	Dim fieldCounter
	Dim pobjRS
	Dim pobjCmd
	Dim pobjCmd_CC
	Dim pstrSQL
	Dim plngPrevID
	Dim pstrssOrderShipped
	Dim xmlOrderPayment
	Dim pstrTRClass
	Dim pstrTRImage
	Dim pblnSelected

On Error Goto 0

		plngPrevID = 0
		set xmlDoc = server.CreateObject("MSXML2.DOMDocument.3.0")
		' Create processing instruction and document root
		Set xmlNode = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-16'")
		Set xmlNode = xmlDoc.insertBefore(xmlNode, xmlDoc.childNodes.Item(0))
		
		' Create document root
		Set xmlRoot = xmlDoc.createElement("orders")
		Set xmlDoc.documentElement = xmlRoot
		xmlRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"

		Call addNode(xmlDoc, xmlRoot, "PageCount", mlngPageCount)
		Call addNode(xmlDoc, xmlRoot, "AbsolutePage", mlngAbsolutePage)
		Call addNode(xmlDoc, xmlRoot, "MaxRecords", mlngMaxRecords)
		Call addNode(xmlDoc, xmlRoot, "RecordCount", prsOrderSummaries.RecordCount)

		If len(mstrOrderBy) > 0 Then
			Call addNode(xmlDoc, xmlRoot, "OrderBy", mstrOrderBy)
		Else
			Call addNode(xmlDoc, xmlRoot, "OrderBy", 1)
		End If

		If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
			Call addNode(xmlDoc, xmlRoot, "SortOrder", "ASC")
		Else
			Call addNode(xmlDoc, xmlRoot, "SortOrder", "DESC")
		End If
		
		'add order display options
		Set xmlOrderStatusOptions = xmlDoc.createElement("orderDetailDisplayOptions")
		xmlRoot.appendChild xmlOrderStatusOptions
		Call addNode(xmlDoc, xmlOrderStatusOptions, "ordersExtra1_Label", cstrOrdersExtra1_Label)
		Call addNode(xmlDoc, xmlOrderStatusOptions, "Insured_Label", cstrInsured_Label)
		Call addNode(xmlDoc, xmlOrderStatusOptions, "PackageWeight_Label", cstrPackageWeight_Label)
		Call addNode(xmlDoc, xmlOrderStatusOptions, "orderTrackingExtra1_Label", cstrOrderTrackingExtra1_Label)
		Call addNode(xmlDoc, xmlOrderStatusOptions, "UseOrderFlags", cblnUseOrderFlags)
	
	    If prsOrderSummaries.RecordCount > 0 Then
	        prsOrderSummaries.MoveFirst

			If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
			If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
			plnglbound = (mlngAbsolutePage - 1) * prsOrderSummaries.PageSize + 1
			plnguBound = mlngAbsolutePage * prsOrderSummaries.PageSize

			If plnguBound > prsOrderSummaries.RecordCount Then plnguBound = prsOrderSummaries.RecordCount
			prsOrderSummaries.AbsolutePosition = plnglbound
			For i = plnglbound To plnguBound

				If plngCurrentOrderID <> prsOrderSummaries.Fields("orderID").Value Then
					plngCurrentOrderID = prsOrderSummaries.Fields("orderID").Value
					plngCurrentOrderNumber = prsOrderSummaries.Fields("orderID").Value
					
					'Create Order Node
					Set xmlOrderDetail = xmlDoc.createElement("order")
					xmlRoot.appendChild xmlOrderDetail

					xmlOrderDetail.setAttribute "uid", plngCurrentOrderID
				
					Call addNode(xmlDoc, xmlOrderDetail, "OrderNumber", getRSFieldValue_Unknown(prsOrderSummaries.Fields("orderID")))
					Call addNode(xmlDoc, xmlOrderDetail, "ssOrderStatus", getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssOrderStatus")))
					Call addNode(xmlDoc, xmlOrderDetail, "LastName", getRSFieldValue_Unknown(prsOrderSummaries.Fields("custLastName")))
					Call addNode(xmlDoc, xmlOrderDetail, "FirstName", getRSFieldValue_Unknown(prsOrderSummaries.Fields("custFirstName")))
					Call addNode(xmlDoc, xmlOrderDetail, "EMail", getRSFieldValue_Unknown(prsOrderSummaries.Fields("custEmail")))
					Call addNode(xmlDoc, xmlOrderDetail, "SumOfQuantity", getRSFieldValue_Unknown(prsOrderSummaries.Fields("SumOfodrdtQuantity")))
					Call addNode(xmlDoc, xmlOrderDetail, "GrandTotal", getRSFieldValue_Unknown(prsOrderSummaries.Fields("orderGrandTotal")))
					Call addNode(xmlDoc, xmlOrderDetail, "DateOrdered", getRSFieldValue_Unknown(prsOrderSummaries.Fields("orderDate")))
					Call addNode(xmlDoc, xmlOrderDetail, "ssOrderFlagged", getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssOrderFlagged")))
					Call addNode(xmlDoc, xmlOrderDetail, "currencySymbol", Replace(formatCurrency("0", 0), "0", ""))
					Call addNode(xmlDoc, xmlOrderDetail, "decimalSeparator", LCase(InStr(1, formatCurrency(0, 2, True), ",") > 0))
					Call addNode(xmlDoc, xmlOrderDetail, "ssDatePaymentReceived", getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssDatePaymentReceived")))
					Call addNode(xmlDoc, xmlOrderDetail, "ssDateOrderShipped", getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssDateOrderShipped")))
					Call addNode(xmlDoc, xmlOrderDetail, "ssOrderStatus", getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssOrderStatus")))
					Call addNode(xmlDoc, xmlOrderDetail, "ssInternalOrderStatus", getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssInternalOrderStatus")))
					Call addNode(xmlDoc, xmlOrderDetail, "ssBackOrderDateExpected", getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssBackOrderDateExpected")))

					pblnSelected = (CStr(plngCurrentOrderID) = CStr(plngssorderID))
					If pblnSelected Then
						pstrTRClass = "Selected"
						Call addNode(xmlDoc, xmlOrderDetail, "ActiveOrder", "1")
					Else
						'pstrTRClass = maryOrderStatuses(correctEmptyValue(getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssOrderStatus")), 0))(1)
						If Len(pstrTRClass) = 0 Then
							If False Then
								pstrTRClass = "Inactive"
							Else
								pstrTRClass = "Active"
							End If
						End If
					End If
					
					'On Error Resume Next
					'pstrTRImage = maryOrderStatuses(correctEmptyValue(getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssOrderStatus")), 0))(2)
					'If Len(pstrTRImage) > 0 Then pstrTRImage = "<img src='" & pstrTRImage & "' border=0 />"

					Call addNode(xmlDoc, xmlOrderDetail, "TRClass", pstrTRClass)
					Call addNode(xmlDoc, xmlOrderDetail, "TRImage", pstrTRImage)
					Call addNode(xmlDoc, xmlOrderDetail, "Checked", pblnSelected Or blnCheckAll)
				End If	'plngCurrentOrderID <> prsOrderSummaries.Fields("orderID").Value
					
				If cblnIncludeProductsInDisplay Then
					If Not isObject(pobjCmd) Then
						pstrSQL = "SELECT OrderItems.UID, OrderItems.ProductID, OrderItems.ProductCode, OrderItems.ProductName, OrderItems.Quantity, OrderItems.BackOrderedQty, Products.Name, OrderItemsAttributes.AttributeName, OrderItemsAttributes.AttributeDetailName, OrderItemsAttributes.CustomAttribute" _
								& " FROM (OrderItemsAttributes RIGHT JOIN OrderItems ON OrderItemsAttributes.OrderItemID = OrderItems.uid) LEFT JOIN Products ON OrderItems.ProductID = Products.uid" _
								& " WHERE OrderID=?" _
								& " ORDER BY OrderItems.uid, OrderItemsAttributes.AttributeName"
						Set pobjCmd  = Server.CreateObject("ADODB.Command")
						With pobjCmd
							Set .ActiveConnection = cnn
							.Commandtype = adCmdText
							.Commandtext = pstrSQL
							.Parameters.Append .CreateParameter("OrderID", adInteger, adParamInput, 4, prsOrderSummaries.Fields("orderID").Value)
						End With
					End If	'Not isObject(pobjCmd)

					'On Error Resume Next
					pobjCmd.Parameters(0).Value = prsOrderSummaries.Fields("orderID").Value
					Set pobjRS = pobjCmd.Execute
					With pobjRS
						Do While Not .EOF
							If plngPrevID <> .Fields("orderID").Value Then
								plngPrevID = .Fields("orderID").Value
								Set xmlOrderItem = xmlDoc.createElement("OrderItems")
								xmlOrderDetail.appendChild xmlOrderItem
								xmlOrderItem.setAttribute "uid", plngPrevID
								Call addNode(xmlDoc, xmlOrderItem, "ProductUID", .Fields("ProductID").Value)
								Call addCDATA(xmlDoc, xmlOrderItem, "ProductCode", stripHTML(getRSFieldValue_Unknown(.Fields("ProductCode"))))
								Call addCDATA(xmlDoc, xmlOrderItem, "ProductName", stripHTML(getRSFieldValue_Unknown(.Fields("ProductName"))))
								Call addCDATA(xmlDoc, xmlOrderItem, "Quantity", stripHTML(getRSFieldValue_Unknown(.Fields("Quantity"))))
								Call addCDATA(xmlDoc, xmlOrderItem, "BackOrderedQty", stripHTML(getRSFieldValue_Unknown(.Fields("BackOrderedQty"))))
								
								Set xmlOrderItemsAttributes = xmlDoc.createElement("OrderItemsAttributes")
								xmlOrderItem.appendChild xmlOrderItemsAttributes
							End If
							
							If Not isNull(.Fields("AttributeName").Value) Then
								Set xmlOrderAttributeDetail = xmlDoc.createElement("OrderItemsAttribute")
								xmlOrderItemsAttributes.appendChild xmlOrderAttributeDetail
								Call addCDATA(xmlDoc, xmlOrderAttributeDetail, "AttributeName", stripHTML(getRSFieldValue_Unknown(.Fields("AttributeName"))))
								Call addCDATA(xmlDoc, xmlOrderAttributeDetail, "AttributeDetailName", stripHTML(getRSFieldValue_Unknown(.Fields("AttributeDetailName"))))
								Call addCDATA(xmlDoc, xmlOrderAttributeDetail, "CustomAttribute", stripHTML(getRSFieldValue_Unknown(.Fields("CustomAttribute"))))
							End If
							.MoveNext
							If .EOF Then Set xmlOrderItem = Nothing
						Loop
					End With
					Set pobjRS = Nothing
				End If	'cblnIncludeProductsInDisplay

				If False Then
				Set xmlOrderPayment = xmlDoc.createElement("OrderPayment")
				xmlOrderDetail.appendChild xmlOrderPayment
				Call addNode(xmlDoc, xmlOrderPayment, "CardType", prsOrderSummaries.Fields("CardType").Value)
				Call addNode(xmlDoc, xmlOrderPayment, "ExpireMonth", prsOrderSummaries.Fields("ExpireMonth").Value)
				Call addNode(xmlDoc, xmlOrderPayment, "ExpireYear", prsOrderSummaries.Fields("ExpireYear").Value)
				
				Select Case CStr(prsOrderSummaries.Fields("PayMethod").Value & "")
					Case "1":	Call addNode(xmlDoc, xmlOrderPayment, "PaymentMethodName", "eCheck")
					Case "2":	Call addNode(xmlDoc, xmlOrderPayment, "PaymentMethodName", "COD")
					Case "3":	Call addNode(xmlDoc, xmlOrderPayment, "PaymentMethodName", "PO")
					Case "4", "5":	Call addNode(xmlDoc, xmlOrderPayment, "PaymentMethodName", "PhoneFax")
					Case "6":	Call addNode(xmlDoc, xmlOrderPayment, "PaymentMethodName", "PayPal")
					Case Else:	Call addNode(xmlDoc, xmlOrderPayment, "PaymentMethodName", prsOrderSummaries.Fields("CardType").Value)
				End Select
				End If

				prsOrderSummaries.MoveNext
				
			Next	'i
			
		End If	'prsOrderSummaries.RecordCount > 0
		
		'Now for the paging information
		Dim xmlOrderPaging
		Set xmlOrderPaging = xmlDoc.createElement("orderPaging")
		xmlRoot.appendChild xmlOrderPaging

		For i=1 to mlngPageCount
			plnglbound = (i-1) * mlngMaxRecords + 1
			plnguBound = i * mlngMaxRecords
			if plnguBound > prsOrderSummaries.RecordCount Then plnguBound = prsOrderSummaries.RecordCount

			If cblnShowPageNumbers Then
				pstrDisplay = i
			Else
				pstrDisplay = plnglbound & " - " & plnguBound
			End If
			Call addNode(xmlDoc, xmlOrderPaging, "Page", pstrDisplay)
		Next
				
		'Response.Write "<textarea rows=50 cols=80>" & xmlDoc.xml & "</textarea>"
		'Response.Write "<hr>" & vbcrlf & vbcrlf & vbcrlf
		
		Response.Write WriteXSL(xmlDoc, ssAdminPath & "OrderManager_Support/OrderSummary_Standard.xsl")

		'Response.Write "<hr>"
		
		If isObject(pobjCmd) Then Set pobjCmd = Nothing
		If isObject(pobjCmd_CC) Then Set pobjCmd_CC = Nothing
		
		Set xmlOrderDetail = Nothing
		Set xmlDoc = Nothing
		
	End Function    'LoadOrderSummaryXML

	'***********************************************************************************************

	Public Function LoadOrderSummaries(arySQLParameters)

	Dim pstrSQL
	Dim p_strWhere
	Dim p_strGroupBy
	Dim p_sqlHaving
	Dim p_OrderBy
	Dim i
	Dim pstrOrderIDs

	'On Error Resume Next

		p_strWhere = getValueFromArray(arySQLParameters, 0, "")
		p_strGroupBy = getValueFromArray(arySQLParameters, 1, "")
		p_sqlHaving = getValueFromArray(arySQLParameters, 2, "")
		p_OrderBy = getValueFromArray(arySQLParameters, 3, "")

		'paging routine
		mlngAbsolutePage = LoadRequestValue("AbsolutePage")
		If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
		mlngMaxRecords = LoadRequestValue("PageSize")
		If len(mlngMaxRecords) = 0 Then
			mlngMaxRecords = clngDefaultRecords	
		Else
			mlngMaxRecords = CLng(mlngMaxRecords)	
		End If

		pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custEmail, 1  AS SumOfodrdtQuantity, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderStatus, ssOrderManager.ssInternalOrderStatus, ssOrderManager.ssOrderFlagged" _
				& " FROM sfCShipAddresses RIGHT JOIN (((sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId" _
				& " " & p_strWhere _
				& " " & p_OrderBy

		pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custEmail, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderStatus, ssOrderManager.ssInternalOrderStatus, ssOrderManager.ssOrderFlagged" _
				& " FROM sfCShipAddresses RIGHT JOIN (((sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId" _
				& " " & p_strWhere _
				& " GROUP BY sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custEmail, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderStatus, ssOrderManager.ssOrderFlagged, ssOrderManager.ssInternalOrderStatus, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipEmail, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCustomers.custCompany, sfCustomers.custAddr1, sfCustomers.custAddr2, sfOrders.orderIsComplete, sfOrders.orderTradingPartner " & p_strGroupBy _
				& " " & p_sqlHaving _
				& " " & p_OrderBy

		pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custEmail, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderStatus, ssOrderManager.ssInternalOrderStatus, ssOrderManager.ssOrderFlagged" _
				& " FROM sfCShipAddresses RIGHT JOIN (((sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId" _
				& " " & p_strWhere _
				& " GROUP BY sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custEmail, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderStatus, ssOrderManager.ssInternalOrderStatus, ssOrderManager.ssOrderFlagged, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipEmail, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCustomers.custCompany, sfCustomers.custAddr1, sfCustomers.custAddr2, sfOrders.orderIsComplete, sfOrders.orderTradingPartner " & p_strGroupBy _
				& " " & p_sqlHaving _
				& " " & p_OrderBy

'debugprint "pstrSQL", pstrSQL				
		If Len(cstrDisplayMemoField) > 0 Then
			pstrSQL = Replace(pstrSQL, "ssOrderManager.ssDateOrderShipped,", "ssOrderManager.ssDateOrderShipped," & cstrDisplayMemoField & ",", 1,2)
		End If

		set	prsOrderSummaries = server.CreateObject("adodb.recordset")
		With prsOrderSummaries
	        .CursorLocation = 3 'adUseClient
	'        .CursorType = 3 'adOpenStatic
	        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
	        .LockType = 1 'adLockReadOnly
			If CBool(mlngMaxRecords <> 0) Then 
				.CacheSize = mlngMaxRecords
				.PageSize = mlngMaxRecords
			End If

			.Open pstrSQL, cnn
			If Err.number = 0 Then
				If mlngMaxRecords = 0 Then
					'mlngMaxRecords = .RecordCount
					If Not .EOF Then .PageSize = .RecordCount
				End If
				plngOrderSummaryCount = .RecordCount
				mlngPageCount = .PageCount
				If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
				
				If Not .EOF Then
					plngAltOrderID = .Fields("orderID").Value
					For i = 1 To .RecordCount
						If Len(pstrOrderIDs) > 0 Then
							pstrOrderIDs = pstrOrderIDs & ", " & .Fields("orderID").Value
						Else
							pstrOrderIDs = .Fields("orderID").Value
						End If
						.MoveNext
					Next 'i
'debugprint "pstrOrderIDs", pstrOrderIDs				
'Response.Write "<fieldset><legend>XML</legend><pre>" & exportOrders(pstrOrderIDs, "xml.xsl") & "</pre></legend>"
					.MoveFirst
					LoadOrderSummaries = True
				Else
					LoadOrderSummaries = False
				End If
			Else
				Call DisplayUpgradeError(Err, "pstrSQL: " & pstrSQL)
				Err.Clear
				Response.Flush
				LoadOrderSummaries = False
			End If
			
		End With

	End Function    'LoadOrderSummaries

	'***********************************************************************************************

	Public Sub ShowOrderSummaries(byVal blnCheckAll)
		'This function replaced with XML display
		Call LoadOrderSummaryXML(False)
	End Sub      'ShowOrderSummaries
%>
