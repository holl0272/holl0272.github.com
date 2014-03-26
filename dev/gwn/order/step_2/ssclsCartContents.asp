<!--#include file="ssmodDiscounts.asp"-->
<!--#include file="ssmodShipping.asp"-->
<!--#include file="ssmodShippingZoneBased.asp"-->
<!--#include file="ssmodBuyersClub.asp"-->
<!--#include file="/ssl/ssGiftCertificateRegister_common.asp"-->
<%
'********************************************************************************
'*   Sandshot Software Random Product Component									*
'*   Release Version   1.0														*
'*   Release Date      May 19, 2002												*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.				*
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/


'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Const enOrderItem_tmpID = 0
Const enOrderItem_prodID = 1
Const enOrderItem_prodName = 2
Const enOrderItem_prodPrice = 3
Const enOrderItem_prodPLPrice = 4
Const enOrderItem_prodSalePrice = 5
Const enOrderItem_prodPLSalePrice = 6
Const enOrderItem_prodSaleIsActive = 7
Const enOrderItem_prodImageSmallPath = 8
Const enOrderItem_prodLink = 9
Const enOrderItem_prodShipIsActive = 10
Const enOrderItem_prodWeight = 11
Const enOrderItem_prodShip = 12
Const enOrderItem_prodStateTaxIsActive = 13
Const enOrderItem_prodCountryTaxIsActive = 14
Const enOrderItem_gwPrice = 15
Const enOrderItem_gwActivate = 16
Const enOrderItem_FILLER = 17
Const enOrderItem_invenbTracked = 18
Const enOrderItem_invenInStock = 19
Const enOrderItem_invenLowFlag = 20

Const enOrderItem_odrdttmpQuantity = 21
Const enOrderItem_odrdttmpBackOrderQTY = 22
Const enOrderItem_odrdttmpGiftWrapQTY = 23

Const enOrderItem_QuantityByProductID = 24
Const enOrderItem_UnitPrice = 25
Const enOrderItem_MTPDiscount = 26
Const enOrderItem_UnitWeight = 27
Const enOrderItem_UnitLength = 28
Const enOrderItem_UnitWidth = 29
Const enOrderItem_UnitHeight = 30
Const enOrderItem_MustShipFreight = 31
Const enOrderItem_AttributeArray = 32
Const enOrderItem_AttributeCount = 33
Const enOrderItem_prodCategoryID = 34
Const enOrderItem_prodManufacturerID = 35
Const enOrderItem_prodVendorID = 36
Const enOrderItem_OrderDetailID = 37
Const enOrderItem_CategoryName = 38
Const enOrderItem_ManufacturerName = 39
Const enOrderItem_VendorName = 40
Const enOrderItem_attributeIDs = 41

Const enOrderItem_prodHandlingFee = 42
Const enOrderItem_prodSetupFee = 43
Const enOrderItem_SpecialTaxFlag_1 = 44
Const enOrderItem_prodFixedShippingCharge = 45
Const enOrderItem_prodSpecialShippingMethods = 46
Const enOrderItem_prodFileName = 47
Const enOrderItem_prodSetupFeeOneTime = 48
Const enOrderItem_prodBasePrice = 49

Const enOrderItems = 49

Const enAttributeItem_attrName = 0
Const enAttributeItem_attrdtID = 1
Const enAttributeItem_attrdtName = 2
Const enAttributeItem_attrdtPrice = 3
Const enAttributeItem_attrdtPLPrice = 4
Const enAttributeItem_attrdtType = 5
Const enAttributeItem_attrdtWeight = 6
Const enAttributeItem_attrPriceChange = 7
Const enAttributeItem_OrderAttributeID = 8
Const enAttributeItem_Length = 8

Dim mclsCartTotal

'**********************************************************
'*	Functions
'**********************************************************

'Function InitializeCart
'Function cartHasItems

'**********************************************************
'*	Begin Page Code
'**********************************************************

Sub cleanup_ssclsCartContents

On Error Resume Next

	Set mclsCartTotal = Nothing

End Sub	'cleanup_ssclsCartContents

'**********************************************************

Function cartHasItems
	If InitializeCart Then
		cartHasItems = Not mclsCartTotal.isEmptyCart
	Else
		cartHasItems = False
	End If
End Function	'cartHasItems

'**********************************************************

Function InitializeCart

	If Not isObject(mclsCartTotal) Then
		Set mclsCartTotal = New clsCartTotal
		With mclsCartTotal

			.City = visitorCity
			.State = visitorState
			.ZIP = visitorZIP
			.Country = visitorCountry
			.isCODOrder = False
			.ShipMethodCode = visitorPreferredShippingCode
			.EstimatedShipping = visitorEstimatedShipping

			.LoadAllShippingMethods = True
			.GetShippingRates = isCheckoutPage Or isOrderPage
			Call DebugRecordSplitTime("LoadCartContents . . . (ssl/SFLib/ssclsCartContents.asp?InitializeCart)")
			.LoadCartContents
			Call DebugRecordSplitTime("LoadCartContents complete")

			'.writeDebugCart

			.MinimumOrderMessage = "<SPAN style=""FONT-SIZE: 12pt; COLOR: #ff0000; FONT-FAMILY: Verdana""><B>Due to the cost of handling small orders a minimum purchase of {MinimumOrderAmount} is required.</B><BR />" _
									& "You need to spend {amountShort} more to check out.<br />Thank you for your understanding!</SPAN><BR />"
			.MinimumOrderAmount = 0

		End With	'mclsCartTotal
	End If

	InitializeCart = True

End Function	'InitializeCart

'**********************************************************

Sub removeCartFromSession
	If InitializeCart Then mclsCartTotal.removeCartFromSession
End Sub

'**********************************************************

Sub checkForUpdateVisitorShippingPreferences
	If Request.Form("updateVisitorShippingPreferences") = "Update" Then
		Call updateVisitorShippingPreferences(Request.Form("visitorState"), Request.Form("visitorZIP"), Request.Form("visitorCountry"), Request.Form("visitorPreferredShippingCode"))
		If InitializeCart Then
			Call setVisitorPreference("visitorEstimatedShipping", mclsCartTotal.Shipping, True)
		End If
	End If
End Sub

'**********************************************************

Call checkForUpdateVisitorShippingPreferences
Call checkForCertificateEntry(Request.Form("Certificate"))

'For testing
If False Then
	Set mclsCartTotal = New clsCartTotal
	With mclsCartTotal

		.City = visitorCity
		.State = visitorState
		.ZIP = visitorZIP
		.Country = visitorCountry
		.isCODOrder = True

		.ShipMethodCode = visitorPreferredShippingCode
		.LoadAllShippingMethods = True

		.LoadCartContents
		.checkInventoryLevels
		'.writeDebugCart

		.displayOrder
	End With	'mclsCartTotal

	Set mclsCartTotal = Nothing
End If	'False

'***********************************************************************************************
'***********************************************************************************************

Class clsCartTotal

Private pstrSessionName
Private cstrContentsDelimeter
Private paryOrderItems
Private paryAttributes
Private paryAvailableShippingMethods

Private pblnCartCreated
Private pblnStockDepleted
Private pblnConnectionOpened
Private pblnEmptyCart
Private pobjCnn
Private pstrCartProductsByID
Private pstrFullCart
Private pstrMiniCart
Private pstrViewCartText

'for calculating complete cart
Private pstrCity
Private pstrCountry
Private pstrCounty
Private pstrState
Private pstrZIP
Private pstrShipMethodCode
Private pstrShipMethodName
Private pblnCompleteCalculation

Private pblnFreeShipping
Private pblnCODOrder
Private pbytPremiumShipping

Private pblnGetShippingRates
Private pblnLoadAllShippingMethods
Private pblnOrderIsShipped
Private pdblSubTotal_LocalTaxable
Private pdblSubTotal_StateTaxable
Private pdblSubTotal_StateTaxable_Special
Private pdblSubTotal_CountryTaxable

Private pdblSubTotal
Private pdblSubTotal_Shipped
Private pdblDiscount
Private pdblDiscount_FirstTimeCustomer
Private pdblSubTotalWithDiscount
Private pdblEstimatedShipping
Private pdblShipping
Private pdblShipping_ProductBased
Private pdblShipping_ProductBased_BO
Private pdblHandling
Private pdblHandling_Order
Private pdblHandling_ProductSpecific
Private pdblHandling_ProductSetup
Private pdblHandling_ProductSetupOneTime
Private pdblCOD
Private pdblStateTax
Private pdblLocalTax
Private pdblCountryTax
Private pdblCartTotal
Private pdblAvailableStoreCredit
Private pdblStoreCredit
Private pdblAmountDue
Private pdblMinimumOrderAmount
Private pstrMinimumOrderMessage

Private pdblTaxRate_local
Private pdblTaxRate_state
Private pdblTaxRate_country
Private pblnReturningCustomer

Private plngOrderItemCount
Private plngUniqueOrderItemCount
Private pblnHasSavedCart

'For saving the order
Private plngOrderID

'***********************************************************************************************

Private Sub class_Initialize()

'////////////////////////////////////////////////////////////////////////////////
'//
'//		USER CONFIGURATION


		pstrSessionName = "cartContents"
		pstrViewCartText = "View Cart"

'//
'////////////////////////////////////////////////////////////////////////////////

	pblnConnectionOpened = False
	cstrContentsDelimeter = "||"
	pblnEmptyCart = True
	pblnCartCreated = False
	pblnOrderIsShipped = False
	pblnFreeShipping = False
	pblnCODOrder = False
	plngOrderItemCount = 0
	plngUniqueOrderItemCount = 0
	pblnLoadAllShippingMethods = False
	pblnGetShippingRates = True
	pblnStockDepleted = False
	pdblMinimumOrderAmount = 0
	pblnCompleteCalculation = True
	pblnReturningCustomer = False
	pdblDiscount_FirstTimeCustomer = 0
	pdblHandling_Order = 0
	pdblEstimatedShipping = 0

End Sub	'class_Initialize

Private Sub class_Terminate()

	On Error Resume Next

	If pblnConnectionOpened Then
		pobjCnn.Close
		Set pobjCnn = Nothing
	End If

End Sub

'***********************************************************************************************

Public Property Let Connection(byRef objCnn)
	If isObject(objCnn) Then Set pobjCnn = objCnn
End Property

Public Sub ResetCart
	pblnCartCreated = False
End Sub

Public Property Get CartCreated
	CartCreated = pblnCartCreated
End Property

Public Property Let isReturningCustomer(byVal Value)
	pblnReturningCustomer = Value
End Property

Public Property Let OrderID(byVal Value)
	plngOrderID = Value
End Property
Public Property Get OrderID
	OrderID = plngOrderID
End Property

Public Property Let MinimumOrderMessage(byVal Value)
	pstrMinimumOrderMessage = Value
End Property

Public Property Let MinimumOrderAmount(byVal Value)
	pdblMinimumOrderAmount = Value
End Property

Public Property Get isEmptyCart
	If Not pblnCartCreated Then Call LoadCartContents
	isEmptyCart = pblnEmptyCart
End Property

Public Property Get isOrderShipped
	If Not pblnCartCreated Then Call LoadCartContents
	isOrderShipped = pblnOrderIsShipped
End Property

Public Property Get isStockDepleted
	If Not pblnCartCreated Then Call LoadCartContents
	isStockDepleted = pblnStockDepleted
End Property

Public Property Get CartProductsByID
	CartProductsByID = pstrCartProductsByID
End Property

Public Property Let Country(byVal strValue)
	pstrCountry = strValue
End Property

Public Property Let City(byVal strValue)
	pstrCity = strValue
End Property

Public Property Let County(byVal strValue)
	pstrCounty = strValue
End Property

Public Property Let State(byVal strValue)
	pstrState = strValue
End Property

Public Property Let ZIP(byVal strValue)
	pstrZIP = strValue
End Property

Public Property Let isFreeShipping(byVal blnValue)
	pblnFreeShipping = blnValue
End Property

Public Property Let isCODOrder(byVal blnValue)
	pblnCODOrder = blnValue
End Property

Public Property Let isPremiumShipping(byVal bytValue)
	pbytPremiumShipping = bytValue
End Property

Public Property Let ShipMethodCode(byVal strValue)
	pstrShipMethodCode = strValue
End Property
Public Property Get ShipMethodCode
	ShipMethodCode = pstrShipMethodCode
End Property

Public Property Let EstimatedShipping(byVal strValue)
	If Len(strValue) > 0 And isNumeric(strValue) Then pdblEstimatedShipping = strValue
End Property

Public Property Let ShipMethodName(byVal strValue)
	pstrShipMethodName = strValue
End Property
Public Property Get ShipMethodName
	ShipMethodName = pstrShipMethodName
End Property

Public Property Get OrderItemCount
	OrderItemCount = plngOrderItemCount
End Property

Public Property Get UniqueOrderItemCount
	UniqueOrderItemCount = plngUniqueOrderItemCount
End Property

Public Property Get OrderItems()
	OrderItems = paryOrderItems
End Property

Public Property Get OrderItem(byVal index)
	OrderItem = paryOrderItems(index)
End Property

Public Property Let LoadAllShippingMethods(byVal blnValue)
	pblnLoadAllShippingMethods = blnValue
End Property

Public Property Let GetShippingRates(byVal blnValue)
	pblnGetShippingRates = blnValue
End Property

Public Property Get AmountDue
	AmountDue = pdblAmountDue
End Property

Public Property Get SubTotal
	SubTotal = pdblSubTotal
End Property

Public Property Get SubTotal_Shipped
	SubTotal_Shipped = pdblSubTotal_Shipped
End Property

Public Property Get Discount
	Discount = pdblDiscount
End Property

Public Property Get Discount_FirstTimeCustomer
	Discount_FirstTimeCustomer = pdblDiscount_FirstTimeCustomer
End Property

Public Property Get SubTotalWithDiscount
	SubTotalWithDiscount = pdblSubTotalWithDiscount
End Property

Public Property Get Shipping
	Shipping = pdblShipping
End Property

Public Property Get Handling
	Handling = pdblHandling
End Property

Public Property Get Handling_Order
	Handling_Order = pdblHandling_Order
End Property

Public Property Get Handling_ProductSpecific
	Handling_ProductSpecific = pdblHandling_ProductSpecific
End Property

Public Property Get Handling_ProductSetup
	Handling_ProductSetup = pdblHandling_ProductSetup
End Property

Public Property Get Handling_ProductSetupOneTime
	Handling_ProductSetupOneTime = pdblHandling_ProductSetupOneTime
End Property

Public Property Get COD
	COD = pdblCOD
End Property

Public Property Get StateTax
	StateTax = pdblStateTax
End Property

Public Property Get LocalTax
	LocalTax = pdblLocalTax
End Property

Public Property Get CountryTax
	CountryTax = pdblCountryTax
End Property

Public Property Get StoreCredit
	StoreCredit = pdblStoreCredit
End Property

Public Property Get AvailableStoreCredit
	AvailableStoreCredit = pdblAvailableStoreCredit
End Property

Public Property Get CartTotal
	CartTotal = pdblCartTotal
End Property

Public Property Get CompleteCalculation
	CompleteCalculation = pblnCompleteCalculation
End Property

'***********************************************************************************************

Private Function InitializeConnection()
'Initializes the connection to the database

Dim database_path
Dim ConnPasswords_RuntimeUserName
Dim ConnPasswords_RuntimePassword
Dim ServerPath
Dim pstrConnection

	pstrConnection = Application("DSN_Name")
	If Len(pstrConnection) = 0 Then	pstrConnection = Session("DSN_Name")
	If Len(pstrConnection) = 0 Then

		database_path = ""

		ConnPasswords_RuntimeUserName = "admin"
		ConnPasswords_RuntimePassword = ""

		pstrConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" _
						& "Data Source=" & database_PATH & ";"
	End If

	Set pobjCnn = CreateObject("ADODB.Connection")
	pobjCnn.Open pstrConnection

	If Err.number <> 0 then
		InitializeConnection = False
	Else
		InitializeConnection = (pobjCnn.State = 1)
	End If
	pblnConnectionOpened = True

End Function		' InitializeConnection

'***********************************************************************************************

Public Sub removeCartFromSession
	Call removeFromSession(pstrSessionName)
	'Response.Write "<h2><font color=red>Removing cart from session</font></h2>"
End Sub

'***********************************************************************************************

Private Sub LoadCartData

Dim i
Dim paryAttribute
Dim paryOrderItem
Dim paryTempCart()
Dim pblnNewOrderItem
Dim pblnNewProduct
Dim plngCounter_Attributes
Dim plngItemCount
Dim plngQuantityByProductID
Dim plngodrdttmpID
Dim plngRecordCount
Dim pobjRS
Dim pobjRSClone
Dim pstrPrevProdID
Dim pstrSQL

	If Len(SessionID) = 0 Then Exit Sub

	'Call removeCartFromSession
	paryOrderItems = getFromSession(pstrSessionName)
	If isArray(paryOrderItems) Then Exit Sub

	If cblnSF5AE Then
		'No Inventory
		pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfTmpOrderDetails.odrdttmpID, sfProducts.prodPrice, sfProducts.prodPLPrice, sfProducts.prodSalePrice, sfProducts.prodPLSalePrice, sfProducts.prodSaleIsActive, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodShipIsActive, sfProducts.prodWeight, sfProducts.prodLength, sfProducts.prodWidth, sfProducts.prodHeight, sfProducts.prodShip, sfProducts.prodStateTaxIsActive, sfProducts.prodCountryTaxIsActive, sfGiftWraps.gwPrice, sfGiftWraps.gwActivate, sfTmpOrderDetails.odrdttmpQuantity, sfTmpOrderDetailsAE.odrdttmpBackOrderQTY, sfTmpOrderDetailsAE.odrdttmpGiftWrapQTY, sfTmpOrdersAE.odrtmpCouponCode, sfTmpOrderAttributes.odrattrtmpAttrID, sfTmpOrderAttributes.odrattrtmpAttrText, sfAttributes.attrName, sfAttributes.attrDisplayStyle, sfAttributeDetail.attrdtName, sfAttributeDetail.attrdtPrice, sfAttributeDetail.attrdtType, sfAttributeDetail.attrdtPLPrice, sfAttributeDetail.attrdtWeight, sfProducts.prodCategoryID, sfProducts.prodManufacturerID, sfProducts.prodVendorID, sfProducts.prodSetupFee, sfProducts.prodHandlingFee, sfProducts.prodFixedShippingCharge, sfProducts.prodSpecialShippingMethods" _
				& " FROM sfAttributes RIGHT JOIN ((sfGiftWraps RIGHT JOIN ((sfProducts RIGHT JOIN (sfTmpOrdersAE RIGHT JOIN (sfTmpOrderDetailsAE INNER JOIN sfTmpOrderDetails ON sfTmpOrderDetailsAE.odrdttmpAEID = sfTmpOrderDetails.odrdttmpID) ON sfTmpOrdersAE.odrtmpSessionID = sfTmpOrderDetails.odrdttmpSessionID) ON sfProducts.prodID = sfTmpOrderDetails.odrdttmpProductID) LEFT JOIN sfTmpOrderAttributes ON sfTmpOrderDetails.odrdttmpID = sfTmpOrderAttributes.odrattrtmpOrderDetailId) ON sfGiftWraps.gwProdID = sfProducts.prodID) LEFT JOIN sfAttributeDetail ON sfTmpOrderAttributes.odrattrtmpAttrID = sfAttributeDetail.attrdtID) ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId" _
				& " WHERE sfTmpOrderDetails.odrdttmpSessionID=" & SessionID _
				& " ORDER BY sfProducts.prodName Asc, sfTmpOrderDetails.odrdttmpID Asc, sfProducts.prodID Asc, sfAttributes.attrDisplayOrder, sfAttributes.attrName, sfAttributeDetail.attrdtOrder, sfAttributeDetail.attrdtName"

		'With InventoryInfo join
		pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfTmpOrderDetails.odrdttmpID, sfProducts.prodPrice, sfProducts.prodPLPrice, sfProducts.prodSalePrice, sfProducts.prodPLSalePrice, sfProducts.prodSaleIsActive, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodShipIsActive, sfProducts.prodWeight, sfProducts.prodLength, sfProducts.prodWidth, sfProducts.prodHeight, sfProducts.prodShip, sfProducts.prodStateTaxIsActive, sfProducts.prodCountryTaxIsActive, sfGiftWraps.gwPrice, sfGiftWraps.gwActivate, sfTmpOrderDetails.odrdttmpQuantity, sfTmpOrderDetailsAE.odrdttmpBackOrderQTY, sfTmpOrderDetailsAE.odrdttmpGiftWrapQTY, sfTmpOrdersAE.odrtmpCouponCode, sfTmpOrderAttributes.odrattrtmpAttrID, sfTmpOrderAttributes.odrattrtmpAttrText, sfAttributes.attrName, sfAttributes.attrDisplayStyle, sfAttributeDetail.attrdtName, sfAttributeDetail.attrdtPrice, sfAttributeDetail.attrdtType, sfAttributeDetail.attrdtPLPrice, sfAttributeDetail.attrdtWeight, sfInventoryInfo.invenbTracked, sfProducts.prodCategoryID, sfProducts.prodManufacturerID, sfProducts.prodVendorID, sfProducts.prodSetupFee, sfProducts.prodSetupFeeOneTime, sfProducts.prodHandlingFee, sfProducts.prodFixedShippingCharge, sfProducts.prodSpecialShippingMethods, sfProducts.prodFileName" _
				& " FROM sfInventoryInfo RIGHT JOIN (sfAttributes RIGHT JOIN ((sfGiftWraps RIGHT JOIN ((sfProducts RIGHT JOIN (sfTmpOrdersAE RIGHT JOIN (sfTmpOrderDetailsAE INNER JOIN sfTmpOrderDetails ON sfTmpOrderDetailsAE.odrdttmpAEID = sfTmpOrderDetails.odrdttmpID) ON sfTmpOrdersAE.odrtmpSessionID = sfTmpOrderDetails.odrdttmpSessionID) ON sfProducts.prodID = sfTmpOrderDetails.odrdttmpProductID) LEFT JOIN sfTmpOrderAttributes ON sfTmpOrderDetails.odrdttmpID = sfTmpOrderAttributes.odrattrtmpOrderDetailId) ON sfGiftWraps.gwProdID = sfProducts.prodID) LEFT JOIN sfAttributeDetail ON sfTmpOrderAttributes.odrattrtmpAttrID = sfAttributeDetail.attrdtID) ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId) ON sfInventoryInfo.invenProdId = sfProducts.prodID" _
				& " WHERE sfTmpOrderDetails.odrdttmpSessionID=" & SessionID _
				& " ORDER BY sfProducts.prodName Asc, sfTmpOrderDetails.odrdttmpID Asc, sfProducts.prodID Asc, sfAttributes.attrDisplayOrder, sfAttributes.attrName, sfAttributeDetail.attrdtOrder, sfAttributeDetail.attrdtName"
	Else
		pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfTmpOrderDetails.odrdttmpID, sfProducts.prodPrice, sfProducts.prodPLPrice, sfProducts.prodSalePrice, sfProducts.prodPLSalePrice, sfProducts.prodSaleIsActive, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodShipIsActive, sfProducts.prodWeight, sfProducts.prodLength, sfProducts.prodWidth, sfProducts.prodHeight, sfProducts.prodShip, sfProducts.prodStateTaxIsActive, sfProducts.prodCountryTaxIsActive, sfTmpOrderDetails.odrdttmpQuantity, sfTmpOrderAttributes.odrattrtmpAttrID, sfTmpOrderAttributes.odrattrtmpAttrText, sfAttributes.attrName, sfAttributes.attrDisplayStyle, sfAttributeDetail.attrdtName, sfAttributeDetail.attrdtPrice, sfAttributeDetail.attrdtType, sfAttributeDetail.attrdtPLPrice, sfAttributeDetail.attrdtWeight, sfProducts.prodCategoryId, sfProducts.prodManufacturerId, sfProducts.prodVendorID, sfProducts.prodSetupFee, sfProducts.prodSetupFeeOneTime, sfProducts.prodHandlingFee, sfProducts.prodFixedShippingCharge, sfProducts.prodSpecialShippingMethods, sfProducts.prodFileName" _
				& " FROM (sfProducts RIGHT JOIN sfTmpOrderDetails ON sfProducts.prodID = sfTmpOrderDetails.odrdttmpProductID) LEFT JOIN (sfAttributes RIGHT JOIN (sfTmpOrderAttributes LEFT JOIN sfAttributeDetail ON sfTmpOrderAttributes.odrattrtmpAttrID = sfAttributeDetail.attrdtID) ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId) ON sfTmpOrderDetails.odrdttmpID = sfTmpOrderAttributes.odrattrtmpOrderDetailId" _
				& " WHERE sfTmpOrderDetails.odrdttmpSessionID=" & SessionID _
				& " ORDER BY sfProducts.prodName, sfTmpOrderDetails.odrdttmpID, sfProducts.prodID, sfAttributes.attrDisplayOrder, sfAttributes.attrName, sfAttributeDetail.attrdtOrder, sfAttributeDetail.attrdtName"
	End If

	'Sandshot modification for custom tax/handling fee implementation
	'pstrSQL = Replace(pstrSQL, ", sfProducts.prodHandlingFee", ", sfProducts.prodHandlingFee, sfProducts.prodILhalfTax", 1, 1)

	'debugprint "pstrSQL", pstrSQL
	'Response.Flush

	Set pobjRS = GetRS(pstrSQL)
	With pobjRS
		If .State = 0 Then
			plngUniqueOrderItemCount = -1
		ElseIf .EOF Then
			plngUniqueOrderItemCount = -1
		Else
			Set pobjRSClone = .Clone

			'Initialize variables
			plngUniqueOrderItemCount = -1
			plngQuantityByProductID = 0
			ReDim paryOrderItems(.RecordCount - 1)
			ReDim paryOrderItem(enOrderItems)

			For i = 0 To .RecordCount - 1
				pblnNewOrderItem = CBool(paryOrderItem(enOrderItem_tmpID) <> .Fields("odrdttmpID").Value)
				pblnNewProduct = CBool(paryOrderItem(enOrderItem_prodName) <> Trim(.Fields("prodName").Value))

				'debugprint i, .Fields("odrdttmpID").Value & " - " & .Fields("attrdtName").Value & "(" & pblnNewOrderItem & ")"
				If pblnNewOrderItem Then
					'New unique entry
					'Set Product Level entries
					ReDim paryOrderItem(enOrderItems)
					paryOrderItem(enOrderItem_tmpID) = .Fields("odrdttmpID").Value
					paryOrderItem(enOrderItem_prodID) = Trim(.Fields("prodID").Value)
					paryOrderItem(enOrderItem_prodName) = Trim(.Fields("prodName").Value)
					paryOrderItem(enOrderItem_prodPrice) = Trim(.Fields("prodPrice").Value)
					paryOrderItem(enOrderItem_prodPLPrice) = Trim(.Fields("prodPLPrice").Value)
					paryOrderItem(enOrderItem_prodSalePrice) = Trim(.Fields("prodSalePrice").Value)
					paryOrderItem(enOrderItem_prodPLSalePrice) = Trim(.Fields("prodPLSalePrice").Value)
					paryOrderItem(enOrderItem_prodSaleIsActive) = CorrectEmptyValue(.Fields("prodSaleIsActive").Value, 0)
					paryOrderItem(enOrderItem_prodImageSmallPath) = Trim(.Fields("prodImageSmallPath").Value)
					paryOrderItem(enOrderItem_prodLink) = Trim(.Fields("prodLink").Value & "")
					paryOrderItem(enOrderItem_prodShipIsActive) = CorrectEmptyValue(.Fields("prodShipIsActive").Value, 0)
					paryOrderItem(enOrderItem_prodWeight) = CorrectEmptyValue(.Fields("prodWeight").Value, 0)
					paryOrderItem(enOrderItem_prodShip) = CorrectEmptyValue(.Fields("prodShip").Value, 0)
					paryOrderItem(enOrderItem_prodStateTaxIsActive) = CorrectEmptyValue(.Fields("prodStateTaxIsActive").Value, 0)
					paryOrderItem(enOrderItem_prodCountryTaxIsActive) = CorrectEmptyValue(.Fields("prodCountryTaxIsActive").Value, 0)

					paryOrderItem(enOrderItem_prodHandlingFee) = CorrectEmptyValue(.Fields("prodHandlingFee").Value, 0)
					paryOrderItem(enOrderItem_prodSetupFee) = CorrectEmptyValue(.Fields("prodSetupFee").Value, 0)
					paryOrderItem(enOrderItem_prodSetupFeeOneTime) = CorrectEmptyValue(.Fields("prodSetupFeeOneTime").Value, 0)
					paryOrderItem(enOrderItem_prodFixedShippingCharge) = CorrectEmptyValue(.Fields("prodFixedShippingCharge").Value, 0)
					paryOrderItem(enOrderItem_prodSpecialShippingMethods) = Trim(.Fields("prodSpecialShippingMethods").Value & "")
					paryOrderItem(enOrderItem_prodFileName) = CorrectEmptyValue(.Fields("prodFileName").Value, "")

					'Zero out the one time per product fee
					If Not pblnNewProduct Then paryOrderItem(enOrderItem_prodSetupFeeOneTime) = 0

					'Sandshot modification for custom tax/handling fee implementation
					'paryOrderItem(enOrderItem_SpecialTaxFlag_1) = CorrectEmptyValue(.Fields("prodILhalfTax").Value, 0)
					'paryOrderItem(enOrderItem_prodHandlingFee) = CorrectEmptyValue(0.05, 0)
					'paryOrderItem(enOrderItem_prodSetupFee) = CorrectEmptyValue(5, 0)
					paryOrderItem(enOrderItem_SpecialTaxFlag_1) = CorrectEmptyValue(1, 0)

					If cblnSF5AE Then
						paryOrderItem(enOrderItem_gwPrice) = CorrectEmptyValue(Trim(.Fields("gwPrice").Value), 0)
						paryOrderItem(enOrderItem_gwActivate) = CorrectEmptyValue(Trim(.Fields("gwActivate").Value), 0)

						paryOrderItem(enOrderItem_invenbTracked) = CorrectEmptyValue(Trim(.Fields("invenbTracked").Value), 0)
						paryOrderItem(enOrderItem_invenInStock) = 0
						paryOrderItem(enOrderItem_invenLowFlag) = 0

						paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) = CorrectEmptyValue(.Fields("odrdttmpBackOrderQTY").Value, 0)
						paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) = CorrectEmptyValue(.Fields("odrdttmpGiftWrapQTY").Value, 0)
					Else
						paryOrderItem(enOrderItem_gwPrice) = 0
						paryOrderItem(enOrderItem_gwActivate) = 0

						paryOrderItem(enOrderItem_invenbTracked) = 0
						paryOrderItem(enOrderItem_invenInStock) = 0
						paryOrderItem(enOrderItem_invenLowFlag) = 0

						paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) = 0
						paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) = 0
					End If

					paryOrderItem(enOrderItem_prodCategoryID) = Trim(.Fields("prodCategoryID").Value)
					paryOrderItem(enOrderItem_prodManufacturerID) = Trim(.Fields("prodManufacturerID").Value)
					paryOrderItem(enOrderItem_prodVendorID) = Trim(.Fields("prodVendorID").Value)

					paryOrderItem(enOrderItem_odrdttmpQuantity) = CDbl(CorrectEmptyValue(.Fields("odrdttmpQuantity").Value, 0))

					paryOrderItem(enOrderItem_UnitLength) = CorrectEmptyValue(.Fields("prodLength").Value, 0)
					paryOrderItem(enOrderItem_UnitWidth) = CorrectEmptyValue(.Fields("prodWidth").Value, 0)
					paryOrderItem(enOrderItem_UnitHeight) = CorrectEmptyValue(.Fields("prodHeight").Value, 0)
					paryOrderItem(enOrderItem_MustShipFreight) = False

					pblnOrderIsShipped = pblnOrderIsShipped Or CBool(paryOrderItem(enOrderItem_prodShipIsActive) = 1)

					plngUniqueOrderItemCount = plngUniqueOrderItemCount + 1
					plngItemCount = plngItemCount + paryOrderItem(enOrderItem_odrdttmpQuantity)

					'Now set the Sell Price
					If paryOrderItem(enOrderItem_prodSaleIsActive) = 0 Then
						paryOrderItem(enOrderItem_UnitPrice) = CDbl(GetPricingLevelPrice(paryOrderItem(enOrderItem_prodPrice), paryOrderItem(enOrderItem_prodPLPrice)))
					Else
						paryOrderItem(enOrderItem_UnitPrice) = CDbl(GetPricingLevelPrice(paryOrderItem(enOrderItem_prodSalePrice), paryOrderItem(enOrderItem_prodPLSalePrice)))
					End If
					paryOrderItem(enOrderItem_prodBasePrice) = paryOrderItem(enOrderItem_UnitPrice)

					If pstrPrevProdID = paryOrderItem(enOrderItem_prodID) Then
						plngQuantityByProductID = plngQuantityByProductID + paryOrderItem(enOrderItem_odrdttmpQuantity)
					Else
						pstrPrevProdID = paryOrderItem(enOrderItem_prodID)
						plngQuantityByProductID = paryOrderItem(enOrderItem_odrdttmpQuantity)
					End If
					paryOrderItem(enOrderItem_QuantityByProductID) = plngQuantityByProductID

					pobjRSClone.Filter = "odrdttmpID=" & paryOrderItem(enOrderItem_tmpID)
					paryOrderItem(enOrderItem_AttributeCount) = pobjRSClone.RecordCount
					ReDim paryAttributes(paryOrderItem(enOrderItem_AttributeCount) - 1)
					plngCounter_Attributes = 0

				End If	'plngodrdttmpID <> .Fields("odrdttmpID").Value

				ReDim paryAttribute(enAttributeItem_Length)
				paryAttribute(enAttributeItem_attrdtID) = Trim(.Fields("odrattrtmpAttrID").Value & "")
				paryAttribute(enAttributeItem_attrName) = Trim(.Fields("attrName").Value & "")
				paryAttribute(enAttributeItem_attrdtPrice) = CorrectEmptyValue(.Fields("attrdtPrice").Value, 0)
				paryAttribute(enAttributeItem_attrdtPLPrice) = Trim(.Fields("attrdtPLPrice").Value & "")
				paryAttribute(enAttributeItem_attrdtType) = CorrectEmptyValue(.Fields("attrdtType").Value, 0)
				If isCustomerDefinedOption(.Fields("attrDisplayStyle").Value) Then
					If Len(Trim(.Fields("odrattrtmpAttrText").Value & "")) > 0 Then
						paryAttribute(enAttributeItem_attrdtName) = Trim(.Fields("odrattrtmpAttrText").Value & "")
					Else
						paryAttribute(enAttributeItem_attrdtName) = emptyCustomerDefinedOptionText
					End If
				Else
					paryAttribute(enAttributeItem_attrdtName) = Trim(.Fields("attrdtName").Value & "")
				End If

				paryAttribute(enAttributeItem_attrdtWeight) = CorrectEmptyValue(.Fields("attrdtWeight").Value, 0)
				'Now every product will have an attribute due to join
				'Use an empty name to test
				If Len(Trim(.Fields("attrName").Value & "")) = 0 Then paryOrderItem(enOrderItem_AttributeCount) = -1

				Select Case paryAttribute(enAttributeItem_attrdtType)
					Case "0"	'=
						paryAttribute(enAttributeItem_attrPriceChange) = 0
					Case "1"	'add
						paryAttribute(enAttributeItem_attrPriceChange) = CDbl(GetPricingLevelPrice(paryAttribute(enAttributeItem_attrdtPrice), paryAttribute(enAttributeItem_attrdtPLPrice)))
					Case "2"	'subtract
						paryAttribute(enAttributeItem_attrPriceChange) = -1 * CDbl(GetPricingLevelPrice(paryAttribute(enAttributeItem_attrdtPrice), paryAttribute(enAttributeItem_attrdtPLPrice)))
					Case Else
						paryAttribute(enAttributeItem_attrPriceChange) = 0
				End Select

				paryAttributes(plngCounter_Attributes) = paryAttribute
				plngCounter_Attributes = plngCounter_Attributes + 1

				.MoveNext

				'Now clean up
				If .EOF Then
					pblnNewOrderItem = True
				ElseIf CBool(paryOrderItem(enOrderItem_tmpID) = .Fields("odrdttmpID").Value) Then
					pblnNewOrderItem = False
				Else
					pblnNewOrderItem = True
				End If
				If pblnNewOrderItem Then
					If paryOrderItem(enOrderItem_AttributeCount) = -1 Then
						paryOrderItem(enOrderItem_AttributeCount) = 0
					Else
						paryOrderItem(enOrderItem_AttributeCount) = plngCounter_Attributes
					End If
					paryOrderItem(enOrderItem_AttributeArray) = paryAttributes
					paryOrderItems(plngUniqueOrderItemCount) = paryOrderItem
				End If

				If False Then
					Response.Write "<fieldset><legend>LoadCartData - Item " & plngUniqueOrderItemCount & "</legend>"
					Response.Write "prodName: " & paryOrderItem(enOrderItem_prodName) & "<br />"
					Response.Write "pblnNewOrderItem: " & pblnNewOrderItem & "<br />"
					If isArray(paryOrderItems(plngUniqueOrderItemCount)) Then
						If isArray(paryOrderItems(plngUniqueOrderItemCount)(enOrderItem_AttributeArray)) Then
							Response.Write "Order item <b>has</b> attributes<br />"
						Else
							Response.Write "Order item has no attributes<br />"
						End If
						Call writeDebugOrderItem(paryOrderItems(plngUniqueOrderItemCount))
					Else
						Response.Write "<b>Order Item Not Set Yet</b><br />"
					End If
					Response.Write "Attribute: " & plngCounter_Attributes  & "<br />"
					Response.Write "enAttributeItem_attrName: " & paryAttribute(enAttributeItem_attrName) & "<br />"
					Response.Write "enAttributeItem_attrdtName: " & paryAttribute(enAttributeItem_attrdtName) & "<br />"
					Response.Write "</fieldset>"
				End If

			Next 'i

			Call closeObj(pobjRSClone)
		End If	'.EOF
	End With	'pobjRS
	Call closeObj(pobjRS)

	If plngUniqueOrderItemCount > -1 Then
		pstrPrevProdID = ""
		ReDim paryTempCart(plngUniqueOrderItemCount)
		For i = plngUniqueOrderItemCount To 0 Step -1
			If pstrPrevProdID = paryOrderItems(i)(enOrderItem_prodID) Then
				paryOrderItems(i)(enOrderItem_MTPDiscount) = paryOrderItems(i+1)(enOrderItem_MTPDiscount)
				paryOrderItems(i)(enOrderItem_QuantityByProductID) = paryOrderItems(i+1)(enOrderItem_QuantityByProductID)
			Else
				If cblnSF5AE Then
					paryOrderItems(i)(enOrderItem_MTPDiscount) = GetMTPrice2(paryOrderItems(i)(enOrderItem_prodID), paryOrderItems(i)(enOrderItem_QuantityByProductID), paryOrderItems(i)(enOrderItem_UnitPrice))
				Else
					paryOrderItems(i)(enOrderItem_MTPDiscount) = paryOrderItems(i)(enOrderItem_UnitPrice)
				End If
				pstrPrevProdID = paryOrderItem(enOrderItem_prodID)
			End If	'pstrPrevProdID = paryOrderItems(i)(enOrderItem_prodID)

			'Now sum the attribute info to the top line
			paryOrderItems(i)(enOrderItem_UnitPrice) = paryOrderItems(i)(enOrderItem_MTPDiscount)
			paryOrderItems(i)(enOrderItem_UnitWeight) = paryOrderItems(i)(enOrderItem_prodWeight)

			paryAttributes = paryOrderItems(i)(enOrderItem_AttributeArray)
			If isArray(paryAttributes) Then
				For plngCounter_Attributes = 0 To UBound(paryAttributes)
					paryAttribute = paryAttributes(plngCounter_Attributes)
					paryOrderItems(i)(enOrderItem_UnitPrice) = paryOrderItems(i)(enOrderItem_UnitPrice) + paryAttribute(enAttributeItem_attrPriceChange)
					paryOrderItems(i)(enOrderItem_UnitWeight) = paryOrderItems(i)(enOrderItem_UnitWeight) + paryAttribute(enAttributeItem_attrdtWeight)
				Next 'plngCounter_Attributes
			End If

			paryTempCart(i) = paryOrderItems(i)

		Next 'i

		paryOrderItems = paryTempCart

		Call saveToSession(pstrSessionName, paryOrderItems, DateAdd("m", 15, Now()))
	End If	'plngUniqueOrderItemCount > -1

End Sub	'LoadCartData

'***********************************************************************************************

Public Sub LoadCartContents

Dim i
Dim paryAttribute
Dim paryOrderItem
Dim paryTempCart()
Dim pblnNewOrderItem
Dim pdblProductExtendedAmount
Dim pdblProductExtendedAmountBO
Dim plngCounter_Attributes
Dim plngItemCount
Dim plngQuantityByProductID
Dim plngodrdttmpID
Dim plngRecordCount
Dim pobjRS
Dim pobjRSClone
Dim pstrPrevProdID
Dim pstrSQL

	If Len(SessionID) = 0 Then Exit Sub

	Call LoadCartData
	If plngUniqueOrderItemCount > -1 Then

		'Initialize the variables
		pdblSubTotal_LocalTaxable = 0
		pdblSubTotal_StateTaxable = 0
		pdblSubTotal_CountryTaxable = 0
		pdblSubTotal_StateTaxable_Special = 0

		pdblSubTotal = 0
		pdblSubTotal_Shipped = 0
		pdblDiscount = 0
		pdblSubTotalWithDiscount = 0
		pdblStateTax = 0
		pdblLocalTax = 0
		pdblCountryTax = 0
		pdblCartTotal = 0
		pdblStoreCredit = 0
		pdblAvailableStoreCredit = 0
		pdblAmountDue = 0
		pdblShipping_ProductBased = 0
		pdblShipping_ProductBased_BO = 0
		pdblHandling_ProductSpecific = 0
		pdblHandling_ProductSetup = 0
		pdblHandling_ProductSetupOneTime = 0

		'Now calculate the order values
		For i = plngUniqueOrderItemCount To 0 Step -1
			plngOrderItemCount = plngOrderItemCount + paryOrderItems(i)(enOrderItem_odrdttmpQuantity)

			pdblProductExtendedAmount = paryOrderItems(i)(enOrderItem_UnitPrice) * paryOrderItems(i)(enOrderItem_odrdttmpQuantity) _
									  + paryOrderItems(i)(enOrderItem_gwPrice) * paryOrderItems(i)(enOrderItem_odrdttmpGiftWrapQTY)

			'add in the setup fees
			pdblProductExtendedAmount = pdblProductExtendedAmount + paryOrderItems(i)(enOrderItem_prodSetupFee)
			pdblProductExtendedAmount = pdblProductExtendedAmount + paryOrderItems(i)(enOrderItem_prodSetupFeeOneTime)

			'pdblProductExtendedAmountBO = paryOrderItems(i)(enOrderItem_UnitPrice) * paryOrderItems(i)(enOrderItem_odrdttmpBackOrderQTY)

			'Now add in the gift wrap for BO items
			If (paryOrderItems(i)(enOrderItem_odrdttmpQuantity) - paryOrderItems(i)(enOrderItem_odrdttmpBackOrderQTY)) > paryOrderItems(i)(enOrderItem_odrdttmpGiftWrapQTY) Then
				pdblProductExtendedAmountBO = pdblProductExtendedAmountBO + paryOrderItems(i)(enOrderItem_gwPrice) * (paryOrderItems(i)(enOrderItem_odrdttmpGiftWrapQTY) - paryOrderItems(i)(enOrderItem_odrdttmpQuantity) + paryOrderItems(i)(enOrderItem_odrdttmpBackOrderQTY))
			End If

			If paryOrderItems(i)(enOrderItem_prodStateTaxIsActive) Then pdblSubTotal_LocalTaxable = pdblSubTotal_LocalTaxable + pdblProductExtendedAmount
			If paryOrderItems(i)(enOrderItem_prodStateTaxIsActive) Then pdblSubTotal_StateTaxable = pdblSubTotal_StateTaxable + pdblProductExtendedAmount
			If paryOrderItems(i)(enOrderItem_SpecialTaxFlag_1) Then pdblSubTotal_StateTaxable_Special = pdblSubTotal_StateTaxable_Special + pdblProductExtendedAmount
			If CBool(paryOrderItems(i)(enOrderItem_prodShipIsActive) = 1) Then
				pdblSubTotal_Shipped = pdblSubTotal_Shipped + pdblProductExtendedAmount
				pblnOrderIsShipped = True
			End If

			If paryOrderItems(i)(enOrderItem_prodCountryTaxIsActive) Then pdblSubTotal_CountryTaxable = pdblSubTotal_CountryTaxable + pdblProductExtendedAmount
			pdblSubTotal = pdblSubTotal + pdblProductExtendedAmount

			pdblShipping_ProductBased = pdblShipping_ProductBased + paryOrderItems(i)(enOrderItem_prodShip) * paryOrderItems(i)(enOrderItem_odrdttmpQuantity)
			pdblShipping_ProductBased_BO = pdblShipping_ProductBased + paryOrderItems(i)(enOrderItem_prodShip) * paryOrderItems(i)(enOrderItem_odrdttmpBackOrderQTY)

			pdblHandling_ProductSetup = pdblHandling_ProductSetup + paryOrderItems(i)(enOrderItem_prodSetupFee)
			pdblHandling_ProductSetupOneTime = pdblHandling_ProductSetupOneTime + paryOrderItems(i)(enOrderItem_prodSetupFeeOneTime)
			pdblHandling_ProductSpecific = pdblHandling_ProductSpecific + CDbl(paryOrderItems(i)(enOrderItem_prodHandlingFee)) * CDbl(paryOrderItems(i)(enOrderItem_odrdttmpQuantity))

			Call ssGiftCertificate_CheckCartForCertificates_MT(paryOrderItems(i))
			Call CheckCartForDiscountClub(paryOrderItems(i))
		Next 'i

		'now calculate any discounts
		pblnReturningCustomer = visitorHasPriorOrders

		Dim pclsPromotion
		Set pclsPromotion = New clsPromotion
		with pclsPromotion
			.Connection = cnn
			.subTotal = pdblSubTotal
			.Promotions = vistorDiscountCodes
			.SetOrderItems paryOrderItems
			.CustID = visitorLoggedInCustomerID
			.CalcBestDiscount
			pdblDiscount = .BestDiscountAmount

			If Not pblnReturningCustomer Then
				pdblDiscount_FirstTimeCustomer = .FirstTimeCustomerDiscountAmount
				pdblDiscount = pdblDiscount + pdblDiscount_FirstTimeCustomer
			End If

			If pdblDiscount > pdblSubTotal Then pdblDiscount = pdblSubTotal
		end with
		Set pclsPromotion = Nothing

		pdblSubTotalWithDiscount = pdblSubTotal - pdblDiscount

		'Calculate handling charge
		If pdblSubTotalWithDiscount < adminHandlingIsActive Then
			If adminHandlingType = 1 Or pblnOrderIsShipped Then
				pdblHandling_Order = adminHandling
			Else
				pdblHandling_Order = 0
			End If
		Else
			pdblHandling_Order = 0
		End If
		'pdblHandling = pdblHandling_Order + pdblHandling_ProductSetup + pdblHandling_ProductSpecific
		pdblHandling = pdblHandling_Order + pdblHandling_ProductSpecific

		'Calculate COD charge
		If pblnCODOrder Then
			pdblCOD = adminCODAmount
		Else
			pdblCOD = 0
		End If

		'now calculate the shipping
		If pdblSubTotal_Shipped > pdblDiscount Then
			pdblSubTotal_Shipped = pdblSubTotal_Shipped - pdblDiscount
		Else
			pdblSubTotal_Shipped = 0
		End If
		Call CalculateShipping

		Call checkDiscountedShipping(pdblSubTotal, pdblShipping, pstrShipMethodCode)

		'now for the tax
		Call getTaxRates(pstrZIP, pstrState, pstrCountry)

		pblnCompleteCalculation = Len(pstrZIP) > 0 And Len(pstrState) > 0 And Len(pstrCountry) > 0

		Dim pdblAmountToTax_Local
		Dim pdblAmountToTax_State
		Dim pdblAmountToTax_Country

		pdblAmountToTax_Local = pdblSubTotal_LocalTaxable - pdblDiscount
		pdblAmountToTax_State = pdblSubTotal_StateTaxable - pdblDiscount
		pdblAmountToTax_Country = pdblSubTotal_CountryTaxable - pdblDiscount

		'Custom tax rule - IL taxes only half the value of flagged items
		'If pstrState = "IL" Then
		'	pdblAmountToTax_State = pdblAmountToTax_State - 0.5 * pdblSubTotal_StateTaxable_Special
		'End If

		If adminTaxShipIsActive Then
			pdblAmountToTax_Local = pdblAmountToTax_Local + pdblShipping
			pdblAmountToTax_State = pdblAmountToTax_State + pdblShipping
			pdblAmountToTax_Country = pdblAmountToTax_Country + pdblShipping
		End If

		If False Then	'tax handing?
			pdblAmountToTax_Local = pdblAmountToTax_Local + pdblHandling
			pdblAmountToTax_State = pdblAmountToTax_State + pdblHandling
			pdblAmountToTax_Country = pdblAmountToTax_Country + pdblHandling
		End If

		pdblLocalTax = Round(pdblAmountToTax_Local * pdblTaxRate_local, 2)
		pdblStateTax = Round(pdblAmountToTax_State * pdblTaxRate_state, 2)
		pdblCountryTax = Round(pdblAmountToTax_Country * pdblTaxRate_country, 2)

		pdblCartTotal = pdblSubTotalWithDiscount + pdblHandling + pdblCOD + pdblShipping + pdblLocalTax + pdblStateTax + pdblCountryTax

		'now calculate any store credits
		Dim pclsssGiftCertificate
		pdblStoreCredit = 0
		pdblAvailableStoreCredit = 0

		If Len(visitorCertificateCodes) > 0 Then
			Set pclsssGiftCertificate = New clsssGiftCertificate
			pclsssGiftCertificate.Connection = cnn
			If pclsssGiftCertificate.validateCertificate(visitorCertificateCodes) Then
				pdblAvailableStoreCredit = pclsssGiftCertificate.CertificateValue

				'Now limit to order amount
				If pdblAvailableStoreCredit > pdblCartTotal Then
					pdblStoreCredit = pdblCartTotal
				Else
					pdblStoreCredit = pdblAvailableStoreCredit
				End If

				mdblssGCNewTotalDue = mdblssGCOriginalTotalDue - mdblssCertificateAmount

			End If
			Set pclsssGiftCertificate = Nothing
		End If

		pdblAmountDue = pdblCartTotal - pdblStoreCredit

	End If	'plngUniqueOrderItemCount > -1

	pblnCartCreated = True

	pblnEmptyCart = CBool(plngUniqueOrderItemCount < 0)	'check against 0 since it is a 0 based array

End Sub	'LoadCartContents

'***********************************************************************************************

Public Function checkInventoryLevels

Dim i
Dim paryAttribute
Dim plngCounter_Attributes
Dim pstrAttributeIDs

	For i = 0 To plngUniqueOrderItemCount
		If paryOrderItems(i)(enOrderItem_invenbTracked) = 1 Then
			paryAttributes = paryOrderItems(i)(enOrderItem_AttributeArray)
			If isArray(paryAttributes) Then
				pstrAttributeIDs = ""	'initialize the attributes
				For plngCounter_Attributes = 0 To UBound(paryAttributes)
					paryAttribute = paryAttributes(plngCounter_Attributes)
					If Len(pstrAttributeIDs) = 0 Then
						pstrAttributeIDs = paryAttribute(enAttributeItem_attrdtID)
					Else
						pstrAttributeIDs = pstrAttributeIDs & "," & paryAttribute(enAttributeItem_attrdtID)
					End If
				Next 'plngCounter_Attributes
			End If

			If Len(pstrAttributeIDs) = 0 Then
				pstrAttributeIDs = "0"
			Else
				pstrAttributeIDs = bubbleSortAttributeIDList(pstrAttributeIDs, ",")
			End If
			paryOrderItems(i)(enOrderItem_attributeIDs) = pstrAttributeIDs
			paryOrderItems(i)(enOrderItem_invenInStock) = GetAvailableQty(paryOrderItems(i)(enOrderItem_prodID), pstrAttributeIDs)

			'Only check if a customer hasn't accepted a back order
			If paryOrderItems(i)(enOrderItem_odrdttmpBackOrderQTY) = 0 Then
				pblnStockDepleted = pblnStockDepleted Or CBool(paryOrderItems(i)(enOrderItem_invenInStock) < paryOrderItems(i)(enOrderItem_odrdttmpQuantity))
			End If

		Else
			paryOrderItems(i)(enOrderItem_invenInStock) = "X"
		End If	'paryOrderItems(i)(enOrderItem_invenbTracked) = 1
	Next 'i

End Function	'checkInventoryLevels

'***********************************************************************************************

Public Function bubbleSortAttributeIDList(byVal strList, byVal strDelimiter)
'use the bubble sort since it is generally the best to use for smaller (ie < 25) items)

Dim i
Dim j
Dim paryList
Dim pstrTemp

	If Len(strList) > 0 Then
		paryList = Split(strList, strDelimiter)
		For i = UBound(paryList) - 1 To 0 Step -1
			For j = 0 To i
				If CLng(paryList(j)) > CLng(paryList(j+1)) Then
					pstrTemp = paryList(j+1)
					paryList(j+1) = paryList(j)
					paryList(j) = pstrTemp
				End If
			Next 'j
		Next 'i

		strList = paryList(0)
		For i = 1 To UBound(paryList)
			strList = strList & strDelimiter & paryList(i)
		Next 'i
	End If	'Len(strList) > 0

	bubbleSortAttributeIDList = strList

End Function	'bubbleSortAttributeIDList

'***********************************************************************************************

Private Sub getTaxRates(byVal strLocale, byVal strState, byVal strCountry)

Dim pstrSQL
Dim pobjCmd
Dim pobjRS

	'tax rates stored in Application for performance reasons
	'Formats: Application("taxRate_locale-strLocale")
	'Formats: Application("taxRate_state-strState")
	'Formats: Application("taxRate_country-strCountry")

	pdblTaxRate_local = 0
	pdblTaxRate_state = 0
	pdblTaxRate_country = 0

	'get the local tax rate
	If Len(strLocale) > 0 Then
		pdblTaxRate_local = Application("taxRate_locale-" & strLocale)
		If Len(CStr(pdblTaxRate_local)) = 0 Then
			pstrSQL = "Select TaxRate From ssTaxTable Where PostalCode=?"
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = pstrSQL
				Set .ActiveConnection = cnn
				.Parameters.Append .CreateParameter("taxDistrict", adVarChar, adParamInput, 10, strLocale)
				Set pobjRS = .Execute
				If pobjRS.EOF Then
					pdblTaxRate_local = 0
				Else
					pdblTaxRate_local = pobjRS.Fields("TaxRate").Value
				End If
				pobjRS.Close
				Set pobjRS = Nothing
			End With	'pobjCmd
			Set pobjCmd = Nothing
			Application("taxRate_locale-" & strLocale) = pdblTaxRate_local
		End If	'Len(CStr(pdblTaxRate_local)) = 0
	End If	'Len(strLocale) > 0

	'get the state tax rate
	If Len(strState) > 0 Then
		If True Then
			pdblTaxRate_state = getStateTaxRate(strState)
		Else
			pdblTaxRate_state = Application("taxRate_state-" & strState)
			If Len(CStr(pdblTaxRate_state)) = 0 Then

				pstrSQL = "Select loclstTax FROM sfLocalesState WHERE loclstLocaleIsActive = 1 AND loclstAbbreviation=?"
				Set pobjCmd  = CreateObject("ADODB.Command")
				With pobjCmd
					.Commandtype = adCmdText
					.Commandtext = pstrSQL
					Set .ActiveConnection = cnn
					.Parameters.Append .CreateParameter("taxDistrict", adVarChar, adParamInput, 3, strState)
					Set pobjRS = .Execute
					If pobjRS.EOF Then
						pdblTaxRate_state = 0
					Else
						pdblTaxRate_state = pobjRS.Fields("loclstTax").Value
					End If
					pobjRS.Close
					Set pobjRS = Nothing
				End With	'pobjCmd
				Set pobjCmd = Nothing
				Application("taxRate_state-" & strState) = pdblTaxRate_state
			End If	'Len(CStr(pdblTaxRate_state)) = 0
		End If
	End If	'Len(strState) > 0

	'get the country tax rate
	If Len(strCountry) > 0 Then
		If True Then
			pdblTaxRate_country = getCountryTaxRate(strCountry)
		Else
			pdblTaxRate_country = Application("taxRate_country-" & strCountry)
			If Len(CStr(pdblTaxRate_country)) = 0 Then
				pstrSQL = "SELECT loclctryTax FROM sfLocalesCountry WHERE  loclctryLocalIsActive = 1 AND loclctryAbbreviation=?"
				Set pobjCmd  = CreateObject("ADODB.Command")
				With pobjCmd
					.Commandtype = adCmdText
					.Commandtext = pstrSQL
					Set .ActiveConnection = cnn
					.Parameters.Append .CreateParameter("taxDistrict", adVarChar, adParamInput, 3, strCountry)
					Set pobjRS = .Execute
					If pobjRS.EOF Then
						pdblTaxRate_country = 0
					Else
						pdblTaxRate_country = pobjRS.Fields("loclctryTax").Value
					End If
					pobjRS.Close
					Set pobjRS = Nothing
				End With	'pobjCmd
				Set pobjCmd = Nothing
				Application("taxRate_country-" & strCountry) = pdblTaxRate_country
			End If	'Len(CStr(pdblTaxRate_country)) = 0
		End If
	End If	' Len(strCountry) > 0

	If False Then
		Response.Write "<fieldset><legend>Tax Rates</legend>"
		Response.Write "Local Tax Rate (" & strLocale & "): " & pdblTaxRate_local & "<br />"
		Response.Write "State Tax Rate (" & strState & "): " & pdblTaxRate_state & "<br />"
		Response.Write "Country Tax Rate (" & strCountry & "): " & pdblTaxRate_country & "<br />"
		Response.Write "</fieldset>"
	End If

End Sub 'getTaxRates

'***********************************************************************************************

Private Sub CalculateShipping

Dim i
Dim pblnValidShippingDestination
Dim pclsShipping
Dim pvntTempShippingMethod

	If pblnFreeShipping Or Not pblnOrderIsShipped Then
		pdblShipping = 0
	Else
		Select Case adminShipType
			Case 1, 4, 5	'Zone Based Shipping
				Set pclsShipping= New clsZoneBasedShipping
				With pclsShipping

					If cblnDebugZBS Then
						Response.Write "Destination Country:" & pstrCountry & "<br />"
						Response.Write "Destination State:" & pstrState & "<br />"
						Response.Write "Destination Zip:" & pstrZIP & "<br />"
						Response.Write "Ship Code:" & pstrShipMethodCode & "<br />"
						Response.Write "Total Purchase:" & pdblSubTotalWithDiscount & "<br />"
						Response.Write "Weight:" & mdblOrderWeight& "<br />"
						Response.Write "Item Count:" & plngOrderItemCount & "<br />"
						Response.Flush
					End If

					.Connection = cnn                        'Set the connection to the StoreFront database
					.ZoneType = 0                            'Set to 0 for Country, 1 for State, 2 for ZIP
					.SetZone(pstrCountry)                    'Set to the shipping Country, State, or ZIP

					If pblnLoadAllShippingMethods Then
						pvntTempShippingMethod = pstrShipMethodCode

						If adminShipType = 1 Then		'Use this line to go by order amount
							'pdblShipping = .GetAnyAvailableShipping(pstrShipMethodCode, pdblSubTotalWithDiscount)
							pdblShipping = .GetAnyAvailableShipping(pstrShipMethodCode, pdblSubTotal_Shipped)
						ElseIf adminShipType = 4 Then	'Use this line to go by weight
							pdblShipping = .GetAnyAvailableShipping(pstrShipMethodCode, mdblOrderWeight)
						ElseIf adminShipType = 5 Then	'Use this line to go by order item count
							pdblShipping = .GetAnyAvailableShipping(pstrShipMethodCode, plngOrderItemCount)
						End If
						If pvntTempShippingMethod <> pstrShipMethodCode Then
							'could manually save to visitorPreferredShippingCode here but that should be done outside the class
						End If
						paryAvailableShippingMethods = .availableRates
					Else
						If adminShipType = 1 Then		'Use this line to go by order amount
							'pdblShipping = .GetRate(pstrShipMethodCode, pdblSubTotalWithDiscount)
							pdblShipping = .GetRate(pstrShipMethodCode, pdblSubTotal_Shipped)
						ElseIf adminShipType = 4 Then	'Use this line to go by weight
							pdblShipping = .GetRate(pstrShipMethodCode, mdblOrderWeight)
						ElseIf adminShipType = 5 Then	'Use this line to go by order item count
							pdblShipping = .GetRate(pstrShipMethodCode, plngOrderItemCount)
						End If
						If CStr(pdblShipping) = "FAIL" And cblnAutomaticallyFindAnyAvailable Then pdblShipping = .GetAnyAvailableShipping(pstrShipMethodCode, pdblSubTotal_Shipped)
					End If

					If isNumeric(pstrShipMethodCode) Then
						pstrShipMethodName = getNameWithID("sfShipping", pstrShipMethodCode, "shipID", "shipMethod", 0)
					Else
						pstrShipMethodName = getNameWithID("sfShipping", pstrShipMethodCode, "shipCode", "shipMethod", 1)
					End If

					'debugprint "pstrShipMethodCode",pstrShipMethodCode
					'debugprint "pstrShipMethodName",pstrShipMethodName
					'debugprint "pdblSubTotalWithDiscount",pdblSubTotalWithDiscount
					'debugprint "pdblShipping",pdblShipping
				End With
				set pclsShipping= Nothing

			Case 2	'Carrier shipping
				'Call LoadOrderItems_SF5(maryOrderItems)
				Call LoadOrderItems_SF5(paryOrderItems)

				Set pclsShipping= New clsShipping
				With pclsShipping

					.Connection = cnn
					.OriginStateAbb = adminOriginState
					.OriginZip = adminOriginZip
					.OriginCountryAbb = adminOriginCountry

					.DestinationStateAbb = pstrState
					.DestinationZIP = pstrZIP
					.DestinationCountryAbb = pstrCountry
					.DestinationCountryName = GetDestinationCountryName(pstrCountry)

					'It is up to you to define/calculate the following
					'Comment out if you do not use
					'.ResidentialDelivery = mblnShipResidential
					'.InsideDelivery = mblnIndoorDeliver

					.OrderItems = maryOrderItems
					.MaxItemWeight = mdblMaxItemWeight
					.TotalOrderWeight = mdblOrderWeight
					.OrderSubtotal = FormatNumber(pdblSubTotalWithDiscount,2,,,false)
					.DeclaredValue = FormatNumber(pdblSubTotalWithDiscount,2,,,false)
					.Insured = False

					.ShippingSelection = Request.Form("ShippingSelection")

					If pstrCountry = "US" Or pstrCountry = "CA" Then
						pblnValidShippingDestination = CBool(Len(pstrZIP) > 0)
					Else
						pblnValidShippingDestination = True
					End If

					If pblnGetShippingRates And pblnValidShippingDestination Then

						If pblnLoadAllShippingMethods Then
							pdblShipping = .GetRates("")
							paryAvailableShippingMethods = .availableRates

							If Len(pstrShipMethodCode) > 0 Then
								If pstrShipMethodCode = .ssShippingMethodCode Then
									'No need to do anything
								ElseIf isArray(paryAvailableShippingMethods) Then
									pstrShipMethodName = ""

									If UBound(paryAvailableShippingMethods) >= 0 Then
										For i = 0 To UBound(paryAvailableShippingMethods)
											If paryAvailableShippingMethods(i)(0) = pstrShipMethodCode Then
												pstrShipMethodName = paryAvailableShippingMethods(i)(1)
												pdblShipping = paryAvailableShippingMethods(i)(2)
												Exit For
											End If
										Next 'i

										'Now use last item if no match
										If Len(pstrShipMethodName) = 0 Then
											pstrShipMethodCode = paryAvailableShippingMethods(UBound(paryAvailableShippingMethods))(0)
											pstrShipMethodName = paryAvailableShippingMethods(UBound(paryAvailableShippingMethods))(1)
											pdblShipping = paryAvailableShippingMethods(UBound(paryAvailableShippingMethods))(2)
										End If
									End If	'UBound(paryAvailableShippingMethods) >= 0
								End If
							End If
						Else
							pdblShipping = .GetRates(pstrShipMethodCode)
							paryAvailableShippingMethods = .availableRates
							pstrShipMethodName = .ssShippingMethodName
						End If
					Else
						pdblShipping = pdblEstimatedShipping
					End If	'pblnGetShippingRates

					'Array(Trim(.ShippingCode), Trim(.ShippingType), .Postage)

					'debugprint "pdblShipping",pdblShipping
					'debugprint "pstrShipMethodCode",pstrShipMethodCode
					'debugprint "pstrShipMethodName",pstrShipMethodName
					If CStr(pdblShipping) = "FAIL" And cblnAutomaticallyFindAnyAvailable Then .GetAnyAvailableShipping pstrShipMethodCode, pstrShipMethodName, pdblShipping
				End With
				Set pclsShipping= Nothing
			Case 3	'Product Based
				If pbytPremiumShipping Then
					pdblShipping = pdblShipping_ProductBased + adminSpcShipAmt
					pstrShipMethodName = "Premium"
				Else
					pdblShipping = pdblShipping_ProductBased
					pstrShipMethodName = "Standard"
				End If
		End Select

		'Backup Method
		If CStr(pdblShipping) = "FAIL" Then
			Select Case 1	'adminShipType2
				Case 1	'Zone Based Shipping
					Set pclsShipping= New clsZoneBasedShipping
					With pclsShipping

						If cblnDebugZBS Then
							Response.Write "Destination Country:" & pstrCountry & "<br />"
							Response.Write "Destination State:" & pstrState & "<br />"
							Response.Write "Destination Zip:" & pstrZIP & "<br />"
							Response.Write "Ship Code:" & pstrShipMethodCode & "<br />"
							Response.Write "Total Purchase:" & pdblSubTotalWithDiscount & "<br />"
							Response.Write "Weight:" & mdblOrderWeight& "<br />"
							Response.Write "Item Count:" & plngOrderItemCount & "<br />"
							Response.Flush
						End If

						.Connection = cnn                        'Set the connection to the StoreFront database
						.ZoneType = 0                            'Set to 0 for Country, 1 for State, 2 for ZIP
						.SetZone(pstrCountry)                    'Set to the shipping Country, State, or ZIP

						'Use this line to go by weight
						'pdblShipping = .GetRate(pstrShipMethodCode, mdblOrderWeight)

						'Use this line to go by order amount
						pdblShipping = .GetRate(pstrShipMethodCode, pdblSubTotalWithDiscount)

						'Use this line to go by order item count
						'pdblShipping = .GetRate(pstrShipMethodCode, plngOrderItemCount)

						'pstrShipMethodName = getNameWithID("sfShipping", pstrShipMethodCode, "shipID", "shipMethod", 0)
						pstrShipMethodName = getNameWithID("ssShippingMethods", pstrShipMethodCode, "ssShippingMethodCode", "ssShippingMethodName", 1)

						'debugprint "pstrShipMethodCode",pstrShipMethodCode
						'debugprint "pstrShipMethodName",pstrShipMethodName
						'debugprint "pdblSubTotalWithDiscount",pdblSubTotalWithDiscount
						'debugprint "pdblShipping",pdblShipping
					End With	'pclsShipping
					set pclsShipping= Nothing

					'Backup Method
					If CStr(pdblShipping) = "FAIL" Then
						If pbytPremiumShipping Then
							pdblShipping = pdblShipping_ProductBased + adminSpcShipAmt
							pstrShipMethodName = "Premium"
						Else
							pdblShipping = pdblShipping_ProductBased
							pstrShipMethodName = "Standard"
						End If
					End If	'CStr(pdblShipping) = "FAIL"

				Case 2	'Carrier shipping
				'	Call ssPostageRate_SF5(pdblShipping, pstrShipMethodCode, adminOriginZip, adminOriginCountry, pstrZIP, pstrCountry, iTotalPur, iPremiumShipping, sTotalPrice)
				Case 3	'Product Based
					If pbytPremiumShipping Then
						pdblShipping = pdblShipping_ProductBased + adminSpcShipAmt
						pstrShipMethodName = "Premium"
					Else
						pdblShipping = pdblShipping_ProductBased
						pstrShipMethodName = "Standard"
					End If
			End Select

		End If	'CStr(pdblShipping) = "FAIL"

		If Len(pstrShipMethodName) = 0 Then pstrShipMethodName = "Est."

	End If	'pblnFreeShipping Or Not pblnOrderIsShipped

End Sub	'CalculateShipping

'***********************************************************************************************

Private Function DeleteOrderDetails(byVal lngOrderID)

Dim pobjCmd

'On Error Resume Next

	If len(lngOrderID) = 0 Then Exit Function

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		Set .ActiveConnection = cnn
		.Commandtype = adCmdText
		.Parameters.Append .CreateParameter("odrdtOrderId", adInteger, adParamInput, 4, lngOrderID)

		'Delete any order attributes
		.Commandtext = "Delete From sfOrderAttributes Where odrattrOrderDetailId In (SELECT sfOrderAttributes.odrattrOrderDetailId FROM sfOrderDetails INNER JOIN sfOrderAttributes ON sfOrderDetails.odrdtOrderId = sfOrderAttributes.odrattrOrderDetailId WHERE sfOrderDetails.odrdtOrderId=?)"
		.Execute , , adExecuteNoRecords

		'Delete any order attributes
		.Commandtext = "Delete From sfOrderDetailsAE Where odrdtAEID In (SELECT sfOrderDetailsAE.odrdtAEID FROM sfOrderDetails INNER JOIN sfOrderDetailsAE ON sfOrderDetails.odrdtOrderId = sfOrderDetailsAE.odrdtAEID WHERE sfOrderDetails.odrdtOrderId=?)"
		If cblnSF5AE Then .Execute , , adExecuteNoRecords

		'Delete individual order details
		.Commandtext = "Delete from sfOrderDetails where odrdtOrderId=" & wrapSQLValue(lngOrderID, False, enDatatype_number)
		.Execute , , adExecuteNoRecords

		'Delete any order attributes
		.Commandtext = "Delete from sfOrdersAE where orderAEID=" & wrapSQLValue(lngOrderID, False, enDatatype_number)
		If cblnSF5AE Then .Execute , , adExecuteNoRecords

	End With	'pobjCmd
	Set pobjCmd = Nothing

	DeleteOrderDetails = CBool(Err.Number = 0)

End Function	'DeleteOrderDetails

'***********************************************************************************************

Public Sub finalizeCart(byVal lngOrderID)

Dim i
Dim j
Dim paryAttribute
Dim paryOrderItem

	plngOrderID = lngOrderID
	If isArray(paryOrderItems) Then
		'Need to remove any order details which may be present from an earlier attempt
		'It is faster to delete and add then to try and match attributes and the ids don't matter here unlike orderID

		For i = 0 To UBound(paryOrderItems)
			paryOrderItem = paryOrderItems(i)
			If isArray(paryOrderItem) Then
				Call setDeleteOrder("odrdttmp", paryOrderItem(enOrderItem_tmpID))	'setDeleteOrder in incGeneral.asp
				If cblnSF5AE Then Call UpdateAvailableQTY(paryOrderItem(enOrderItem_prodID), paryOrderItem(enOrderItem_attributeIDs), paryOrderItem(enOrderItem_odrdttmpQuantity))

			End If
		Next 'i

	End If	'isArray(paryOrderItems)

End Sub	'finalizeCart

'***********************************************************************************************

Public Function hasSavedCart(byVal lngCustID)

Dim pobjCmd
Dim pobjRS

	If Len(CStr(pblnHasSavedCart)) = 0 Then

		pblnHasSavedCart = False

		If Len(lngCustID) > 0 And isNumeric(lngCustID) Then
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = "Select odrdtsvdID From sfSavedOrderDetails Where odrdtsvdCustID=?"
				Set .ActiveConnection = cnn
				.Parameters.Append .CreateParameter("odrdtsvdCustID", adInteger, adParamInput, 4, lngCustID)
				Set pobjRS = .Execute
				If Not pobjRS.EOF Then pblnHasSavedCart = True
				pobjRS.Close
				Set pobjRS = Nothing
			End With	'pobjCmd
			Set pobjCmd = Nothing
		End If
	End If

	hasSavedCart = pblnHasSavedCart

End Function	'hasSavedCart

'***********************************************************************************************

Public Function hasDownloadableItems()

Dim i
Dim paryOrderItem
Dim pblnResult

	pblnResult = False
	If isArray(paryOrderItems) Then
		For i = 0 To UBound(paryOrderItems)
			paryOrderItem = paryOrderItems(i)
			If isArray(paryOrderItem) Then
				If Len(paryOrderItem(enOrderItem_prodFileName)) > 0 Then
					pblnResult = True
					Exit For
				End If
			End If
		Next 'i
	End If	'isArray(paryOrderItems)

	hasDownloadableItems = pblnResult

End Function	'hasDownloadableItems

'***********************************************************************************************

Public Sub saveOrderItemsToOrderDetails(byVal lngOrderID)

Dim i
Dim j
Dim paryAttribute
Dim paryOrderItem

	plngOrderID = lngOrderID
	If isArray(paryOrderItems) Then
		'Need to remove any order details which may be present from an earlier attempt
		'It is faster to delete and add then to try and match attributes and the ids don't matter here unlike orderID
		Call DeleteOrderDetails(lngOrderID)
		For i = 0 To UBound(paryOrderItems)
			paryOrderItem = paryOrderItems(i)
			If isArray(paryOrderItem) Then
				paryOrderItems(i)(enOrderItem_OrderDetailID) = setOrderDetail(lngOrderID, paryOrderItem)
				'Now for the attributes
				If paryOrderItem(enOrderItem_AttributeCount) > 0 Then
					paryAttributes = paryOrderItem(enOrderItem_AttributeArray)

					For j = 0 To paryOrderItem(enOrderItem_AttributeCount) - 1
						Call setOrderAttribute(paryOrderItem(enOrderItem_OrderDetailID), paryAttributes(j))
					Next 'j
				End If	'isArray(paryOrderItemAttributes)

			End If
		Next 'i

	End If	'isArray(paryOrderItems)

End Sub	'saveOrderItemsToOrderDetails

'***********************************************************************************************

Private Function CategoryName(byRef aryOrderDetail)
	If Len(aryOrderDetail(enOrderItem_CategoryName)) = 0 Then aryOrderDetail(enOrderItem_CategoryName) = getNameWithID("sfCategories", aryOrderDetail(enOrderItem_prodCategoryID), "catID", "catName", 0)
	CategoryName = aryOrderDetail(enOrderItem_CategoryName)
End Function	'CategoryName

'***********************************************************************************************

Private Function ManufacturerName(byRef aryOrderDetail)
	If Len(aryOrderDetail(enOrderItem_ManufacturerName)) = 0 Then aryOrderDetail(enOrderItem_ManufacturerName) = getNameWithID("sfManufacturers", aryOrderDetail(enOrderItem_prodManufacturerID), "mfgID", "mfgName", 0)
	ManufacturerName = aryOrderDetail(enOrderItem_ManufacturerName)
End Function	'ManufacturerName

'***********************************************************************************************

Private Function VendorName(byRef aryOrderDetail)
	If Len(aryOrderDetail(enOrderItem_VendorName)) = 0 Then aryOrderDetail(enOrderItem_VendorName) = getNameWithID("sfVendors", aryOrderDetail(enOrderItem_prodVendorID), "vendID", "vendName", 0)
	VendorName = aryOrderDetail(enOrderItem_VendorName)
End Function	'VendorName

'***********************************************************************************************

Private Function setOrderDetail(byRef lngOrderID, byRef aryOrderDetail)
'aryOrderDetail - See enOrderItem

Dim pblnResult
Dim plngodrdtID
Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	If Len(lngOrderID) = 0 Or Not isNumeric(lngOrderID) Or Not isArray(aryOrderDetail) Then
		setOrderDetail = -1
		Exit Function
	End If

	plngodrdtID = aryOrderDetail(enOrderItem_OrderDetailID)
	If Len(plngodrdtID) = 0 Then
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			Set .ActiveConnection = cnn
			.Commandtype = adCmdText
			.Commandtext = "Insert Into sfOrderDetails (odrdtOrderId, odrdtProductID) Values (?,?)"
			'.Commandtype = adCmdStoredProc

			.Parameters.Append .CreateParameter("odrdtOrderId", adInteger, adParamInput, 4, lngOrderID)
			addParameter pobjCmd, "odrdtProductID", adWChar, SessionID, 50, 2
			.Execute , , adExecuteNoRecords

			.Commandtext = "Select odrdtID From sfOrderDetails Where odrdtOrderId=? And odrdtProductID=?"
			Set pobjRS = .Execute
			If Not pobjRS.EOF Then
				plngodrdtID = pobjRS.Fields("odrdtID").Value
				If vDebug = 1 Then Call WriteCommandParameters(pobjCmd, plngodrdtID, "setOrderDetail")
			Else
				plngodrdtID = -1
				setOrderDetail = plngodrdtID
				Exit Function
			End If
			Call closeObj(pobjRS)
		End With	'pobjCmd
		Set pobjCmd = Nothing
	End If	'Len(plngodrdtID) = 0 Then

	If Len(plngodrdtID) > 0 Then
		aryOrderDetail(enOrderItem_OrderDetailID) = plngodrdtID
		pstrSQL = "Update sfOrderDetails Set odrdtQuantity=?, odrdtSubTotal=?, odrdtCategory=?, odrdtManufacturer=?, odrdtVendor=?, odrdtProductName=?, odrdtPrice=?, odrdtProductId=? Where odrdtID=?"
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			Set .ActiveConnection = cnn
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			'.Commandtype = adCmdStoredProc

			.Parameters.Append .CreateParameter("odrdtQuantity", adInteger, adParamInput, 4, aryOrderDetail(enOrderItem_odrdttmpQuantity))
			addParameter pobjCmd, "odrdtSubTotal", adWChar, aryOrderDetail(enOrderItem_odrdttmpQuantity) * aryOrderDetail(enOrderItem_UnitPrice), 50, 2
			addParameter pobjCmd, "odrdtCategory", adWChar, CategoryName(aryOrderDetail), 50, 2
			addParameter pobjCmd, "odrdtManufacturer", adWChar, ManufacturerName(aryOrderDetail), 50, 2
			addParameter pobjCmd, "odrdtVendor", adWChar, VendorName(aryOrderDetail), 50, 2
			addParameter pobjCmd, "odrdtProductName", adWChar, aryOrderDetail(enOrderItem_prodName), 255, 2
			addParameter pobjCmd, "odrdtPrice", adWChar, aryOrderDetail(enOrderItem_UnitPrice), 50, 2
			addParameter pobjCmd, "odrdtProductId", adWChar, aryOrderDetail(enOrderItem_prodID), 50, 2
			.Parameters.Append .CreateParameter("odrdtID", adInteger, adParamInput, 4, plngodrdtID)

			.Execute , , adExecuteNoRecords
			If vDebug = 1 Then Call WriteCommandParameters(pobjCmd, "-", "setOrderDetail-Update sfOrderDetails")

		End With	'pobjCmd
		Set pobjCmd = Nothing

		If cblnSF5AE Then
			pstrSQL = "Insert Into sfOrderDetailsAE (odrdtaeID,odrdtGiftWrapQty,odrdtGiftWrapPrice,odrdtAttDetailID,odrdtBackOrderQty) Values (?,?,?,?,?)"
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				Set .ActiveConnection = cnn
				.Commandtype = adCmdText
				.Commandtext = pstrSQL
				'.Commandtype = adCmdStoredProc

				.Parameters.Append .CreateParameter("odrdtaeID", adInteger, adParamInput, 4, plngodrdtID)
				.Parameters.Append .CreateParameter("odrdtGiftWrapQty", adInteger, adParamInput, 4, aryOrderDetail(enOrderItem_odrdttmpGiftWrapQTY))
				addParameter pobjCmd, "odrdtGiftWrapPrice", adWChar, aryOrderDetail(enOrderItem_odrdttmpGiftWrapQTY) * aryOrderDetail(enOrderItem_gwPrice), 50, 2
				addParameter pobjCmd, "odrdtAttDetailID", adWChar, aryOrderDetail(enOrderItem_attributeIDs), 255, 2
				.Parameters.Append .CreateParameter("odrdtBackOrderQty", adInteger, adParamInput, 4, aryOrderDetail(enOrderItem_odrdttmpBackOrderQTY))

				.Execute , , adExecuteNoRecords
				If vDebug = 1 Then Call WriteCommandParameters(pobjCmd, "-", "setOrderDetail-Insert Into sfOrderDetailsAE")

			End With	'pobjCmd
			Set pobjCmd = Nothing
		End If	'If cblnSF5AE

	End If	'Len(plngodrdtID) > 0

	setOrderDetail = plngodrdtID

End Function	'setOrderDetail

'***********************************************************************************************

Private Function setOrderAttribute(byRef lngodrattrID, byVal aryAttributes)
'aryAttributes - See enAttributeItem

Dim pobjCmd
Dim pblnResult

	pblnResult = False

	'Check for invalid entries
	If Len(lngodrattrID) > 0 And Not isNumeric(lngodrattrID) Then
		setOrderAttribute = pblnResult
		Exit Function
	End If

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		Set .ActiveConnection = cnn
		.Commandtype = adCmdText
		.Commandtext = "Insert Into sfOrderAttributes (odrattrOrderDetailId,odrattrAttribute,odrattrName,odrattrPrice,odrattrType) Values (?,?,?,?,?)"
		'.Commandtype = adCmdStoredProc

		.Parameters.Append .CreateParameter("odrattrOrderDetailId", adInteger, adParamInput, 4, lngodrattrID)
		addParameter pobjCmd, "odrattrAttribute", adWChar, aryAttributes(enAttributeItem_attrdtName), 255, 2
		addParameter pobjCmd, "odrattrName", adLongVarWChar, aryAttributes(enAttributeItem_attrName), 2147483646, 2
		addParameter pobjCmd, "odrattrPrice", adWChar, aryAttributes(enAttributeItem_attrdtPrice), 50, 2
		.Parameters.Append .CreateParameter("odrattrType", adInteger, adParamInput, 4, aryAttributes(enAttributeItem_attrdtType))
		.Execute , , adExecuteNoRecords

		pblnResult = True

	End With	'pobjCmd
	Set pobjCmd = Nothing

	setOrderAttribute = pblnResult

End Function	'setOrderAttribute

'***********************************************************************************************

Public Sub writeDebugCart

Dim i
Dim j
Dim paryAttribute
Dim paryOrderItem

	If isArray(paryOrderItems) Then
		Response.Write "<table border=1 cellpadding=2 cellspacing=0>"
		Response.Write "<tr>"
		Response.Write "<th colspan=10>writeDebugCart - cart contents for cartID <b>" & SessionID & "</b></th>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<th>Product</th>"
		Response.Write "<th>Unit Price</th>"
		Response.Write "<th>Qty</th>"
		Response.Write "<th>BO Qty</th>"
		Response.Write "<th>GW Qty</th>"
		Response.Write "<th>Product Qty</th>"
		Response.Write "<th>MTP Discount</th>"
		Response.Write "<th>GW Price</th>"
		Response.Write "<th>Inventory Tracked</th>"
		Response.Write "<th>Qty On Hand</th>"
		Response.Write "<th>Weight</th>"
		Response.Write "</tr>"
		For i = 0 To UBound(paryOrderItems)
			paryOrderItem = paryOrderItems(i)
			If isArray(paryOrderItem) Then
				Response.Write "<tr>"
				Response.Write "<td>"
				Response.Write paryOrderItem(enOrderItem_prodName) & "<br />"

				'Now for the attributes
				If paryOrderItem(enOrderItem_AttributeCount) > 0 Then
					paryAttributes = paryOrderItem(enOrderItem_AttributeArray)
					For j = 0 To paryOrderItem(enOrderItem_AttributeCount) - 1
						Response.Write "--" & paryAttributes(j)(enAttributeItem_attrName) & ": " & paryAttributes(j)(enAttributeItem_attrdtName) & "<br />"
					Next 'j
				Else
					Response.Write "-- No Attributes<br />"
				End If	'isArray(paryOrderItemAttributes)

				Response.Write "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_UnitPrice) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_odrdttmpQuantity) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_QuantityByProductID) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_MTPDiscount) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_gwPrice) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_invenbTracked) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_invenInStock) & "</td>"
				Response.Write "<td>" & paryOrderItem(enOrderItem_prodWeight) & "</td>"
				Response.Write "</tr>"
			End If
		Next 'i

		'Now for the cart summary
		Response.Write "<tr>"
		Response.Write "<td colspan=11 align=right>"

		Response.Write "<table border=1 cellpadding=2 cellspacing=0>"
		Response.Write "<colgroup align=left /><colgroup align=right />"
		Response.Write "<tr><td>pdblSubTotal: </td><td>" & pdblSubTotal & "</td></tr>"
		Response.Write "<tr><td>pdblDiscount: </td><td>" & pdblDiscount & "</td></tr>"
		Response.Write "<tr><td>-- pdblDiscount_FirstTimeCustomer: </td><td>" & pdblDiscount_FirstTimeCustomer & "</td></tr>"
		Response.Write "<tr><td>-- pdblSubTotal_Shipped: </td><td>" & pdblSubTotal_Shipped & "</td></tr>"
		Response.Write "<tr><td colspan=2>&nbsp;&nbsp;-Promotion Codes: " & vistorDiscountCodes & "</td></tr>"
		Response.Write "<tr><td>pdblSubTotalWithDiscount: </td><td>" & pdblSubTotalWithDiscount & "</td></tr>"
		Response.Write "<tr><td>pdblHandling: </td><td>" & pdblHandling & "</td></tr>"
		Response.Write "<tr><td>pdblHandling_Order: </td><td>" & pdblHandling_Order & "</td></tr>"
		Response.Write "<tr><td>pdblHandling_ProductSetup: </td><td>" & pdblHandling_ProductSetup & "</td></tr>"
		Response.Write "<tr><td>pdblHandling_ProductSetupOneTime: </td><td>" & pdblHandling_ProductSetupOneTime & "</td></tr>"
		Response.Write "<tr><td>pdblHandling_ProductSpecific: </td><td>" & pdblHandling_ProductSpecific & "</td></tr>"
		Response.Write "<tr><td>pdblCOD: </td><td>" & pdblCOD & "</td></tr>"
		Response.Write "<tr><td>pdblShipping: </td><td>" & pdblShipping & "</td></tr>"
		Response.Write "<tr><td colspan=2>&nbsp;&nbsp;-ShipMethodCode: " & pstrShipMethodCode & "</td></tr>"
		Response.Write "<tr><td colspan=2>&nbsp;&nbsp;-ShipMethodName: " & pstrShipMethodName & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
		Response.Write "<tr><td colspan=2>&nbsp;&nbsp;-Free Shipping: " & pblnFreeShipping & "</td></tr>"
		Response.Write "<tr><td colspan=2>&nbsp;&nbsp;-State: " & pstrState & "</td></tr>"
		Response.Write "<tr><td colspan=2>&nbsp;&nbsp;-ZIP: " & pstrZIP & "</td></tr>"
		Response.Write "<tr><td colspan=2>&nbsp;&nbsp;-Country: " & pstrCountry & "</td></tr>"
		Response.Write "<tr><td>pdblLocalTax: </td><td>" & pdblLocalTax & "</td></tr>"
		Response.Write "<tr><td>pdblStateTax: </td><td>" & pdblStateTax & "</td></tr>"
		Response.Write "<tr><td>pdblCountryTax: </td><td>" & pdblCountryTax & "</td></tr>"
		Response.Write "<tr><td>pdblCartTotal: </td><td>" & pdblCartTotal & "</td></tr>"
		Response.Write "<tr><td>pdblStoreCredit: </td><td>" & pdblStoreCredit & "</td></tr>"
		Response.Write "<tr><td>pdblAvailableStoreCredit: </td><td>" & pdblAvailableStoreCredit & "</td></tr>"
		Response.Write "<tr><td colspan=2>&nbsp;&nbsp;-GC Codes: " & visitorCertificateCodes & "</td></tr>"
		Response.Write "<tr><td>pdblAmountDue: </td><td>" & pdblAmountDue & "</td></tr>"
		Response.Write "</table>"

		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table>"
	End If

End Sub	'writeDebugCart

'***********************************************************************************************

Public Sub writeDebugOrderItem(byRef aryOrderItem)

Dim i
Dim j
Dim paryAttribute
Dim paryOrderItem

	If isArray(aryOrderItem) Then
		Response.Write "<table border=1 cellpadding=2 cellspacing=0>"
		Response.Write "<tr>"
		Response.Write "<th colspan=10>writeDebugOrderItem</th>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<th>Product</th>"
		Response.Write "<th>Unit Price</th>"
		Response.Write "<th>Base Price</th>"
		Response.Write "<th>Qty</th>"
		Response.Write "<th>BO Qty</th>"
		Response.Write "<th>GW Qty</th>"
		Response.Write "<th>Product Qty</th>"
		Response.Write "<th>MTP Discount</th>"
		Response.Write "<th>GW Price</th>"
		Response.Write "<th>Inventory Tracked</th>"
		Response.Write "<th>Qty On Hand</th>"
		Response.Write "</tr>"

		paryOrderItem = aryOrderItem
		If isArray(paryOrderItem) Then
			Response.Write "<tr>"
			Response.Write "<td>"
			Response.Write paryOrderItem(enOrderItem_prodName) & "<br />"

			'Now for the attributes
			Response.Write "--Attributes: " & paryOrderItem(enOrderItem_AttributeCount) & "<br />"
			Response.Write "--isArray: " & isArray(paryOrderItem(enOrderItem_AttributeArray)) & "<br />"
			If paryOrderItem(enOrderItem_AttributeCount) > 0 Then
				paryAttributes = paryOrderItem(enOrderItem_AttributeArray)
				For j = 0 To paryOrderItem(enOrderItem_AttributeCount) - 1
					Response.Write "--" & paryAttributes(j)(enAttributeItem_attrName) & ": " & paryAttributes(j)(enAttributeItem_attrdtName) & "<br />"
				Next 'j
			Else
				Response.Write "-- No Attributes<br />"
			End If	'isArray(paryOrderItemAttributes)

			Response.Write "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_UnitPrice) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_prodBasePrice) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_odrdttmpQuantity) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_QuantityByProductID) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_MTPDiscount) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_gwPrice) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_invenbTracked) & "</td>"
			Response.Write "<td>" & paryOrderItem(enOrderItem_invenInStock) & "</td>"
			Response.Write "</tr>"
		End If

		Response.Write "</table>"
	End If

End Sub	'writeDebugOrderItem

'***********************************************************************************************

Private Sub CreateCartHTML

Dim i
Dim pstrHTML
Dim pstrProduct
Dim pstrProdLink

	If Not isObject(pobjRS) Then Call LoadCartContents

	If Not pblnEmptyCart Then

		'create the full cart display
		pstrHTML = "<table id='tblFullCart' border=0 cellspacing=0 cellpadding=1>" & vbcrlf _
					& "<tr>" & vbcrlf _
					& "  <th>product</th>" & vbcrlf _
					& "  <th align=right>unit&nbsp;price</th>" & vbcrlf _
					& "  <th align=right>qty</th>" & vbcrlf _
					& "  <th align=right>price</th>" & vbcrlf _
					& "</tr>" & vbcrlf

		For i = 1 To UBound(paryOrderItems)

			'build the link
			If Len(paryOrderItems(i)(6)) > 0 Then
				pstrProdLink = "<a href=" & Chr(34) & Server.HTMLEncode(paryOrderItems(i)(6)) & Chr(34) & ">" & Replace(paryOrderItems(i)(1)," ","&nbsp;") & "</a>"
			Else
				pstrProdLink = Replace(paryOrderItems(i)(1)," ","&nbsp;")
			End If

			'build the attribute
			If Len(paryOrderItems(i)(2)) > 0 Then
				pstrProduct = pstrProdLink & vbcrlf & "<br />&nbsp;&nbsp;" & Replace(paryOrderItems(i)(2),cstrContentsDelimeter,"<br />" & vbcrlf & "&nbsp;&nbsp;")
			Else
				pstrProduct = pstrProdLink
			End If
			pstrHTML = pstrHTML _
					& "<tr>" & vbcrlf _
					& "  <td>" & pstrProduct & "</td>" & vbcrlf _
					& "  <td align=right>" & FormatCurrency(paryOrderItems(i)(3)) & "</td>" & vbcrlf _
					& "  <td align=right>" & paryOrderItems(i)(4) & "</td>" & vbcrlf _
					& "  <td align=right>" & FormatCurrency(paryOrderItems(i)(5)) & "</td>" & vbcrlf _
					& "</tr>" & vbcrlf
		Next 'i
		pstrHTML = pstrHTML _
					& "<tr>" & vbcrlf _
					& "  <td colspan=4><hr></td>" & vbcrlf _
					& "</tr>" & vbcrlf _
					& "<tr>" & vbcrlf _
					& "  <td colspan=2>&nbsp;</td>" & vbcrlf _
					& "  <td colspan=2 align=right>Sub&nbsp;Total:&nbsp;" & FormatCurrency(paryOrderItems(0)(2)) & "</td>" & vbcrlf _
					& "</tr>" & vbcrlf _
					& "<tr>" & vbcrlf _
					& "  <td colspan=4 align=center><A class='' href='order.asp' title='View your cart and checkout'>" & pstrViewCartText & "</A></td>" & vbcrlf _
					& "</tr>" & vbcrlf

		'pstrHTML = pstrHTML _
		'		 & "<tr><td colspan=4 align=center><A href=""Shipping Calculator"" onclick=""var mScreenHeight = window.screen.availHeight; var mScreenWidth = window.screen.availWidth; window.open('ShippingCalculator.asp','SearchResults','toolbar=0,location=0,directories=0,status=1,menubar=No,copyhistory=0,scrollbars=1,height=' + mScreenHeight/1.5 + ',width=' + mScreenWidth/2 + ',screenY=' + mScreenHeight/1.5 + ',screenX=' + mScreenWidth/2 + ',top=0,left=' + mScreenWidth/2.05 + ',resizable'); return false;"">Estimate Shipping</A></td></tr>" & vbcrlf

		pstrHTML = pstrHTML _
					& "</table>" & vbcrlf

		pstrFullCart = pstrHTML

		'Response.Write "Cart: paryOrderItems(0)(0) = " & paryOrderItems(0)(0) & "<br />"
		Select Case paryOrderItems(0)(0)
			Case 1:	pstrProduct = "One item"
			Case 2:	pstrProduct = "Two items"
			Case 3:	pstrProduct = "Three items"
			Case 4:	pstrProduct = "Four items"
			Case 5:	pstrProduct = "Five items"
			Case 6:	pstrProduct = "Six items"
			Case 7:	pstrProduct = "Seven items"
			Case 8:	pstrProduct = "Eight items"
			Case 9:	pstrProduct = "Nine items"
			Case Else: pstrProduct = paryOrderItems(0)(0) & " items in cart"
		End Select

		'create the mini cart display
		pstrHTML = "<table id='tblMiniCart' border=0 cellspacing=0 cellpadding=1>" & vbcrlf _
					& "<tr><td align=center>" & pstrProduct & " in cart</td></tr>" & vbcrlf _
					& "<tr><td align=center>Sub&nbsp;Total:&nbsp;" & FormatCurrency(paryOrderItems(0)(2)) & "</td></tr>" & vbcrlf _
					& "<tr><td align=center><A class='' href='order.asp' title='View your cart and checkout'>" & pstrViewCartText & "</A></td></tr>" & vbcrlf

		'pstrHTML = pstrHTML _
		'		 & "<tr><td colspan=4 align=center><A href=""Shipping Calculator"" onclick=""var mScreenHeight = window.screen.availHeight; var mScreenWidth = window.screen.availWidth; window.open('ShippingCalculator.asp','SearchResults','toolbar=0,location=0,directories=0,status=1,menubar=No,copyhistory=0,scrollbars=1,height=' + mScreenHeight/1.5 + ',width=' + mScreenWidth/2 + ',screenY=' + mScreenHeight/1.5 + ',screenX=' + mScreenWidth/2 + ',top=0,left=' + mScreenWidth/2.05 + ',resizable'); return false;"">Estimate Shipping</A></td></tr>" & vbcrlf

		pstrHTML = pstrHTML _
					& "</table>" & vbcrlf

		pstrMiniCart = pstrHTML

	End If

End Sub	'CreateCartHTML

'***********************************************************************************************

Public Sub DisplayMiniCartContents

	If Not pblnCartCreated Then Call CreateCartHTML
	Response.Write pstrMiniCart

End Sub	'DisplayMiniCartContents

'***********************************************************************************************

Public Sub DisplayFullCartContents

	If Not pblnCartCreated Then Call CreateCartHTML
	Response.Write pstrFullCart

End Sub	'DisplayFullCartContents

'***********************************************************************************************

	Public Sub displayVisitorShippingPreferences
	%>
	<!--#include file="ssclsCartContents_visitorPreferences.asp"-->
	<%
	End Sub	'displayVisitorShippingPreferences

	'***********************************************************************************************

	Public Sub displayOrderSummaryCompact
	%>
	<!--#include file="ssclsCartContents_orderSummary_Compact.asp"-->
	<%
	End Sub	'displayOrderSummaryCompact

	'***********************************************************************************************

	Public Sub displayOrderSummary
	%>
	<!--#include file="ssclsCartContents_orderSummary.asp"-->
	<%
	End Sub	'displayOrderSummary

	'***********************************************************************************************

	Public Sub displayOrder
	%>
	<!--#include file="ssclsCartContents_OrderView.asp"-->
	<%
	End Sub	'displayOrder

	'***********************************************************************************************

	Public Sub displayOrder_CheckoutView
	%>
	<!--#include file="ssclsCartContents_CheckoutView.asp"-->
	<%
	End Sub	'displayOrder_CheckoutView

	'***********************************************************************************************

	Public Sub displayOrder_CartSummaryShort

	Response.Write "Cart Summary Short"

	End Sub	'displayOrder_CheckoutView

	'***********************************************************************************************

	Public Sub displayOrder_CartSummaryDetailed

	Response.Write "Cart Summary Detailed"

	End Sub	'displayOrder_CheckoutView

	'***********************************************************************************************

	Public Sub WriteGoogleAnalyticsEommerceTrackingItems()
	'Format for output
	'UTM:I|[order-id]|[sku/code]|[productname]|[category]|[price]|[quantity]

	Dim paryOrderItem

		For i = 0 To plngUniqueOrderItemCount
			paryOrderItem = paryOrderItems(i)
			Response.Write "UTM:I|" & plngOrderID & "|" & paryOrderItem(enOrderItem_prodID) & "|" & paryOrderItem(enOrderItem_prodName) & "|" & paryOrderItem(enOrderItem_CategoryName) & "|" & paryOrderItem(enOrderItem_UnitPrice) & "|" & paryOrderItem(enOrderItem_odrdttmpQuantity) & vbcrlf
		Next 'i

	Response.Write "Cart Summary Detailed"

	End Sub	'WriteGoogleAnalyticsEommerceTrackingItems

End Class	'clsCartTotal

'***********************************************************************************************
'***********************************************************************************************

%>
