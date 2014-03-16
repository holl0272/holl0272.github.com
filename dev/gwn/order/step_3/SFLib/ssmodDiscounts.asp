<%
'********************************************************************************
'*   Promotion Manager for StoreFront 5.0										*
'*   Release Version:	2.01.004 												*
'*   Release Date:		August 10, 2003											*
'*   Revision Date:		May 2, 2004												*
'*																				*
'*   Release 2.01.004 (May 2, 2004)												*
'*	   - Added section for TaxRate Manager support								*
'*																				*
'*   Release 2.01.003 (January 18, 2004)										*
'*	   - Added additional remote debugging code									*
'*	   - Bug Fix - Buy1Get1 Free Gift bug introduced in 2.01.002				*
'*																				*
'*   Release 2.01.002 (January 14, 2004)										*
'*	   - Added additional remote debugging code									*
'*	   - Bug Fix - updated Free Gift, minimum order amount to correctly exclude	*
'*				   or include only specified products							*
'*																				*
'*   Release 2.01.001 (December 14, 2003)										*
'*	   - Added additional remote debugging code									*
'*	   - Bug Fix - updated free shipping bug introduced in 2.00.004				*
'*	   - Enhancement - added club product discount - not supported				*
'*																				*
'*   Release 2.00.004 (October 18, 2003)										*
'*	   - Added additional remote debugging code									*
'*	   - Bug Fix - free product display issue									*
'*																				*
'*   Release 2.00.003 (September 29, 2003)										*
'*	   - Updated promotion registration form/link interface to help with display*
'*	   - Bug Fix - updated code to account for trailing spaces in product codes	*
'*																				*
'*   Release 2.00.002 (September 5, 2003)										*
'*	   - Added DHTML to submit buttons to indicate they've been clicked			*
'*	   - Optimized code for use with large numbers of active promotions			*
'*	   - Merged all promotion registration functions to this file for security	*
'*	   - Implemented security routine to protect against SQL Injection attacks	*
'*	   - Bug Fix - Updated promotion usage recording to database				*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.				*
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	Const mblnCaseInsensitive = True						'Set to true to make promotion codes case insensitive
	Const cbytFreeGiftOfferLocation = 1						'Set to 0 to show with discount text; Set to 1 to show just above sub total text
	Const cblnAutoCheckFreeGift = True						'Set True to automatically select free gift on order.asp; false to require customer to select
	Const cblnShowRegistrationBoxOnOrderSummary = False		'Set to true to show registration box on order.asp; False not to
	Const cblnShowRegistrationLinkOnOrderSummary = True		'Set to true to show link to registration pop-box on order.asp; False not to
	Const cblnShowDiscountDetailsOnOrderSummary = True
	
	Const cstrDiscountAppliedMessage = "Discount Applied"
	
	'SPECIAL Settings
	
	'Promotion Registration Page Settings
	Const mblnRedirectToProduct = True		'Set to true to redirect to 1st product in product list
	'Const cstrRegistrationRedirectPage = "search.asp"		'Set to page you want the customer to be redirected to (if selected above)
	Const cstrRegistrationRedirectPage = ""		'Set to page you want the customer to be redirected to (if selected above)

'/
'/////////////////////////////////////////////////

Dim mlngNumClubCertificatesToCreate:	mlngNumClubCertificatesToCreate = -1
Dim maryDiscountClubsToCreate()
Dim maryDiscountClubProducts	'will hold an array of (prodID, attrDetailID, promoCode, expDate)

'**********************************************************
'	Developer notes
'**********************************************************

'Key to maryDiscountSummary
'0 - code
'1 - discount
'2 - combineable
'3 - sTaxable
'4 - cTaxable
'5 - use
'6 - Title
'7 - offerFreeGiftAutomatically
'8 - FreeProductID
'9 - MinSubTotal
'10 - FreeShipping Code | FreeShipping Limit

'**********************************************************
'*	Page Level variables
'**********************************************************

Const enDiscountUse = 5
Const enPromoTitle = 6
Const enofferFreeGiftAutomatically = 7
Const enDiscountFreeProduct = 8
Const enDiscountFreeShipping = 10

Dim cblnDebugPromotionManagerAddon
Dim mblnAlreadyCalculated		'This is here because SF will perform repeated queries to the database for the same information
Dim mstrDiscountText
Dim mcurDiscountAmount
Dim maryDiscountSummary
Dim mstrDiscountMessage
Dim mstrPromotionRegistrationMessage
Dim mblnSuccessfulRegistration
Dim mstrPromotedProducts
Dim mstrPromotionCode
Dim mblnAlreadyDisplayed

Dim msngSTaxRate, mcurSTax
Dim msngCTaxRate, mcurCTax
Dim mblnTaxShipIsActive

'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'**********************************************************

mblnAlreadyCalculated = False
mblnAlreadyDisplayed = False
mblnSuccessfulRegistration = False

cblnDebugPromotionManagerAddon = CBool(Len(Session("ssDebug_PromotionManager")) > 0)
'cblnDebugPromotionManagerAddon = True

Call CheckForPromotionRegistration

'**********************************************************
'**********************************************************

Class clsPromotion

Dim cstrDelimiter
Dim cstrSSTextBasedAttributeDelimiter

'working variables
dim pConnection
dim prsOrder
dim prsPromo

Dim paryOrder
Dim enOrder_ID
Dim enOrder_ProductID
Dim enOrder_Quantity
Dim enOrder_SellPrice
Dim enOrder_SaleIsActive
Dim enOrder_StateTaxIsActive
Dim enOrder_CountryTaxIsActive
Dim enOrder_DiscountAmount
Dim enOrder_DiscountRunningTotal
Dim enOrder_BasePrice
Dim enOrder_ArrayLength
'key to paryOrder
' 0 - odrdttmpID
' 1 - odrdttmpProductID
' 2 - odrdttmpQuantity
' 3 - prodSellPrice - includes attributes/sale price if applicable
' 4 - prodSaleIsActive
' 5 - prodStateTaxIsActive
' 6 - prodCountryTaxIsActive
' 7 - Discount Amount (array corresponding to promotions)
' 8 - Running Total For Discounts - Used so that multiple discounts can't exceed cost of item

dim pblnComplete

'variable for  handling Order parameters
Dim pstrPromotions
Dim paryPromotions

Dim pcurLocalSubTotal
dim pcursubTotal, pcursubTotalMinusSaleItems
dim pcursubTotalSTaxable
dim pcursubTotalCTaxable
Dim pBestDiscountAmount
Dim pstrExclusiveCode
Private plngCustID

Private Sub class_Terminate()

On Error Resume Next

	prsPromo.Close
	set prsPromo = nothing
	If Err.number <> 0 Then Err.Clear

End Sub

Private Sub class_Initialize()

	cstrDelimiter = ";"

	pblnComplete = False
	plngCustID = -1
	
	'Initialize enumerations
	enOrder_ID = 0
	enOrder_ProductID = 1
	enOrder_Quantity = 2
	enOrder_SellPrice = 3
	enOrder_SaleIsActive = 4
	enOrder_StateTaxIsActive = 5
	enOrder_CountryTaxIsActive = 6
	enOrder_DiscountAmount = 7
	enOrder_DiscountRunningTotal = 8
	enOrder_BasePrice = 9
	enOrder_ArrayLength = 9
	
End Sub	'class_Initialize

'***********************************************************************************************

Public Property Let Connection(objConnection)
	Set pConnection = objConnection
End Property

Public Property Let CustID(byVal Value)
	If isNumeric(Value) And Len(Value) > 0 Then plngCustID = Value
End Property

Public Property Let Promotions(strPromotions)
	pstrPromotions = strPromotions
	Call safePromotionCodes
End Property

Public Property Let SubTotal(curSubTotal)
	pcursubTotal = CDbl(curSubTotal)
End Property

Public Property Get SubTotal()
	subTotal = pcurSubTotal
End Property

Public Property Get subTotalMinusSaleItems()
	subTotalMinusSaleItems = pcursubTotalMinusSaleItems
End Property

Public Property Get subTotalSTaxable()
	subTotalSTaxable = pcursubTotalSTaxable
End Property

Public Property Get subTotalCTaxable()
	subTotalCTaxable = pcursubTotalCTaxable
End Property

Public Property Get ExclusiveCode()
	ExclusiveCode = pstrExclusiveCode
End Property

'***********************************************************************************************

Private Function WrapString(strToWrap,blnWrap)

dim strTemp

	if len(strToWrap) = 0 then Exit Function
	strTemp = strToWrap
	if blnWrap then
		if left(strTemp,1) <> cstrDelimiter then strTemp = cstrDelimiter & strTemp
		if right(strTemp,1) <> cstrDelimiter then strTemp = strTemp & cstrDelimiter
	else
		if left(strTemp,1) = cstrDelimiter then strTemp = right(strTemp,len(strTemp)-1)
		if right(strTemp,1) = cstrDelimiter then strTemp = left(strTemp,len(strTemp)-1)
	end if

	WrapString = strTemp
	
End Function	'WrapString

'***********************************************************************************************

Private Sub safePromotionCodes()

Dim pstrPromoCode
Dim i
Dim char

	paryPromotions = Split(pstrPromotions, cstrDelimiter)
	For i = 0 to UBound(paryPromotions)
		If Not isPromotionCodeSafe(Trim(paryPromotions(i))) Then paryPromotions(i) = ""
	Next 'i
	
End Sub 'safePromotionCodes

'***********************************************************************************************

Private Function IsAvailableForUseByThisCustomer(byVal strPromoCode)

Dim pblnResult
Dim pobjCmd
Dim pobjRS

	If plngCustID = -1 Or Len(strPromoCode) = 0 Then
		pblnResult = True
	Else
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "SELECT Count(sfOrders.orderID) AS CountOforderID, NumUsesByCustomer" _
						 & " FROM (sfOrders INNER JOIN ordersDiscounts ON sfOrders.orderID = ordersDiscounts.OrderID) INNER JOIN Promotions ON ordersDiscounts.PromotionID = Promotions.PromotionID" _
						 & " GROUP BY sfOrders.orderCustId, Promotions.PromoCode, Promotions.NumUsesByCustomer" _
						 & " HAVING sfOrders.orderCustId=? AND Promotions.PromoCode=?"
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, plngCustID)
			.Parameters.Append .CreateParameter("PromoCode", adVarChar, adParamInput, Len(strPromoCode), strPromoCode)
			Set pobjRS = .Execute
			If pobjRS.EOF Then
				pblnResult = False
			Else
				If isNumeric(pobjRS.Fields("NumUsesByCustomer").Value) And Len(pobjRS.Fields("NumUsesByCustomer").Value & "") > 0 Then
					pblnResult = pobjRS.Fields("NumUsesByCustomer").Value > pobjRS.Fields("CountOforderID").Value
				Else
					pblnResult = True
				End If
			End If	'pobjRS.EOF
			pobjRS.Close
			Set pobjRS = Nothing

		End With	'pobjCmd
		Set pobjCmd = Nothing
	End If

	IsAvailableForUseByThisCustomer = pblnResult

End Function	'IsAvailableForUseByThisCustomer

'***********************************************************************************************

Private Function LoadPromotions

Dim pstrSQL_Access
Dim pstrSQL_SQLServer
Dim pstrSQL
Dim i
Dim pstrSQLWhere

	If isArray(paryPromotions) Then
		For i = 0 To UBound(paryPromotions)
			If Len(paryPromotions(i)) > 0 Then
				If Len(paryPromotions(i)) > 0 Then
					pstrSQLWhere = pstrSQLWhere & " OR PromoCode = '" & paryPromotions(i) & "'"
				End If
			End If
		Next 'i
	End If

	pstrSQL_SQLServer = "SELECT * " _
						& "FROM Promotions " _
						& "WHERE " _
						& "(" _
						& "  (MaxUses IS NULL) AND " _
						& "  (StartDate <= GETDATE() OR StartDate IS NULL) AND " _
						& "  (EndDate >= GETDATE() OR EndDate IS NULL) AND " _
						& "  (Inactive = 0) AND " _
						& "  (Duration IS NULL OR ((Promotions.Duration)>=Convert(int,[StartDate]-'" & Date() & "')))" _
						& "  " _
						& " OR" _
						& "  " _
						& "  (StartDate <= GETDATE() OR StartDate IS NULL) AND " _
						& "  (EndDate >= GETDATE() OR EndDate IS NULL) AND " _
						& "  (Inactive = 0) AND " _
						& "  (Duration IS NULL OR ((Promotions.Duration)>=Convert(int,[StartDate]-'" & Date() & "'))) AND " _
						& "  (NumUses IS NULL)" _
						& "  " _
						& " OR" _
						& "" _
						& "  (MaxUses > NumUses) AND " _
						& "  (StartDate <= GETDATE() OR StartDate IS NULL) AND " _
						& "  (EndDate >= GETDATE() OR EndDate IS NULL) AND " _
						& "  (Inactive = 0) AND " _
						& "  (Duration IS NULL OR ((Promotions.Duration)>=Convert(int,[StartDate]-'" & Date() & "')))" _
						& ")" _
						& "" _
						& "AND" _
						& "" _
						& "  (ApplyAutomatically=1 " & pstrSQLWhere & ")"

	pstrSQL_Access = "SELECT * " _
					& "FROM Promotions " _
					& "WHERE " _
					& "(" _
					& "  (MaxUses IS NULL) AND " _
					& "  (StartDate <= Date() OR StartDate IS NULL) AND " _
					& "  (EndDate >= Date() OR EndDate IS NULL) AND " _
					& "  (Inactive = 0) AND " _
					& "  (Duration IS NULL OR Duration >= StartDate - Date())" _
					& "  " _
					& " OR" _
					& "  " _
					& "  (StartDate <= Date() OR StartDate IS NULL) AND " _
					& "  (EndDate >= Date() OR EndDate IS NULL) AND " _
					& "  (Inactive = 0) AND " _
					& "  (Duration IS NULL OR Duration >= StartDate - Date()) AND " _
					& "  (NumUses IS NULL)" _
					& "  " _
					& " OR" _
					& "" _
					& "  (MaxUses > NumUses) AND " _
					& "  (StartDate <= Date() OR StartDate IS NULL) AND " _
					& "  (EndDate >= Date() OR EndDate IS NULL) AND " _
					& "  (Inactive = 0) AND " _
					& "  (Duration IS NULL OR Duration >= StartDate - Date())" _
					& ")" _
					& "" _
					& "AND" _
					& "" _
					& "  (ApplyAutomatically=-1 " & pstrSQLWhere & ")"

	If cblnSQLDatabase Then
		pstrSQL = pstrSQL_SQLServer
	Else
		pstrSQL = pstrSQL_Access
	End If
	
	On Error Resume Next
	If Err.number <> 0 Then Err.Clear
	
	set prsPromo = CreateObject("adodb.recordset")
	with prsPromo
		.ActiveConnection = pConnection
		.CursorLocation = 2 'adUseClient
		.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		'debugprint "pstrSQL", pstrSQL
		If Err.number <> 0 Then
			If InStr(1, LCase(Err.Description), "getdate") > 0 Then
				Err.Clear
				If .State = 1 Then .Close
				.Open pstrSQL_Access, cnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					Application("AppDatabase") = "Access"
					If cblnDebugPromotionManagerAddon Then Response.Write("<hr><font color=black>Promotions RecordCount: " & prsPromo.RecordCount & "</font><br />")
					LoadPromotions = (.RecordCount > 0)
					Exit Function
				End If
			End If
			
			If cblnDebugPromotionManagerAddon Then
				Response.Write("<hr><font color=black>Error " & Err.number & ": " & Err.Description & "</font><br />")
				Response.Write("<hr><font color=black>LoadPromotions sql: " & sql & "</font><br />")
			End If
			LoadPromotions = False
		Else
			If cblnDebugPromotionManagerAddon Then Response.Write("<hr><font color=black>Promotions RecordCount: " & prsPromo.RecordCount & "</font><br />")
			LoadPromotions = (prsPromo.RecordCount > 0)
		End If
	end with

End Function	'LoadPromotions

'***********************************************************************************************

Public Function SetOrderItems(byRef aryOrderItems)

Dim i, j
Dim pdblProductExtendedAmount
Dim pdblProductExtendedAmountBO
Dim pdblSubTotal
Dim pdblSubTotal_LocalTaxable
Dim pdblSubTotal_StateTaxable
Dim pdblSubTotal_CountryTaxable
Dim pdblSubTotal_StateTaxable_Special
Dim plngItemCount
Dim plngOrderItemCount
Dim plngUniqueOrderItemCount

	'Initialize the variables
	pdblSubTotal = 0
	pdblSubTotal_LocalTaxable = 0
	pdblSubTotal_StateTaxable = 0
	pdblSubTotal_CountryTaxable = 0
	pdblSubTotal_StateTaxable_Special = 0

	'Now calculate the order values
	plngUniqueOrderItemCount = UBound(aryOrderItems)
	For i = 0 To plngUniqueOrderItemCount
		plngOrderItemCount = plngOrderItemCount + aryOrderItems(i)(enOrderItem_odrdttmpQuantity)
		
		pdblProductExtendedAmount = aryOrderItems(i)(enOrderItem_UnitPrice) * aryOrderItems(i)(enOrderItem_odrdttmpQuantity) _
									+ aryOrderItems(i)(enOrderItem_gwPrice) * aryOrderItems(i)(enOrderItem_odrdttmpGiftWrapQTY)
		pdblProductExtendedAmountBO = aryOrderItems(i)(enOrderItem_UnitPrice) * aryOrderItems(i)(enOrderItem_odrdttmpBackOrderQTY)
					
		'Now add in the gift wrap for BO items
		If (aryOrderItems(i)(enOrderItem_odrdttmpQuantity) - aryOrderItems(i)(enOrderItem_odrdttmpBackOrderQTY)) > aryOrderItems(i)(enOrderItem_odrdttmpGiftWrapQTY) Then
			pdblProductExtendedAmountBO = pdblProductExtendedAmountBO + aryOrderItems(i)(enOrderItem_gwPrice) * (aryOrderItems(i)(enOrderItem_odrdttmpGiftWrapQTY) - aryOrderItems(i)(enOrderItem_odrdttmpQuantity) + aryOrderItems(i)(enOrderItem_odrdttmpBackOrderQTY))
		End If
		
		If aryOrderItems(i)(enOrderItem_prodStateTaxIsActive) Then pdblSubTotal_LocalTaxable = pdblSubTotal_LocalTaxable + pdblProductExtendedAmount
		If aryOrderItems(i)(enOrderItem_prodStateTaxIsActive) Then pdblSubTotal_StateTaxable = pdblSubTotal_StateTaxable + pdblProductExtendedAmount
		If aryOrderItems(i)(enOrderItem_SpecialTaxFlag_1) Then pdblSubTotal_StateTaxable_Special = pdblSubTotal_StateTaxable_Special + pdblProductExtendedAmount
		If aryOrderItems(i)(enOrderItem_prodCountryTaxIsActive) Then pdblSubTotal_CountryTaxable = pdblSubTotal_CountryTaxable + pdblProductExtendedAmount

		pdblSubTotal = pdblSubTotal + pdblProductExtendedAmount
		
	Next 'i

	'Now populate the discount array, each item in the cart is added as a unique array item
	plngItemCount = - 1
	ReDim p_aryTemp(plngOrderItemCount - 1,enOrder_ArrayLength)

	For i = 0 to plngUniqueOrderItemCount
		For j = 1 To aryOrderItems(i)(enOrderItem_odrdttmpQuantity)
			plngItemCount = plngItemCount + 1

			p_aryTemp(plngItemCount,enOrder_ID) = aryOrderItems(i)(enOrderItem_tmpID)
			p_aryTemp(plngItemCount,enOrder_ProductID) = aryOrderItems(i)(enOrderItem_prodID)
			p_aryTemp(plngItemCount,enOrder_Quantity) = 1
			p_aryTemp(plngItemCount,enOrder_SellPrice) = aryOrderItems(i)(enOrderItem_UnitPrice)
			p_aryTemp(plngItemCount,enOrder_SaleIsActive) = aryOrderItems(i)(enOrderItem_prodSaleIsActive)
			p_aryTemp(plngItemCount,enOrder_StateTaxIsActive) = aryOrderItems(i)(enOrderItem_prodStateTaxIsActive)
			p_aryTemp(plngItemCount,enOrder_CountryTaxIsActive) = aryOrderItems(i)(enOrderItem_prodCountryTaxIsActive)
			p_aryTemp(plngItemCount,enOrder_DiscountAmount) = aryOrderItems(i)(enOrderItem_prodName)

		Next 'j
	Next

	pcurLocalSubTotal = 0
	pcursubTotalMinusSaleItems = 0

	redim paryOrder(plngItemCount,enOrder_ArrayLength)
	For i = 0 to plngItemCount
		For j = 0 to 6
			paryOrder(i,j) = p_aryTemp(i,j)
		Next
		paryOrder(i,enOrder_DiscountAmount) = Array("")
		paryOrder(i,enOrder_BasePrice) = aryOrderItems(i)(enOrderItem_prodBasePrice)
		
		If Not p_aryTemp(i,enOrder_SaleIsActive) Then pcursubTotalMinusSaleItems = pcursubTotalMinusSaleItems + p_aryTemp(i,enOrder_SellPrice) * p_aryTemp(i,enOrder_Quantity)
		pcurLocalSubTotal = pcurLocalSubTotal + p_aryTemp(i,enOrder_SellPrice) * p_aryTemp(i,enOrder_Quantity)
	Next

	If cblnDebugPromotionManagerAddon Then
		Response.Write "<hr><font color=black><fieldset><legend>Load Order Items to Promotion Array</legend>"
		Response.Write "Unique order items: " & plngUniqueOrderItemCount & "<br />"
		Response.Write "Item count: " & plngItemCount & "<br />"
		Response.Write "<ol>"
		For i = 0 To plngItemCount
		 Response.Write "<li>" & p_aryTemp(i,7) & " (" & p_aryTemp(i,1) & ") - " & p_aryTemp(i,3) & "</li>"
		Next 'i
		Response.Write "</ol>"
		
		Response.Write "</fieldset></font>"
	End If

	SetOrderItems = (plngItemCount > -1)
	
End Function	'SetOrderItems

'***********************************************************************************************

Public Function UpdateUsage(byRef objcnn, byVal lngPromotionID)

Dim pstrSQL
Dim pobjRS
Dim p_lngNumUses

'On Error Resume Next

	if len(lngPromotionID) > 0 then
		
		pstrSQL = "Select NumUses from Promotions where PromotionID=" & lngPromotionID
		Set pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 3 'adUseServer
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			'debugprint "pstrSQL", pstrSQL
			If Not .EOF Then
				p_lngNumUses = Trim(.Fields("NumUses").Value & "")
				If Len(p_lngNumUses) = 0 Then 
					p_lngNumUses = 1
				Else
					p_lngNumUses = CLng(p_lngNumUses) + 1
				End If
				
				pstrSQL = "Update Promotions Set NumUses=" & p_lngNumUses & " where PromotionID=" & lngPromotionID
				'debugprint "pstrSQL", pstrSQL
				objcnn.Execute pstrSQL,,128
			End If
			.Close
		End With
		Set pobjRS = Nothing
	end if
	
	UpdateUsage = (Err.number=0)

End Function	'UpdateUsage

'***********************************************************************************************

Public Property Get SubTotalBeforeDiscount
	SubTotalBeforeDiscount = pcursubTotal
End Property

Public Property Get DiscountAmount
	DiscountAmount = pBestDiscountAmount
End Property

Public Property Get SubTotalWithDiscount
	SubTotalWithDiscount = pcursubTotal - pBestDiscountAmount
End Property

'***********************************************************************************************

Public Property Get BestDiscountAmount
	BestDiscountAmount = pBestDiscountAmount
End Property

'***********************************************************************************************

Public Sub SaveDiscountsToDatabase(byRef objcnn, byVal lngOrderID)

Dim i
Dim p_aryCodes
Dim plngNumCodes
Dim p_lngPromotionID
Dim p_dblDiscountAmount
Dim pobjCommand

	On Error Resume Next
	If Err.number <> 0 Then Err.Clear
	
	If Len(lngOrderID) = 0 Then Exit Sub
	
	If isArray(maryDiscountSummary) Then
		plngNumCodes = UBound(maryDiscountSummary)

		For i = 0 To plngNumCodes
			If maryDiscountSummary(i)(enDiscountUse) Then
				p_lngPromotionID = maryDiscountSummary(i)(11)
				p_dblDiscountAmount = maryDiscountSummary(i)(1)
				'For better performance use the Command Object
				'save to both tables for reporting
				If True Then
					Set pobjCommand = CreateObject("ADODB.Command")
					With pobjCommand
						.CommandType = 1	'adCmdText
						.ActiveConnection = objcnn
						.CommandText = "Insert Into ordersDiscounts (OrderID,PromotionID,DiscountAmount) Values (?,?,?)"
						.Parameters.Append .CreateParameter("OrderID", adInteger, adParamInput, 4, lngOrderID)
						.Parameters.Append .CreateParameter("PromotionID", adInteger, adParamInput, 4, p_lngPromotionID)
						.Parameters.Append .CreateParameter("DiscountAmount", adDouble, adParamInput, 4, p_dblDiscountAmount)

						.Execute , , adExecuteNoRecords
					End With
					Set pobjCommand = Nothing
				Else
					pstrSQL = "Insert Into ordersDiscounts (OrderID, PromotionID, DiscountAmount) " _
							& "Values (" & lngOrderID & "," & p_lngPromotionID & "," & p_dblDiscountAmount & ")"
					objcnn.Execute pstrSQL,,128
				End If
				Call UpdateUsage(objcnn, p_lngPromotionID)
			End If
		Next 'i
	End If

End Sub	'SaveDiscountsToDatabase

'***********************************************************************************************

Public Sub SortOrder(byVal blnLayerProducts)

Dim i
Dim j
Dim paryTemp(0,9)

	If cblnDebugPromotionManagerAddon Then
		Response.Write("<table border=1 cellspacing=0 cellpadding=3><tr><th colspan=4>Order Items - preSort</th></tr><br />")
		For i = 0 to uBound(paryOrder)
			Response.Write("<font color=black><tr><td>Item " & i & "</td><td>" & paryOrder(i,1) & "</td><td align=right>" & FormatCurrency(paryOrder(i,3),2) & "</td></tr></font>")
		Next
		Response.Write("</table>")
	End If	'cblnDebugPromotionManagerAddon

	For i = 0 To UBound(paryOrder)
		For j = 0 To UBound(paryOrder)-1
			If paryOrder(j,enOrder_SellPrice) < paryOrder(j+1,enOrder_SellPrice) Then Call SwapItem(paryOrder, j, j+1)
		Next 'j
	Next 'i
	
	If cblnDebugPromotionManagerAddon Then
		Response.Write("<table border=1 cellspacing=0 cellpadding=3><tr><th colspan=4>Order Items - postSort</th></tr><br />")
		For i = 0 to uBound(paryOrder)
			Response.Write("<font color=black><tr><td>Item " & i & "</td><td>" & paryOrder(i,enOrder_ProductID) & "</td><td align=right>" & FormatCurrency(paryOrder(i,enOrder_SellPrice),2) & "</td></tr></font>")
		Next
		Response.Write("</table>")
	End If	'cblnDebugPromotionManagerAddon

	If blnLayerProducts Then
		For i = 0 To UBound(paryOrder)
			For j = 0 To UBound(paryOrder)-1
				If paryOrder(j,enOrder_ProductID) > paryOrder(j+1,enOrder_ProductID) Then Call SwapItem(paryOrder, j, j+1)
			Next 'j
		Next 'i
		
		If cblnDebugPromotionManagerAddon Then
			Response.Write("<table border=1 cellspacing=0 cellpadding=3><tr><th colspan=4>Order Items - postSort Layered Products</th></tr><br />")
			For i = 0 to uBound(paryOrder)
				Response.Write("<font color=black><tr><td>Item " & i & "</td><td>" & paryOrder(i,enOrder_ProductID) & "</td><td align=right>" & FormatCurrency(paryOrder(i,enOrder_SellPrice),2) & "</td></tr></font>")
			Next
			Response.Write("</table>")
		End If	'cblnDebugPromotionManagerAddon

	End If	'blnLayerProducts

End Sub	'SortOrder

'***********************************************************************************************

Public Sub SwapItem(byRef arySource, byVal lngIndex1, byVal lngIndex2)

Dim i
Dim pvntTemp

	For i = 0 to enOrder_ArrayLength
		pvntTemp = arySource(lngIndex1,i)
		arySource(lngIndex1,i) = arySource(lngIndex2,i)
		arySource(lngIndex2,i) = pvntTemp
	Next	'i
	
End Sub	'SwapItem

'***********************************************************************************************

Public Function CalcBestDiscount()

dim p_lngRecordCount
dim i,k
dim plngPromotionCount
dim p_aPossible
'key to p_aPossible
' 0 - pstrPromoCode
' 1 - Discount
' 2 - Combineable
' 3 - pcurSAmountToReduceBy
' 4 - pcurCAmountToReduceBy
' 5 - PromotionID
' 6 - PromoTitle
' 7 - offerFreeGiftAutomatically
' 8 - FreeProductID

Dim pstrPromoCode

Dim ExclusiveAmount
Dim CombinedAmount
Dim pstrCombined
Dim pblnProductLevelDiscount
Dim pblnBuy1Get1
Dim pblnFreeShipping

	If pblnComplete Then Exit Function	'No need to redo the calculation if done before

	If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black>Checking potential discounts . . . </font><br />")

	If LoadPromotions Then
		If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black>Promotions Loaded - calculating order qualifications and individual discount amounts </font><br />")
		plngPromotionCount = -1
		p_lngRecordCount = prsPromo.RecordCount
		redim p_aPossible(p_lngRecordCount-1, 11)
		For i = 0 to uBound(paryOrder)
			ReDim paryDiscounts(p_lngRecordCount-1)
			paryOrder(i,enOrder_DiscountAmount) = paryDiscounts
		Next

		for i = 0 to p_lngRecordCount-1
			pstrPromoCode = trim(prsPromo.Fields("PromoCode").Value)
			If cblnDebugPromotionManagerAddon Then
				Response.Write("<h4><font color=black><b>Possible PromoCode " & i & ": " & pstrPromoCode & "</b></font></h4>")
				Response.Write("<table border=1 cellspacing=0 cellpadding=2>")
				Response.Write("<tr><td><font color=black size=-1>&nbsp;&nbsp;&nbsp;&nbsp;<i>Apply Automatically " & prsPromo.Fields("ApplyAutomatically").Value & "</i></font></td></tr>")
				Response.Write("<tr><td><font color=black size=-1>&nbsp;&nbsp;&nbsp;&nbsp;<i>Min subTotal " & prsPromo.Fields("MinSubTotal").Value & "</i></font></td></tr>")
			End If
			
			'set to true for case insensitive promotions
			If mblnCaseInsensitive Then
				pstrPromoCode = LCase(pstrPromoCode)
				pstrPromotions = LCase(pstrPromotions)
			End If
			
			If instr(1,pstrPromotions,";" & pstrPromoCode & ";") <> 0 or prsPromo.Fields("ApplyAutomatically").Value then
				plngPromotionCount = plngPromotionCount + 1
				p_aPossible(plngPromotionCount,0) = pstrPromoCode
				p_aPossible(plngPromotionCount,5) = trim(prsPromo.Fields("PromotionID").Value)
				p_aPossible(plngPromotionCount,6) = trim(prsPromo.Fields("PromoTitle").Value)
				p_aPossible(plngPromotionCount,7) = trim(prsPromo.Fields("offerFreeGiftAutomatically").Value)
				p_aPossible(plngPromotionCount,8) = trim(prsPromo.Fields("FreeProductID").Value & "")
				p_aPossible(plngPromotionCount,11) = IsAvailableForUseByThisCustomer(pstrPromoCode)
				
				pblnFreeShipping = Len(Trim(prsPromo.Fields("FreeShippingCode").Value & "")) > 0
				If pblnFreeShipping Then
					p_aPossible(plngPromotionCount,10) = Trim(prsPromo.Fields("FreeShippingCode").Value & "") & "|" & Trim(prsPromo.Fields("FreeShippingLimit").Value & "")
				Else
					p_aPossible(plngPromotionCount,10) = ""
				End If
				
				If isNumeric(trim(prsPromo.Fields("MinSubTotal").Value)) Then
					p_aPossible(plngPromotionCount,9) = CDbl(trim(prsPromo.Fields("MinSubTotal").Value))
				Else
					p_aPossible(plngPromotionCount,9) = 0
				End If
				
				'Determine if this is a product level discount
				pblnProductLevelDiscount = CBool(Len(Trim( prsPromo.Fields("ProductID").Value & prsPromo.Fields("ProductIDExclusion").Value _
														   & prsPromo.Fields("Category").Value & prsPromo.Fields("CategoryExclusion").Value _
														   & prsPromo.Fields("Manufacturer").Value & prsPromo.Fields("ManufacturerExclusion").Value _
														   & prsPromo.Fields("Vendor").Value & prsPromo.Fields("VendorExclusion").Value _
														   & prsPromo.Fields("FreeProductID").Value & "" _
														  ))> 0)
														  
				'Determine if this is a buy 1 get 1 free
				pblnBuy1Get1 = CBool(Len(Trim(prsPromo.Fields("buyX").Value & prsPromo.Fields("getY").Value & prsPromo.Fields("FreeProductID").Value & "" ))> 0)
														  
				'Can't be a product level if it is free shipping or Buy1Get1
				If pblnFreeShipping Or pblnBuy1Get1 Then pblnProductLevelDiscount = False				
														  
				If cblnDebugPromotionManagerAddon Then
					Response.Write("<tr><td><font color=black size=-1>&nbsp;&nbsp;&nbsp;&nbsp;pblnProductLevelDiscount: " & pblnProductLevelDiscount & "</font></td></tr>")
					Response.Write("<tr><td><font color=black size=-1>&nbsp;&nbsp;&nbsp;&nbsp;pblnBuy1Get1: " & pblnBuy1Get1 & "</font></td></tr>")
					Response.Write("<tr><td><font color=black size=-1>&nbsp;&nbsp;&nbsp;&nbsp;pblnFreeShipping: " & pblnFreeShipping & "</font></td></tr>")
					Response.Write("</table>")
				End If

				If pblnBuy1Get1 then
					Call SortOrder(CBool(prsPromo.Fields("likeItem").Value <> 0))
					Call CalcBuy1Get1(plngPromotionCount, p_aPossible)

					If prsPromo.Fields("offerFreeGiftAutomatically").Value Then
						p_aPossible(plngPromotionCount,7) = prsPromo.Fields("getY").Value
					Else
						p_aPossible(plngPromotionCount,7) = -1
					End If
				ElseIf pblnFreeShipping then
					If QualifiesForFreeShipping(plngPromotionCount) Then
					Else
						p_aPossible(plngPromotionCount,10) = ""
					End If
					p_aPossible(plngPromotionCount,1) = 0
				ElseIf pblnProductLevelDiscount then
					p_aPossible(plngPromotionCount,1) = CalcProductDiscount(plngPromotionCount)
				Else
					p_aPossible(plngPromotionCount,1) = CalcOrderLevelDiscount(plngPromotionCount)
				End If
				
				p_aPossible(plngPromotionCount,2) = not prsPromo("Combineable").Value
			Else
				If cblnDebugPromotionManagerAddon Then Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;- <font color=black><i>Not registered for</i></font><br />")			
			end if
			prsPromo.MoveNext
			
			If cblnDebugPromotionManagerAddon Then
				Response.Write("<table border=1 cellspacing=0 cellpadding=2>")
				Response.Write("<tr><td colspan=2><font color=black><b>Possible PromoCode " & i & ": " & pstrPromoCode & "</b></font></td></tr>")
				Response.Write("<tr>")
				Response.Write("<th>Use</th>")
				Response.Write("<th>Amount</th>")
				Response.Write("</tr>")
				Response.Write("<tr>")
				Response.Write("<td align=center><font color=black size=-1>" & p_aPossible(i, 11) & "</font></td>")
				Response.Write("<td align=center><font color=black size=-1>" & p_aPossible(i, 1) & "</font></td>")
				Response.Write("</tr>")
				Response.Write("</table>")
			End If
			
		Next 'i

		If cblnDebugPromotionManagerAddon Then 
			Response.Write("<h4><font color=black>Now calculating best discount . . .</font></h4>")
			For i = 0 To plngPromotionCount
				Response.Write("<font color=black>Codes which will be checked - " & i & ": " & p_aPossible(i,0) & "</font><br />")
			next 'i
		End If

		ExclusiveAmount = 0
		CombinedAmount = 0

		Dim plngProductCounter
		ReDim maryDiscountSummary(plngPromotionCount)
		
		'Now determine which method gives the biggest discount
		For i = 0 To plngPromotionCount
			'maryDiscountSummary(i) = Array("code", "discount", "combineable", "sTaxable", "cTaxable", "use", "Title", "offerFreeGiftAutomatically", "FreeProductID", "MinSubTotal", "FreeShippingCode")
			
			maryDiscountSummary(i) = Array("", 0, False, 0, 0, False, "", "", "", "", "", "")
			
			maryDiscountSummary(i)(0) = p_aPossible(i,0)
			maryDiscountSummary(i)(2) = p_aPossible(i,2)
			maryDiscountSummary(i)(enPromoTitle) = p_aPossible(i,6)
			
			If isNull(p_aPossible(i,7)) Then
				maryDiscountSummary(i)(enofferFreeGiftAutomatically) = False
			Else
				maryDiscountSummary(i)(enofferFreeGiftAutomatically) = CBool(p_aPossible(i,7))
			End If
			
			maryDiscountSummary(i)(enDiscountFreeProduct) = p_aPossible(i,8)
			maryDiscountSummary(i)(9) = p_aPossible(i,9)
			maryDiscountSummary(i)(enDiscountFreeShipping) = p_aPossible(i,10)
			maryDiscountSummary(i)(11) = p_aPossible(i,5)	'promotionID
			maryDiscountSummary(i)(enDiscountUse) = p_aPossible(i, 11)
			
			For plngProductCounter = 0 to uBound(paryOrder)
				maryDiscountSummary(i)(1) = CDbl(maryDiscountSummary(i)(1)) + CDbl(paryOrder(plngProductCounter,enOrder_DiscountAmount)(i))
			Next
			
			If maryDiscountSummary(i)(2) Then	'Check for exclusive
				If maryDiscountSummary(i)(1) >= ExclusiveAmount Then
					ExclusiveAmount = maryDiscountSummary(i)(1)
					pstrExclusiveCode = maryDiscountSummary(i)(0)
				End If
			Else
				CombinedAmount = CDbl(CombinedAmount) + maryDiscountSummary(i)(1)
			End If	'maryDiscountSummary(i)(2) - Check for exclusive

			If cblnDebugPromotionManagerAddon Then
				Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;- <font color=black>pblnProductLevelDiscount: " & pblnProductLevelDiscount & "</font><br />")
				Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;- <font color=black>pblnBuy1Get1: " & pblnBuy1Get1 & "</font><br />")
				Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;- <font color=black>Exclusive: " & maryDiscountSummary(i)(2) & "</font><br />")
			End If
			
		Next 'i

		If cblnDebugPromotionManagerAddon Then 
			Response.Write("<font color=black>Exclusive Amount: " & ExclusiveAmount & "</font><br />")
			Response.Write("<font color=black>Exclusive Code: " & pstrExclusiveCode & "</font><br />")
			Response.Write("<font color=black>Combined Amount: " & CombinedAmount & "</font><br />")
		End If

		'Mark the promotions as valid which present the biggest discounts
		pBestDiscountAmount = 0
		Dim pcursubTotal_Local:	pcursubTotal_Local = 0
		Dim pcurItemPriceCheck:	pcurItemPriceCheck = 0
		pcursubTotalSTaxable = 0
		pcursubTotalCTaxable = 0
		
		If plngPromotionCount < 0 Then
			For plngProductCounter = 0 to uBound(paryOrder)
				If paryOrder(plngProductCounter,enOrder_StateTaxIsActive) Then pcursubTotalSTaxable = CDbl(pcursubTotalSTaxable) + paryOrder(plngProductCounter,enOrder_SellPrice)
				If paryOrder(plngProductCounter,enOrder_CountryTaxIsActive) Then pcursubTotalCTaxable = CDbl(pcursubTotalCTaxable) + paryOrder(plngProductCounter,enOrder_SellPrice)
			Next 'plngProductCounter
		End	If 'plngPromotionCount < 0
		
		For i = 0 To plngPromotionCount
		
			If cblnDebugPromotionManagerAddon And False Then 
				Response.Write("<font color=black>pcursubTotal: " & pcursubTotal & "</font><br />")
				Response.Write("<font color=black>MinSubTotal: " & maryDiscountSummary(i)(9) & "</font><br />")
				Response.Write("<font color=black>MinSubTotal <= pcursubTotal: " & CBool(maryDiscountSummary(i)(9) <= CDbl(pcursubTotal)) & "</font><br />")
				Response.Write("<font color=black>Exclusive Amount: " & ExclusiveAmount & "</font><br />")
				Response.Write("<font color=black>Combined Amount: " & CombinedAmount & "</font><br />")
			End If

			If maryDiscountSummary(i)(2) Then	'Check for exclusive
				maryDiscountSummary(i)(enDiscountUse) = maryDiscountSummary(i)(enDiscountUse) And CBool(ExclusiveAmount >= CombinedAmount) And CBool(maryDiscountSummary(i)(0) = pstrExclusiveCode) And CBool(maryDiscountSummary(i)(9) <= CDbl(pcursubTotal))
			Else
				maryDiscountSummary(i)(enDiscountUse) = maryDiscountSummary(i)(enDiscountUse) And CBool(CombinedAmount >= ExclusiveAmount) And CBool(maryDiscountSummary(i)(9) <= CDbl(pcursubTotal))
			End If

			If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black>Use this code - PromoCode " & maryDiscountSummary(i)(0) & ": " & maryDiscountSummary(i)(1) & ": " & maryDiscountSummary(i)(enDiscountUse) & "</font><br />")
			
			'Sum up the discounts applied to taxable items
			maryDiscountSummary(i)(3) = 0
			maryDiscountSummary(i)(4) = 0
			
			For plngProductCounter = 0 to uBound(paryOrder)
				If i = 0 Then
					pcursubTotal_Local = CDbl(pcursubTotal_Local) + paryOrder(plngProductCounter,enOrder_SaleIsActive)
					pcurItemPriceCheck = 0
					paryOrder(plngProductCounter,enOrder_DiscountRunningTotal) = 0
				End If
				
				pcurItemPriceCheck = pcurItemPriceCheck + paryOrder(plngProductCounter,enOrder_DiscountAmount)(i)
				If pcurItemPriceCheck < maryDiscountSummary(i)(3) Then
					pcurItemPriceCheck = maryDiscountSummary(i)(3) - paryOrder(plngProductCounter,enOrder_DiscountRunningTotal)
					paryOrder(plngProductCounter,enOrder_DiscountRunningTotal) = paryOrder(plngProductCounter,enOrder_SellPrice)
				Else
					paryOrder(plngProductCounter,enOrder_DiscountRunningTotal) = pcurItemPriceCheck
					pcurItemPriceCheck = paryOrder(plngProductCounter,enOrder_DiscountAmount)(i)
				End If
				
				'Check to see if item is excluded from State Tax
				If paryOrder(plngProductCounter,enOrder_StateTaxIsActive) Then
					If i = 0 Then pcursubTotalSTaxable = CDbl(pcursubTotalSTaxable) + paryOrder(plngProductCounter,enOrder_SellPrice)
					maryDiscountSummary(i)(3) = CDbl(maryDiscountSummary(i)(3)) + paryOrder(plngProductCounter,enOrder_DiscountAmount)(i)
				End If
				
				'Check to see if item is excluded from Country Tax
				If paryOrder(plngProductCounter,enOrder_CountryTaxIsActive) Then
					If i = 0 Then pcursubTotalCTaxable = pcursubTotalCTaxable + paryOrder(plngProductCounter,enOrder_SellPrice)
					maryDiscountSummary(i)(4) = CDbl(maryDiscountSummary(i)(4)) + paryOrder(plngProductCounter,enOrder_DiscountAmount)(i)
				End If
				
				'Response.Write("<font color=black>Item " & plngProductCounter & " " & paryOrder(plngProductCounter,1) & ": <i>" & paryOrder(plngProductCounter,3) & " - " & paryOrder(plngProductCounter,7)(i) & "<i></font><br />")
			Next
			'Response.Write("<font color=black>xxxPromoCode " & maryDiscountSummary(i)(0) & ": " & maryDiscountSummary(i)(enDiscountUse) & "</font><br />")
			'Response.Write("<font color=black>xxxPromoCode " & maryDiscountSummary(i)(0) & ": " & maryDiscountSummary(i)(enPromoTitle) & "</font><br />")

			If maryDiscountSummary(i)(enDiscountUse) Then
				pBestDiscountAmount = CDbl(pBestDiscountAmount) + CDbl(maryDiscountSummary(i)(1))
				If maryDiscountSummary(i)(1) > 0 Then
					mstrDiscountText = mstrDiscountText & maryDiscountSummary(i)(enPromoTitle) & " - Save " & FormatCurrency(maryDiscountSummary(i)(1),2) & "<br />"
				ElseIf Len(maryDiscountSummary(i)(enDiscountFreeShipping)) > 0 Then
					mstrDiscountText = mstrDiscountText & maryDiscountSummary(i)(enPromoTitle) & "<br />"
				End If
				pcursubTotalSTaxable = pcursubTotalSTaxable - maryDiscountSummary(i)(3)
				pcursubTotalCTaxable = pcursubTotalCTaxable - maryDiscountSummary(i)(4)
				If cblnDebugPromotionManagerAddon Then
					Response.Write("<h4><font color=black>pBestDiscountAmount (" & maryDiscountSummary(i)(0) & ") - " & FormatCurrency(pBestDiscountAmount) & "</font></h4>")
				End If	'cblnDebugPromotionManagerAddon
			End If
			
			'Response.Write("<font color=black>PromoCode " & maryDiscountSummary(i)(0) & ": " & maryDiscountSummary(i)(enDiscountUse) & "</font><br />")
			
			If Len(maryDiscountSummary(i)(enDiscountFreeShipping)) > 0 Then
			'	debugprint "free shipping", maryDiscountSummary(i)(enDiscountFreeShipping)
			End If
			
			'Set use to False if value is $0, no free productID, and no shipping
			If (maryDiscountSummary(i)(1) <= 0) And (Len(maryDiscountSummary(i)(enDiscountFreeShipping)) = 0)  And (Len(maryDiscountSummary(i)(enDiscountFreeProduct)) = 0) Then
				If cblnDebugPromotionManagerAddon Then
					Response.Write("<br /><table border=1 cellspacing=0 cellpadding=2>")
					Response.Write("<tr><th colspan=3>Changing usage to false</th></tr>")
					Response.Write("<tr><th>Promotion (0)</th><th>Discount (1)</th><th>maryDiscountSummary(i)(enDiscountFreeShipping)</th></tr>")
					Response.Write("<font color=black><tr><td>" & maryDiscountSummary(i)(0) & " " & maryDiscountSummary(i)(enPromoTitle) & "</td><td>&nbsp;" & maryDiscountSummary(i)(1) & "</td><td>&nbsp;" & maryDiscountSummary(i)(enDiscountFreeShipping) & "</td></tr></font>")
					Response.Write("</table>")
				End If
				maryDiscountSummary(i)(enDiscountUse) = False
			End If
			
		Next 'i

		If cblnDebugPromotionManagerAddon Then
			'maryDiscountSummary(i) = Array("code", "discount", "combineable", "sTaxable", "cTaxable", "use", "Title", "offerFreeGiftAutomatically", "FreeProductID")
			Response.Write("<br /><table border=1 cellspacing=0 cellpadding=2><tr><th colspan=7>Discount Summary</th></tr>")
			Response.Write("<tr><th>Promotion (0)</th><th>Discount (1)</th><th>Exclusive (2)</th><th>Use (5)</th><th>offerFreeGiftAutomatically</th><th>FreeProductID</th><th>FreeShipping</th></tr>")
			For i = 0 to uBound(maryDiscountSummary)
				Response.Write("<font color=black><tr><td>" & maryDiscountSummary(i)(0) & " " & maryDiscountSummary(i)(enPromoTitle) & "</td><td>" & FormatCurrency(maryDiscountSummary(i)(1),2) & "</td><td align=center>" & maryDiscountSummary(i)(2) & "</td><td align=center>" & maryDiscountSummary(i)(enDiscountUse) & "</td><td align=center>" & maryDiscountSummary(i)(enofferFreeGiftAutomatically) & "</td><td align=center>&nbsp;" & maryDiscountSummary(i)(enDiscountFreeProduct) & "</td><td align=center>&nbsp;" & maryDiscountSummary(i)(enDiscountFreeShipping) & "</td></tr></font>")
			Next
			'Response.Write("<tr><td colspan=2>&nbsp;</td><td align=right>Total:&nbsp;</td><td align=right><font color=black>" & FormatCurrency(pcurDiscount,2) & "</td></tr></font>")
			Response.Write("</table><hr>")
		End If	'cblnDebugPromotionManagerAddon
	

		If cblnDebugPromotionManagerAddon Then 
			Response.Write("<hr>")
			Response.Write("<font color=black>subTotal: " & pcursubTotal_Local & "</font><br />")
			Response.Write("<font color=black>subTotal State Taxable: " & pcursubTotalSTaxable & "</font><br />")
			Response.Write("<font color=black>subTotal Country Taxable: " & pcursubTotalCTaxable & "</font><br />")
			
			For i = 0 To plngPromotionCount
				Response.Write("<font color=black>PromoCode " & i & ": " & maryDiscountSummary(i)(0) & " - " & maryDiscountSummary(i)(enPromoTitle) & " - " & maryDiscountSummary(i)(1) & "</font><br />")
				For plngProductCounter = 0 to uBound(paryOrder)
					Response.Write("<font color=black>Item " & plngProductCounter & " " & paryOrder(plngProductCounter,enOrder_ProductID) & ": <i>" & paryOrder(plngProductCounter,enOrder_SellPrice) & " - " & paryOrder(plngProductCounter,enOrder_DiscountAmount)(i) & "<i></font><br />")
				Next
				Response.Write("<hr>")
			Next 'i
			
		End If

	Else
		If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black>Error Loading Promotions</font><br />")

		For plngProductCounter = 0 to uBound(paryOrder)
			If paryOrder(plngProductCounter,enOrder_StateTaxIsActive) Then pcursubTotalSTaxable = CDbl(pcursubTotalSTaxable) + paryOrder(plngProductCounter,enOrder_SellPrice)
			If paryOrder(plngProductCounter,enOrder_CountryTaxIsActive) Then pcursubTotalCTaxable = CDbl(pcursubTotalCTaxable) + paryOrder(plngProductCounter,enOrder_SellPrice)
		Next 'plngProductCounter

		pBestDiscountAmount = 0
	End If	'LoadPromotions
	
	pblnComplete = True
	CalcBestDiscount = pBestDiscountAmount
	
End Function	'CalcBestDiscount

'***********************************************************************************************

Public Function FirstTimeCustomerDiscountAmount

Dim pblnResult

	If FirstTimeCustomerDiscount_IsPercent Then
		'Only apply the discount against the discounted total
		pblnResult = (CDbl(pcursubTotal) - CDbl(pBestDiscountAmount)) * FirstTimeCustomerDiscount / 100
	Else
		pblnResult = FirstTimeCustomerDiscount
	End If
	
	FirstTimeCustomerDiscountAmount = Round(pblnResult, 2)

End Function	'FirstTimeCustomerDiscountAmount

'***********************************************************************************************

Private Function CalcOrderLevelDiscount(lngIndex)

Dim p_curSubTotalToUse
Dim i
Dim pcurProductPrice
Dim plngQuantity
Dim plngMaxItems
Dim plngqtyToApplyTo
Dim pcursubDiscount
Dim pcurDiscount
Dim paryDiscounts
Dim pblnFlatDiscountApplied

'On Error Resume Next

	if pcursubTotal >= CDbl(prsPromo.Fields("MinSubTotal").Value) then
	
		pblnFlatDiscountApplied = False
		For i = 0 to uBound(paryOrder)
			If Not (prsPromo.Fields("ExcludeSaleItems").Value AND paryOrder(i,enOrder_SaleIsActive)) Then
				If (pcursubTotal) >= CDbl(prsPromo.Fields("MinSubTotal").Value) Then

					if prsPromo.Fields("ApplyToBasePrice").Value then
						pcurProductPrice = paryOrder(i,enOrder_BasePrice)
					else
						pcurProductPrice = paryOrder(i,enOrder_SellPrice)
					end if

					plngQuantity = paryOrder(i,enOrder_Quantity)

					'this section has been edited in preparation for limiting product based promotions to a max number of items
					plngMaxItems = Trim(prsPromo.Fields("productCountLimit").Value & "")
					'plngMaxItems = ""
					If Len(plngMaxItems) = 0 Then
						plngqtyToApplyTo = plngQuantity
					Else
						If plngQuantity > plngMaxItems Then
							plngqtyToApplyTo = plngMaxItems
						Else
							plngqtyToApplyTo = plngQuantity
						End If
					End If
					
					if prsPromo.Fields("Percentage").Value then
						pcursubDiscount = CDbl(prsPromo.Fields("Discount").Value)/100 * pcurProductPrice * plngqtyToApplyTo
					else
						If pblnFlatDiscountApplied Then
							pcursubDiscount = 0
						Else
							pcursubDiscount = CDbl(prsPromo.Fields("Discount").Value) * plngqtyToApplyTo
							pblnFlatDiscountApplied = True
						End If
					end if
					
					'limit discount by order line item
					If Len(Trim(prsPromo.Fields("MaxAllowableValuePerItem").Value & "")) > 0 Then
						If pcursubDiscount >= CDbl(prsPromo.Fields("MaxAllowableValuePerItem").Value) Then pcursubDiscount = CDbl(prsPromo.Fields("MaxAllowableValuePerItem").Value)
					End If
					
					'limit discount by total order
					If Len(Trim(prsPromo.Fields("MaxAllowableValue").Value & "")) > 0 Then
						If CDbl(pcurDiscountTotal + pcursubDiscount) >= CDbl(prsPromo.Fields("MaxAllowableValue").Value) Then
							pcursubDiscount = CDbl(prsPromo.Fields("MaxAllowableValue").Value) - pcurDiscountTotal
						End If
					End If

					paryDiscounts = paryOrder(i,enOrder_DiscountAmount)
					'ReDim Preserve paryDiscounts(lngIndex)
					paryDiscounts(lngIndex) = pcursubDiscount
					pcurDiscount = CDbl(pcurDiscount) + cDbl(pcursubDiscount)
					paryOrder(i,enOrder_DiscountAmount) = paryDiscounts
				else
					pcurDiscount = 0
				end if
			End If
		Next
	
		pcurDiscount= Round(pcurDiscount, 2)

	else
		pcurDiscount = 0
	end if
	
	CalcOrderLevelDiscount = pcurDiscount
	
	If cblnDebugPromotionManagerAddon Then
		Response.Write("<br /><table border=1 cellspacing=0 cellpadding=2><tr><th colspan=4>Order Level Discount - " & prsPromo.Fields("PromoCode").Value & "<br />" & prsPromo.Fields("PromoTitle").Value & "</th></tr>")
		Response.Write("<tr><th>Item</th><th>Product</th><th>Unit Price</th><th>Discount</th></tr>")
		For i = 0 to uBound(paryOrder)
			Response.Write("<font color=black><tr><td>Item " & i & "</td><td>" & paryOrder(i,enOrder_ProductID) & "</td><td align=right>" & FormatCurrency(paryOrder(i,enOrder_SellPrice), 3) & "</td><td align=right>" & FormatCurrency(paryOrder(i,enOrder_DiscountAmount)(lngIndex),2) & "</td></tr></font>")
		Next
		Response.Write("<tr><td colspan=2>&nbsp;</td><td align=right>Total:&nbsp;</td><td align=right><font color=black>" & FormatCurrency(pcurDiscount,2) & "</td></tr></font>")
		If (pcursubTotal) < CDbl(prsPromo.Fields("MinSubTotal").Value) Then
			Response.Write("<tr><td colspan=4 align=center><font color=red>No discount applied. subTotal of " & FormatCurrency(pcursubTotal,2) & " is less than minimum of " & FormatCurrency(CDbl(prsPromo.Fields("MinSubTotal").Value),2) & "</font></td></tr></font>")
		End If
		Response.Write("</table><hr>")
	End If	'cblnDebugPromotionManagerAddon
	
End Function	'CalcOrderLevelDiscount

'***********************************************************************************************

Private Function MakeProductSearchSQL(strFieldName, strSourceValue, strEquality, strValueWrapper, strANDOR)

Dim pstrTemp

	pstrTemp = Trim(strSourceValue & "")
	If True Then
		If strEquality = "=" Then
			If Len(pstrTemp) > 0 Then pstrTemp = " " & strANDOR & "(" & strFieldName & " In (" & strValueWrapper & Replace(pstrTemp, cstrDelimiter, strValueWrapper & "," & strValueWrapper) & strValueWrapper & ")" & ")"
		Else
			If Len(pstrTemp) > 0 Then pstrTemp = " " & strANDOR & "(" & strFieldName & " Not In (" & strValueWrapper & Replace(pstrTemp, cstrDelimiter, strValueWrapper & "," & strValueWrapper) & strValueWrapper & ")" & ")"
		End If
	Else
		If Len(pstrTemp) > 0 Then pstrTemp = " " & strANDOR & " ((" & strFieldName & strEquality & strValueWrapper & Replace(pstrTemp, cstrDelimiter, "" & strValueWrapper & ") OR (" & strFieldName & strEquality & strValueWrapper & "") & strValueWrapper & "))"
	End If
	MakeProductSearchSQL = pstrTemp
	
End Function	'MakeProductSearchSQL

'***********************************************************************************************

Private Function CalcBuy1Get1(byVal lngIndex, byRef aryDiscount)
'This function deals with buy1get1 free promotions and free product with order 

dim pcursubDiscount
dim pcurDiscount
dim pstrProductID
Dim pcurDiscountTotal
dim pcurProductPrice
dim i
dim pblnProductMatch
dim plngMaxItems
Dim pstrApplicableProductIDs
Dim pstrFreeProductIDs
Dim paryDiscounts
Dim pstrPrevProductID
Dim plngNextFreeID
Dim plngFreeItemCounter
Dim plngQualifyingItemCounter
Dim pblnThisItemDiscounted
Dim pblnLikeItem
Dim pdblSubTotalWithoutFreeProduct
Dim pblnQualifiesForFreeProduct
Dim plngBuyX
Dim pdblSubTotal_Qualifying

Dim plngNumberOfFreeQualifyingProductsInCart
Dim plngNumberOfFreeQualifyingUniqueProductsInCart

'On Error Resume Next

	If cblnDebugPromotionManagerAddon Then Response.Write("<h4><font color=black>Checking for CalcBuy1Get1 . . .</font></h4>")

	plngNumberOfFreeQualifyingProductsInCart = 0
	plngNumberOfFreeQualifyingUniqueProductsInCart = 0
	pcurDiscount = 0
	pcurDiscountTotal = 0
	pstrApplicableProductIDs = GetQualifyingProducts
	pblnQualifiesForFreeProduct = False

	pstrFreeProductIDs = WrapString(Trim(prsPromo.Fields("FreeProductID").Value & ""),True)

	plngMaxItems = Trim(prsPromo.Fields("productCountLimit").Value & "")
	If Len(plngMaxItems) = 0 Then
		plngMaxItems = 0
	Else
		plngMaxItems = CLng(plngMaxItems)
	End If

	If Len(prsPromo.Fields("likeItem").Value & "") > 0 Then
		pblnLikeItem = CBool(prsPromo.Fields("likeItem").Value)
	Else
		pblnLikeItem = False
	End If

	If Len(Trim(prsPromo.Fields("buyX").Value & "")) > 0 Then
		plngBuyX = prsPromo.Fields("buyX").Value	
	Else
		plngBuyX = 1
	End If
	plngNextFreeID = 0 + plngBuyX	

	plngFreeItemCounter = 0
	plngQualifyingItemCounter = 0
	pblnThisItemDiscounted = False

	If cblnDebugPromotionManagerAddon Then
		Response.Write("<table border=1 cellspacing=0 cellpadding=2><tr><th colspan=7>CalcBuy1Get1 Discount (" & prsPromo.Fields("PromoCode").Value & ")</th></tr>")
		Response.Write("<tr><th colspan=7><hr></th></tr>")
		Response.Write("<tr><td colspan=7>&nbsp;&nbsp;buyX:" & prsPromo.Fields("buyX").Value & "</td></tr>")
		Response.Write("<tr><td colspan=7>&nbsp;&nbsp;getY:" & prsPromo.Fields("getY").Value & "</td></tr>")
		Response.Write("<tr><td colspan=7>&nbsp;&nbsp;FreeShippingLimit:" & prsPromo.Fields("FreeShippingLimit").Value & "</td></tr>")
		Response.Write("<tr><td colspan=7>&nbsp;&nbsp;likeItem:" & pblnLikeItem & "</td></tr>")
		Response.Write("<tr><td colspan=7>&nbsp;&nbsp;FreeProductID:" & pstrFreeProductIDs & "</td></tr>")
		Response.Write("<tr><td colspan=7>&nbsp;&nbsp;offerFreeGiftAutomatically:" & prsPromo.Fields("offerFreeGiftAutomatically").Value & "</td></tr>")
		Response.Write("<tr><td colspan=7>&nbsp;&nbsp;productCountLimit:" & prsPromo.Fields("productCountLimit").Value & "</td></tr>")
		Response.Write("<tr><th colspan=7><hr></th></tr>")
		Response.Write("<tr><th>Counter</th><th>Product ID</th><th>Match</th><th>Next ID</th><th>Discounted</th><th>Qualifying Items</th><th>Free Items</th></tr>")
	End If	'cblnDebugPromotionManagerAddon

	pdblSubTotal_Qualifying = 0
	For i = 0 to uBound(paryOrder)
		
		pstrProductID = WrapString(paryOrder(i,enOrder_ProductID),True)
		'pblnProductMatch = instr(1,pstrApplicableProductIDs,pstrProductID) > 0
		pblnProductMatch = isProductMatch(pstrProductID, pstrApplicableProductIDs)

		If CBool(instr(1,pstrFreeProductIDs,pstrProductID) > 0) Then
			pdblSubTotalWithoutFreeProduct = CDbl(pdblSubTotalWithoutFreeProduct) + paryOrder(i,enOrder_SellPrice)
		End If

		If pblnProductMatch then
			If Not (prsPromo.Fields("ExcludeSaleItems").Value AND paryOrder(i,enOrder_SaleIsActive)) Then pdblSubTotal_Qualifying = pdblSubTotal_Qualifying + paryOrder(i,enOrder_SellPrice)

			If Not (prsPromo.Fields("ExcludeSaleItems").Value AND paryOrder(i,enOrder_SaleIsActive)) Then
				If (pcursubTotal) >= CDbl(prsPromo.Fields("MinSubTotal").Value) Then
					plngNumberOfFreeQualifyingProductsInCart = plngNumberOfFreeQualifyingProductsInCart + 1

					pblnThisItemDiscounted = CBool(i >= plngNextFreeID)
					If plngQualifyingItemCounter = 0 Then
						plngQualifyingItemCounter = prsPromo.Fields("getY").Value

						plngNextFreeID = i + plngBuyX	

						If pstrPrevProductID <> paryOrder(i,enOrder_ProductID) Then
							pstrPrevProductID = paryOrder(i,enOrder_ProductID)
							plngNumberOfFreeQualifyingUniqueProductsInCart = plngNumberOfFreeQualifyingUniqueProductsInCart + 1
							If pblnLikeItem Then
								plngNextFreeID = i + plngBuyX	
							End If
						End If
						pblnThisItemDiscounted = False
					Else
						If pstrPrevProductID <> paryOrder(i,enOrder_ProductID) Then
							pstrPrevProductID = paryOrder(i,enOrder_ProductID)
							plngNumberOfFreeQualifyingUniqueProductsInCart = plngNumberOfFreeQualifyingUniqueProductsInCart + 1
							If pblnLikeItem Then
								plngNextFreeID = i + plngBuyX	
								pblnThisItemDiscounted = False
							End If
						End If
					End If
					
					If CBool(plngMaxItems > 0) Then
						pblnThisItemDiscounted = pblnThisItemDiscounted And CBool(plngFreeItemCounter < plngMaxItems)
					End If
				
					If pblnThisItemDiscounted Then
						plngQualifyingItemCounter = plngQualifyingItemCounter - 1
						plngFreeItemCounter = plngFreeItemCounter + 1
						pcurProductPrice = paryOrder(i,enOrder_SellPrice)
						if prsPromo.Fields("Percentage").Value then
							pcursubDiscount = Round((CDbl(prsPromo.Fields("Discount").Value)/100) * pcurProductPrice,2)
						else
							pcursubDiscount = Round(CDbl(prsPromo.Fields("Discount").Value),2)
						end if
					Else
						pcursubDiscount = 0
					End If	'pblnThisItemDiscounted
					
					'limit discount by order line item
					If Len(Trim(prsPromo.Fields("MaxAllowableValuePerItem").Value & "")) > 0 Then
						If pcursubDiscount >= CDbl(prsPromo.Fields("MaxAllowableValuePerItem").Value) Then pcursubDiscount = CDbl(prsPromo.Fields("MaxAllowableValuePerItem").Value)
					End If
					
					'limit discount by total order
					If Len(Trim(prsPromo.Fields("MaxAllowableValue").Value & "")) > 0 Then
						If CDbl(pcurDiscountTotal + pcursubDiscount) >= CDbl(prsPromo.Fields("MaxAllowableValue").Value) Then
							pcursubDiscount = CDbl(prsPromo.Fields("MaxAllowableValue").Value) - pcurDiscountTotal
						End If
					End If
					pcurDiscountTotal = pcurDiscountTotal + pcursubDiscount

					paryDiscounts = paryOrder(i,enOrder_DiscountAmount)
					'ReDim Preserve paryDiscounts(lngIndex)
					paryDiscounts(lngIndex) = pcursubDiscount
					pcurDiscount = pcurDiscount + cDbl(pcursubDiscount)
					paryOrder(i,enOrder_DiscountAmount) = paryDiscounts
				else
					pcurDiscount = 0
				end if
			End If
		Else
			pblnThisItemDiscounted = False
			plngNextFreeID = plngNextFreeID + 1
		End If	'pblnProductMatch
		
		If cblnDebugPromotionManagerAddon Then
			If pblnThisItemDiscounted Then
				Response.Write("<tr bgcolor=yellow>")
			Else
				Response.Write("<tr>")
			End If
			Response.Write("<td>" & i & "</td><td>" & paryOrder(i,enOrder_ProductID) & "</td><td>" & pblnProductMatch & "</td><td>" & plngNextFreeID & "</td><td>" & pblnThisItemDiscounted & " (" & FormatCurrency(paryOrder(i,enOrder_DiscountAmount)(lngIndex),2) & ")</td><td>" & plngQualifyingItemCounter & "</td><td>" & plngFreeItemCounter & "</td></tr>")
		End If	'cblnDebugPromotionManagerAddon
		
	Next	'i
	
	If cblnDebugPromotionManagerAddon Then
		Response.Write("</table>")
	End If	'cblnDebugPromotionManagerAddon
	
	'Now check for the free product with order
	'The above loop was necessary to determine qualification - ie. necessary items in cart meeting select criteria
	
	pblnQualifiesForFreeProduct = plngNumberOfFreeQualifyingProductsInCart >= plngBuyX And Len(pstrFreeProductIDs) > 0
	
	If Len(prsPromo.Fields("MinSubTotal").Value & "") > 0 Then
		If CDbl(prsPromo.Fields("MinSubTotal").Value) > CDbl(pdblSubTotal_Qualifying) Then pblnQualifiesForFreeProduct = False
	End If

	If cblnDebugPromotionManagerAddon Then
		Response.Write("<table border=1 cellspacing=0 cellpadding=2><tr><th>CalcBuy1Get1 Discount - Free Products (" & prsPromo.Fields("PromoCode").Value & ")</th></tr>")
		Response.Write("<tr><th><hr></th></tr>")
		Response.Write("<tr><td>&nbsp;&nbsp;Number Of Free Qualifying Products In Cart:" & plngNumberOfFreeQualifyingProductsInCart & "</td></tr>")
		Response.Write("<tr><td>&nbsp;&nbsp;Number Of Free Qualifying Unique Products In Cart:" & plngNumberOfFreeQualifyingUniqueProductsInCart & "</td></tr>")
		Response.Write("<tr><td>&nbsp;&nbsp;Qualifying SubTotal:" & pdblSubTotal_Qualifying & "</td></tr>")
		Response.Write("<tr><td>&nbsp;&nbsp;Min SubTotal:" & prsPromo.Fields("MinSubTotal").Value & "</td></tr>")
		Response.Write("<tr><td>&nbsp;&nbsp;Qualifies For Free Product:" & pblnQualifiesForFreeProduct & "</td></tr>")
	End If
	
	If pblnQualifiesForFreeProduct Then

		pcursubDiscount = 0
		
		If cblnDebugPromotionManagerAddon Then
			Response.Write("<table border=1 cellspacing=0 cellpadding=2><tr><th colspan=7>CalcBuy1Get1 Discount - Free Products</th></tr>")
			Response.Write("<tr><th colspan=7><hr></th></tr>")
			Response.Write("<tr><td colspan=7>&nbsp;&nbsp;FreeProductID:" & Replace(pstrFreeProductIDs & "", ";", " ") & "</td></tr>")
			Response.Write("<tr><td colspan=7>&nbsp;&nbsp;offerFreeGiftAutomatically:" & prsPromo.Fields("offerFreeGiftAutomatically").Value & "</td></tr>")
			Response.Write("<tr><td colspan=7>&nbsp;&nbsp;productCountLimit:" & prsPromo.Fields("productCountLimit").Value & "</td></tr>")
			Response.Write("<tr><th colspan=7><hr></th></tr>")
			Response.Write("<tr><th>Counter</th><th>Product ID</th><th>Match</th><th>Next ID</th><th>Discounted</th><th>Qualifying Items</th><th>Free Items</th></tr>")
		End If	'cblnDebugPromotionManagerAddon
		
		plngFreeItemCounter = 0
		For i = 0 to uBound(paryOrder)
			
			'Reset discounts since they are being recalculated for free product
			paryOrder(i,enOrder_DiscountAmount)(lngIndex) = 0
			
			pstrProductID = WrapString(paryOrder(i,enOrder_ProductID),True)
			'pblnProductMatch = instr(1,pstrFreeProductIDs,pstrProductID) > 0
			pblnProductMatch = isProductMatch(pstrProductID, pstrFreeProductIDs)

			'If cblnDebugPromotionManagerAddon Then Response.Write "Product match " & paryOrder(i,enOrder_ProductID) & ": " & pblnProductMatch & "<br />"
			
			If pblnProductMatch then

				If Not (prsPromo.Fields("ExcludeSaleItems").Value AND paryOrder(i,enOrder_SaleIsActive)) Then
					If (pcursubTotal) >= CDbl(prsPromo.Fields("MinSubTotal").Value) Then
						
						pblnThisItemDiscounted = True
						If CBool(plngMaxItems > 0) Then
							pblnThisItemDiscounted = pblnThisItemDiscounted And CBool(plngFreeItemCounter < plngMaxItems)
						End If
						
						'If cblnDebugPromotionManagerAddon Then debugprint "pblnThisItemDiscounted " & paryOrder(i,enOrder_ProductID),pblnThisItemDiscounted
						'If cblnDebugPromotionManagerAddon Then debugprint "plngFreeItemCounter " & plngMaxItems,plngFreeItemCounter
					
						If pblnThisItemDiscounted Then
							plngFreeItemCounter = plngFreeItemCounter + 1
							pcurProductPrice = paryOrder(i,enOrder_SellPrice)
							if prsPromo.Fields("Percentage").Value then
								pcursubDiscount = Round((CDbl(prsPromo.Fields("Discount").Value)/100) * pcurProductPrice,2)
							else
								pcursubDiscount = Round(CDbl(prsPromo.Fields("Discount").Value),2)
							end if
						Else
							pcursubDiscount = 0
						End If	'pblnThisItemDiscounted
						
						'limit discount by order line item
						If Len(Trim(prsPromo.Fields("MaxAllowableValuePerItem").Value & "")) > 0 Then
							If pcursubDiscount >= CDbl(prsPromo.Fields("MaxAllowableValuePerItem").Value) Then pcursubDiscount = CDbl(prsPromo.Fields("MaxAllowableValuePerItem").Value)
						End If
						
						'limit discount by total order
						If Len(Trim(prsPromo.Fields("MaxAllowableValue").Value & "")) > 0 Then
							If CDbl(pcurDiscountTotal + pcursubDiscount) >= CDbl(prsPromo.Fields("MaxAllowableValue").Value) Then
								pcursubDiscount = CDbl(prsPromo.Fields("MaxAllowableValue").Value) - pcurDiscountTotal
							End If
						End If
						pcurDiscountTotal = pcurDiscountTotal + pcursubDiscount

						paryOrder(i,enOrder_DiscountAmount)(lngIndex) = pcursubDiscount
						pcurDiscount = pcurDiscount + cDbl(pcursubDiscount)
					else
						pcurDiscount = 0
					end if
				End If
			Else
				pblnThisItemDiscounted = False
				plngNextFreeID = plngNextFreeID + 1
			End If	'pblnProductMatch
			
			If cblnDebugPromotionManagerAddon Then
				If pblnThisItemDiscounted Then
					Response.Write("<tr bgcolor=yellow>")
				Else
					Response.Write("<tr>")
				End If
				Response.Write("<td>" & i & "</td><td>" & paryOrder(i,enOrder_ProductID) & "</td><td>" & pblnProductMatch & "</td><td>" & plngNextFreeID & "</td><td>" & pblnThisItemDiscounted & " (" & FormatCurrency(paryOrder(i,enOrder_DiscountAmount)(lngIndex),2) & ")</td><td>" & plngQualifyingItemCounter & "</td><td>" & plngFreeItemCounter & "</td></tr>")
			End If	'cblnDebugPromotionManagerAddon
			
		Next
		
		If cblnDebugPromotionManagerAddon Then
			Response.Write("<br /><table border=1 cellspacing=0 cellpadding=2><tr><th colspan=4>CalcBuy1Get1 Discount - " & prsPromo.Fields("PromoCode").Value & "<br />" & prsPromo.Fields("PromoTitle").Value & "</th></tr>")
			Response.Write("<tr><th>Item</th><th>Product</th><th>Unit Price</th><th>Discount</th></tr>")
			For i = 0 to uBound(paryOrder)
				Response.Write("<font color=black><tr><td>Item " & i & "</td><td>" & paryOrder(i,enOrder_ProductID) & "</td>")
				If isNumeric(paryOrder(i,enOrder_SellPrice)) Then
					Response.Write("<td align=right>" & FormatCurrency(paryOrder(i,enOrder_SellPrice),2) & "</td>")
				Else
					Response.Write("<td align=right>" & paryOrder(i,enOrder_SellPrice) & "</td>")
				End If

				If isNumeric(paryOrder(i,enOrder_DiscountAmount)(lngIndex)) Then
					Response.Write("<td align=right>" & FormatCurrency(paryOrder(i,enOrder_DiscountAmount)(lngIndex),2) & "</td>")
				Else
					Response.Write("<td align=right>" & paryOrder(i,enOrder_DiscountAmount)(lngIndex) & "</td>")
				End If
				Response.Write("</tr></font>")
			Next
			Response.Write("<tr><td colspan=2>&nbsp;</td><td align=right>Total:&nbsp;</td><td align=right><font color=black>" & FormatCurrency(pcurDiscount,2) & "</td></tr></font>")
			Response.Write("</table><hr>")
		End If	'cblnDebugPromotionManagerAddon
		
		aryDiscount(lngIndex, 11) = pblnQualifiesForFreeProduct
	Else
		If plngNumberOfFreeQualifyingProductsInCart >= plngBuyX And Len(pstrFreeProductIDs) = 0 Then
			aryDiscount(lngIndex, 11) = True
		Else
			pcurDiscount = 0
			For i = 0 to uBound(paryOrder)
				paryOrder(i,enOrder_DiscountAmount)(lngIndex) = 0
			Next 'i
			aryDiscount(lngIndex, 11) = False
		End If
	End If	'pblnQualifiesForFreeProduct

	aryDiscount(lngIndex, 1) = Round(pcurDiscount,2)

	CalcBuy1Get1 = Round(pcurDiscount,2)

End Function	'CalcBuy1Get1

'***********************************************************************************************

Private Function isProductMatch(byVal strProduct, byVal strAllowed)

Dim i
Dim paryProductsToCheck
Dim pblnProductMatch
Dim pstrAllowed
Dim pstrToCheck

	pstrToCheck = LCase(strProduct)
	strAllowed = LCase(strAllowed)
	
	If Instr(1, strAllowed, "*") > 0 Then
		paryProductsToCheck = Split(strAllowed, ";")

		For i = 1 To UBound(paryProductsToCheck) - 1
			pstrAllowed = Replace(paryProductsToCheck(i), "*", "")
			pblnProductMatch = CBool(instr(1, pstrToCheck, pstrAllowed) > 0)
			If pblnProductMatch Then Exit For
		Next 'i
	Else
		pblnProductMatch = CBool(instr(1, strAllowed, pstrToCheck) > 0)
	End If

	If cblnDebugPromotionManagerAddon Then
		Response.Write "<fieldset><legend></legend>"
		Response.Write "strAllowed: " & strAllowed & "<br />"
		Response.Write "strProduct: " & strProduct & "<br />"
		Response.Write "pblnProductMatch: " & pblnProductMatch & "<br />"
		Response.Write "</fieldset>"
	End If

	isProductMatch = pblnProductMatch
		
End Function	'isProductMatch

'***********************************************************************************************

Private Function CalcProductDiscount(byVal lngIndex)

dim i
Dim paryDiscounts
dim pblnProductMatch
dim pcurDiscount
Dim pcurDiscountTotal
dim pcurProductPrice
dim pcursubDiscount
Dim pdblSubTotal_Local
dim plngMaxItems
dim plngMaxOrderItems
dim plngqtyToApplyTo
dim plngQuantity
dim plngQuantityAppliedTo
Dim pstrApplicableProductIDs
dim pstrProductID
Dim pstrPrevProductID

'On Error Resume Next

	If cblnDebugPromotionManagerAddon Then
		Response.Write("<fieldset><legend>Checking for CalcProductDiscount . . .</legend><font color=black>")
		Response.Write "pcursubTotal = " & Round(pcursubTotal,2)
	End If

	pcurDiscount = 0
	pcurDiscountTotal = 0
	
	pstrApplicableProductIDs = GetQualifyingProducts

	plngMaxItems = Trim(prsPromo.Fields("productCountLimit").Value & "")
	If Len(plngMaxItems) = 0 Then
		plngMaxItems = 0
	Else
		plngMaxItems = CLng(plngMaxItems)
	End If

	'Now calculate working order subtotal; it may be different due to excluded products
	pdblSubTotal_Local = 0
	For i = 0 to uBound(paryOrder)
		pstrProductID = WrapString(paryOrder(i, enOrder_ProductID), True)
		'pblnProductMatch = instr(1, pstrApplicableProductIDs, pstrProductID) > 0
		pblnProductMatch = isProductMatch(pstrProductID, pstrApplicableProductIDs)
		If pblnProductMatch Then pdblSubTotal_Local = pdblSubTotal_Local + paryOrder(i, enOrder_Quantity) * paryOrder(i, enOrder_SellPrice)
	Next	'i

	plngMaxItems = Trim(prsPromo.Fields("productCountLimit").Value & "")
	plngMaxOrderItems = Trim(prsPromo.Fields("productCountLimit").Value & "")
	
	For i = 0 to uBound(paryOrder)
		
		pstrProductID = WrapString(paryOrder(i, enOrder_ProductID), True)
		'pblnProductMatch = instr(1, pstrApplicableProductIDs, pstrProductID) > 0
		pblnProductMatch = isProductMatch(pstrProductID, pstrApplicableProductIDs)
		If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black>" & pstrProductID & " match: <i>" & pblnProductMatch & "<i></font><br />")

		If pblnProductMatch then

			If pdblSubTotal_Local >= CDbl(prsPromo.Fields("MinSubTotal").Value) Then

				pcurProductPrice = paryOrder(i,enOrder_SellPrice)
				plngQuantity = paryOrder(i,enOrder_Quantity)
				If Len(plngMaxItems) = 0 Then
					plngqtyToApplyTo = plngQuantity
				Else
					If plngQuantity < plngMaxItems Then
						plngqtyToApplyTo = plngQuantity
					Else
						plngqtyToApplyTo = plngMaxItems
					End If
					
					If CDbl(plngQuantityAppliedTo) >= CDbl(plngMaxItems) Then
						plngqtyToApplyTo = 0
					Else
						plngQuantityAppliedTo = plngQuantityAppliedTo + plngqtyToApplyTo
					End If
					
				End If
				
				if prsPromo.Fields("Percentage").Value then
					pcursubDiscount = (CDbl(prsPromo.Fields("Discount").Value)/100) * pcurProductPrice * plngqtyToApplyTo
				else
					pcursubDiscount = CDbl(prsPromo.Fields("Discount").Value) * plngqtyToApplyTo
				end if
				
				'limit discount by order line item
				If Len(Trim(prsPromo.Fields("MaxAllowableValuePerItem").Value & "")) > 0 Then
					If pcursubDiscount >= CDbl(prsPromo.Fields("MaxAllowableValuePerItem").Value) Then pcursubDiscount = CDbl(prsPromo.Fields("MaxAllowableValuePerItem").Value)
				End If
				
				'limit discount by total order
				If Len(Trim(prsPromo.Fields("MaxAllowableValue").Value & "")) > 0 Then
					If CDbl(pcurDiscountTotal + pcursubDiscount) >= CDbl(prsPromo.Fields("MaxAllowableValue").Value) Then
						pcursubDiscount = CDbl(prsPromo.Fields("MaxAllowableValue").Value) - pcurDiscountTotal
					End If
				End If
				pcurDiscountTotal = pcurDiscountTotal + pcursubDiscount

				paryDiscounts = paryOrder(i,enOrder_DiscountAmount)
				paryDiscounts(lngIndex) = pcursubDiscount
				pcurDiscount = pcurDiscount + cDbl(pcursubDiscount)
				paryOrder(i,enOrder_DiscountAmount) = paryDiscounts
			else
				pcurDiscount = 0
			end if

		End If	'pblnProductMatch
		
		If cblnDebugPromotionManagerAddon Then
		End If	'cblnDebugPromotionManagerAddon
		
	Next	'i
	
	If cblnDebugPromotionManagerAddon Then
		Response.Write("<br /><table border=1 cellspacing=0 cellpadding=2><tr><th colspan=4>Product Level Discount - " & prsPromo.Fields("PromoCode").Value & "<br />" & prsPromo.Fields("PromoTitle").Value & "</th></tr>")
		Response.Write("<tr><th>Item</th><th>Product</th><th>Unit Price</th><th>Discount</th></tr>")
		For i = 0 to uBound(paryOrder)
			Response.Write("<font color=black><tr><td>Item " & i & "</td><td>" & paryOrder(i,enOrder_ProductID) & "</td><td align=right>" & FormatCurrency(paryOrder(i,enOrder_SellPrice),2) & "</td><td align=right>" & FormatCurrency(paryOrder(i,enOrder_DiscountAmount)(lngIndex),2) & "</td></tr></font>")
		Next
		Response.Write("</table>")
		If pdblSubTotal_Local < CDbl(prsPromo.Fields("MinSubTotal").Value) And pcursubTotal >= CDbl(prsPromo.Fields("MinSubTotal").Value) Then Response.Write "<h4><font color=red>No discout applied due to excluded products</font></h4>"
		Response.Write("Min. Sub Total:&nbsp;" & prsPromo.Fields("MinSubTotal").Value & "<br />")
		Response.Write("Order Sub Total:&nbsp;" & FormatCurrency(pcursubTotal,2) & "<br />")
		Response.Write("Order Sub Total (minus exclusions):&nbsp;" & FormatCurrency(pdblSubTotal_Local,2) & "<br />")
		Response.Write("Max. Discount:&nbsp;" & prsPromo.Fields("MaxAllowableValue").Value & "<br />")
		Response.Write("Discount Total:&nbsp;" & FormatCurrency(pcurDiscount,2) & "<br />")
		Response.Write "</font></fieldset>"
	End If	'cblnDebugPromotionManagerAddon

	CalcProductDiscount = Round(pcurDiscount,2)

End Function	'CalcProductDiscount

'***********************************************************************************************

Private Function QualifiesForFreeShipping(lngIndex)

Dim pstrApplicableProductIDs
dim i
dim pblnProductMatch
dim pstrProductID
Dim pblnQualifiesForFreeShipping

'On Error Resume Next

	If cblnDebugPromotionManagerAddon Then Response.Write("<br /><font color=black>Checking if QualifiesForFreeShipping . . .</font><br />")

	pblnQualifiesForFreeShipping = False
	pstrApplicableProductIDs = GetQualifyingProducts

	For i = 0 to uBound(paryOrder)
		pstrProductID = WrapString(paryOrder(i,enOrder_ProductID),True)
		
		'pblnProductMatch = instr(1,pstrApplicableProductIDs,pstrProductID) > 0
		pblnProductMatch = isProductMatch(pstrProductID, pstrApplicableProductIDs)

		If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black>" & paryOrder(i,enOrder_ProductID) & " match: <i>" & pblnProductMatch & "<i></font><br />")
		If pblnProductMatch then
			If Not (prsPromo.Fields("ExcludeSaleItems").Value AND paryOrder(i,enOrder_SaleIsActive)) Then
				If (pcursubTotal) >= CDbl(prsPromo.Fields("MinSubTotal").Value) Then
					pblnQualifiesForFreeShipping = True
					Exit For
				end if
			End If
		End If	'pblnProductMatch
	Next
	
	If cblnDebugPromotionManagerAddon Then
		Response.Write("<br /><table border=1 cellspacing=0 cellpadding=2><tr><th colspan=2>Free Shipping - " & prsPromo.Fields("PromoCode").Value & "<br />" & prsPromo.Fields("PromoTitle").Value & "</th></tr>")
		Response.Write("<tr><td>Free Shipping</td><td>" & pblnQualifiesForFreeShipping & "</td></tr>")
		Response.Write("</table><hr>")
	End If	'cblnDebugPromotionManagerAddon
	
	QualifiesForFreeShipping = pblnQualifiesForFreeShipping

End Function	'QualifiesForFreeShipping

'***********************************************************************************************

Private Function GetQualifyingProducts()

dim i
Dim pstrPrevProdID
Dim pstrApplicableProductIDs
Dim pstrSQL_ProductID
Dim pstrSQL_ProductIDExclusion
Dim pstrSQL_Category
Dim pstrSQL_CategoryExclusion
Dim pstrSQL_Manufacturer
Dim pstrSQL_ManufacturerExclusion
Dim pstrSQL_Vendor
Dim pstrSQL_VendorExclusion
Dim pstrSQL
Dim pstrSQL_Where
Dim pstrSQL_Where_Include
Dim pstrSQL_Where_Exclude
Dim pobjRS

'On Error Resume Next

	If cblnDebugPromotionManagerAddon Then
		Response.Write "<table border=1 cellpadding=2 cellspacing=0>"
		Response.Write("<tr><td><font color=black>Getting Qualifying Products . . .</font></td></tr>")
	End If

	pstrSQL_ProductID = MakeProductSearchSQL("sfProducts.prodID", prsPromo.Fields("ProductID").Value, "=", "'", "")
	If cblnSF5AE Then
		pstrSQL_Category = MakeProductSearchSQL("sfSub_Categories.subcatCategoryId", prsPromo.Fields("Category").Value, "=", "", "")
	Else
		pstrSQL_Category = MakeProductSearchSQL("prodCategoryId", prsPromo.Fields("Category").Value, "=", "", "")
	End If
	
	pstrSQL_Manufacturer = MakeProductSearchSQL("prodManufacturerId", prsPromo.Fields("Manufacturer").Value, "=", "", "")
	pstrSQL_Vendor = MakeProductSearchSQL("prodVendorId", prsPromo.Fields("Vendor").Value, "=", "", "")
	pstrSQL_ProductIDExclusion = MakeProductSearchSQL("sfProducts.prodID", prsPromo.Fields("ProductIDExclusion").Value, "<>", "'", "")
	pstrSQL_CategoryExclusion = MakeProductSearchSQL("prodCategoryId", prsPromo.Fields("CategoryExclusion").Value, "<>", "", "")
	pstrSQL_ManufacturerExclusion = MakeProductSearchSQL("prodManufacturerId", prsPromo.Fields("ManufacturerExclusion").Value, "<>", "", "")
	pstrSQL_VendorExclusion = MakeProductSearchSQL("prodVendorId", prsPromo.Fields("VendorExclusion").Value, "<>", "", "")

	If Len(pstrSQL_ProductID) > 0 Then
		If Len(pstrSQL_Where_Include) > 0 Then
			pstrSQL_Where_Include = pstrSQL_Where_Include & " OR " & pstrSQL_ProductID
		Else
			pstrSQL_Where_Include = pstrSQL_ProductID
		End If
	End If
	
	If Len(pstrSQL_Category) > 0 Then
		If Len(pstrSQL_Where_Include) > 0 Then
			pstrSQL_Where_Include = pstrSQL_Where_Include & " OR " & pstrSQL_Category
		Else
			pstrSQL_Where_Include = pstrSQL_Category
		End If
	End If
	
	If Len(pstrSQL_Manufacturer) > 0 Then
		If Len(pstrSQL_Where_Include) > 0 Then
			pstrSQL_Where_Include = pstrSQL_Where_Include & " OR " & pstrSQL_Manufacturer
		Else
			pstrSQL_Where_Include = pstrSQL_Manufacturer
		End If
	End If
	
	If Len(pstrSQL_Vendor) > 0 Then
		If Len(pstrSQL_Where_Include) > 0 Then
			pstrSQL_Where_Include = pstrSQL_Where_Include & " OR " & pstrSQL_Vendor
		Else
			pstrSQL_Where_Include = pstrSQL_Vendor
		End If
	End If
	'Response.Write("<font color=black>pstrSQL_Where_Include: <i>" & pstrSQL_Where_Include & "<i></font><br />")
	
	If Len(pstrSQL_ProductIDExclusion) > 0 Then
		If Len(pstrSQL_Where_Exclude) > 0 Then
			pstrSQL_Where_Exclude = pstrSQL_Where_Exclude & " And " & pstrSQL_ProductIDExclusion
		Else
			pstrSQL_Where_Exclude = pstrSQL_ProductIDExclusion
		End If
	End If
	
	If Len(pstrSQL_CategoryExclusion) > 0 Then
		If Len(pstrSQL_Where_Exclude) > 0 Then
			pstrSQL_Where_Exclude = pstrSQL_Where_Exclude & " OR " & pstrSQL_CategoryExclusion
		Else
			pstrSQL_Where_Exclude = pstrSQL_CategoryExclusion
		End If
	End If
	
	If Len(pstrSQL_ManufacturerExclusion) > 0 Then
		If Len(pstrSQL_Where_Exclude) > 0 Then
			pstrSQL_Where_Exclude = pstrSQL_Where_Exclude & " OR " & pstrSQL_ManufacturerExclusion
		Else
			pstrSQL_Where_Exclude = pstrSQL_ManufacturerExclusion
		End If
	End If
	
	If Len(pstrSQL_VendorExclusion) > 0 Then
		If Len(pstrSQL_Where_Exclude) > 0 Then
			pstrSQL_Where_Exclude = pstrSQL_Where_Exclude & " OR " & pstrSQL_VendorExclusion
		Else
			pstrSQL_Where_Exclude = pstrSQL_VendorExclusion
		End If
	End If
	'Response.Write("<font color=black>pstrSQL_Where_Exclude: <i>" & pstrSQL_Where_Exclude & "<i></font><br />")
	
	pstrSQL_Where = pstrSQL_Where_Include
	If Len(pstrSQL_Where_Exclude) > 0 Then
		If Len(pstrSQL_Where) > 0 Then
			pstrSQL_Where = pstrSQL_Where & " AND " & pstrSQL_Where_Exclude
		Else
			pstrSQL_Where = pstrSQL_Where_Exclude
		End If
	End If
	If Len(pstrSQL_Where) > 0 Then
		pstrSQL_Where = " WHERE ((prodEnabledIsActive=1) AND " & pstrSQL_Where & ")"
	Else
		pstrSQL_Where = " WHERE (prodEnabledIsActive=1)"
	End If
	'Response.Write("<font color=black>pstrSQL_Where: <i>" & pstrSQL_Where & "<i></font><br />")
	
	If cblnSF5AE Then
		pstrSQL = "SELECT sfProducts.prodID" _
				& " FROM sfSub_Categories INNER JOIN (sfProducts RIGHT JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) ON sfSub_Categories.subcatID = sfSubCatDetail.subcatCategoryId " _
				& pstrSQL_Where _
				& " Order By sfProducts.prodID"
	Else
		pstrSQL = "SELECT sfProducts.prodID FROM sfProducts" _
				& pstrSQL_Where _
				& " Order By sfProducts.prodID"
	End If
	
	Set pobjRS = CreateObject("ADODB.RecordSet")
	pobjRS.Open pstrSQL, cnn, 0, 1, &H0001
	If cblnDebugPromotionManagerAddon Then
		If pobjRS.EOF Then
			Response.Write("<tr><td><font color=black>No Applicable Products Found</font></td></tr>")
			Response.Write("<tr><td><font color=black>pstrSQL: " & pstrSQL & "</font></td></tr>")
		Else
			Response.Write("<tr><td><font color=black>Applicable Products Found: Yes</font></td></tr>")
		End If
	End If
	Do While Not pobjRS.EOF
		If pstrPrevProdID <> Trim(pobjRS.Fields("prodID").Value) Then
			pstrPrevProdID = Trim(pobjRS.Fields("prodID").Value)
			pstrApplicableProductIDs = pstrApplicableProductIDs & cstrDelimiter & pstrPrevProdID
		End If
		pobjRS.MoveNext
	Loop
	pstrApplicableProductIDs = pstrApplicableProductIDs & cstrDelimiter
	pobjRS.Close
	Set pobjRS = Nothing

	If cblnDebugPromotionManagerAddon Then
		Response.Write("<tr><td><font color=black size=-1>ProductID: <i>" & prsPromo.Fields("ProductID").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>ProductIDExclusion: <i>" & prsPromo.Fields("ProductIDExclusion").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>Category: <i>" & prsPromo.Fields("Category").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>CategoryExclusion: <i>" & prsPromo.Fields("CategoryExclusion").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>Manufacturer: <i>" & prsPromo.Fields("Manufacturer").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>ManufacturerExclusion: <i>" & prsPromo.Fields("ManufacturerExclusion").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>Vendor: <i>" & prsPromo.Fields("Vendor").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>VendorExclusion: <i>" & prsPromo.Fields("VendorExclusion").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>FreeProductID: <i>" & prsPromo.Fields("FreeProductID").Value & "<i></font></td></tr>")
		Response.Write("<tr><td><font color=black size=-1>Applicable Products: " & Replace(pstrApplicableProductIDs & "", ";", " ") & "</font></td></tr>")
		Response.Write("</table>")
	End If	'cblnDebugPromotionManagerAddon

	GetQualifyingProducts = pstrApplicableProductIDs

End Function	'GetQualifyingProducts

'***********************************************************************************************

Private Function ConvertBoolean(vntSource)

On Error Resume Next

	If Len(vntSource & "") = 0 Then
		ConvertBoolean = False
	Else
		ConvertBoolean = CBool(vntSource)
	End If

End Function

End Class	'clsPromotion

'***********************************************************************************************
'***********************************************************************************************

Function getFreeProducts(byVal strProductIDs, byRef lngNumFreeProductsAvailable)

Dim pstrSQL
Dim pobjRS
Dim paryProductIDs
Dim i

	paryProductIDs = Split(Trim(strProductIDs & ""), ";")
	For i = 0 To UBound(paryProductIDs)
		If Len(paryProductIDs(i)) > 0 Then
			If Len(pstrSQL) > 0 Then
				pstrSQL = pstrSQL & " OR prodID='" & Replace(paryProductIDs(i), "'", "''") & "'"
			Else
				pstrSQL = "prodID='" & Replace(paryProductIDs(i), "'", "''") & "'"
			End If
		End If
	Next 'i

	If Len(pstrSQL) > 0 Then
		pstrSQL = "Select prodID, prodName From sfProducts Where " & pstrSQL
		set pobjRS = CreateObject("adodb.recordset")
		with pobjRS
			.ActiveConnection = cnn
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			
			If cblnDebugPromotionManagerAddon Then
				Response.Write("<font color=black size=-1>getFreeProducts<i></font><br />")
				Response.Write("<font color=black size=-1>pstrSQL: <i>" & pstrSQL & "<i></font><br />")
				Response.Write("<font color=black size=-1>.EOF: <i>" & .EOF & "<i></font><br />")
			End If	'cblnDebugPromotionManagerAddon

			If Not .EOF Then
				lngNumFreeProductsAvailable = .RecordCount

				ReDim paryTemp(lngNumFreeProductsAvailable - 1, 1)
				For i = 0 To lngNumFreeProductsAvailable - 1
					paryTemp(i,0) = Trim(.Fields("prodID").Value & "")
					paryTemp(i,1) = Trim(.Fields("prodName").Value & "")
					If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black size=-1>" & i & ": " & paryTemp(i,0) & " - " & paryTemp(i,1) & "<i></font><br />")
					.MoveNext
				Next 'i
				getFreeProducts = paryTemp
			Else
				lngNumFreeProductsAvailable = 0
			End If
			.Close

			If cblnDebugPromotionManagerAddon Then
				Response.Write("<font color=black size=-1>getFreeProducts<i></font><br />")
				Response.Write("<font color=black size=-1>pstrSQL: <i>" & pstrSQL & "<i></font><br />")
				Response.Write("<font color=black size=-1>lngNumFreeProductsAvailable: <i>" & lngNumFreeProductsAvailable & "<i></font><br />")
			End If	'cblnDebugPromotionManagerAddon

		End With
		set pobjRS = Nothing
	End If

End Function	'getFreeProducts

'***********************************************************************************************

Function wrapFreeProducts(strProductIDs)

Dim pstrTempSelectedProducts

	pstrTempSelectedProducts = Trim(strProductIDs)
	pstrTempSelectedProducts = Replace(pstrTempSelectedProducts, ", ", "|")
	If Len(pstrTempSelectedProducts) > 0 Then pstrTempSelectedProducts = "|" & pstrTempSelectedProducts & "|"
	
	wrapFreeProducts = pstrTempSelectedProducts

End Function	'wrapFreeProducts

'***********************************************************************************************

Sub OfferFreeGift(aryDiscountSummary)

Dim paryFreeProducts
Dim plngNumFreeProductsAvailable
Dim i
Dim pstrSelectedFreeProducts
Dim pstrChecked
Dim pblnProductSelected
Dim pstrInputType
Dim pstrTempSelectedProducts

	If cblnDebugPromotionManagerAddon Then
		Response.Write("<br /><font color=black>Offering Free Gift . . .</font><br />")
		Response.Write("<font color=black>isArray(aryDiscountSummary): <i>" & isArray(aryDiscountSummary) & "<i></font><br />")
		Response.Write("<font color=black>aryDiscountSummary(enDiscountFreeProduct): <i>" & aryDiscountSummary(enDiscountFreeProduct) & "<i></font><br />")
	End If
	
	If isArray(aryDiscountSummary) Then paryFreeProducts = getFreeProducts(aryDiscountSummary(enDiscountFreeProduct), plngNumFreeProductsAvailable)

	If cblnDebugPromotionManagerAddon Then
		Response.Write("<font color=black>plngNumFreeProductsAvailable: <i>" & plngNumFreeProductsAvailable & "<i></font><br />")
		Response.Write("<font color=black>cbytFreeGiftOfferLocation: <i>" & cbytFreeGiftOfferLocation & "<i></font><br />")
	End If
	
	If plngNumFreeProductsAvailable > 0 Then
		pblnProductSelected = False
		pstrSelectedFreeProducts = Trim(Request.Form("FreeProduct"))
		If Len(pstrSelectedFreeProducts) > 0 Then
			Call setVisitorPreference("visitorSelectedFreeProducts", pstrSelectedFreeProducts, True)
		Else
			pstrSelectedFreeProducts = visitorSelectedFreeProducts
		End If
		pstrTempSelectedProducts = wrapFreeProducts(pstrSelectedFreeProducts)
		
		If cblnDebugPromotionManagerAddon Then
			Response.Write("<font color=black>pstrTempSelectedProducts: <i>" & pstrTempSelectedProducts & "<i></font><br />")
			Response.Write("<font color=black>SelectedFreeProducts: <i>" & visitorSelectedFreeProducts & "<i></font><br />")
		End If

		If cbytFreeGiftOfferLocation = 1 Then Response.Write "<form name=frmFreeProduct id=frmFreeProduct action='order.asp' method=post>"
%>
<script language="javascript" type="text/javascript">

function checkFreeGiftSelection(theItem)
{
var theForm = theItem.form;
var plngCount = 0;
var cnumFreeItems = 2;
var i;

	if (theForm.FreeProduct.checked==undefined)
	{
		for (i=0; i < theForm.FreeProduct.length;i++)
		{
			if (theForm.FreeProduct[i].checked){plngCount ++;}
	//	document.frmData.custID[i].checked = blnCheck;
		}
		
		if (plngCount > cnumFreeItems)
		{
			alert("You have already selected your free items. \n \n If you wish to change your selection please deselect an item first.");
			for (i=0; i < theForm.FreeProduct.length;i++)
			{
				if (theForm.FreeProduct[i].value == theItem.value){theForm.FreeProduct[i].checked = false;}
			}
		}
	}

}
</script>

		<table border="1" cellpadding="0" cellspacing="0" id="tblFreeGiftBorder">
		<tr><td>
		<table border="0" cellpadding="2" cellspacing="0" width="100%" id="tblFreeGiftSelction">
		  <tr><th class="tdTopBanner" align="center"><%= aryDiscountSummary(enPromoTitle) %></th></tr>
<%
		For i = 0 To plngNumFreeProductsAvailable - 1
			If InStr(1, pstrTempSelectedProducts, "|" & Trim(paryFreeProducts(i, 0)) & "|") > 0 Then
			'If pstrSelectedFreeProducts = paryFreeProducts(i, 0) Then
				pstrChecked = "checked"
				pblnProductSelected = True
			Else
				pstrChecked = ""
			End If
			
			If cblnAutoCheckFreeGift Then
				If Not pblnProductSelected And (i = plngNumFreeProductsAvailable - 1) Then
					pstrChecked = "checked"
					pblnProductSelected = True
					Call setVisitorPreference("visitorSelectedFreeProducts", paryFreeProducts(i, 0), True)
				End If
			End If
			
			If (aryDiscountSummary(enofferFreeGiftAutomatically) > 1) Or (plngNumFreeProductsAvailable = 1) Then
				pstrInputType = "checkbox"
			Else
				pstrInputType = "radio"
			End If

			If plngNumFreeProductsAvailable = 1 And pstrChecked = "checked" Then
%>
		  <tr><td><%= paryFreeProducts(i, 1) %></td></tr>
<%
			Else
%>
		  <tr>
		    <td align="left">
		      <input type="<%= pstrInputType %>" name="FreeProduct" id="FreeProduct<%= i %>" value="<%= paryFreeProducts(i, 0) %>" <%= pstrChecked %> onclick="return checkFreeGiftSelection(this);">
		      <label for="FreeProduct<%= i %>"><%= paryFreeProducts(i, 1) %></label>
		    </td>
		  </tr>
<%
			End If
		Next 'i
		
		If False Then	'this is just here if you want to add a <HR>
%>
		  <tr><td><hr width="90%>" </td></tr>
<%
		End If
		
		If plngNumFreeProductsAvailable = 1 And pstrChecked = "checked" Then
		Else
		If pblnProductSelected Then
%>
		  <tr><td><input type="submit" name="btnFreeProductSubmit" id="btnFreeProductSubmit0" value="Update your free product selection"></td></tr>
<%
		Else
			If aryDiscountSummary(enofferFreeGiftAutomatically) > 1 Then
%>
		  <tr><td><input type="submit" name="btnFreeProductSubmit" id="btnFreeProductSubmit1" value="Select your <%= aryDiscountSummary(enofferFreeGiftAutomatically) %> free products"></td></tr>
<%
			Else
%>
		  <tr><td><input type="submit" name="btnFreeProductSubmit" id="btnFreeProductSubmit2" value="Select your free product"></td></tr>
<%
			End If
		End If 'pblnProductSelected
		End If	'plngNumFreeProductsAvailable = 1 And pstrChecked = "checked"
%>
		</table>
		</td></tr>
		</table>
<%
		If cbytFreeGiftOfferLocation = 1 Then Response.Write "</form>"

	End If	'isArray(paryFreeProducts)

End Sub	'OfferFreeGift

'***********************************************************************************************

Sub showFreeGift(aryDiscountSummary)

Dim paryFreeProducts
Dim i
Dim pstrSelectedFreeProducts
Dim pstrChecked
Dim pblnProductSelected
Dim plngNumFreeProductsAvailable
Dim pstrTempSelectedProducts
Dim plngNumSelectedProducts

	If isArray(aryDiscountSummary) Then paryFreeProducts = getFreeProducts(aryDiscountSummary(enDiscountFreeProduct), plngNumFreeProductsAvailable)
	If cblnDebugPromotionManagerAddon Then Response.Write "<h4>Sub showFreeGift</h4>"
	If cblnDebugPromotionManagerAddon Then Response.Write "- isArray(paryFreeProducts): " & isArray(paryFreeProducts) & "<br />"
	If isArray(paryFreeProducts) Then
		pblnProductSelected = False
		pstrSelectedFreeProducts = visitorSelectedFreeProducts
		If cblnDebugPromotionManagerAddon Then Response.Write "- pstrSelectedFreeProducts: " & pstrSelectedFreeProducts & "<br />"
		If Len(pstrSelectedFreeProducts) = 0 Then Exit Sub
		
		pstrTempSelectedProducts = wrapFreeProducts(pstrSelectedFreeProducts)
		plngNumSelectedProducts = 0
		If cblnDebugPromotionManagerAddon Then Response.Write "- pstrTempSelectedProducts: " & pstrTempSelectedProducts & "<br />"
		If cblnDebugPromotionManagerAddon Then Response.Write "- plngNumFreeProductsAvailable: " & plngNumFreeProductsAvailable & "<br />"
		For i = 0 To plngNumFreeProductsAvailable - 1
			If cblnDebugPromotionManagerAddon Then Response.Write "- paryFreeProducts(0, " & i & "): " & paryFreeProducts(i, 0) & "<br />"
			If cblnDebugPromotionManagerAddon Then Response.Write "- InStr: " & InStr(1, pstrTempSelectedProducts, "|" & Trim(paryFreeProducts(i, 0)) & "|") & "<br />"
			If InStr(1, pstrTempSelectedProducts, "|" & Trim(paryFreeProducts(i, 0)) & "|") > 0 Then
			'If pstrSelectedFreeProducts = paryFreeProducts(i, 0) Then
				If cblnDebugPromotionManagerAddon Then Response.Write "- paryFreeProducts(1, " & i & "): " & paryFreeProducts(i, 1) & "<br />"
				If Len(pstrChecked) = 0 Then
					pstrChecked = paryFreeProducts(i, 1)
				Else
					pstrChecked = pstrChecked & "<br />" & paryFreeProducts(i, 1)
				End If
				plngNumSelectedProducts = plngNumSelectedProducts + 1
			End If
		Next 'i
		
		If cblnDebugPromotionManagerAddon Then Response.Write "- pstrChecked: " & pstrChecked & "<br />"
		If Len(pstrChecked) > 0 Then
%>
		<table border="1" cellpadding="0" cellspacing="0" id="tblFreeProductBorder">
		<tr><td>
		<table border="0" cellpadding="2" cellspacing="0" width="100%" id="tblFreeProduct">
		  <tr><th class="tdTopBanner" align="center"><%= aryDiscountSummary(enPromoTitle) %></th></tr>
		  <% If plngNumSelectedProducts = 1 Then %>
		  <tr><th>You have selected the following free gift:</th></tr>
		  <% Else %>
		  <tr><th>You have selected the following free gifts:</th></tr>
		  <% End If %>
		  <tr><td><hr width="90%>" </td></tr>
		  <tr><td align="center"><%= pstrChecked %></td></tr>
		  <tr><td><hr width="90%>" </td></tr>
		</table>
		</td></tr>
		</table>
<%
		End If	'Len(pstrChecked) > 0 Then
	End If	'isArray(paryFreeProducts)

End Sub	'showFreeGift

'***********************************************************************************************

Sub setEmail_PromotionManager(strEmailConfirmationBody, aryDiscountSummary)

Dim paryFreeProducts
Dim i,j
Dim pstrSelectedFreeProducts
Dim pstrChecked
Dim pblnProductSelected
Dim plngNumFreeProductsAvailable
Dim pstrTempSelectedProducts
Dim plngNumSelectedProducts

	If isArray(aryDiscountSummary) Then
		For i = 0 to uBound(aryDiscountSummary)
			If aryDiscountSummary(i)(enDiscountUse) Then
				paryFreeProducts = getFreeProducts(aryDiscountSummary(i)(enDiscountFreeProduct), plngNumFreeProductsAvailable)
				If isArray(paryFreeProducts) Then
					pstrSelectedFreeProducts = visitorSelectedFreeProducts
					If Len(pstrSelectedFreeProducts) = 0 Then Exit Sub
					
					pstrTempSelectedProducts = wrapFreeProducts(pstrSelectedFreeProducts)
					plngNumSelectedProducts = 0
					For j = 0 To plngNumFreeProductsAvailable - 1
						If InStr(1, pstrTempSelectedProducts, "|" & Trim(paryFreeProducts(j, 0)) & "|") > 0 Then
							If plngNumSelectedProducts = 0 Then
								strEmailConfirmationBody = strEmailConfirmationBody & vbcrlf & "Promotion: " & aryDiscountSummary(i)(enPromoTitle) & vbcrlf
								strEmailConfirmationBody = strEmailConfirmationBody & "You have selected the following free gift(s):" & vbcrlf
							End If
							strEmailConfirmationBody = strEmailConfirmationBody & paryFreeProducts(j, 1) & vbcrlf
							plngNumSelectedProducts = plngNumSelectedProducts + 1
						End If
					Next 'j

					If plngNumSelectedProducts > 0 Then strEmailConfirmationBody = strEmailConfirmationBody & vbcrlf
				End If	'isArray(paryFreeProducts)
			End If	'aryDiscountSummary(i)(enDiscountUse)
		Next
	End If

End Sub	'setEmail_PromotionManager

'***********************************************************************************************

Sub saveFreeGift(byVal lngOrderID)

Dim paryFreeProducts
Dim pobjRS
Dim pstrSelectedFreeProducts
Dim i,j
Dim pstrSQL
Dim plngNumFreeProductsAvailable
Dim pstrTempSelectedProducts

	If cblnDebugPromotionManagerAddon Then
		Response.Write "<fieldset><legend>saveFreeGift</legend>"
		Response.Write("lngOrderID: " & lngOrderID & "<br />")
		Response.Write("isArray(maryDiscountSummary): " & isArray(maryDiscountSummary) & "<br />")
		Response.Write("visitorSelectedFreeProducts: " & visitorSelectedFreeProducts & "<br />")
	End If

	If isArray(maryDiscountSummary) Then
		For i = 0 to uBound(maryDiscountSummary)
			If maryDiscountSummary(i)(enDiscountUse) Then
				paryFreeProducts = getFreeProducts(maryDiscountSummary(i)(enDiscountFreeProduct), plngNumFreeProductsAvailable)
				If isArray(paryFreeProducts) Then
					pstrSelectedFreeProducts = visitorSelectedFreeProducts
					If Len(pstrSelectedFreeProducts) = 0 Then Exit Sub
					
					pstrTempSelectedProducts = wrapFreeProducts(pstrSelectedFreeProducts)

					For j = 0 To plngNumFreeProductsAvailable - 1
						If InStr(1, pstrTempSelectedProducts, "|" & Trim(paryFreeProducts(j, 0)) & "|") > 0 Then
							If False Then
								pstrSQL = "Insert Into sfTmpOrderDetails (odrdttmpQuantity, odrdttmpProductID, odrdttmpSessionID, odrdttmpHttpReferer, odrdttmpShipping) " _
										& " Values(1, '" & Replace(paryFreeProducts(j, 0), "'", "''") & "', " & SessionID & ", '', 0)"
								debugprint "pstrSQL",pstrSQL
								cnn.Execute pstrSQL,,128
							Else
								Redim paryProduct(6)
								'debugprint "get6ProdValues",paryFreeProducts(j, 0)
								'paryProduct = get6ProdValues(paryFreeProducts(j, 0))			

								Set pobjRS = CreateObject("ADODB.RecordSet")
								With pobjRS
									.CursorLocation = adUseClient
									.Open "sfOrderDetails Order By odrdtID", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
								
									.AddNew
									.Fields("odrdtOrderId").Value = Trim(lngOrderID)
									.Fields("odrdtQuantity").Value = 1
									.Fields("odrdtSubTotal").Value = "0"
									.Fields("odrdtCategory").Value = ""
									.Fields("odrdtManufacturer").Value = ""
									.Fields("odrdtVendor").Value = ""
									.Fields("odrdtProductName").Value = Trim(paryFreeProducts(j, 1))
									.Fields("odrdtPrice").Value = "0"
									.Fields("odrdtProductId").Value = paryFreeProducts(j, 0)
									.Update
									
									If cblnDebugPromotionManagerAddon Then
										Response.Write("j: " & j & "<br />")
										Response.Write("Free Product " & paryFreeProducts(j, 0) & ": " & Trim(paryFreeProducts(j, 1)) & "<br />")
										Response.Write("odrdtID: " & .Fields("odrdtID").Value & "<br />")
									End If
									
									'bookMark = .AbsolutePosition 
									'.Requery 
									'.AbsolutePosition = bookMark
									
									If Application("AppName") = "StoreFrontAE" Then
										pstrSQL = "Insert Into sfOrderDetailsAE (odrdtAEID, odrdtGiftWrapPrice, odrdtGiftWrapQTY, odrdtBackOrderQTY, odrdtAttDetailID) " _
												& " Values(" & .Fields("odrdtID") & ", '0', 0, 0, '0')"
										cnn.Execute pstrSQL,,128
									End IF
								End With
								Set pobjRS = Nothing
								
							End If
						End If
					Next 'j
					
				End If	'isArray(paryFreeProducts)
			End If	'maryDiscountSummary(i)(enDiscountUse)
		Next	'i
	End If	'isArray(maryDiscountSummary)

	If cblnDebugPromotionManagerAddon Then
		Response.Write("</fieldset>")
	End If

End Sub	'saveFreeGift

'***********************************************************************************************

Sub ShowOrderDiscounts()

Dim i,j
Dim pblnFreeProductPromotion
Dim pblnFreeShipping
Dim pblnFound
Dim pstrOut

	pblnFreeProductPromotion = False
	pblnFreeShipping = False
	
	pblnFreeProductPromotion = False
	If isOrderPage Then
		If cbytFreeGiftOfferLocation = 1 Then
			If isArray(maryDiscountSummary) Then
				pblnFound = False
				For i = 0 to uBound(maryDiscountSummary)
					pblnFreeProductPromotion = pblnFreeProductPromotion OR (CBool(maryDiscountSummary(i)(enDiscountUse)) And maryDiscountSummary(i)(enofferFreeGiftAutomatically) And CBool(Len(maryDiscountSummary(i)(enDiscountFreeProduct)) > 0))
					pblnFreeShipping = pblnFreeShipping OR (maryDiscountSummary(i)(enDiscountUse) And Len(maryDiscountSummary(i)(enDiscountFreeShipping)) > 0)
					
					If cblnDebugPromotionManagerAddon Then
						Response.Write "<b>Promotion " & i + 1 & " - " & maryDiscountSummary(i)(0) & "</b><br />"
						Response.Write "Use: " & maryDiscountSummary(i)(enDiscountUse) & "<br />"
						Response.Write "FreeProduct: " & maryDiscountSummary(i)(enDiscountFreeProduct) & "<br />"
						Response.Write "offer Free Gift Automatically: " & maryDiscountSummary(i)(enofferFreeGiftAutomatically) & "<br />"
					End If
					
					If pblnFreeProductPromotion Then
						Call OfferFreeGift(maryDiscountSummary(i))
					End If
				Next
				If cblnDebugPromotionManagerAddon Then debugprint "pblnFreeProductPromotion",pblnFreeProductPromotion
			End If
		End If	'cbytFreeGiftOfferLocation = 1
	Else

		pblnFreeProductPromotion = False
		
		If isArray(maryDiscountSummary) Then
		
			'This section shows the selected free gift
			For i = 0 to uBound(maryDiscountSummary)
				If UBound(maryDiscountSummary(i)) >= 8 Then
					pblnFreeProductPromotion = pblnFreeProductPromotion OR (maryDiscountSummary(i)(enDiscountUse) And maryDiscountSummary(i)(enofferFreeGiftAutomatically) And Len(maryDiscountSummary(i)(enDiscountFreeProduct)) > 0)
				Else
					pblnFreeProductPromotion = pblnFreeProductPromotion OR (maryDiscountSummary(i)(enDiscountUse) And maryDiscountSummary(i)(enofferFreeGiftAutomatically))
				End If
				pblnFreeShipping = pblnFreeShipping OR (maryDiscountSummary(i)(enDiscountUse) And Len(maryDiscountSummary(i)(enDiscountFreeShipping)) > 0)
				If pblnFreeProductPromotion Then showFreeGift maryDiscountSummary(i)
			Next
		
			'Key to maryDiscountSummary
			'0 - code
			'1 - discount
			'2 - combineable
			'3 - sTaxable
			'4 - cTaxable
			'5 - use
			'6 - Title
			'7 - offerFreeGiftAutomatically
			'8 - FreeProductID
			'9 - MinSubTotal
			'10 - FreeShipping Code | FreeShipping Limit
			For i = 0 to uBound(maryDiscountSummary)
				If maryDiscountSummary(i)(enDiscountUse) Then
					pblnFreeProductPromotion = (maryDiscountSummary(i)(enDiscountUse) And maryDiscountSummary(i)(enofferFreeGiftAutomatically) And Len(maryDiscountSummary(i)(enDiscountFreeProduct)) > 0)
					If pblnFreeProductPromotion Then
						'Do nothing as this was done above
						'showFreeGift maryDiscountSummary(i)(enDiscountFreeProduct)
					ElseIf mcurDiscountAmount = 0 And Len(maryDiscountSummary(i)(enDiscountFreeProduct)) = 0 And Not pblnFreeShipping Then
						'do nothing
					ElseIf mstrDiscountMessage = "OfferFreeGift" Then
						Call showFreeGift
					Else
						If cblnShowDiscountDetailsOnOrderSummary Then
							pstrOut = "<table cellpadding=4 cellspacing=0 border=0>" _
									& "<tr><th valign=bottom align=left>Promotion</th><th valign=bottom><font color=red>" & cstrDiscountAppliedMessage & "!</font></th></tr>"
							If mcurDiscountAmount > 0 Then
								pstrOut = pstrOut & "<tr><td valign=bottom>" & mstrDiscountText & "</td><td valign=bottom align=right><b>" & FormatCurrency(mcurDiscountAmount) & "&nbsp;</B></td></tr>"
							Else
								pstrOut = pstrOut & "<tr><td valign=bottom>" & mstrDiscountText & "</td><td valign=bottom align=right><b>" & "" & "&nbsp;</b></td></tr>"
							End If
							pstrOut = pstrOut & "</table>"
						Else
							pstrOut = cstrDiscountAppliedMessage & ": <b>" & FormatCurrency(mcurDiscountAmount) & "</B>! "
						End If
						mblnAlreadyDisplayed = True	'disable the display to prevent duplicates
					End If
				End If
			Next	'i
			
			Response.Write pstrOut
			
		End If	'isArray(maryDiscountSummary)
	End If	'order.asp check

End Sub	'ShowOrderDiscounts

'***********************************************************************************************

Sub checkDiscountedShipping(byVal dblSubTotal, byRef dblShipping, byVal strShipMethod)

Dim i
Dim pstrFreeShippingMethod
Dim pdblFreeShippingLimit
Dim paryFreeShipping
Dim pdblFreeShippingDiscount

	pdblFreeShippingDiscount = 0
	If isArray(maryDiscountSummary) Then
		For i = 0 to uBound(maryDiscountSummary)
			If Len(maryDiscountSummary(i)(enDiscountFreeShipping)) > 0 And maryDiscountSummary(i)(enDiscountUse) Then
				paryFreeShipping = Split(maryDiscountSummary(i)(enDiscountFreeShipping), "|")
				If UBound(paryFreeShipping) >= 0 Then pstrFreeShippingMethod = paryFreeShipping(0)
				If UBound(paryFreeShipping) >= 1 Then pdblFreeShippingLimit = paryFreeShipping(1)
				
				If cblnDebugPromotionManagerAddon Then
					Response.Write("<font color=black>Check Discounted Shipping</font><br />")
					Response.Write("<font color=black>Free Shipping Method: " & pstrFreeShippingMethod & "</font><br />")
					Response.Write("<font color=black>Free Shipping Limit: " & pdblFreeShippingLimit & "</font><br />")
					Response.Write("<font color=black>dblShipping: " & dblShipping & "</font><br />")
					Response.Write("<font color=black>strShipMethod: " & strShipMethod & "</font><br />")
					Response.Write("<font color=black>minSubTotal: " &  maryDiscountSummary(i)(9) & "</font><br />")
					Response.Write("<font color=black>SubTotal: " &  dblSubTotal & "</font><br />")
				End If

				If (dblSubTotal) >= maryDiscountSummary(i)(9) Then
					If strShipMethod = pstrFreeShippingMethod Then
						pdblFreeShippingDiscount = CDbl(dblShipping)
						If isNumeric(pdblFreeShippingLimit) Then
							If pdblFreeShippingDiscount > CDbl(pdblFreeShippingLimit) Then pdblFreeShippingDiscount = CDbl(pdblFreeShippingLimit)
						End If
						dblShipping = CDbl(dblShipping) - pdblFreeShippingDiscount
						If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black>Shipping Discount: " & pdblFreeShippingDiscount & "</font><br />")
						If cblnDebugPromotionManagerAddon Then Response.Write("<font color=black>Shipping Lowered to: " & dblShipping & "</font><br />")
					End If
				End If
				
				'Now if there is no discount should deactivate this discount
				If pdblFreeShippingDiscount <= 0 Then
					maryDiscountSummary(i)(enDiscountUse) = False
				Else
					maryDiscountSummary(i)(enPromoTitle) = maryDiscountSummary(i)(enPromoTitle) & " - Save " & FormatCurrency(pdblFreeShippingDiscount,2)
				End If
			End If
		Next
	End If

End Sub	'checkDiscountedShipping

'***********************************************************************************************

Sub SaveDiscounts(byVal lngOrderID, byVal subTotal)
'Note: keep subTotal for compatibility with prior installations
Dim pclsPromo

	Set pclsPromo = New clsPromotion
	Call pclsPromo.SaveDiscountsToDatabase(cnn, lngOrderID)
	Call SaveDiscountClub()
	Set pclsPromo = Nothing

End Sub 'SaveDiscounts

'***********************************************************************************************

Function getPromotionCode(byRef strPromotionCode)

Dim pstrPromoCode

	'get PromoCode from querystring or form variables
	pstrPromoCode = Request.QueryString("PromoCode")
	If len(pstrPromoCode) = 0 Then pstrPromoCode = Request.Form("PromoCode")
	
	If isPromotionCodeSafe(pstrPromoCode) Then
		strPromotionCode = pstrPromoCode
		getPromotionCode = True
	Else
		getPromotionCode = False
	End If

End Function 'getPromotionCode

'***********************************************************************************************

Function isPromotionCodeSafe(byVal strPromotionCode)

	'Note the - added since prom frequently contain them
	'DO NOT permit the apostrophe character (') !!!
	isPromotionCodeSafe = allPermissibleCharacters(strPromotionCode, cstrPermissibleCharacters & "-")

End Function 'isPromotionCodeSafe

'***********************************************************************************************

Sub CheckForPromotionRegistration()

Dim pstrPromoCode
Dim prsPromo

	'get PromoCode from querystring or form variables
	 If getPromotionCode(mstrPromotionCode) Then
		If len(mstrPromotionCode) = 0 Then
			mstrPromotionRegistrationMessage = ""
		Else
			mstrPromotionRegistrationMessage = LoadPromo(mstrPromotionCode, prsPromo)
			If Len(mstrPromotionRegistrationMessage) = 0 Then
				mstrPromotionRegistrationMessage = "<H4>Thank You! I have signed you up for " & prsPromo.Fields("PromoTitle").value & "</H4>" _
							& Replace(prsPromo.Fields("PromoRules").value & "" ,vbcrlf,"<br />")	
				prsPromo.Close
			Else
				mstrPromotionRegistrationMessage = "<font class=""Error"">" & mstrPromotionRegistrationMessage & "</font>"
				mblnSuccessfulRegistration = False
			End If
			
			set prsPromo = nothing
			mblnSuccessfulRegistration = True
		End If
	Else
		mstrPromotionRegistrationMessage = "You have entered an invalid code."
	End If

End Sub 'CheckForPromotionRegistration

'************************************************************************************************

Function LoadPromo(byVal strPromoCode, byRef objrsPromo)

dim pstrMessage
dim pdtNow
dim pstrPromoTitle
dim pdtEndDate
dim pdtStartDate
dim pblnValidCode
	
	pdtNow = Now()
	
	If len(strPromoCode) = 0 Then
		pstrMessage = "No promotion code was entered."
	Else
		Set objrsPromo = CreateObject("ADODB.RECORDSET")
		With objrsPromo
			.ActiveConnection = cnn
			.CursorLocation = 2 'adUseClient
			.Open "Select * from Promotions where PromoCode='" & strPromoCode & "'", cnn, 3, 1	'adOpenStatic, adLockReadOnly
		
			If .EOF Then
				pstrMessage = "We're sorry but the promotion code you entered <b>" & strPromoCode & "</b> is invalid."
			Else	'Valid Promotion Code so check for expiration
				If mblnCaseInsensitive Then
					If instr(1,LCase(Trim(.Fields("PromoCode").Value)),LCase(strPromoCode)) > 0 Then
						strPromoCode = Trim(.Fields("PromoCode").Value)
					End If
				End If
				
				If instr(1,Trim(.Fields("PromoCode").Value),strPromoCode) > 0 Then
					pstrPromoTitle = .Fields("PromoTitle").Value
					pdtStartDate = .Fields("StartDate").Value
					pdtEndDate = .Fields("EndDate").Value
					If len(pdtEndDate & "") = 0 Then
						If len(.Fields("Duration").Value & "") > 0 Then	pdtEndDate = pdtStartDate + .Fields("Duration").Value
					End If
				
					If .Fields("Inactive").Value Then
						pstrMessage = "We're sorry but the " & pstrPromoTitle & " promotion is no longer valid."
					Else
						If pdtStartDate > pdtNow Then
							pstrMessage = "We're sorry but the " & pstrPromoTitle & " promotion does not start until " & FormatDateTime(pdtStartDate) & "."
						Else
							If len(pdtEndDate & "") > 0 Then
								If pdtEndDate < pdtNow Then pstrMessage = "We're sorry but the " & pstrPromoTitle & " promotion has expired."
							Else
								If len(.Fields("MaxUses").Value & "") > 0 Then
									If .Fields("NumUses").Value >= .Fields("MaxUses").Value Then pstrMessage = "We're sorry but the " & pstrPromoTitle & " promotion has already been redeemed the maximum number of times."
								End If
							End If
						End If
					End If
					mstrPromotedProducts = Trim(.Fields("ProductID").value & "")
				Else
					pstrMessage = "We're sorry but the promotion code you entered <b>" & strPromoCode & "</b> is invalid."
				End If
			End If
		End With
	End If
	
	'Set the promotion to the session variables
	If len(pstrMessage) = 0 Then Call SetPromotionCodeToSession(strPromoCode)
	
	LoadPromo = pstrMessage

End Function	'LoadPromo

'***********************************************************************************************

Sub SetPromotionCodeToSession(byVal strPromoCode)

Dim pstrTempPromo

	If Len(strPromoCode) = 0 Then Exit Sub

	pstrTempPromo = vistorDiscountCodes
	If len(pstrTempPromo) = 0 Then
		pstrTempPromo = ";" & strPromoCode & ";"
		Call setVisitorPreference("vistorDiscountCodes", pstrTempPromo, True)
	Else
		If instr(1,pstrTempPromo,";" & strPromoCode & ";",1) = 0 Then
			pstrTempPromo = pstrTempPromo & strPromoCode & ";"
			Call setVisitorPreference("vistorDiscountCodes", pstrTempPromo, True)
		End If
	End If

End Sub	'SetPromotionCodeToSession

'***********************************************************************************************
' Added for Discount Club
'***********************************************************************************************

Sub InitializeDiscountClubProducts

	If cblnDebugPromotionManagerAddon Then Response.Write "Initializing discount club products<br />"
	
	If False Then
		ReDim maryDiscountClubProducts(5)
		maryDiscountClubProducts(0) = Array("CLUB", "43", "GoldClub", DateAdd("y", 1, Date()))
		maryDiscountClubProducts(1) = Array("CLUB", "45", "GoldClub", DateAdd("y", 2, Date()))
		maryDiscountClubProducts(2) = Array("CLUB", "47", "GoldClub", "")
		maryDiscountClubProducts(3) = Array("CLUB", "42", "PlatinumClub", DateAdd("y", 1, Date()))
		maryDiscountClubProducts(4) = Array("CLUB", "46", "PlatinumClub", DateAdd("y", 2, Date()))
		maryDiscountClubProducts(5) = Array("CLUB", "44", "PlatinumClub", "")
	Else
		ReDim maryDiscountClubProducts(1)
		maryDiscountClubProducts(0) = Array("WF002", "White", "10PercentOff", DateAdd("y", 1, Date()))
		maryDiscountClubProducts(1) = Array("WF002", "4 Pack", "10PercentOff", DateAdd("y", 1, Date()))
	End If

End Sub	'InitializeDiscountClubProducts

'***********************************************************************************************

Sub CheckCartForDiscountClub(byRef aryOrderItem)
'Purpose: Checks cart contents during final cart roll-up in confirm.asp and collects gift certificate products into array for future processing

Dim i
Dim paryAttributes
Dim pblnMatch
Dim plngClubCounter
Dim pstrClubAttrToCheck
Dim pstrClubProductToCheck
Dim plngProductQty
Dim pstrProductID

	Call InitializeDiscountClubProducts
	
	pstrProductID = aryOrderItem(enOrderItem_prodID)
	
	If cblnDebugPromotionManagerAddon Then Response.Write "CheckCartForDiscountClub - <b>" & pstrProductID & "</b><br />"

	plngProductQty = aryOrderItem(enOrderItem_odrdttmpQuantity)
	
	For plngClubCounter = 0 To UBound(maryDiscountClubProducts)
		pstrClubProductToCheck = maryDiscountClubProducts(plngClubCounter)(0)
		pstrClubAttrToCheck = maryDiscountClubProducts(plngClubCounter)(1)
		pblnMatch = False
		
		If cblnDebugPromotionManagerAddon Then Response.Write "Checking product <b>" & pstrProductID & "</b> to see if it is a discount club purchase <b>" & pstrClubProductToCheck & "</b> - Result: " & CBool(pstrClubProductToCheck = pstrProductID) & "<br />"
	
		If CBool(pstrClubProductToCheck = pstrProductID) Then
		
			If aryOrderItem(enOrderItem_AttributeCount) = 0 Then
				pblnMatch = True
			Else
				For i = 0 To aryOrderItem(enOrderItem_AttributeCount) - 1
					paryAttributes = paryOrderItem(enOrderItem_AttributeArray)
					If cblnDebugPromotionManagerAddon Then Response.Write " - Checking attribute <b>" & paryAttributes(j)(enAttributeItem_attrdtID) & "</b> to see if it is a discount club purchase <b>" & pstrClubAttrToCheck & "</b> - Result: " & CBool(pstrClubAttrToCheck = paryAttributes(j)(enAttributeItem_attrdtID)) & "<br />"
					If paryAttributes(j)(enAttributeItem_attrdtID) = pstrClubAttrToCheck Then
						pblnMatch = True
						Exit For
					End If
				Next 'i
			End If	'aryOrderItem(enOrderItem_AttributeCount) = 0
		End If	'pstrClubProductToCheck = pstrProductID
				
		If pblnMatch Then
		
			mlngNumClubCertificatesToCreate = mlngNumClubCertificatesToCreate + 1
			ReDim Preserve maryDiscountClubsToCreate(mlngNumClubCertificatesToCreate)

			maryDiscountClubsToCreate(mlngNumClubCertificatesToCreate) = Array(maryDiscountClubProducts(plngClubCounter)(2), maryDiscountClubProducts(plngClubCounter)(3))
			
			If cblnDebugPromotionManagerAddon Then 
				Response.Write "DiscountClubsToCreate: " & mlngNumClubCertificatesToCreate & "<br />"
				Response.Write "- Code: " & maryDiscountClubProducts(mlngNumClubCertificatesToCreate)(2) & "<br />"
				Response.Write "- Exp.: " & maryDiscountClubProducts(mlngNumClubCertificatesToCreate)(3) & "<br />"
			End If	'cblnDebugPromotionManagerAddon
			
			Exit For
		End If	'pblnMatch
	        
	Next 'plngClubCounter

	If cblnDebugPromotionManagerAddon Then 
		Response.Write "DiscountClubsToCreate: " & mlngNumClubCertificatesToCreate & "<br />"
		If mlngNumClubCertificatesToCreate >= 0 Then
			Response.Write "Code: " & maryDiscountClubsToCreate(mlngNumClubCertificatesToCreate)(0) & "<br />"
			Response.Write "Exp.: " & maryDiscountClubsToCreate(mlngNumClubCertificatesToCreate)(1) & "<br />"
		Else
			Response.Write "- Match found: " & pblnMatch & "<br />"
		End If
	End If	'cblnDebugPromotionManagerAddon
	
End Sub	'CheckCartForDiscountClub

'***********************************************************************************************

Sub SaveDiscountClub()

Dim i
Dim pstrSQL

	If mlngNumClubCertificatesToCreate >= 0 Then
		For i = 0 To mlngNumClubCertificatesToCreate
			If CBool(Application("AppDatabase") = "Access") Then
				pstrSQL = "Update sfCustomers Set clubCode='" & maryDiscountClubsToCreate(i)(0) & "', " _
						& " clubExpDate=#" & maryDiscountClubsToCreate(i)(1) & "# " _
						& "Where custID=" & visitorLoggedInCustomerID
			Else
				pstrSQL = "Update sfCustomers Set clubCode='" & maryDiscountClubsToCreate(i)(0) & "', " _
						& " clubExpDate='" & maryDiscountClubsToCreate(i)(1) & "' " _
						& "Where custID=" & visitorLoggedInCustomerID
			End If
			cnn.Execute pstrSQL,,128
			If cblnDebugPromotionManagerAddon Then
				Response.Write "ssPromotionManager_SaveDiscountClub - SQL: " & pstrSQL & "<br />"
			End If
		Next 'i
	End If

End Sub	'SaveDiscountClub

%>
