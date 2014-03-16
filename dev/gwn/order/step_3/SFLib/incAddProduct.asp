<!--#include file="ssmodNotifyMe.asp"-->
<%
'********************************************************************************
'*
'*   search_results.asp - 
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins are search_results.asp APPVERSION: 50.4014.0.3
'*   from LaGarde, Incorporated. It has been heavily modified by Sandshot Software
'*   with permission from LaGarde, Incorporated who has NOT released any of the 
'*   original copyright protections. As such, this is a derived work covered by
'*   the copyright provisions of both companies.
'*
'*   LaGarde, Incorporated Copyright Statement                                                                           *
'*   The contents of this file is protected under the United States copyright
'*   laws and is confidential and proprietary to LaGarde, Incorporated.  Its 
'*   use ordisclosure in whole or in part without the expressed written 
'*   permission of LaGarde, Incorporated is expressly prohibited.
'*   (c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'*   
'*   Sandshot Software Copyright Statement
'*   The contents of this file are protected by United States copyright laws 
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************

Const cblnScrollToTop = True

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim maryCartAdditions
	Const enCartItem_prodID = 0
	Const enCartItem_prodName = 1
	Const enCartItem_AttributeArray = 2
	Const enCartItem_QtyToAdd = 3
	Const enCartItem_QtyInCart = 4
	Const enCartItem_QtyInStock = 5
	Const enCartItem_QtyToGW = 6
	Const enCartItem_invenbTracked = 7
	Const enCartItem_invenbBackOrder = 8
	Const enCartItem_Upsell = 9
	Const enCartItem_QtyAdded = 10
	Const enCartItem_AddType = 11
	Const enCartItem_Result = 12
	Const enCartItem_ResponseMessage = 13
	Const enCartItem_tmpOrderDetailId = 14
	Const enCartItem_AskForBackOrder = 15

Const cstrSSMPOAttributeDelimiter_Default = "IMPOI"	'this must match the value found in ssmodAttributeExtender

Dim mblnMPOPage_ssMPO
Dim mblnMPODuplicates_ssMPO

'**********************************************************
'*	Functions
'**********************************************************

'Function getCartActionText(byRef aryCartAddition)

'**********************************************************
'*	Begin Page Code
'**********************************************************

Call initializeMPOSettings

'**********************************************************
'*	Begin Function Definitions
'**********************************************************

Function getCartActionText(byRef aryCartAddition)
	If aryCartAddition(enCartItem_AddType) Then
		If aryCartAddition(enCartItem_QtyAdded) > 1 Then
			getCartActionText = "have been saved to your " & CartName & "."
		Else
			getCartActionText = "has been saved to your " & CartName & "."
		End If
	Else
		If aryCartAddition(enCartItem_QtyAdded) > 1 Then
			getCartActionText = "have been added to your order."
		Else
			getCartActionText = "has been added to your order."
		End If
	End If	
End Function	'getCartActionText

'*******************************************************************************************************

Function getAttributes_Duplicate(byVal strValue, byVal blnMPODuplicates_ssMPO, byVal lngCounter_ssMPO)

Dim pstrTemp
Dim paryTemp

	pstrTemp = Trim(strValue)
	
	If vDebug = 1 Then
		Response.Write "<hr><h4>duplicate</h4>"
		Response.Write "blnMPODuplicates_ssMPO: " & blnMPODuplicates_ssMPO & "<br />"
		Response.Write "strValue: " & strValue & "<br />"
		Response.Write "lngCounter_ssMPO: " & lngCounter_ssMPO & "<br />"
	End If

	If blnMPODuplicates_ssMPO And Len(pstrTemp) > 0 Then
		paryTemp = Split(pstrTemp, ",")
		If isArray(paryTemp) Then
			pstrTemp = paryTemp(lngCounter_ssMPO)
		End If
	End If
	
	If vDebug = 1 Then		
		Response.Write "<strong>Value Out: " & pstrTemp & "</strong><br />"
		Response.Write "<br /><hr><br />"
		Response.Flush
	End If
	
	getAttributes_Duplicate = pstrTemp

End Function	'getAttributes_Duplicate

'**********************************************************************************************************

Function getFractionalQty(byVal pstrProductID)

Dim pstrTempQtyFraction

	pstrTempQtyFraction = Request.Form("QUANTITY_FRACTION." & pstrProductID)
	If Len(pstrTempQtyFraction) = 0 Or Not isNumeric(pstrTempQtyFraction) Then pstrTempQtyFraction = 0
	
	If pstrTempQtyFraction = 0 Then
		getFractionalQty = 0
	Else
		getFractionalQty = FormatNumber(pstrTempQtyFraction, 3, False)
	End If

End Function	'getFractionalQty

'**********************************************************************************************************

Sub initializeMPOSettings

	mblnMPOPage_ssMPO = (Len(Request.Form("ssMPOPage")) > 0)
	cstrSSMPOAttributeDelimiter = Request.Form("ssMPO_Delimiter")
	If Len(cstrSSMPOAttributeDelimiter) = 0 Then cstrSSMPOAttributeDelimiter = cstrSSMPOAttributeDelimiter_Default
	
End Sub	'initializeMPOSettings

'**********************************************************************************************************

Sub AddItemToCartAdditions(byVal lngCounter, byVal strProductID, byVal lngQty, byVal aryAttributes)

	If Not isArray(maryCartAdditions) Then ReDim maryCartAdditions(5)
	If UBound(maryCartAdditions) < lngCounter Then ReDim Preserve maryCartAdditions(lngCounter)
	maryCartAdditions(lngCounter) = Array(strProductID, getProductInfo(strProductID, enProduct_Name), aryAttributes, lngQty, 0, 0, 0, 0, 0, "", 0, mblnSaveCart, "", "", "", False)

End Sub

'**********************************************************************************************************

Function numberProductsToAddToCart()
'Purpose: determine how many items in Request.Form represent products to add
'Returns: long integer; -1 for false, number of products, base 0
'		  sets values for mblnMPOPage_ssMPO which is defined globally in this file
'		  sets values for mblnMPODuplicates_ssMPO which is defined globally in this file

Dim j, attrCounter
Dim paryTempAttr
Dim paryTempProdID
Dim paryTempQty
Dim plngCounter
Dim plngItemsToAdd
Dim plngPos1
Dim pstrFormItem
Dim pstrQTYAttr
Dim pstrTempProductID
Dim pstrTempQty
Dim vItem
	
	If Len(cstrSSMPOAttributeDelimiter) = 0 Then cstrSSMPOAttributeDelimiter = cstrSSMPOAttributeDelimiter_Default
	
	mblnMPOPage_ssMPO = (Len(Request.Form("ssMPOPage")) > 0)
	mblnMPODuplicates_ssMPO = (Len(Request.Form("ssAttrQTY")) > 0)
	
	If vDebug = 1 Then
		Response.Write "<fieldset><legend>numberProductsToAddToCart</legend>" & vbcrlf
		Response.Write "<b>Form Contents</b>"
		For Each pstrFormItem In Request.Form
			Response.Write pstrFormItem & ": " & Request.Form(pstrFormItem) & "<br />" & vbcrlf
		Next 'pstrFormItem
		Response.Write "<hr>"
		Response.Write "MPO Page: " & mblnMPOPage_ssMPO & "<br />"
		Response.Write "mblnMPODuplicates_ssMPO: " & mblnMPODuplicates_ssMPO & "<br />"
		Response.Write "cstrSSTextBasedAttributeHTMLDelimiter: " & cstrSSTextBasedAttributeHTMLDelimiter & "<br />"
		Response.Write "cstrSSMPOAttributeDelimiter: " & cstrSSMPOAttributeDelimiter & "<br />"
		Response.Write "mblnMPODuplicates_ssMPO: " & mblnMPODuplicates_ssMPO & "<br />"
		Response.Write "</fieldset>" & vbcrlf
		Response.Flush
	End If
	
	plngCounter = -1
	pstrTempQty = Trim(Request.Form("QUANTITY"))

	pstrTempProductID = Trim(Request.Form("PRODUCT"))
	'Scenario: Detail Page, Attributes As Qty
	If mblnMPODuplicates_ssMPO Then
		If vDebug = 1 Then Response.Write "<h4>Scenario to check: Detail Page, Attributes As Qty</h4>" & vbcrlf

		pstrTempProductID = Request.Form("PRODUCT_ID")
		pstrTempQty = Request.Form("QUANTITY")
		paryTempQty = Split(pstrTempQty, ",")
	
		For j = 0 To UBound(paryTempQty)
			pstrTempQty = Trim(paryTempQty(j))
			If Len(pstrTempQty) = 0 Then pstrTempQty = 0
			pstrTempQty = pstrTempQty + getFractionalQty(pstrTempProductID)
			If pstrTempQty > 0 And Len(pstrTempProductID) > 0 Then
				plngCounter = plngCounter + 1

				pstrQTYAttr = Trim(Request.Form("ssAttrQTY"))
				If isNumeric(pstrQTYAttr) Then pstrQTYAttr = pstrQTYAttr - 1
				paryTempAttr = getAttributesFromRequest(pstrTempProductID, j)
				If isArray(paryTempAttr) Then
					paryTempAttr(pstrQTYAttr) = Split(paryTempAttr(pstrQTYAttr), ",")(j)
				End If	'isArray(paryTempAttr)

				Call AddItemToCartAdditions(plngCounter, pstrTempProductID, pstrTempQty, paryTempAttr)
			End If
		Next 'j

	'Scenario: Quick Order Page
	ElseIf InStr(1, pstrTempProductID, ",") > 0 Then
		If vDebug = 1 Then Response.Write "<h4>Scenario to check: Quick Order Page</h4>" & vbcrlf
		paryTempQty = Split(pstrTempQty, ",")
		paryTempProdID = Split(pstrTempProductID, ",")
		For j = 0 To UBound(paryTempQty)
			pstrTempProductID = Trim(paryTempProdID(j))
			pstrTempQty = Trim(paryTempQty(j))
			If Len(pstrTempQty) = 0 Then pstrTempQty = 0
			pstrTempQty = pstrTempQty + getFractionalQty(pstrTempProductID)
			If pstrTempQty > 0 And Len(pstrTempProductID) > 0 Then
				plngCounter = plngCounter + 1
				Call AddItemToCartAdditions(plngCounter, pstrTempProductID, pstrTempQty, getAttributesFromRequest(pstrTempProductID, j))
			End If
		Next 'j

	'Scenario: Standard Product Entry
	ElseIf InStr(1, pstrTempQty, ",") = 0 Then
		If vDebug = 1 Then Response.Write "<h4>Scenario to check: Standard Product Entry</h4>" & vbcrlf
		pstrTempProductID = Trim(Request.Form("PRODUCT_ID"))
		If Len(pstrTempQty) = 0 Then pstrTempQty = 0
		'now add in the fractional qty
		pstrTempQty = pstrTempQty + getFractionalQty(pstrTempProductID)
		If pstrTempQty > 0 Then
			plngCounter = plngCounter + 1
			Call AddItemToCartAdditions(plngCounter, pstrTempProductID, pstrTempQty, getAttributesFromRequest(pstrTempProductID, j))
		End If
	
	'Scenario: Multiple, identical (sans attributes) products from same page
	ElseIf Len(pstrTempQty) > 0 Then
		If vDebug = 1 Then Response.Write "<h4>Scenario to check: Multiple, identical (sans attributes) products from same page</h4>" & vbcrlf
		pstrTempProductID = Trim(Request.Form("PRODUCT_ID"))
		paryTempQty = Split(pstrTempQty,",")
		mblnMPODuplicates_ssMPO = True

		If getProductInfo(pstrTempProductID, enProduct_IsActive) Then
			For j = 0 To UBound(paryTempQty)
				pstrTempQty = Trim(paryTempQty(j))
				If Len(pstrTempQty) = 0 Then pstrTempQty = 0
				'now add in the fractional qty
				pstrTempQty = pstrTempQty + getFractionalQty(pstrTempProductID)
				If pstrTempQty > 0 Then
					plngCounter = plngCounter + 1
					Call AddItemToCartAdditions(plngCounter, pstrTempProductID, pstrTempQty, getAttributesFromRequest(pstrTempProductID, j))
				End If
			Next 'j
		End If	'getProductInfo(pstrTempProductID, enProduct_IsActive)
	End If
	
	'Scenario: Multiple different products from same page
	If vDebug = 1 Then Response.Write "<h4>Scenario to check: Multiple different products from same page</h4>" & vbcrlf
	For each vItem in Request.Form
		pstrTempProductID = ""
		pstrTempQty = ""
		If instr(1,vItem,"QUANTITY.") <> 0 Then
			pstrTempQty = Trim(Request.Form(vItem))
			If len(pstrTempQty) > 0 And isNumeric(pstrTempQty) Then
				paryTempQty = Split(pstrTempQty,",")
				pstrTempQty = ""

				For j = 0 To UBound(paryTempQty)
					If isNumeric(Trim(paryTempQty(j))) Then 
						If Len(pstrTempQty) = 0 Then pstrTempQty = 0
						pstrTempQty = pstrTempQty + CInt(Trim(paryTempQty(j)))
					End If
				Next 'j
				
				If Len(pstrTempQty) > 0 Then
					plngPos1 = instr(1,vItem,".")
					pstrTempProductID = Right(vItem,len(vItem)-plngPos1)
					plngCounter = plngCounter + 1
					If vDebug = 1 Then Response.Write "<h4>Scenario: Multiple different products from same page</h4>" & vbcrlf
					Call AddItemToCartAdditions(plngCounter, pstrTempProductID, pstrTempQty, getAttributesFromRequest(pstrTempProductID, j))
				End If

			End If
		End If
	Next
	
	'Shrink down the array if necessary
	If UBound(maryCartAdditions) > plngCounter Then ReDim Preserve maryCartAdditions(plngCounter)

	If vDebug = 1 Then
		Response.Write "<fieldset><legend>Products To Add To Cart</legend>" & vbcrlf
		Response.Write "Number of items to add: " & plngCounter & "<br />"
		Response.Write "<ol>"
		For j = 0 To UBound(maryCartAdditions)
			Response.Write "<li>"
			Response.Write "Product: " & maryCartAdditions(j)(enCartItem_prodID) & "<br />" & vbcrlf
			Response.Write "Quantity: " & maryCartAdditions(j)(enCartItem_QtyToAdd) & "<br />" & vbcrlf
			paryTempQty = maryCartAdditions(j)(enCartItem_AttributeArray)
			If isArray(paryTempQty) Then
				Response.Write "Attributes<ul>"
				For attrCounter = 0 To UBound(paryTempQty)
					Response.Write "<li>Attribute " & attrCounter & ": " & paryTempQty(attrCounter) & "</li>" & vbcrlf
				Next 'attrCounter
				Response.Write "</ul>"
			End If	'isArray(paryTempQty)
			Response.Write "</li>"
		Next 'j
		Response.Write "<ol>"
		Response.Write "</fieldset>" & vbcrlf
	End If
	
	numberProductsToAddToCart = plngCounter

End Function	'numberProductsToAddToCart

'**********************************************************************************************************

Function getAttributesFromRequest(byVal strProductID, byVal lngCounter)

Dim paryAttributes
Dim pvntNumAttributes

	pvntNumAttributes = getProductInfo(strProductID, enProduct_AttrNum)
	If Len(pvntNumAttributes) = 0 Then
		'Do Nothing
	ElseIf  CLng(pvntNumAttributes) > 0 Then
		Call getProductToAddsAttributes(strProductID, paryAttributes, getProductInfo(strProductID, enProduct_AttrNum), False, lngCounter)
	End If
	
	getAttributesFromRequest = paryAttributes
	
End Function	'getAttributesFromRequest

'**********************************************************************************************************
	
Sub getProductToAdd(byRef strProductID, byRef lngQuantity, byVal lngCounter, byVal blnOldPage)
'Purpose: determine how many items in Request.Form represent products to add
'Returns: Nothing
'		  sets values for strProductID which is byRef
'		  sets values for lngQuantity which is byRef

Dim formItem
Dim paryItem_ssMPO

	If Not blnOldPage Then
		strProductID = maryCartAdditions(lngCounter)(enCartItem_prodID)
		lngQuantity = maryCartAdditions(lngCounter)(enCartItem_QtyToAdd)
	Else
		For Each formItem In Request.Form
			If InStr(formItem, "PRODUCT_ID") Then 
				strProductID = Trim(Request.Form(formItem))
				Exit For
			End	If
		Next
	End If

End Sub	'getProductToAdd

'**********************************************************************************************************
	
Sub getProductToAddsAttributes(byVal strProductID, byRef aProdAttr, byVal iProdAttrNum, byVal blnOldPage, byVal lngCounter)
'Purpose: fills aProdAttr with attribute selection from Request.Form
'Returns: Nothing
'		  sets values for aProdAttr which is byRef

	If HasAttributes Then
		Call GetAttributes(strProductID, aProdAttr, iProdAttrNum, lngCounter)
	ElseIf blnOldPage Then 
		ReDim aProdAttr(3)
		aProdAttr(0) = Trim(Request.Form("AttributeA"))
		aProdAttr(1) = Trim(Request.Form("AttributeB"))
		aProdAttr(2) = Trim(Request.Form("AttributeC"))

		aProdAttr = getAttrID(strProductID, aProdAttr)
	
		If CDbl(trim(Ubound(aProdAttr))) < CDbl(trim(iProdAttrNum)) Then
			Response.Write "<p><b><font face=""verdana"" size=""2"">Product no longer has one of the attributes listed in the product pages. Please contact store owner.</font></b>"
			' Redirect to error page later --
			Response.End
		End If	 
	End If  ' End Attributes If

End Sub	'getProductToAddsAttributes

'**********************************************************************************************************

Function getAttrID(byVal strProductID, byVal aProdAttr)
'Purpose: Returns and array of product's attributes' id(s) information for a given product ID and attribute names
' This is to accommodate old product pages, so only three attributes are needed
'Input: prodID, array of attribute names
'Return: array

Dim iArrayBound
Dim paryProduct
Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	Set pobjCmd = CreateObject("ADODB.Command")
	With pobjCmd
		If True Then
			pstrSQL = "SELECT attrdtID" _
					& " FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributeDetail.attrdtAttributeId = sfAttributes.attrID" _
					& " WHERE trim(attrProdId)=? AND (trim(sfAttributeDetail.attrdtName)=? OR trim(sfAttributeDetail.attrdtName)=? OR trim(sfAttributeDetail.attrdtName)=?)"
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
		Else
			pstrSQL = "spGetAttrID"
			.Commandtype = adCmdStoredProc
			.Commandtext = pstrSQL
		End IF
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProdID)
		.Parameters.Append .CreateParameter("attrA", adVarChar, adParamInput, 255, aProdAttr(0))
		.Parameters.Append .CreateParameter("attrB", adVarChar, adParamInput, 255, aProdAttr(1))
		.Parameters.Append .CreateParameter("attrC", adVarChar, adParamInput, 255, aProdAttr(2))
		Set pobjRS = .Execute
	End With	'pobjCmd
		
	With pobjRS
		If .EOF Then
			If vDebug = 1 Then Response.Write "<p>Empty Recordset in getAttrID. Product " & strProductID & " possibly not activated."
		Else

  			iCounter = 0
  			iArrayBound = .RecordCount
  			If vDebug = 1 Then Response.Write "<br />Array Bound in getAttrID" & iArrayBound
  			ReDim paryProduct(iArrayBound)
			  			  			
  			Do While NOT .EOF
				paryProduct(iCounter) = .Fields("attrdtID").Value
				If vDebug = 1 Then Response.Write "<br />Array ID = " & paryProduct(iCounter)
  				iCounter = iCounter + 1
  				.MoveNext   		
   			Loop

		End If	'.EOF
	End With	'pobjRS
	
	closeObj(pobjRS)
	closeObj(pobjCmd)

	getAttrID = paryProduct
	
End Function	'getAttrID

'**********************************************************************************************************

'-------------------------------------------------------------------
' This combines identical items in saved cart
'-------------------------------------------------------------------
Sub setCombineProducts(iCustID)

Dim	sLocalSQL,sLocalSQL2,rsSaved,iTotalRecords, tmpProdID, i, j, tmpProdQuantity, tmpProdSvdID, rsDetails, rsDetails2, cmpProdID
Dim aProd, rsCheckExists, sCheckExists, sAttrString, sAttrString2	
	sLocalSQL = "SELECT odrdtsvdID, odrdtsvdQuantity, odrdtsvdProductID FROM sfSavedOrderDetails WHERE odrdtsvdCustID = " & makeInputSafe(iCustID)
	
	Set rsSaved = CreateObject("ADODB.RecordSet")
	rsSaved.Open sLocalSQL, cnn, adOpenKeySet, adLockOptimistic, adCmdText
	iTotalRecords = rsSaved.RecordCount
	i = 0
	Redim aProd(3,iTotalRecords-1)
	
	Do While Not rsSaved.EOF
		tmpProdID		= Trim(rsSaved.Fields("odrdtsvdProductID"))
		tmpProdQuantity = Trim(rsSaved.Fields("odrdtsvdQuantity"))
		tmpProdSvdID	= Trim(rsSaved.Fields("odrdtsvdID"))
	
		aProd(0,i) = tmpProdID
		aProd(1,i)	= tmpProdQuantity
		aProd(2,i) = tmpProdSvdID		
		
	i = i + 1
	rsSaved.MoveNext
	Loop
	
	
	For i = 0 to UBOUND(aProd,2)
		tmpProdID = aProd(0,i)
		tmpProdSvdID = aProd(2,i)
		
		For j = i + 1 to UBOUND(aProd,2)
				cmpProdID = aProd(0,j)
				If cmpProdID = tmpProdID Then

					sLocalSQL = "SELECT odrattrsvdID, odrattrsvdOrderDetailId, odrattrsvdAttrID "_
					& "FROM sfSavedOrderDetails INNER JOIN sfSavedOrderAttributes on sfSavedOrderAttributes.odrattrsvdOrderDetailId = sfSavedOrderDetails.odrdtsvdID "_
					& "WHERE sfSavedOrderAttributes.odrattrsvdOrderDetailId = " & aProd(2,j)
						
					sLocalSQL2 = "SELECT odrattrsvdID, odrattrsvdOrderDetailId, odrattrsvdAttrID "_
					& "FROM sfSavedOrderDetails INNER JOIN sfSavedOrderAttributes on sfSavedOrderAttributes.odrattrsvdOrderDetailId = sfSavedOrderDetails.odrdtsvdID "_
					& "WHERE sfSavedOrderAttributes.odrattrsvdOrderDetailId = " & aProd(2,i)	
					
					sCheckExists = "SELECT odrdtsvdID FROM sfSavedOrderDetails WHERE sfSavedOrderDetails.odrdtsvdID = " & aProd(2,i)
					
					Set rsDetails = CreateObject("ADODB.RecordSet")
					rsDetails.Open sLocalSQL, cnn, adOpenKeySet, adLockOptimistic, adCmdText
					
					Set rsCheckExists = CreateObject("ADODB.RecordSet")
						rsCheckExists.Open sCheckExists, cnn, adOpenKeySet, adLockOptimistic, adCmdText
						If NOT rsCheckExists.EOF Then
								Set rsDetails2 = CreateObject("ADODB.RecordSet")
								rsDetails2.Open sLocalSQL2, cnn, adOpenKeySet, adLockOptimistic, adCmdText
							
								If rsDetails.EOF AND rsDetails2.EOF Then
									' combine the two since there are no attributes
									Call setUpdateQuantity("odrdtsvd",aProd(1,j),tmpProdSvdID)
									Call setDeleteOrder("odrdtsvd",aProd(2,j))	
								Else
									' compare the attributes
									Do While Not rsDetails.EOF
										sAttrString = sAttrString & "[" & Trim(rsDetails.Fields("odrattrsvdAttrID")) & "]"							
										rsDetails.MoveNext
									Loop
									Do While Not rsDetails2.EOF
										sAttrString2 = sAttrString2 & "[" & Trim(rsDetails2.Fields("odrattrsvdAttrID")) & "]"
										rsDetails2.MoveNext
									Loop
							
									If sAttrString	= sAttrString2 Then
										Call setUpdateQuantity("odrdtsvd",aProd(1,j),tmpProdSvdID)
										Call setDeleteOrder("odrdtsvd",aProd(2,j))
									End If	
								End If	
						End If					
				End If					
			Next
	Next		
closeObj(rsDetails)
closeObj(rsCheckExists)
closeObj(rsSaved)				
End Sub

'-----------------------------------------------------------------------
' Deletes saved customer row
'-----------------------------------------------------------------------
Sub DeleteCustRow(iCustID)
	Dim rsDelete, sSQL
	
	sSQL = "DELETE FROM sfCustomers WHERE custID= " & makeInputSafe(iCustID)	& " AND custFirstName = 'Saved Cart Customer'"
	Set rsDelete = cnn.Execute(sSQL)
	closeObj(rsDelete)
End Sub

'---------------------------------------------------------------
' To see if it is a saved cart customer
' Returns a boolean value
'---------------------------------------------------------------
Function CheckSavedCartCustomer(iCustID)
	Dim sSQL, rsTmp, bTruth
	sSQL = "SELECT custFirstName FROM sfCustomers WHERE custID=" & makeInputSafe(iCustID)
	
	bTruth = false
	
	Set rsTmp = CreateObject("ADODB.RecordSet")
		 rsTmp.Open sSQL,cnn,adOpenDynamic,adLockOptimistic,adCmdText
		 If NOT rsTmp.EOF Then
		 	If trim(rsTmp.Fields("custFirstName")) = "Saved Cart Customer" Then
		 		bTruth = true
		 		
		 	Else
		 		bTruth = false
		 	End If		
		 End If
		
	closeobj(rsTmp)	
	CheckSavedCartCustomer = bTruth
End Function

'--------------------------------------------------------
' Checks if Customer exists in customer table
'--------------------------------------------------------
Function CheckCustomerExists(iCustID)
	Dim sSQL, rsCust, bExists
	sSQL = "SELECT custID FROM sfCustomers WHERE custID = " & makeInputSafe(iCustID)
	Set rsCust = CreateObject("ADODB.RecordSet")
		rsCust.Open sSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If NOT rsCust.EOF Then
			If CLng(rsCust.Fields("custID")) > 0 Then
				bExists = true
			Else
				bExists = false	
			End If
		Else
			bExists = false	
		End If
		
	CheckCustomerExists = bExists
End Function

'-------------------------------------------------------------------
' Subroutine setUpdateSavedCartCustID
'-------------------------------------------------------------------
Sub setUpdateSavedCartCustID(iCustID,iDeletedCustID)
	Dim sSQL, rsTmpCust
	sSQL = "Select odrdtsvdCustID FROM sfSavedOrderDetails WHERE odrdtsvdCustID=" & makeInputSafe(iDeletedCustID)
	Set rsTmpCust = CreateObject("ADODB.RecordSet")		
		rsTmpCust.Open sSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText					
			Do While NOT rsTmpCust.EOF
					rsTmpCust.Fields("odrdtsvdCustID")	= trim(iCustID)
					rsTmpCust.Update	
					rsTmpCust.MoveNext
			Loop
		closeobj(rsTmpCust)		
End Sub

'*******************************************************************************************************

Sub setCartAdditionResultsToSession()
	Session("CartAdditions") = maryCartAdditions
	If vDebug = 1 Then Response.Write "<p>Cart Addition Results Saved To Session</p>"
End Sub

'*******************************************************************************************************

Function getCartAdditionResultsFromSession()
	getCartAdditionResultsFromSession = Session("CartAdditions")
End Function

'*******************************************************************************************************

Function clearCartAdditionResultsFromSession()
	Session.Contents.Remove("CartAdditions")
End Function

'*******************************************************************************************************

Function isCartInSession()
	If Not isArray(maryCartAdditions) Then maryCartAdditions = getCartAdditionResultsFromSession
	isCartInSession = isArray(maryCartAdditions)
End Function

'*******************************************************************************************************

Sub writeCartItem(byRef aryCartItem)
	Response.Write "<fieldset style=""background: white""><legend>Cart Item " & aryCartItem(enCartItem_prodID) & "</legend>"
	Response.Write "enCartItem_prodName: " & aryCartItem(enCartItem_prodName) & "<br />"
	Response.Write "enCartItem_QtyToAdd: " & aryCartItem(enCartItem_QtyToAdd) & "<br />"
	Response.Write "enCartItem_QtyInCart: " & aryCartItem(enCartItem_QtyInCart) & "<br />"
	Response.Write "enCartItem_QtyInStock: " & aryCartItem(enCartItem_QtyInStock) & "<br />"
	Response.Write "enCartItem_QtyAdded: " & aryCartItem(enCartItem_QtyAdded) & "<br />"
	Response.Write "enCartItem_QtyToGW: " & aryCartItem(enCartItem_QtyToGW) & "<br />"
	Response.Write "enCartItem_invenbTracked: " & aryCartItem(enCartItem_invenbTracked) & "<br />"
	Response.Write "enCartItem_invenbBackOrder: " & aryCartItem(enCartItem_invenbBackOrder) & "<br />"
	Response.Write "enCartItem_Upsell: " & aryCartItem(enCartItem_Upsell) & "<br />"
	Response.Write "enCartItem_AddType: " & aryCartItem(enCartItem_AddType) & "<br />"
	Response.Write "enCartItem_Result: " & aryCartItem(enCartItem_Result) & "<br />"
	Response.Write "enCartItem_ResponseMessage: " & aryCartItem(enCartItem_ResponseMessage) & "<br />"
	Response.Write "enCartItem_tmpOrderDetailId: " & aryCartItem(enCartItem_tmpOrderDetailId) & "<br />"
	Response.Write "enCartItem_AskForBackOrder: " & aryCartItem(enCartItem_AskForBackOrder) & "<br />"
	Response.Write "</fieldset>"
End Sub

'*******************************************************************************************************

Sub WriteThankYouMessage()

Dim paryCartItem
Dim pblnAddToCart
Dim plngCartItemCounter
Dim plngNumCartItems
Dim plngNumItemsAdded
Dim plngOrderQty
Dim plngBOQty
Dim pstrURL

	'On Error Resume Next

	maryCartAdditions = getCartAdditionResultsFromSession
	Call DebugRecordSplitTime("Loaded cart addition results from session . . . " & isArray(maryCartAdditions))
	If isArray(maryCartAdditions) Then

		If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
			pstrURL = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING") & "&"
		Else
			pstrURL = Request.ServerVariables("URL") & "?"
		End If
		
		'strip any of our actions
		pstrURL = Replace(pstrURL, "&btnAction=NoThanks", "")
		pstrURL = Replace(pstrURL, "&btnAction=BackOrder", "")

		plngNumItemsAdded = 0
		plngNumCartItems = UBound(maryCartAdditions)
		pblnAddToCart = Not maryCartAdditions(0)(enCartItem_AddType)
		For plngCartItemCounter = 0 To plngNumCartItems
			paryCartItem = maryCartAdditions(plngCartItemCounter)
			plngNumItemsAdded = plngNumItemsAdded + paryCartItem(enCartItem_QtyAdded)
			'Call writeCartItem(paryCartItem)
%>
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="100%" align="center" id="tblThankyouMessage">
  <tr>
	<td>
	  <table width="100%" border="0" cellspacing="1" cellpadding="3">
		<tr>
		  <td class="tdContent2" align="center">
			<div id="ssThanks"><hr noshade size="1" width="90%" /></div>
			<% If paryCartItem(enCartItem_QtyAdded) = 0 Then %>
			  <font color="red">
			  <font class="Content_Large"><b>We're Sorry!</b></font><br /><br />
			  <b><%= paryCartItem(enCartItem_prodName) %> Stock Availability</b>
			  <b><%= paryCartItem(enCartItem_ResponseMessage) %></b>
			  <p style="color:red">Sorry, we don't have your requested quantity in stock.
			  <% If paryCartItem(enCartItem_QtyInCart) > 0 Then %> There are <b><%= paryCartItem(enCartItem_QtyInCart) %></b> of these item(s) in your cart.<% End If %>
			  <% If paryCartItem(enCartItem_invenbBackOrder) = 1 Then %> If you would like us to send the product as soon as it arrives please click the Back Order button below.<% End If %>
			  </p>
			  <b><%= paryCartItem(enCartItem_Upsell) %></b>
			  </font>
			<% ElseIf paryCartItem(enCartItem_QtyAdded) = -1 Then %>
			  <font class="Content_Large"><b><font color="red">We're Sorry!</font></b></font><br />
			  <b><%= paryCartItem(enCartItem_ResponseMessage) %></b><br />
			  <b><%= paryCartItem(enCartItem_Upsell) %></b>
			<% ElseIf paryCartItem(enCartItem_QtyToAdd) <> paryCartItem(enCartItem_QtyAdded) Then %>
			  <font class="Content_Large"><b><font color="red">We're Sorry!</font></b></font><br />
			  <p style="color:red">Sorry, we don't have your requested quantity in stock.</p>
			  <b><%= paryCartItem(enCartItem_QtyAdded) %>&nbsp;<%= paryCartItem(enCartItem_prodName) %>&nbsp;<%= getCartActionText(paryCartItem) %></b><br /><br />
			  <% If paryCartItem(enCartItem_invenbBackOrder) = 1 Then %> If you would like us to send the product as soon as it arrives please click the Back Order button below.<% End If %>
			  <b><%= paryCartItem(enCartItem_Upsell) %></b>
			<% Else %>
			  <font class="Content_Large"><b>Thank You!</b></font><br />
			  <b><%= paryCartItem(enCartItem_QtyAdded) %>&nbsp;<%= paryCartItem(enCartItem_prodName) %>&nbsp;<%= getCartActionText(paryCartItem) %></b><br /><br />
			  <b><%= paryCartItem(enCartItem_Upsell) %></b>
			<% End If %>
			<hr noshade size="1" width="90%">
			<%
			 If plngCartItemCounter = plngNumCartItems Then
				If isOrderPage Then
			%>
			<a href="<%= SearchPath %>"><img src="<%= C_BTN04 %>" alt="Continue Shopping" border="0"></a>
			<%
			   ElseIf plngNumItemsAdded > 0 Then
					If pblnAddToCart Then
						Response.Write "<div id=btnssCheckout align=center>"
						Response.Write "<a href='javascript:hideThankYou();'>Close</a> "
						Response.Write "<a href='" & C_HomePath & "order.asp'><img src='images/buttons/checkout.gif' alt='Checkout' border='0' name=imgAddToCart id=imgAddToCart /></a>"
						'Call DisplayMiniCart_Thankyou
						Response.Write "</div>"
					Else
						Response.Write "<div id=btnssCheckout align=center><a href='" & C_HomePath & "savecart.asp'><img src='images/buttons/vscart.gif' alt='Wish List' border='0' /></a></div>"
					End If
			   End If 'isOrderPage
			 End If 'plngCartItemCounter = plngNumCartItems
			%>
		  </td>        
		</tr>
		<tr>
		  <td align="center" class="tdContent2">
			<%
			If paryCartItem(enCartItem_AskForBackOrder) And False Then
				'Need to set order qty, bo qty
				plngOrderQty = paryCartItem(enCartItem_QtyToAdd) + paryCartItem(enCartItem_QtyInCart)
				plngBOQty = plngOrderQty - paryCartItem(enCartItem_QtyInStock)
				%><a href="<%= pstrURL %>btnAction=BackOrder&amp;BackOrderPos=<%= paryCartItem(enCartItem_tmpOrderDetailId) %>&amp;BackOrderCount=<%= plngBOQty %>&amp;OrderQty=<%= plngOrderQty %>">Back Order</a><%
			End If

			If paryCartItem(enCartItem_AskForBackOrder) And ConvertToBoolean(getConfigurationSettingFromCache("NotifyMeEnabled", False), False) Then
				%>
				<script language="javascript" type="text/javascript">
				<!--
				function switchNotifyMeDisplay(blnDisplay)
				{
					if (blnDisplay)
					{
						document.getElementById("aNotifyMe").style.display="none";
						document.getElementById("divNotifyMe").style.display="";
					}else{
						document.getElementById("aNotifyMe").style.display="";
						document.getElementById("divNotifyMe").style.display="none";
					}
				}
				
				function validateNotifyMe(theForm)
				{
					if (!isValidEmailAddress(theForm.notifyEmail.value))
					{
						alert("Please enter an email address.");
						theForm.notifyEmail.focus();
						return false;
					}
					
					return true;
				}
				-->
				</script>
				<br /><a id="aNotifyMe" href="<%= pstrURL %>btnAction=notifyMe&amp;BackOrderPos=<%= paryCartItem(enCartItem_tmpOrderDetailId) %>" onclick="switchNotifyMeDisplay(true);return false;">Notify Me by email when this arrives!</a>
				<div id="divNotifyMe" style="display:none">
				<form name="frmNotifyMe" id="frmNotifyMe" action="<%= pstrURL %>" method="get" onsubmit="return validateNotifyMe(this);"><% Call writeQuerystringParametersToHiddenValues(Array("notifyLastName", "notifyFirstName", "notifyEmail", "btnAction")) %>
				<input type="hidden" name="BackOrderPos" id="BackOrderPos" value="<%= paryCartItem(enCartItem_tmpOrderDetailId) %>">
				<table border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td><label for="notifyLastName">Last Name:</label> </td>
				    <td><input type="text" name="notifyLastName" id="notifyLastName" value="" /></td>
				  </tr>
				  <tr>
				    <td><label for="notifyFirstName">First Name:</label> </td>
				    <td><input type="text" name="notifyFirstName" id="notifyFirstName" value="" /></td>
				  </tr>
				  <tr>
				    <td><label for="notifyEmail">E-Mail Address:</label> </td>
				    <td><input type="text" name="notifyEmail" id="notifyEmail" value="" /></td>
				  </tr>
				  <tr>
				    <td colspan="2">
				      <input type="submit" name="btnAction" id="btnAction" value="notifyMe">
				      &nbsp;<a href="javascript: switchNotifyMeDisplay(false);">Cancel</a>
				    </td>
				  </tr>
				</table>
				</form>
				</div>
				<%
			End If
			%>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
</table>
<% 
		Next 'plngCartItemCounter
		Call writeScrollToTopJavascript 
		Call clearCartAdditionResultsFromSession
	ElseIf Len(maryCartAdditions) > 0 Then
%>
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="100%" align="center">
  <tr>
	<td>
	  <table width="100%" border="0" cellspacing="1" cellpadding="3">
		<tr>
		  <td class="tdContent2" align="center">
			<hr noshade size="1" width="90%" />
			<%= maryCartAdditions %>
			<hr noshade size="1" width="90%">
		  </td>        
		</tr>
	  </table>
	</td>
  </tr>
</table>
<% 
		Call writeScrollToTopJavascript 
		Call clearCartAdditionResultsFromSession
	Else
'		Call CheckAcceptBackOrder
	End If	'isArray(maryCartAdditions)
	
End Sub	'WriteThankYouMessage

'*******************************************************************************************************

Sub CheckAcceptBackOrder()

Dim plngtmpOrderDetailId
Dim plngBackOrderCount
Dim plngOrderQty
Dim pobjCmd
Dim pobjRS

	If Request.Form("btnAction") = "Back Order" Then

	ElseIf Request.Querystring("btnAction") = "BackOrder" Then
		plngtmpOrderDetailId = Request.QueryString("BackOrderPos")
		If Len(plngtmpOrderDetailId) = 0 Or Not isNumeric(plngtmpOrderDetailId) Then Exit Sub

		plngBackOrderCount = Request.QueryString("BackOrderCount")
		plngOrderQty = Request.QueryString("OrderQty")
		
		Set pobjCmd = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			Set .ActiveConnection = cnn
			.Commandtext = "Update sfTmpOrderDetails Set odrdttmpQuantity=? WHERE odrdttmpID=?"

			If Len(plngOrderQty) > 0 And isNumeric(plngOrderQty) Then
				.Parameters.Append .CreateParameter("Quantity", adInteger, adParamInput, 4, plngOrderQty)
				.Parameters.Append .CreateParameter("odrdttmpID", adInteger, adParamInput, 4, plngtmpOrderDetailId)
				.Execute , , adExecuteNoRecords
			End If
			
			If Len(plngBackOrderCount) > 0 And isNumeric(plngBackOrderCount) Then
				.Commandtext = "Update sfTmpOrderDetailsAE Set odrdttmpBackOrderQTY=? WHERE odrdttmpAEID=?"
				.Parameters("Quantity").Value = plngBackOrderCount
				.Execute , , adExecuteNoRecords
			End If
			
			'This is to figure out what was back-ordered for display
			.Commandtext = "SELECT sfTmpOrderDetails.odrdttmpProductID, sfTmpOrderAttributes.odrattrtmpAttrID, sfProducts.prodName FROM (sfTmpOrderDetails INNER JOIN sfTmpOrderAttributes ON sfTmpOrderDetails.odrdttmpID = sfTmpOrderAttributes.odrattrtmpOrderDetailId) INNER JOIN sfProducts ON sfTmpOrderDetails.odrdttmpProductID = sfProducts.prodID WHERE sfTmpOrderDetails.odrdttmpID=?"
			.Parameters.Delete "Quantity"
			Set pobjRS = .Execute
			If Not pobjRS.EOF Then

				If Not isOrderPage Then
					maryCartAdditions = "<div id=btnssCheckout align=center>" _
									  & "<font class=""Content_Large"">" & pobjRS.Fields("prodName").Value & " has been added to your cart.</p>" _
									  & "<a href='" & C_HomePath & "order.asp'><img src='images/buttons/checkout.gif' alt='Checkout' border='0' name=imgAddToCart id=imgAddToCart /></a>" _
									  & "</div>"
					setCartAdditionResultsToSession
				End If	'isOrderPage
			End If
			pobjRS.Close
			Set pobjRS = Nothing
		End With
		Set pobjCmd = Nothing
	ElseIf Request.QueryString("btnAction") = "notifyMe" And Not isCartInSession Then

		plngtmpOrderDetailId = Request.QueryString("BackOrderPos")
		If Len(plngtmpOrderDetailId) = 0 Or Not isNumeric(plngtmpOrderDetailId) Then Exit Sub
	
		Dim notifyLastName
		Dim notifyFirstName
		Dim notifyEmail

		notifyLastName = Trim(Request.QueryString("notifyLastName"))
		notifyFirstName = Trim(Request.QueryString("notifyFirstName"))
		notifyEmail = Trim(Request.QueryString("notifyEmail"))
		
		If Len(notifyEmail) > 0 Then
			If saveNotifyMe(plngtmpOrderDetailId, notifyLastName, notifyFirstName, notifyEmail) Then
				maryCartAdditions = "<div id=btnssCheckout align=center>" _
									& "<font class=""Content_Large"">Thank You! You will be notified at <em>" & notifyEmail & "</em> when the product is back in stock.</p>" _
									& "</div>"
				setCartAdditionResultsToSession
			End If
		Else

		End If
		
	End If	'Request.Querystring("btnAction") = "BackOrder" 
	
End Sub	'CheckAcceptBackOrder

'*******************************************************************************************************

Function SearchPath

Dim pstrSearchPath

	pstrSearchPath = Request.Cookies("sfSearch")("SearchPath")
	If Len(pstrSearchPath) > 0 Then
		If InStr(LCase(pstrSearchPath), "login.asp") <> 0 Then pstrSearchPath = "search.asp"
	Else
		pstrSearchPath = "search.asp"
	End If
	
	SearchPath = pstrSearchPath

End Function	'SearchPath

'*******************************************************************************************************

Sub writeScrollToTopJavascript

	If Not cblnScrollToTop Then Exit Sub
%>
<script language="javascript" type="text/javascript">
<!--
function popupEliminatorScrollToTop()
{
	var elem = document.getElementById("tblThankyouMessage");
	if (elem != null)
	{
		elem.scrollIntoView(true);
		elem.focus();
	}

	//window.scroll(0,0);
}

function hideThankYou()
{
	document.getElementById("tblThankyouMessage").style.display="none";
	var elem = document.getElementById("divCategoryTrail");
	if (elem != null)
	{
		elem.scrollIntoView(true);
		elem.focus();
	}
}

window.onload=popupEliminatorScrollToTop
-->
</script>
<%
End Sub	'writeScrollToTopJavascript
%>