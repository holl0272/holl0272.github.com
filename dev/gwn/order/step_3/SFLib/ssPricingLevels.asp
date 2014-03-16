<%
'********************************************************************************
'*   Pricing Level					                                            *
'*   Release Version:	1.01.003												*
'*   Release Date:		November 1, 2002										*
'*   Revision Date:		March 22, 2004											*
'*                                                                              *
'*   Release 1.01.003	(March 22, 2004)										*
'*    - bug fix - added function back in for Product Bots - lost in prior merge *
'*                                                                              *
'*   Release 1.01.002	(Decmber 28, 2003)										*
'*    - Note: Merged fix below back into this release file			            *
'*    - bug fix - incorrect pricing was being saved to sfOrdersDetail           *
'*                                                                              *
'*   Release 1.01.001	(November 6, 2003)										*
'*    - Note: Altered version system to use common implementation	            *
'*    - Note: Updated debugging to use common implementation		            *
'*                                                                              *
'*   Release 1.1  (September 15, 2002)                                          *
'*    - bug fix - corrected SQL adjustment for search                           *
'*                                                                              *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'**********************************************************
'	Developer notes
'**********************************************************

'for testing
'Session("custPricingLevel") = "1"
'Session("custPricingLevel") = ""

'**********************************************************
'*	Page Level variables
'**********************************************************

Const cblnUsePricingLevel = True

Dim cblnDebugPricingLevel
Dim mstrPricingLevel
Dim mblnAllowSalePrice

'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'*	Begin Page Code
'**********************************************************

Call setPricingLevel

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/


'/
'/////////////////////////////////////////////////

cblnDebugPricingLevel = Len(Session("ssDebug_PricingLevel")) > 0
If cblnDebugPricingLevel And mstrPricingLevel >= 0 Then debugprint "mstrPricingLevel",mstrPricingLevel

'mblnAllowSalePrice = True And mstrPricingLevel = -1	'show sale price only if no pricing level
mblnAllowSalePrice = True 'And Len(mstrPricingLevel) = 0	'show sale price only if no pricing level
	
'**********************************************************
'*	Begin Function Definitions
'**********************************************************


'********************************************************************************

Sub setPricingLevel()

	mstrPricingLevel = Session("custPricingLevel")
	'The below line is only to be used when an incompatible (read older) version of Customer Manager was installed and pricing levels show up off by one
	'If Len(mstrPricingLevel) > 0 And isNumeric(mstrPricingLevel) Then mstrPricingLevel = CLng(mstrPricingLevel) - 1

	'The below line is only to be used for the Master Template where 0 --> no pricing level, 1 and up are pricing levels
	If Len(mstrPricingLevel) = 0 or Not isNumeric(mstrPricingLevel) Then
		mstrPricingLevel = -1
	Else
		mstrPricingLevel = CLng(mstrPricingLevel) - 1
	End If

End Sub	'setPricingLevel

'********************************************************************************

Sub AdjustRecordPricingLevel(byRef objRS, byVal strSwitch)

'	On Error Resume Next

	If cblnUsePricingLevel Then
		If isObject(objRS) Then
			If Not objRS.EOF Then
				Select Case strSwitch
					Case "SearchProduct"
						Do While Not objRS.Eof
							objRS.Fields("prodPrice").Value = GetPricingLevelPrice(objRS.Fields("prodPrice").Value, objRS.Fields("prodPLPrice").Value)
							objRS.Fields("prodSalePrice").Value = GetPricingLevelPrice(objRS.Fields("prodSalePrice").Value, objRS.Fields("prodPLSalePrice").Value)
							objRS.MoveNext
						Loop
						objRS.MoveFirst
					Case "SearchAttDetail"
						Do While Not objRS.Eof
							objRS.Fields("attrdtPrice").Value = GetPricingLevelPrice(objRS.Fields("attrdtPrice").Value, objRS.Fields("attrdtPLPrice").Value)
							objRS.MoveNext
						Loop
						objRS.MoveFirst
					Case "mtprices"
						Do While Not objRS.Eof
							objRS.Fields("mtValue").Value = GetPricingLevelPrice(objRS.Fields("mtValue").Value, objRS.Fields("mtPLValue").Value)
							objRS.MoveNext
						Loop
						objRS.MoveFirst
					Case "Detail"
						Do While Not objRS.Eof
							objRS.Fields("prodPrice").Value = GetPricingLevelPrice(objRS.Fields("prodPrice").Value, objRS.Fields("prodPLPrice").Value)
							objRS.Fields("prodSalePrice").Value = GetPricingLevelPrice(objRS.Fields("prodSalePrice").Value, objRS.Fields("prodPLSalePrice").Value)
							objRS.MoveNext
						Loop
						objRS.MoveFirst
					Case "DetailAttr"
						Do While Not objRS.Eof
							objRS.Fields("attrdtPrice").Value = GetPricingLevelPrice(objRS.Fields("attrdtPrice"), objRS.Fields("attrdtPLPrice"))
							objRS.MoveNext
						Loop
						objRS.MoveFirst
					Case Else
						
				End Select
			End If
		End If
	End If
	
End Sub	'AdjustRecordPricingLevel

'********************************************************************************

Function AdjustSQLPricingLevel(byVal strSQL, byVal strSwitch)

Dim pstrTempSQL

	pstrTempSQL = strSQL
	If cblnUsePricingLevel Then
		Select Case strSwitch
			Case "SearchProduct"
				'modified specifically for use with Attribute Extender??
				'pstrTempSQL = Replace(pstrTempSQL,"prodCategoryId, prodShortDescription","prodCategoryId, prodShortDescription, sfProducts.prodPLPrice,sfProducts.prodPLSalePrice ",1,1,1)

				If Instr(1,pstrTempSQL,"sfProducts.prodShortDescription") > 0 Then
					pstrTempSQL = Replace(pstrTempSQL,"sfProducts.prodShortDescription like","zzzroducts.prodShortDescription like",1,1,1)
					'pstrTempSQL = Replace(pstrTempSQL,"sfProducts.prodShortDescription","sfProducts.prodShortDescription, sfProducts.prodPLPrice,sfProducts.prodPLSalePrice ",1,1,1)
					pstrTempSQL = Replace(pstrTempSQL,"prodShortDescription","prodShortDescription, sfProducts.prodPLPrice,sfProducts.prodPLSalePrice ",1,1,1)
					pstrTempSQL = Replace(pstrTempSQL,"zzzroducts.prodShortDescription like","sfProducts.prodShortDescription like",1,1,1)
				Else
					pstrTempSQL = Replace(pstrTempSQL,"prodShortDescription","zzzShortDescription",1,1,1)
					pstrTempSQL = Replace(pstrTempSQL,"zzzShortDescription","prodShortDescription, sfProducts.prodPLPrice,sfProducts.prodPLSalePrice ",1,1,1)
				End If
			Case "SearchAttDetail"
				pstrTempSQL = Replace(pstrTempSQL,"attrdtOrder","attrdtOrder, attrdtPLPrice",1,1)
			Case "Detail"
				If Instr(1,pstrTempSQL,"sfProducts.prodSalePrice") > 0 Then
					pstrTempSQL = Replace(pstrTempSQL,"sfProducts.prodSalePrice,","sfProducts.prodSalePrice, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice, ",1,1)
				Else
					pstrTempSQL = Replace(pstrTempSQL,"prodSalePrice,","prodSalePrice, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice, ",1,1)
				End If
			Case "DetailAttr"
				pstrTempSQL = Replace(pstrTempSQL,"attrdtPrice,","attrdtPrice, attrdtPLPrice, ",1,1)
			Case Else
				'pstrTempSQL = strSQL
		End Select
	End If
	
	If cblnDebugPricingLevel Then
		Response.Write strSwitch & " - pstrTempSQL: " & pstrTempSQL & "<br />"
	End If


	AdjustSQLPricingLevel = pstrTempSQL

End Function	'AdjustSQLPricingLevel

'********************************************************************************

Function GetDetailProductSQL(blnSFAE)

	If blnSFAE Then	

		GetDetailProductSQL = " SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodShortDescription, sfProducts.prodImageSmallPath," _
							& " sfProducts.prodImageLargePath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodAttrNum," _
							& " sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodMessage, sfProducts.prodDescription," _
							& " sfProducts.prodPLPrice,sfProducts.prodPLSalePrice," _
							& " sfSub_Categories.CatHierarchy FROM (sfProducts INNER JOIN sfsubCatdetail ON sfProducts.prodID = sfsubCatdetail.ProdID)" _
							& " INNER JOIN sfSub_Categories ON sfsubCatdetail.subcatCategoryId = sfSub_Categories.subcatID WHERE sfProducts.prodID = '" & txtProdId & "'"
	Else
		GetDetailProductSQL = "SELECT prodID, prodName, prodShortDescription, prodImageSmallPath, " _
							& "prodImageLargePath, prodLink, prodPrice, prodAttrNum, catName, " _
							& "prodSaleIsActive, prodSalePrice, prodMessage, prodDescription, " _
							& " sfProducts.prodPLPrice,sfProducts.prodPLSalePrice," _
							& "FROM sfProducts " _
							& "INNER JOIN sfCategories ON sfProducts.prodCategoryID = sfCategories.catID " _
							& "WHERE prodID = '" & txtProdId & "'"
	End If
	
End Function	'GetDetailProductSQL

'********************************************************************************

Function GetDetailProductAttrSQL()

	GetDetailProductAttrSQL = "SELECT attrName, attrdtId, attrdtName, attrdtPrice, attrdtType, attrdtOrder, " _
							& " attrdtPrice" & Session("custPricingLevel") & " as SpecialPricing_attr " _
							& "FROM sfAttributes " _
							& "INNER JOIN sfAttributeDetail ON sfAttributes.attrId = sfAttributeDetail.attrdtAttributeId " _
							& "WHERE attrProdId = '" & rsProdDetail("prodId") & "' ORDER BY AttrName, attrdtOrder"
	
End Function	'GetDetailProductAttrSQL

'********************************************************************************

Function GetPricingLevelPrice(byVal strDefaultPrice, byVal strPrices)

Dim paryTemp
Dim pstrTemp

	If Len(mstrPricingLevel) > 0 Then
		If Len(strPrices) > 0 Then
			paryTemp = Split(strPrices,";")
			If mstrPricingLevel >=0 And UBound(paryTemp) >= CLng(mstrPricingLevel) Then
				pstrTemp = Trim(paryTemp(mstrPricingLevel))
				If Not isNumeric(pstrTemp) Then pstrTemp = ""	'just in case bad data has been added
			End If
		End If
	End If
	
	If Len(pstrTemp) = 0 Then pstrTemp = strDefaultPrice

	If cblnDebugPricingLevel Then
		Response.Write "<fieldset><legend>GetPricingLevelPrice</legend>"
		Response.Write "Pricing Level: " & mstrPricingLevel & "<br />"
		Response.Write "Default Price: " & strDefaultPrice & "<br />"
		Response.Write "Prices: " & strPrices & "<br />"
		Response.Write "Price Out: " & pstrTemp
		Response.Write "</fieldset>"
	End If
	
	'Add check for non-numeric entries, just in case
	If Len(pstrTemp) = 0 Or Not isNumeric(pstrTemp) Then pstrTemp = 0

	GetPricingLevelPrice = pstrTemp
	
End Function	'GetPricingLevelPrice

'********************************************************************************

Function getProduct(byVal strProdID)
'Purpose: Returns and array of product information for a given product ID
'Input: prodID
'Return: array (3)
'			(0) - Product Name
'			(1) - Sell Price
'			(2) - Attribute Count
'			(3) - Empty

Dim paryProduct(3)
Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		If True Then
			pstrSQL = "SELECT prodName, prodNamePlural, prodPrice, prodAttrNum, prodSaleIsActive, prodSalePrice, prodPLPrice, prodPLSalePrice" _
					& " FROM sfProducts" _
					& " WHERE prodEnabledIsActive=1 AND prodID = ?"
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
		Else
			pstrSQL = "spGetProduct"
			.Commandtype = adCmdStoredProc
			.Commandtext = pstrSQL
		End IF
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProdID)
		Set pobjRS = .Execute
	End With	'pobjCmd
		
	With pobjRS
		If .EOF Then
			If vDebug = 1 Then Response.Write "<p>Empty Recordset in pobjRS. Product " & sProdID & " possibly not activated."
		Else
			paryProduct(0) = .Fields("prodName").Value
			' Check if sale price is active 
			If .Fields("prodSaleIsActive").Value = 1 And mblnAllowSalePrice Then
				paryProduct(1) = GetPricingLevelPrice(.Fields("prodSalePrice").Value, .Fields("prodPLSalePrice").Value)
			Else 	
				paryProduct(1) = GetPricingLevelPrice(.Fields("prodPrice").Value, .Fields("prodPLPrice").Value)
			End If		
			paryProduct(2) = .Fields("prodAttrNum").Value

			If cblnDebugPricingLevel Then
				Response.Write "<fieldset><legend>getProduct (" & strProdID & ")</legend>"
				Response.Write "Name: paryProduct(0): " & paryProduct(0) & "<br />"
				Response.Write "Sell Price: paryProduct(1): " & paryProduct(1) & "<br />"
				Response.Write "Attribute Num: paryProduct(2): " & paryProduct(2) & "<br />"
				Response.Write "</fieldset>"
			End If

		End If	'.EOF
	End With	'pobjRS
	
	closeObj(pobjRS)
	closeObj(pobjCmd)

	getProduct = paryProduct

End Function	'getProduct

'********************************************************************************

Function getProductInfoPL(byRef objRS, byVal strSwitch)

'	On Error Resume Next

	If cblnUsePricingLevel Then
		If isObject(objRS) Then
			If Not objRS.EOF Then
				Select Case strSwitch
						Case "prodPrice":	getProductInfoPL = GetPricingLevelPrice(objRS.Fields("prodPrice").Value, objRS.Fields("prodPLPrice").Value)
					Case "prodSalePrice":	getProductInfoPL = GetPricingLevelPrice(objRS.Fields("prodSalePrice").Value, objRS.Fields("prodPLSalePrice").Value)
					Case Else
						
				End Select
			End If
		End If
	Else
		getProductInfoPL = objRS.Fields(strSwitch).Value
	End If
	
End Function	'getProductInfoPL

'********************************************************************************

Function getAttrDetails(byVal lngAttrID)
'Purpose: Returns and array of attribute information for a given attrID
'Input: attrID
'Return: array (3)
'			(0) - Attribute Detail Name
'			(1) - Attribute Price
'			(2) - Attribute Type
'			(3) - Attribute Name
'			(4) - Empty

Dim paryProduct(4)
Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	Call CleanAttributeID(lngAttrID)

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		If True Then
			pstrSQL = "SELECT attrName, attrdtName, attrdtPrice, attrdtType, attrdtPLPrice" _
					& " FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId" _
					& " WHERE attrdtID = ?"
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
		Else
			pstrSQL = "spGetAttrDetails"
			.Commandtype = adCmdStoredProc
			.Commandtext = pstrSQL
		End IF
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("attrID", adInteger, adParamInput, 4, lngAttrID)
		Set pobjRS = .Execute
	End With	'pobjCmd
		
	With pobjRS
		If .EOF Then
			If vDebug = 1 Then Response.Write "<p>Empty Recordset in getAttrDetails. Attribute: " & lngAttrID & "."
		Else

			paryProduct(0) = .Fields("attrdtName").Value
			paryProduct(1) = GetPricingLevelPrice(.Fields("attrdtPrice").Value, .Fields("attrdtPLPrice").Value)
			paryProduct(2) = .Fields("attrdtType").Value
			paryProduct(3) = .Fields("attrName").Value

			If cblnDebugPricingLevel Then
				Response.Write "<fieldset><legend>getAttrDetails (" & lngAttrID & ")</legend>"
				Response.Write "Attribute Detail Name: paryProduct(0): " & paryProduct(0) & "<br />"
				Response.Write "Attribute Price: paryProduct(1): " & paryProduct(1) & "<br />"
				Response.Write "Attribute Type: paryProduct(2): " & paryProduct(2) & "<br />"
				Response.Write "Attribute Name: paryProduct(3): " & paryProduct(3) & "<br />"
				Response.Write "</fieldset>"
			End If

		End If	'.EOF
	End With	'pobjRS
	
	closeObj(pobjRS)
	closeObj(pobjCmd)

	getAttrDetails = paryProduct

End Function	'getAttrDetails

'********************************************************************************

Function get6ProdValues(byVal strProdID)
'Purpose: Returns and array of product information for a given product ID
'Input: prodID
'Return: array (5)
'			(0) - Category
'			(1) - Manufacturer
'			(2) - Vendor
'			(3) - Product Name
'			(4) - Attribute Count
'			(5) - Sell Price

Dim paryProduct(5)
Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		If True Then
			pstrSQL = "SELECT prodCategoryID, prodManufacturerID, prodVendorID, prodName, prodPrice, prodAttrNum, prodSaleIsActive, prodSalePrice, prodPLPrice, prodPLSalePrice" _
					& " FROM sfProducts" _
					& " WHERE prodEnabledIsActive=1 AND prodID = ?"
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
		Else
			pstrSQL = "spGet6ProdValues"
			.Commandtype = adCmdStoredProc
			.Commandtext = pstrSQL
		End IF
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProdID)
		Set pobjRS = .Execute
	End With	'pobjCmd
		
	With pobjRS
		If .EOF Then
			If vDebug = 1 Then Response.Write "<p>Empty Recordset in pobjRS. Product " & sProdID & " possibly not activated."
		Else
  			paryProduct(0) = getNameWithID("sfCategories", .Fields("prodCategoryID").Value, "catID", "catName", 0)
  			paryProduct(1) = getNameWithID("sfManufacturers", .Fields("prodManufacturerID").Value, "mfgID", "mfgName", 0)
  			paryProduct(2) = getNameWithID("sfVendors", .Fields("prodVendorID").Value, "vendID", "vendName", 0)
  			paryProduct(3) = .Fields("prodName").Value
  			paryProduct(4) = .Fields("prodAttrNum").Value
			If .Fields("prodSaleIsActive").Value = 1 And mblnAllowSalePrice Then
				paryProduct(5) = GetPricingLevelPrice(.Fields("prodSalePrice").Value, .Fields("prodPLSalePrice").Value)
			Else 	
				paryProduct(5) = GetPricingLevelPrice(.Fields("prodPrice").Value, .Fields("prodPLPrice").Value)
			End If

			If cblnDebugPricingLevel Then
				Response.Write "<fieldset><legend>getProduct (" & strProdID & ")</legend>"
				Response.Write "Category: paryProduct(0): " & paryProduct(0) & "<br />"
				Response.Write "Manufacturer: paryProduct(1): " & paryProduct(1) & "<br />"
				Response.Write "Vendor: paryProduct(2): " & paryProduct(2) & "<br />"
				Response.Write "Name: paryProduct(3): " & paryProduct(3) & "<br />"
				Response.Write "Attribute Num: paryProduct(4): " & paryProduct(4) & "<br />"
				Response.Write "Sell Price: paryProduct(5): " & paryProduct(5) & "<br />"
				Response.Write "</fieldset>"
			End If

		End If	'.EOF
	End With	'pobjRS
	
	closeObj(pobjRS)
	closeObj(pobjCmd)

	get6ProdValues = paryProduct

End Function	'get6ProdValues
%>
