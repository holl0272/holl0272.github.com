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

Const enProduct_ID = 1
Const enProduct_Name = 2
Const enProduct_ShortDescription = 3
Const enProduct_Description = 4
Const enProduct_ImageSmallPath = 5
Const enProduct_ImageLargePath = 6
Const enProduct_Link = 7
Const enProduct_Price = 8
Const enProduct_SalePrice = 9
Const enProduct_CategoryID = 10
Const enProduct_MfgID = 11
Const enProduct_VendorID = 12
Const enProduct_NamePlural = 13
Const enProduct_IsActive = 14
Const enProduct_SaleIsActive = 15

Const enProduct_PLPrice = 16
Const enProduct_PLSalePrice = 17
Const enProduct_ShipIsActive = 18
Const enProduct_Weight = 19
Const enProduct_Ship = 20
Const enProduct_StateTaxIsActive = 21
Const enProduct_CountryTaxIsActive = 22

'AE specific values
Const enProduct_gwPrice = 23
Const enProduct_gwActivate = 24
Const enProduct_invenbTracked = 25
Const enProduct_invenInStock = 26
Const enProduct_invenbStatus = 27
Const enProduct_invenbBackOrder = 28
Const enProduct_invenbNotify = 29

Const enProduct_Message = 30
Const enProduct_AttrNum = 31

Const enProduct_prodMinQty = 32
Const enProduct_prodIncrement = 33

Const enProduct_Height = 34
Const enProduct_Length = 35
Const enProduct_Width = 36
Const enProduct_SellPrice = 37
Const enProduct_sfProductID = 38
Const enProduct_AdditionalImages = 39
Const enProduct_EnableReviews = 40
Const enProduct_EnableAlsoBought = 41
Const enProduct_HandlingFee = 42
Const enProduct_SetupFee = 43
Const enProduct_LimitQtyToMTP = 50
Const enProduct_relatedProducts = 51
Const enProduct_DisplayAdditionalImagesInWindow = 53
Const enProduct_pageName = 54
Const enProduct_metaTitle = 55
Const enProduct_metaDescription = 56
Const enProduct_metaKeywords = 57
Const enProduct_attributes = 58
Const enProduct_SetupFeeOneTime = 59

'Non-sfProduct table values
Const enProduct_Exists = 44
Const enProduct_CategoryName = 45
Const enProduct_MfgName = 46
Const enProduct_VendorName = 47
Const enProduct_AttributeArray = 48
Const enProduct_attributeIDs = 49
Const enProduct_MTP = 52

Const enProduct_ArrayLength = 59	'Make sure you set Dim maryProduct(59) below

Const enAttribute_ID = 0
Const enAttribute_Name = 1
Const enAttribute_DisplayStyle = 2
Const enAttribute_Display = 3
Const enAttribute_Image = 4
Const enAttribute_SKU = 5
Const enAttribute_URL = 6
Const enAttribute_Extra = 7
Const enAttribute_DetailArray = 8
Const enAttribute_DisplayTemplate = 9

Const enAttributeDetail_ID = 0
Const enAttributeDetail_Name = 1
Const enAttributeDetail_Price = 2
Const enAttributeDetail_Type = 3
Const enAttributeDetail_Image = 4
Const enAttributeDetail_PLPrice = 5
Const enAttributeDetail_FileName = 6
Const enAttributeDetail_Default = 7
Const enAttributeDetail_Display = 8
Const enAttributeDetail_SKU = 9
Const enAttributeDetail_URL = 10
Const enAttributeDetail_Weight = 11
Const enAttributeDetail_Extra = 12
Const enAttributeDetail_Extra1 = 13
Const enAttributeDetail_SellPrice = 14

Dim mstrCurrentProductID

'Wrapped in error handler since it is possible for maryProduct to be called prior to this declaration appearing
On Error Resume Next
Dim maryProduct(59)
If Err.number <> 0 Then Err.Clear

'**********************************************************
'*	Functions
'**********************************************************

'Function getProductInfo(byVal strProdID, byVal bytCase)

'**********************************************************
'*	Begin Page Code
'**********************************************************

Function getProductInfo(byVal strProdID, byVal bytCase)

Dim i
Dim pstrSQL
Dim pobjCmd
Dim pobjRS

	If Len(strProdID) = 0 Then
		If bytCase = enProduct_Exists Then
			getProductInfo = False
		Else
			getProductInfo = ""
		End If
		Exit Function
	End If

	If mstrCurrentProductID <> strProdID Then
		If cblnSF5AE Then
			pstrSQL = "SELECT prodName, prodNamePlural, prodShortDescription, prodImageSmallPath, prodMessage, prodDescription, prodAdditionalImages, relatedProducts, prodImageLargePath, prodLink, prodPrice, prodSalePrice, prodEnabledIsActive, prodSaleIsActive, prodAttrNum, prodPLPrice, prodPLSalePrice, prodShipIsActive, prodWeight, prodHeight, prodLength, prodWidth, prodShip, prodStateTaxIsActive, prodCountryTaxIsActive, prodIncrement, prodMinQty, prodCategoryId, prodManufacturerId, prodVendorId, sfProductID, prodEnableReviews, prodEnableAlsoBought, prodHandlingFee, prodSetupFee, prodSetupFeeOneTime, prodLimitQtyToMTP, prodDisplayAdditionalImagesInWindow, pageName, metaTitle, metaDescription, metaKeywords," _
					& " sfGiftWraps.gwActivate, sfGiftWraps.gwPrice, sfInventoryInfo.invenbTracked, sfInventoryInfo.invenbStatus, sfInventoryInfo.invenbBackOrder, sfInventoryInfo.invenbNotify" _
					& " FROM sfGiftWraps RIGHT JOIN (sfInventoryInfo RIGHT JOIN sfProducts ON sfInventoryInfo.invenProdId = sfProducts.prodID) ON sfGiftWraps.gwProdID = sfProducts.prodID" _
					& " Where prodID=?"

			'NOTE: split for SQL Server due to record size limitation (ie. truncates data)
			'Solution is to separate prodMessage, prodDescription, prodAdditionalImages, relatedProducts into another query
			If cblnSQLDatabase Then pstrSQL = Replace(pstrSQL, "prodMessage, prodDescription, prodAdditionalImages, relatedProducts, ", "")
		Else
			pstrSQL = "SELECT prodID, prodCategoryId, prodMessage, prodDescription, prodAdditionalImages, relatedProducts, prodManufacturerId, prodVendorId, prodName, prodNamePlural, prodShortDescription, prodImageSmallPath, prodImageLargePath, prodLink, prodPrice, prodWeight, prodShip, prodShipIsActive, prodCountryTaxIsActive, prodStateTaxIsActive, prodEnabledIsActive, prodAttrNum, prodSaleIsActive, prodSalePrice, prodDateAdded, prodDateModified, prodLength, prodWidth, prodHeight, prodFileName, prodPLPrice, prodPLSalePrice, prodMaxDownloads, prodDownloadValidFor, sortCat, sortMfg, sortVend, sfProductID, prodEnableReviews, prodEnableAlsoBought, prodHandlingFee, prodSetupFee, prodMinQty, prodIncrement, prodLimitQtyToMTP, prodFixedShippingCharge, prodSetupFeeOneTime, prodSpecialShippingMethods, prodDisplayAdditionalImagesInWindow, pageName, metaTitle, metaDescription, metaKeywords" _
					& " FROM sfproducts" _
					& " Where prodID=?"
		End If
				
		Call DebugRecordTime("<b>Loading product <em>" & strProdID & "</em></b>")
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProdID)
			Set pobjRS = .Execute
		End With	'pobjCmd
		Call DebugRecordTime("Product retrieved from database . . . loading to array")

		'Call hack_preReadRecordset(pobjRS)
		With pobjRS
			If Not isArray(maryProduct) Then
				ReDim maryProduct(enProduct_ArrayLength)
			End If
			maryProduct(enProduct_Exists) = CBool(Not .EOF)
			If maryProduct(enProduct_Exists) Then
				mstrCurrentProductID = strProdID
				maryProduct(enProduct_ID) = strProdID
				
				maryProduct(enProduct_Name) = Trim(.Fields("prodName").Value & "")
				maryProduct(enProduct_NamePlural) = Trim(.Fields("prodNamePlural").Value & "")
				maryProduct(enProduct_ShortDescription) = Trim(.Fields("prodShortDescription").Value & "")
				maryProduct(enProduct_ImageSmallPath) = Trim(.Fields("prodImageSmallPath").Value & "")
				maryProduct(enProduct_ImageLargePath) = Trim(.Fields("prodImageLargePath").Value & "")
				maryProduct(enProduct_Link) = Trim(.Fields("prodLink").Value & "")
				maryProduct(enProduct_Price) = Trim(.Fields("prodPrice").Value & "")
				maryProduct(enProduct_SalePrice) = Trim(.Fields("prodSalePrice").Value & "")
				maryProduct(enProduct_CategoryID) = Trim(.Fields("prodCategoryId").Value & "")
				maryProduct(enProduct_MfgID) = Trim(.Fields("prodManufacturerId").Value & "")
				maryProduct(enProduct_VendorID) = Trim(.Fields("prodVendorId").Value & "")
				maryProduct(enProduct_IsActive) = CorrectEmptyValue(.Fields("prodEnabledIsActive").Value, 0)
				maryProduct(enProduct_SaleIsActive) = CorrectEmptyValue(.Fields("prodSaleIsActive").Value, 0)
				maryProduct(enProduct_AttrNum) = CorrectEmptyValue(Trim(.Fields("prodAttrNum").Value), 0)
				maryProduct(enProduct_PLPrice) = Trim(.Fields("prodPLPrice").Value & "")
				maryProduct(enProduct_PLSalePrice) = Trim(.Fields("prodPLSalePrice").Value & "")
				maryProduct(enProduct_ShipIsActive) = Trim(.Fields("prodShipIsActive").Value & "")
				maryProduct(enProduct_Weight) = Trim(.Fields("prodWeight").Value & "")
				maryProduct(enProduct_Ship) = Trim(.Fields("prodShip").Value & "")
				maryProduct(enProduct_StateTaxIsActive) = Trim(.Fields("prodStateTaxIsActive").Value & "")
				maryProduct(enProduct_CountryTaxIsActive) = Trim(.Fields("prodCountryTaxIsActive").Value & "")
				
				maryProduct(enProduct_prodMinQty) = CorrectEmptyValue(Trim(.Fields("prodMinQty").Value), 1)
				maryProduct(enProduct_prodIncrement) = CorrectEmptyValue(Trim(.Fields("prodIncrement").Value), 1)
				If maryProduct(enProduct_prodMinQty) < 1 Then maryProduct(enProduct_prodMinQty) = 1
				If maryProduct(enProduct_prodIncrement) < 1 Then maryProduct(enProduct_prodIncrement) = 1

				maryProduct(enProduct_sfProductID) = .Fields("sfProductID").Value
				maryProduct(enProduct_EnableReviews) = CorrectEmptyValue(Trim(.Fields("prodEnableReviews").Value), 0)
				maryProduct(enProduct_EnableAlsoBought) = CorrectEmptyValue(Trim(.Fields("prodEnableAlsoBought").Value), 0)
				maryProduct(enProduct_HandlingFee) = CorrectEmptyValue(Trim(.Fields("prodHandlingFee").Value), 0)
				maryProduct(enProduct_SetupFee) = CorrectEmptyValue(Trim(.Fields("prodSetupFee").Value), 0)
				maryProduct(enProduct_SetupFeeOneTime) = CorrectEmptyValue(Trim(.Fields("prodSetupFeeOneTime").Value), 0)

				maryProduct(enProduct_pageName) = Trim(.Fields("pageName").Value & "")
				maryProduct(enProduct_metaTitle) = CorrectEmptyValue(Trim(.Fields("metaTitle").Value), stripHTML(maryProduct(enProduct_Name)))
				maryProduct(enProduct_metaDescription) = CorrectEmptyValue(Trim(.Fields("metaDescription").Value), stripHTML(maryProduct(enProduct_ShortDescription)))
				maryProduct(enProduct_metaKeywords) = CorrectEmptyValue(Trim(.Fields("metaKeywords").Value), stripHTML(maryProduct(enProduct_Name)))

				maryProduct(enProduct_Height) = CorrectEmptyValue(Trim(.Fields("prodHeight").Value), 1)
				maryProduct(enProduct_Length) = CorrectEmptyValue(Trim(.Fields("prodLength").Value), 1)
				maryProduct(enProduct_Width) = CorrectEmptyValue(Trim(.Fields("prodWidth").Value), 1)
				maryProduct(enProduct_DisplayAdditionalImagesInWindow) = CorrectEmptyValue(Trim(.Fields("prodDisplayAdditionalImagesInWindow").Value), 0)
				
				'Adjust for Pricing Levels
				maryProduct(enProduct_Price) = GetPricingLevelPrice(maryProduct(enProduct_Price), maryProduct(enProduct_PLPrice))
				maryProduct(enProduct_SalePrice) = GetPricingLevelPrice(maryProduct(enProduct_SalePrice), maryProduct(enProduct_PLSalePrice))

				If maryProduct(enProduct_SaleIsActive) Then
					maryProduct(enProduct_SellPrice) = maryProduct(enProduct_SalePrice)
				Else
					maryProduct(enProduct_SellPrice) = maryProduct(enProduct_Price)
				End If

				If cblnSF5AE Then
					'Set the gift wrap
					maryProduct(enProduct_gwActivate) = CorrectEmptyValue(Trim(.Fields("gwActivate").Value), 0)
					If maryProduct(enProduct_gwActivate) = 1 Then
						maryProduct(enProduct_gwPrice) = CorrectEmptyValue(Trim(.Fields("gwPrice").Value), 0)
					Else
						maryProduct(enProduct_gwPrice) = "X"
					End If 

					'Set the inventory info
					maryProduct(enProduct_invenbTracked) = CorrectEmptyValue(Trim(.Fields("invenbTracked").Value), 0)
					If maryProduct(enProduct_invenbTracked) <> 1 Then
						maryProduct(enProduct_invenbStatus) = 0
						maryProduct(enProduct_invenbBackOrder) = 0
						maryProduct(enProduct_invenbNotify) = 0
					Else
						maryProduct(enProduct_invenbStatus) = CorrectEmptyValue(Trim(.Fields("invenbStatus").Value), 0)
						maryProduct(enProduct_invenbBackOrder) = CorrectEmptyValue(Trim(.Fields("invenbBackOrder").Value), 0)
						maryProduct(enProduct_invenbNotify) = CorrectEmptyValue(Trim(.Fields("invenbNotify").Value), 0)
					End If
					maryProduct(enProduct_invenInStock) = 0

					maryProduct(enProduct_LimitQtyToMTP) = CorrectEmptyValue(Trim(.Fields("prodLimitQtyToMTP").Value), 0)
				Else
					maryProduct(enProduct_gwPrice) = "X"
					maryProduct(enProduct_invenbStatus) = 0
					maryProduct(enProduct_invenbBackOrder) = 0
					maryProduct(enProduct_invenbNotify) = 0
					maryProduct(enProduct_invenInStock) = 0
					maryProduct(enProduct_LimitQtyToMTP) = 0
				End If

				'NOTE: split for SQL Server due to record size limitation (ie. truncates data)
				'Solution is to separate prodMessage, prodDescription, prodAdditionalImages, relatedProducts into another query
				If cblnSQLDatabase Then
					pstrSQL = "SELECT prodMessage, prodDescription, prodAdditionalImages, relatedProducts" _
							& " FROM sfProducts" _
							& " Where prodID=?"
					closeObj(pobjRS)
					pobjCmd.Commandtext = pstrSQL
					Set pobjRS = pobjCmd.Execute
				End If
				
				maryProduct(enProduct_Description) = Trim(pobjRS.Fields("prodDescription").Value & "")
				maryProduct(enProduct_Message) = Trim(pobjRS.Fields("prodMessage").Value & "")
				maryProduct(enProduct_AdditionalImages) = Trim(pobjRS.Fields("prodAdditionalImages").Value & "")
				maryProduct(enProduct_RelatedProducts) = Trim(pobjRS.Fields("relatedProducts").Value & "")

				maryProduct(enProduct_CategoryName) = getCategoryItem(maryProduct(enProduct_CategoryID), "Name")
				maryProduct(enProduct_MfgName) = getMfgVendItem(maryProduct(enProduct_MfgID), "Name", True)
				maryProduct(enProduct_VendorName) = getMfgVendItem(maryProduct(enProduct_VendorID), "Name", False)
	
				If ssDebug_ProductDisplay Then Call writeProductInfo
			Else
				'Reset array
				'For i = 0 To UBound(maryProduct)
				'	maryProduct(i) = ""
				'Next 'i
				'maryProduct(enProduct_Exists) = False
				
				maryProduct(enProduct_ID) = strProdID
			
			End If	'maryProduct(enProduct_Exists)
		End With	'pobjRS
		closeObj(pobjRS)
		Set pobjCmd = Nothing
		Call DebugRecordTime("Product array loaded")
		
	Else
		'Do Nothing
	End If	'mstrCurrentProductID <> strProdID
	
	Select Case bytCase
		Case -1:
		Case -2:
				Set rs = CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM sfAttributes WHERE attrProdId = '" & makeInputSafe(sProdID) & "'"
				rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
				
				sDetails = ""
				If Not (rs.EOF And rs.BOF) Then
				    iCounter = 1
				    While Not rs.EOF
				        sDetails = sDetails & "<br />" & rs("attrName") & "<br />" & getAttributeDetails(rs("attrID"), iCounter, sCase)
				        rs.MoveNext
				        iCounter = iCounter + 1
				    Wend
				End If
				getProductInfo = sDetails
		Case enProduct_MTP:
			If isArray(maryProduct(enProduct_MTP)) Then
				getProductInfo = maryProduct(enProduct_MTP)
			Else
				maryProduct(enProduct_MTP) = LoadMTPrices(strProdID)
				getProductInfo = maryProduct(enProduct_MTP)
			End If
		Case enProduct_attributes:
			If maryProduct(enProduct_AttrNum) > 0 Then
				If isArray(maryProduct(enProduct_attributes)) Then
					getProductInfo = maryProduct(enProduct_attributes)
				Else
					maryProduct(enProduct_attributes) = LoadAttributes(strProdID)
					getProductInfo = maryProduct(enProduct_attributes)
				End If
			End If
		Case Else
			getProductInfo = maryProduct(bytCase)
	End Select
	
End Function	'getProductInfo

'**********************************************************

Sub writeProductInfo()

	Response.Write "<fieldset><legend>Product information for " & maryProduct(enProduct_ID) & "</legend>"
	Response.Write "enProduct_IsActive: " & maryProduct(enProduct_IsActive) & "<br />"
	Response.Write "IsActive: " & CBool(maryProduct(enProduct_IsActive)) & "<br />"
	Response.Write "enProduct_Name: " & maryProduct(enProduct_Name) & "<br />"
	Response.Write "enProduct_NamePlural: " & maryProduct(enProduct_NamePlural) & "<br />"
	Response.Write "enProduct_ShortDescription: " & maryProduct(enProduct_ShortDescription) & "<br />"
	Response.Write "enProduct_Description: " & maryProduct(enProduct_Description) & "<br />"
	Response.Write "enProduct_Message: " & maryProduct(enProduct_Message) & "<br />"
	Response.Write "enProduct_ImageSmallPath: " & maryProduct(enProduct_ImageSmallPath) & "<br />"
	Response.Write "enProduct_ImageLargePath: " & maryProduct(enProduct_ImageLargePath) & "<br />"
	Response.Write "enProduct_Link: " & maryProduct(enProduct_Link) & "<br />"
	Response.Write "enProduct_Price: " & maryProduct(enProduct_Price) & "<br />"
	Response.Write "enProduct_SalePrice: " & maryProduct(enProduct_SalePrice) & "<br />"
	Response.Write "enProduct_CategoryName: " & maryProduct(enProduct_CategoryName) & "<br />"
	Response.Write "enProduct_MfgName: " & maryProduct(enProduct_MfgName) & "<br />"
	Response.Write "enProduct_VendorName: " & maryProduct(enProduct_VendorName) & "<br />"
	Response.Write "enProduct_SaleIsActive: " & maryProduct(enProduct_SaleIsActive) & "<br />"
	Response.Write "enProduct_AttrNum: " & maryProduct(enProduct_AttrNum) & "<br />"
	Response.Write "enProduct_PLPrice: " & maryProduct(enProduct_PLPrice) & "<br />"
	Response.Write "enProduct_PLSalePrice: " & maryProduct(enProduct_PLSalePrice) & "<br />"
	Response.Write "enProduct_ShipIsActive: " & maryProduct(enProduct_ShipIsActive) & "<br />"
	Response.Write "enProduct_Weight: " & maryProduct(enProduct_Weight) & "<br />"
	Response.Write "enProduct_Ship: " & maryProduct(enProduct_Ship) & "<br />"
	Response.Write "enProduct_StateTaxIsActive: " & maryProduct(enProduct_StateTaxIsActive) & "<br />"
	Response.Write "enProduct_CountryTaxIsActive: " & maryProduct(enProduct_CountryTaxIsActive) & "<br />"
	Response.Write "enProduct_gwActivate: " & maryProduct(enProduct_gwActivate) & "<br />"
	Response.Write "enProduct_gwPrice: " & maryProduct(enProduct_gwPrice) & "<br />"
	Response.Write "enProduct_invenbTracked: " & maryProduct(enProduct_invenbTracked) & "<br />"
	Response.Write "enProduct_invenbStatus: " & maryProduct(enProduct_invenbStatus) & "<br />"
	Response.Write "enProduct_invenbBackOrder: " & maryProduct(enProduct_invenbBackOrder) & "<br />"
	Response.Write "enProduct_invenbNotify: " & maryProduct(enProduct_invenbNotify) & "<br />"
	Response.Write "enProduct_invenInStock: " & maryProduct(enProduct_invenInStock) & "<br />"
	Response.Write "enProduct_prodMinQty: " & maryProduct(enProduct_prodMinQty) & "<br />"
	Response.Write "enProduct_prodIncrement: " & maryProduct(enProduct_prodIncrement) & "<br />"

	Response.Write "enProduct_pageName: " & maryProduct(enProduct_pageName) & "<br />"
	Response.Write "enProduct_metaTitle: " & maryProduct(enProduct_metaTitle) & "<br />"
	Response.Write "enProduct_metaDescription: " & maryProduct(enProduct_metaDescription) & "<br />"
	Response.Write "enProduct_metaKeywords: " & maryProduct(enProduct_metaKeywords) & "<br />"

	Response.Write "</fieldset>"
	
End Sub	'writeProductInfo

'***********************************************************************************************

Function GetProductName(byVal strProdID)

Dim pstrSQL
Dim pobjCmd
Dim pobjRS
	
	pstrSQL = "Select ProdName FROM sfProducts WHERE ProdID=?"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtext = pstrSQL
		.Commandtype = adCmdText
		'.Commandtype = adCmdStoredProc
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("ProdID", adWChar, adParamInput, 50, checkFieldLength(strProdID, 50, 0))
		Set pobjRS = .Execute
		
		If Not pobjRS.EOF Then GetProductName = pobjRS.Fields("ProdName").Value
		Call closeObj(pobjRS)

	End With	'pobjCmd
	Call closeObj(pobjCmd)
	
End function	'GetProductName

'***********************************************************************************************

Function getProductInventoryLevel(byRef aryAttributes)

Dim i
Dim paryAttribute
Dim paryAttributes
Dim plngCounter_Attributes
Dim pstrAttributeIDs

	If maryProduct(enProduct_invenbTracked) = 1 Then
		maryProduct(enProduct_AttributeArray) = aryAttributes
		paryAttributes = maryProduct(enProduct_AttributeArray)
		If isArray(paryAttributes) Then
			pstrAttributeIDs = ""	'initialize the attributes
			For plngCounter_Attributes = 0 To UBound(paryAttributes)
				paryAttribute = paryAttributes(plngCounter_Attributes)
				If Len(pstrAttributeIDs) = 0 Then
					If isArray(paryAttribute) Then
						pstrAttributeIDs = paryAttribute(enAttributeItem_attrdtID)
					Else
						pstrAttributeIDs = paryAttribute
					End If
				Else
					If isArray(paryAttribute) Then
						pstrAttributeIDs = pstrAttributeIDs & "," & paryAttribute(enAttributeItem_attrdtID)
					Else
						If plngCounter_Attributes = UBound(paryAttributes) And Len(paryAttribute) = 0 Then
						Else
							pstrAttributeIDs = pstrAttributeIDs & "," & paryAttribute
						End If
					End If
				End If
			Next 'plngCounter_Attributes
		End If
		
		If Len(pstrAttributeIDs) = 0 Then
			pstrAttributeIDs = "0"
		Else
			pstrAttributeIDs = bubbleSortList(pstrAttributeIDs, ",")
		End If
		maryProduct(enProduct_attributeIDs) = pstrAttributeIDs
		maryProduct(enProduct_invenInStock) = GetAvailableQty(maryProduct(enProduct_ID), pstrAttributeIDs)
		
	Else
		maryProduct(enProduct_invenInStock) = "X"
	End If	'maryProduct(enProduct_invenbTracked) = 1
	
	getProductInventoryLevel = maryProduct(enProduct_invenInStock)
	
	If True Then
		Response.Write "<fieldset><legend>getProductInventoryLevel for " & maryProduct(enProduct_ID) & "</legend>" _
					 & "<b>Attribute IDs</b>: " & maryProduct(enProduct_attributeIDs) & "<br />" _
					 & "<b>Track</b>: " & maryProduct(enProduct_invenbTracked) & "<br />" _
					 & "<b>Qty In Stock</b>: " & maryProduct(enProduct_invenInStock) & "<br />" _
					 & "<b>Allow Backorder</b>: " & maryProduct(enProduct_invenbBackOrder) & "<br />" _
					 & "</fieldset>"
	End If
	
End Function	'getProductInventoryLevel

'***********************************************************************************************

Function getAttributeDetails(attrID, iCounter, sCase)
	Set rs = CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM sfAttributeDetail WHERE attrdtAttributeId = " & attrID & " ORDER BY attrdtOrder"
'	rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	rs.Open SQL, cnn, adOpenStatic, adLockBatchOptimistic, adCmdText 
	Call AdjustRecordPricingLevel(rs, "DetailAttr")
	
    If sCase = 14 Then
        sTemp = "<select name=""attr" & iCounter & """>"
    Else
        sTemp = ""
    End If
    
    sChecked = "CHECKED"
    
    While Not rs.EOF
        sAmount = ""
        Select Case rs("attrdtType")
            Case 1
                sAmount = " (add " & FormatCurrency(rs("attrdtPrice")) & ")"
            Case 2
                sAmount = " (subtract " & FormatCurrency(rs("attrdtPrice")) & ")"
        End Select
        
        If sCase = 14 Then
            sTemp = sTemp & "<option value=" & rs("attrdtID") & ">" & rs("attrdtName") & sAmount & "</option>"
        Else
            sTemp = sTemp & "<input type=""radio"" " & sChecked & " name=""attr" & iCounter & """ value=""" & rs("attrdtID") & """>" & rs("attrdtName") & sAmount & "<br />"
        End If
        rs.MoveNext
        sChecked = ""
    Wend
    If sCase = 14 Then
        sTemp = sTemp & "</select>"
    End If
    closeObj(rs)
    getAttributeDetails = sTemp
End Function



'***********************************************************************************************
'	GIFT WRAP
'***********************************************************************************************

	Sub ShowGiftWrap(byVal strProdID)

	Dim gwprice

		gwprice = getProductInfo(strProdID, enProduct_gwPrice)
		If gwprice <> "X" then	
			Response.Write "<br /><INPUT name=chkGiftWrap type=checkbox value = 1 >"
			Response.Write "Gift wrap (add " & formatcurrency(gwprice) & " per item)"	
		End if

	End Sub	'ShowGiftWrap

	'***********************************************************************************************

	Function GetGiftWrapPrice(byVal strProdID)
		GetGiftWrapPrice =  getProductInfo(strProdID, enProduct_gwPrice)
	End Function

'***********************************************************************************************
'	MTP
'***********************************************************************************************

Function calculateMTPrice(byVal lngIndex)

Dim paryMTP
Dim pdblPriceOut

	pdblPriceOut = CDbl(maryProduct(enProduct_Price))
	'paryTemp(i) = Array(.Fields("mtQUANTITY").Value, .Fields("mtvalue").Value, CBool(.Fields("mtType").Value = "Amount"), .Fields("mtPLValue").Value)
	paryMTP = maryProduct(enProduct_MTP)
	
	If lngIndex <= UBound(paryMTP) Then
		If paryMTP(lngIndex)(2) then
			pdblPriceOut = pdblPriceOut - CDbl(GetPricingLevelPrice(paryMTP(lngIndex)(1), paryMTP(lngIndex)(3)))
		Else
			pdblPriceOut = pdblPriceOut * (1 - CDbl(GetPricingLevelPrice(paryMTP(lngIndex)(1), paryMTP(lngIndex)(3)))/100)
			pdblPriceOut = Round(pdblPriceOut, 2)
		End If
		If pdblPriceOut < 0 Then pdblPriceOut = 0
	End If

	calculateMTPrice = pdblPriceOut

End Function	'calculateMTPrice

'***************************************************************************************************************************************************************************
Function GetMTPrice3(byVal strProductID, byVal lngQty)

Dim i
Dim paryMTP
Dim pdblPriceOut

	pdblPriceOut = CDbl(getProductInfo(strProductID, enProduct_Price))
	
	'paryTemp(i) = Array(.Fields("mtQUANTITY").Value, .Fields("mtvalue").Value, CBool(.Fields("mtType").Value = "Amount"), .Fields("mtPLValue").Value)
	paryMTP = getProductInfo(strProductID, enProduct_MTP)
	
	'Step to location where MTP > current qty
	For i = 0 To UBound(paryMTP)
		If paryMTP(i)(0) > lngQty Then Exit For
	Next 'i

	'i should be one too many
	i = i - 1
	
	If i <= UBound(paryMTP) Then
		If paryMTP(i)(2) then
			pdblPriceOut = pdblPriceOut - CDbl(GetPricingLevelPrice(paryMTP(i)(1), paryMTP(i)(3)))
		Else
			pdblPriceOut = pdblPriceOut * (1 - CDbl(GetPricingLevelPrice(paryMTP(i)(1), paryMTP(i)(3)))/100)
			pdblPriceOut = Round(pdblPriceOut, 2)
		End If
		If pdblPriceOut < 0 Then pdblPriceOut = 0
	End If

	GetMTPrice3 = pdblPriceOut

End Function	'GetMTPrice3

'***************************************************************************************************************************************************************************

Function hasMTP(byVal strProdID) 

Dim pobjCmd
Dim pobjRS
	
	Set pobjCmd = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select mtQuantity FROM sfMTPrices WHERE mtprodid=?"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProdID)
		Set pobjRS = .Execute
	End With	'pobjCmd
	Set pobjCmd = Nothing

	hasMTP = Not pobjRS.EOF

	Call CloseObj(pobjRS)

End Function	'hasMTP

'***************************************************************************************************************************************************************************

Function LoadAttributes(byVal strProductID)

Dim i
Dim prevAttrID
Dim paryAttributes
Dim paryAttributeDetails
Dim plngAttrPosition
Dim plngAttrDetailPosition
Dim plngDefaultNumAttributeDetails
Dim pobjCmd
Dim pobjRS

	plngAttrPosition = -1
	plngDefaultNumAttributeDetails = 8
	
	Set pobjCmd = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "SELECT sfAttributes.*, sfAttributeDetail.*" _
					 & " FROM sfAttributes INNER JOIN sfAttributeDetail ON sfAttributes.attrId = sfAttributeDetail.attrdtAttributeId" _
					 & " WHERE sfAttributes.attrProdId=?" _
					 & " ORDER BY sfAttributes.attrDisplayOrder, sfAttributes.attrName, sfAttributeDetail.attrdtOrder"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("attrProdId", adVarChar, adParamInput, Len(strProductID), strProductID)
		Set pobjRS = .Execute
	End With	'pobjCmd
	Set pobjCmd = Nothing

	With pobjRS
		If Not .EOF then
			ReDim paryAttributes(maryProduct(enProduct_AttrNum) - 1)
			Do While Not .EOF
				If prevAttrID <> .Fields("attrID").Value Then
					prevAttrID = .Fields("attrID").Value
					If plngAttrPosition > -1 Then
						If plngAttrDetailPosition < UBound(paryAttributeDetails) Then ReDim Preserve paryAttributeDetails(plngAttrDetailPosition)
						paryAttributes(plngAttrPosition)(enAttribute_DetailArray) = paryAttributeDetails
					End If
					
					'Move to next attribute
					plngAttrPosition = plngAttrPosition + 1
					plngAttrDetailPosition = -1
					ReDim paryAttributeDetails(plngDefaultNumAttributeDetails)
					If UBound(paryAttributes) < plngAttrPosition Then ReDim Preserve paryAttributes(plngAttrPosition)	'this line added just in case the number of attributes are more than saved in the detail record
					paryAttributes(plngAttrPosition) = Array(prevAttrID, Trim(.Fields("attrName").Value & ""), Trim(.Fields("attrDisplayStyle").Value & ""), Trim(.Fields("attrDisplay").Value & ""), Trim(.Fields("attrImage").Value & ""), Trim(.Fields("attrSKU").Value & ""), Trim(.Fields("attrURL").Value & ""), Trim(.Fields("attrExtra").Value & ""), "", "")

					'0-prevAttrID
					'1-Trim(.Fields("attrName").Value & "")
					'2-Trim(.Fields("attrDisplayStyle").Value & "")
					'3-Trim(.Fields("attrDisplay").Value & "")
					'4-Trim(.Fields("attrImage").Value & "")
					'5-'2-Trim(.Fields("attrSKU").Value & "")
					'6-Trim(.Fields("attrURL").Value & "")
					'7-Trim(.Fields("attrExtra").Value & "")
					'8-"" - placeholder for attribute details array
					'9-"" - placeholder for attribute template

				End If	'prevAttrID <> .Fields("attrID").Value

				plngAttrDetailPosition = plngAttrDetailPosition + 1
				If plngAttrDetailPosition > UBound(paryAttributeDetails) Then ReDim Preserve paryAttributeDetails(plngAttrDetailPosition)
				paryAttributeDetails(plngAttrDetailPosition) = Array(.Fields("attrdtID").Value, Trim(.Fields("attrdtName").Value & ""), Trim(.Fields("attrdtPrice").Value & ""), Trim(.Fields("attrdtType").Value & ""), Trim(.Fields("attrdtImage").Value & ""), Trim(.Fields("attrdtPLPrice").Value & ""), Trim(.Fields("attrdtFileName").Value & ""), Trim(.Fields("attrdtDefault").Value & ""), Trim(.Fields("attrdtDisplay").Value & ""), Trim(.Fields("attrdtSKU").Value & ""), Trim(.Fields("attrdtURL").Value & ""), Trim(.Fields("attrdtWeight").Value & ""), Trim(.Fields("attrdtExtra").Value & ""), Trim(.Fields("attrdtExtra1").Value & ""), GetPricingLevelPrice(.Fields("attrdtPrice").Value, .Fields("attrdtPLPrice").Value))
				'0-.Fields("attrdtID").Value
				'1-Trim(.Fields("attrdtName").Value & "")
				'2-Trim(.Fields("attrdtPrice").Value & "")
				'3-Trim(.Fields("attrdtType").Value & "")
				'4-Trim(.Fields("attrdtImage").Value & "")
				'5-Trim(.Fields("attrdtPLPrice").Value & "")
				'6-Trim(.Fields("attrdtFileName").Value & "")
				'7-Trim(.Fields("attrdtDefault").Value & "")
				'8-Trim(.Fields("attrdtDisplay").Value & "")
				'9-Trim(.Fields("attrdtSKU").Value & "")
				'10-Trim(.Fields("attrdtURL").Value & "")
				'11-Trim(.Fields("attrdtWeight").Value & "")
				'12-Trim(.Fields("attrdtExtra").Value & "")
				'13-Trim(.Fields("attrdtExtra1").Value & "")
				'14-GetPricingLevelPrice(.Fields("attrdtPrice").Value
				'15-.Fields("attrdtPLPrice").Value)
				.MoveNext
			Loop
			If plngAttrDetailPosition < UBound(paryAttributeDetails) Then ReDim Preserve paryAttributeDetails(plngAttrDetailPosition)
			paryAttributes(plngAttrPosition)(enAttribute_DetailArray) = paryAttributeDetails
			
		End If	'Not .EOF
	End With	'pobjRS
	Call CloseObj(pobjRS)
	If UBound(paryAttributes) > plngAttrPosition Then ReDim Preserve paryAttributes(plngAttrPosition)	'this line added just in case the number of attributes are less than saved in the detail record
	
	If False Then	'True	False
		Dim j
		Response.Write "<fieldset><legend>Attributes</legend>"
		Response.Write "<ul>"			
		For i = 0 To plngAttrPosition
			Response.Write "<li><strong>" & paryAttributes(i)(enAttribute_Name) & " (" & paryAttributes(i)(enAttribute_ID) & ")</strong>"
			Response.Write "<br />DisplayStyle: " & paryAttributes(i)(enAttribute_DisplayStyle) & "</strong>"
			Response.Write "</li>"
			Response.Flush
			Response.Write "<ul>"
			For j = 0 To UBound(paryAttributes(i)(enAttribute_DetailArray))
				Response.Write "<li><strong>" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Name) & " (" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & ")</strong>"
				Response.Write "<br />Price: " & paryAttributes(i)(enAttributeDetail_Price) & "</strong>"
				Response.Write "</li>"
			Next 'j
			Response.Write "</ul>"
		Next 'i
		Response.Write "</ul>"			
		Response.Write "</fieldset>"			
	End If

	LoadAttributes = paryAttributes

End Function	'LoadAttributes

'***************************************************************************************************************************************************************************

Function LoadMTPrices(byVal strProductID)

Dim i
Dim paryTemp
Dim pobjCmd
Dim pobjRS

	Set pobjCmd = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select mtQUANTITY, mtvalue, mtType, mtPLValue FROM sfMTPrices WHERE mtprodid=? ORDER By mtIndex ASC"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("mtprodid", adVarChar, adParamInput, 50, strProductID)
		Set pobjRS = .Execute
	End With	'pobjCmd
	Set pobjCmd = Nothing

	With pobjRS
		If .EOF then
			ReDim paryTemp(0)
			paryTemp(0) = Array(1, 0, True, 0)
		Else
			ReDim paryTemp(5)

			i = -1
			'Response.Write "<ul>"
			Do While Not .EOF
				'Response.Write "<li>" & pobjRS.Fields("mtQUANTITY").Value & " - " & pobjRS.Fields("mtvalue").Value & "</li>"
				i = i + 1
				If UBound(paryTemp) < i Then ReDim Preserve paryTemp(i)
				paryTemp(i) = Array(.Fields("mtQUANTITY").Value, .Fields("mtvalue").Value, CBool(.Fields("mtType").Value = "Amount"), .Fields("mtPLValue").Value)
				.MoveNext 
			Loop
			If UBound(paryTemp) > i Then ReDim Preserve paryTemp(i)
			'Response.Write "</ul>"
		End if
	End With
	
	Call CloseObj(pobjRS)

	LoadMTPrices = paryTemp

End Function	'LoadMTPrices

'***************************************************************************************************************************************************************************

Sub ShowMTPricesLink(sProdId) 
	If hasMTP(strProdID) Then Response.Write "<br /> <a href='javascript:show_page(" & chr(34) & "MTPrices.asp?sProdId=" & strProdID & chr(34) & ")'>Check Volume Discounts</a>" & vbcrlf	
End Sub

'***************************************************************************************************************************************************************************

Function WriteMTPrices(aryPrices, dblPrice)

Dim i
Dim pstrLine
Dim pstrOutput
Dim pdblPrice

	If isArray(aryPrices) Then
		For i = 0 To UBound(aryPrices)
			If isArray(aryPrices(i)) Then
				If aryPrices(i)(2) Then
					pdblPrice = dblPrice - aryPrices(i)(1)
				Else
					pdblPrice = dblPrice * (1 - aryPrices(i)(1))
				End If
				
				If UBound(aryPrices) = 0 Then
					pstrLine = FormatCurrency(pdblPrice,2) + "<br />"
				Else
					pstrLine = aryPrices(i)(0) & " + " & FormatCurrency(pdblPrice,2) + "<br />"
				End If
			Else
				If UBound(aryPrices) = 0 Then
					pstrLine = FormatCurrency(pdblPrice,2) + "<br />"
				Else
					pstrLine = "1 + " & FormatCurrency(dblPrice,2) + "<br />"
				End If
			End If
			pstrOutput = pstrOutput & pstrLine
		Next 'i
	Else
		pstrOutput = "1 + " & FormatCurrency(dblPrice,2) + "<br />"
	End If

	WriteMTPrices = pstrOutput

End Function	'WriteMTPrices

'***************************************************************************************************************************************************************************

Function WriteMTPricingTable(byVal strProductID, byVal strTableTitle)

Dim i
Dim pstrLine
Dim pstrOutput
Dim pdblPrice

	Dim aryPrices, dblPrice
	aryPrices = getProductInfo(strProductID, enProduct_MTP)
	dblPrice = getProductInfo(strProductID, enProduct_Price)

	If isArray(aryPrices) Then
		pstrOutput = "<table border=""0"" cellpadding=""2"" cellspacing=""0"" class=""MTP"">"
		pstrOutput = pstrOutput & "<tr><th colspan=""" & UBound(aryPrices) + 2 & """ align=""center"" class=""MTP"">" & strTableTitle & "</th></tr>"

		pstrOutput = pstrOutput & "<tr>"
		pstrOutput = pstrOutput & "<th align=""center"" class=""MTP"">1</th>"
		For i = 0 To UBound(aryPrices)
			If isArray(aryPrices(i)) Then
				If UBound(aryPrices) = 0 Then
					pstrLine = FormatCurrency(pdblPrice,2)
				Else
					pstrLine = aryPrices(i)(0) & "+"
				End If
			Else
				If UBound(aryPrices) = 0 Then
					pstrLine = FormatCurrency(pdblPrice,2)
				Else
					pstrLine = "1 + "
				End If
			End If
			
			pstrOutput = pstrOutput & "<th align=""center"" class=""MTP"">" & pstrLine & "</th>"
		Next 'i
		pstrOutput = pstrOutput & "</tr>"


		pstrOutput = pstrOutput & "<tr>"
		pstrOutput = pstrOutput & "<td class=""MTP"">" & FormatCurrency(dblPrice,2) & "</td>"
		For i = 0 To UBound(aryPrices)
			If isArray(aryPrices(i)) Then
				If aryPrices(i)(2) Then
					pdblPrice = dblPrice - aryPrices(i)(1)
				Else
					pdblPrice = dblPrice * (1 - aryPrices(i)(1)/100)
				End If
				
				If UBound(aryPrices) = 0 Then
					pstrLine = FormatCurrency(pdblPrice,2)
				Else
					pstrLine = FormatCurrency(pdblPrice,2)
				End If
			Else
				If UBound(aryPrices) = 0 Then
					pstrLine = FormatCurrency(pdblPrice,2)
				Else
					pstrLine = FormatCurrency(dblPrice,2)
				End If
			End If
			
			pstrOutput = pstrOutput & "<td class=""MTP"">" & pstrLine & "</td>"
		Next 'i
		pstrOutput = pstrOutput & "</tr>"
		pstrOutput = pstrOutput & "</table>"
	Else
		pstrOutput = "1 + " & FormatCurrency(dblPrice,2) + "<br />"
	End If

	WriteMTPricingTable = pstrOutput

End Function	'WriteMTPricingTable

'***********************************************************************************************
'	INVENTORY
'***********************************************************************************************

	Function CheckBackOrder(byVal strProdID)
		CheckBackOrder =  getProductInfo(strProdID, enProduct_invenbBackOrder)
	End Function

	Function CheckShowStatus(byVal strProdID)
		CheckShowStatus = getProductInfo(strProdID, enProduct_invenbStatus)
	End Function

	Function CheckInventoryTracked(byVal strProdID)
		CheckInventoryTracked =  getProductInfo(strProdID, enProduct_invenbTracked)
	End Function

	Function CheckNotification(byVal strProdID)
		CheckNotification =  getProductInfo(strProdID, enProduct_invenbNotify)
	End Function

	'***********************************************************************************************

	Function GetAvailableQty(byVal strProductID, byVal AttIDs)

	Dim pstrSQL
	Dim pobjCmd
	Dim pobjRS
	Dim pstrResult
		
		If CheckInventoryTracked(strProductID) = 0 then 
			pstrResult = "X" 'No inventoryinfo record
		Else
			Set pobjCmd = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = "Select invenInstock FROM sfInventory WHERE invenProdID=? AND  invenAttDetailID=?"
				Set .ActiveConnection = cnn
				
				.Parameters.Append .CreateParameter("invenProdID", adVarChar, adParamInput, 50, strProductID)
				.Parameters.Append .CreateParameter("invenAttDetailID", adVarChar, adParamInput, 255, AttIDs)

				Set pobjRS = .Execute
				If pobjRS.EOF then 
					pstrResult = "X" 'Inventory record missing
				Else
					pstrResult = pobjRS.Fields("invenInstock").Value
				end if 
				CloseObj (pobjRS)

			End With	'pobjCmd
			Set pobjCmd = Nothing
		End If
		
		GetAvailableQty = pstrResult
		
	End Function	'getavailableqty

	'***********************************************************************************************

	Function getInventoryRecordID(byVal strProductID, byVal AttIDs)

	Dim pstrSQL
	Dim pobjCmd
	Dim pobjRS
		
		If CheckInventoryTracked(strProductID) = 0 then 
			getInventoryRecordID = -1 'No inventoryinfo record
		Else
			Set pobjCmd = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = "Select invenId FROM sfInventory WHERE invenProdID=? AND  invenAttDetailID=?"
				Set .ActiveConnection = cnn
				
				.Parameters.Append .CreateParameter("invenProdID", adVarChar, adParamInput, 50, strProductID)
				.Parameters.Append .CreateParameter("invenAttDetailID", adVarChar, adParamInput, 255, AttIDs)

				Set pobjRS = .Execute
				If pobjRS.EOF then 
					getInventoryRecordID = -1 'Inventory record missing
				Else
					getInventoryRecordID = pobjRS.Fields("invenId").Value
				end if 
				CloseObj (pobjRS)

			End With	'pobjCmd
			Set pobjCmd = Nothing
		End If
		
	End Function	'getInventoryRecordID

	'***********************************************************************************************

	Function CheckInStock(byVal strProdID)

	Dim pobjRS
	Dim pobjCmd

		Set pobjCmd = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "Select Sum(invenInstock) as instock FROM sfInventory WHERE invenProdID=?"
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("invenProdID", adVarChar, adParamInput, 50, strProdID)
			Set pobjRS = .Execute
		End With	'pobjCmd
		Set pobjCmd = Nothing

		If pobjRS.EOF Then
			CheckInStock = "X"
		Else
			CheckInStock= pobjRS.Fields("Instock").Value
		End if
		closeObj(pobjRS)
		
	End Function	'CheckInStock

	'***********************************************************************************************

Sub ShowProductInventory (strProduct,sType)

Dim sPath,jsvar

	If CheckInventoryTracked(strProduct) = 1 Then 
		If CheckShowStatus(strProduct) = 1 Then
			Select case CheckInStock(strProduct)
				Case "X" 'inventory not tracked for this product
					
					
				Case 0 'inventory tracked
					If sType = "Dynamic" Then
						Response.Write "<br /> Out of Stock!"
						If checkbackorder(strProduct) = 1 then	
							Response.Write "<br /> Click " & chr(34) & "Add to Cart" & chr(34) & " to BackOrder!"
						End If
					End If
						
				Case else
					sPath = "StockInfo.asp?sProdId=" & strProduct
					jsvar = "javascript:show_page('" & sPath & "')"
						
					If sType ="Dynamic" then
						Response.Write "<br /><a href=" & chr(34) & jsvar & chr(34) & ">Check Stock!</a>"	
					eLse
						Response.Write  "<br /><a href=" & chr(34) & jsvar & chr(34) & ">Stock Information</a>"	
					end If	
									
			End  Select
		End If
	End If
	
End Sub

Private Function GetFullPath(Vdata,justMain) 
Dim sSql ,X
Dim iCatId
Dim sFirst
Dim rsCat,rsSubCat
Dim arrTemp ,bMain
bMain = false
 if left(vData,4)= "none" then
  bMain = True
  arrTemp = split(vdata,"-")
  vdata = arrtemp(1)
 elseif vData = "" then
   GetFullPath = "" 
  exit function
 elseif instr(Vdata,"-") = 0  then
    vData = vData 
 end if 
  arrTemp = split(vData,"-")
 Set rsCat = CreateObject("ADODB.RecordSet")
 Set rsSubCat = CreateObject("ADODB.RecordSet")
  rsSubCat.Open "sfSub_Categories",cnn,adOpenStatic,adLockReadOnly ,adcmdtable 
   For X = 0 To UBound(arrTemp)
     rsSubCat.Requery
     if arrTemp(X)<> "" then
      rsSubCat.Find "SubCatId = " & CInt(arrTemp(X))
      GetFullPath = GetFullPath & rsSubCat("SubCatName") & "-"
     end if
   Next
  sSql  = "Select catName From sfCategories Where catId =" & rsSubCat("subcatCategoryId")   
 rsCat.Open sSql,cnn,adOpenStatic,adLockReadOnly ,adcmdText
 if justmain = 1 then
    GetFullPath = rsCat("catName")
 else 
   if bMain = True Then
      GetFullPath = rsCat("catName")
   else
     GetFullPath = rsCat("catName") & "-" &  Left(GetFullPath, Len(GetFullPath) - 1)
   end if 
 end if
 Set rsCat = Nothing
 Set rsSubCat = Nothing
 Exit Function
End Function

%>


