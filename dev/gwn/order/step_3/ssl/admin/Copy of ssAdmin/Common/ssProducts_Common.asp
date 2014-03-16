<% 
'********************************************************************************
'*   Common Support File For StoreFront 5.0 add-ons                             *
'*   Release Version:	2.00.001                                                *
'*   Release Date:		August 21, 2003				                            *
'*   Revision Date:		November 9, 2004                                        *
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

'**********************************************************
'*	Global variables
'**********************************************************

Dim maryAttributeTypes
Dim maryImageFields
Dim maryPricingLevels
Dim mbytSummaryTableHeight
Dim clngDefaultMaxRecords
Dim cblnDeleteExistingProductAttributesOnCopy:	cblnDeleteExistingProductAttributesOnCopy = False

'**********************************************************
'*	Functions
'**********************************************************

'Function getProductUIDByCode(byVal strProductCode)
'Function getProductInfo(byRef lngProductUID, byRef strProductCode, byRef strProdName)
'Function setAttribute(byVal lngProductUID, byRef aryAttributeInfo)
'Function setAttributeDetail(byVal lngAttributeID, byRef aryAttributeInfo)
'Function setAttributeDetailSortOrder(byVal arySortOrder)
'Function setGiftWrap(byVal lngProductUID, byVal bytActivate, byVal strPrice)
'Function createProductCategoryAssignments(byVal strProductID, byVal aryCategoryUIDs)
'Function createProductCategoryAssignment(byVal strProductID, byVal lngCategoryID)
'Function deleteProductCategoryAssignments(byVal strProductID)
'Function DeleteAllProducts()
'Function DeleteProduct(byVal lngUID)
'Function DeleteAttributeCategory(byVal lngUID, byVal lngProductUID)
'Function DeleteAttributeDetail(byVal lngUID)
'Function DeleteInventoryByAttributeDetail(byVal lngUID)
'Function updateVolumePricing(byVal strProductID, byVal strBreakLevel, byVal strAmount, byVal strDollarOrPercent)
'Function SetImagePath(byVal strImage)

'Function inventoryLowQty(byVal strProductUID, byVal strAttributeIDs)
'Function inventoryQty(byVal strProductUID, byVal strAttributeIDs)
'Function isInventoryTracked(byVal strProductUID)
'Sub updateInventoryQtyDelta(byVal strProductUID, byVal strAttributeIDs, byVal qtyDelta)
'Function updateCategoryHasProductsStatus
'Function updateSubCategoryHasProductsStatus(byVal strSubCatID, byVal strHasProds)

'******************************************************************************************************************************************************************
'******************************************************************************************************************************************************************

Function getProductUIDByCode(byVal strProductCode)
'Input: ProductUID, array of ProductUIDs, Flag to delete existing or add to them
'Output: Returns uid for existing product; -1 for no match
'Action: either finds an attribute detail matching the attribute information or creates one

Dim pstrSQL
Dim pobjRS

	If Err.number <> 0 Then Err.Clear

	If Len(strProductCode) = 0 Then
		getProductUIDByCode = -1
	Else
		pstrSQL = "SELECT sfProductID FROM sfProducts WHERE prodID=" & wrapSQLValue(strProductCode, False, enDatatype_string)
		Set pobjRS = GetRS(pstrSQL)
		If pobjRS.EOF Then
			getProductUIDByCode = -1
		Else
			getProductUIDByCode = pobjRS.Fields("sfProductID").Value
		End If
		pobjRS.Close
		Set pobjRS = Nothing
	End If

End Function	'getProductUIDByCode

'****************************************************************************************************************************************************************

Function getProductInfo(byRef strProductCode, byRef strProdName)

Dim pstrSQL
Dim pobjRS
Dim pblnFound

	If Err.number <> 0 Then Err.Clear

	If Len(strProductCode) > 0 Then
		pstrSQL = "SELECT prodName FROM sfProducts WHERE prodID=" & wrapSQLValue(strProductCode, False, enDatatype_string)
		Set pobjRS = GetRS(pstrSQL)
		If pobjRS.EOF Then
			pblnFound = False
		Else
			pblnFound = True
			strProdName = pobjRS.Fields("prodName").Value
		End If
		pobjRS.Close
		Set pobjRS = Nothing
	Else
		pblnFound = False
	End If

	getProductInfo = pblnFound
	
End Function	'getProductInfo

'****************************************************************************************************************************************************************

Function hasProductPricingLevels()

	If Err.number <> 0 Then Err.Clear
	If Not cblnAddon_ProductPricing Then
		getProductPricingLevels = False
		Exit Function
	End If

	If isArray(maryPricingLevels) Then
		hasProductPricingLevels = True
	ElseIf Not isObject(maryPricingLevels) Then
		hasProductPricingLevels = getProductPricingLevels
	Else
		hasProductPricingLevels = isArray(maryPricingLevels)
	End If
	
End Function	'hasProductPricingLevels

'****************************************************************************************************************************************************************

Function getProductPricingLevels()

Dim i
Dim pobjRS
Dim pblnFound

	If Err.number <> 0 Then Err.Clear
	If Not cblnAddon_ProductPricing Then
		getProductPricingLevels = False
		Exit Function
	End If

	Set pobjRS = GetRS("Select PricingLevelID, PricingLevelName from PricingLevels Order By PricingLevelID")
	If pobjRS.EOF Then
		pblnFound = False
	Else
		ReDim maryPricingLevels(pobjRS.RecordCount - 1)
		For i = 1 To pobjRS.RecordCount
			maryPricingLevels(i - 1) = Trim(pobjRS.Fields("PricingLevelName").Value & "")
			pobjRS.MoveNext
		Next 'i
		pblnFound = True
	End If
	Call ReleaseObject(pobjRS)
	
	getProductPricingLevels = pblnFound
	
End Function	'getProductPricingLevels

'****************************************************************************************************************************************************************

Function setAttribute(byVal lngProductUID, byRef aryAttributeInfo)
'Input: ProductID, array of attribute information
'					0 - Name
'					1 - Type
'					2 - Required
'					3 - uid - optional
'Output: Attribute uid
'Action: either finds an attribute matching the attribute information or creates one

Dim pstrSQL
Dim pobjRS
Dim plngAttributeID

Dim pstrName
Dim pbytType
Dim pbytRequired

	If Len(lngProductUID) = 0 Then
		setAttribute = -1
		Exit Function
	ElseIf Not isArray(aryAttributeInfo) Then
		If Len(aryAttributeInfo) = 0 Then
			setAttribute = -1
			Exit Function
		End If
		pstrName = aryAttributeInfo
	Else
		pstrName = aryAttributeInfo(0)
	End If
	
	pbytType = getArrayValue(aryAttributeInfo, 1, 0)
	pbytRequired = getArrayValue(aryAttributeInfo, 2, 1)
	plngAttributeID = getArrayValue(aryAttributeInfo, 3, 0)

	If Len(CStr(plngAttributeID)) = 0 Then plngAttributeID = -1

	If plngAttributeID < 1 Then
	pstrSQL = "SELECT uid, ProductID, Name, Type, Required" _
			& " FROM Attributes" _
			& " WHERE ProductID=" & wrapSQLValue(lngProductUID, False, enDatatype_number) _
			& "       AND Name = " & wrapSQLValue(pstrName, False, enDatatype_string)
	Else
		pstrSQL = "SELECT uid, ProductID, Name, Type, Required" _
				& " FROM Attributes" _
				& " WHERE uid=" & wrapSQLValue(plngAttributeID, False, enDatatype_number) _
	End If

	Set pobjRS = GetRS(pstrSQL)
	If pobjRS.EOF Then
		pstrSQL = "Insert Into Attributes (ProductID, Name, Type, Required) Values (" _
				& wrapSQLValue(lngProductUID, False, enDatatype_number) & ", " _
				& wrapSQLValue(pstrName, False, enDatatype_string) & ", " _
				& wrapSQLValue(pbytType, False, enDatatype_number) & ", " _
				& wrapSQLValue(pbytRequired, False, enDatatype_number) & ")"
		cnn.Execute pstrSQL,,128
		
		pobjRS.Close
		pstrSQL = "SELECT uid" _
				& " FROM Attributes" _
				& " WHERE ProductID=" & wrapSQLValue(lngProductUID, False, enDatatype_number) _
				& "       AND Name = " & wrapSQLValue(pstrName, False, enDatatype_string) _
				& "       AND Type = " & wrapSQLValue(pbytType, False, enDatatype_number) _
				& "       AND Required = " & wrapSQLValue(pbytRequired, False, enDatatype_number)
		Set pobjRS = GetRS(pstrSQL)
	Else
		pstrSQL = "Update Attributes Set" _
				& " Name = " & wrapSQLValue(pstrName, False, enDatatype_string) & ", " _
				& " Type = " & wrapSQLValue(pbytType, False, enDatatype_number) & ", " _
				& " Required = " & wrapSQLValue(pbytRequired, False, enDatatype_number) & " " _
				& " Where uid=" & wrapSQLValue(pobjRS.Fields("uid").Value, False, enDatatype_number)	
		cnn.Execute pstrSQL,,128
	End If
	
	If pobjRS.EOF Then
		setAttribute = -1
	Else
		setAttribute = pobjRS.Fields("uid").Value
	End If

End Function	'setAttribute

'****************************************************************************************************************************************************************

Function setAttributeDetail(byVal lngAttributeID, byRef aryAttributeInfo)
'Input: ProductID, array of attribute detail information
'Output: Attribute uid
'Action: either finds an attribute detail matching the attribute information or creates one

Dim pstrSQL
Dim pobjRS

Dim plngAttributeID
Dim pstrName
Dim pstrPrice
Dim pstrWeight
Dim pbytPriceType
Dim pbytWeightType
Dim plngAttributeOrder
Dim pstrSmallImage
Dim pstrLargeImage
Dim pstrFileLocation

	If Err.number <> 0 Then Err.Clear

	If Len(lngAttributeID) = 0 Then
		setAttributeDetail = -1
		Exit Function
	ElseIf Not isArray(aryAttributeInfo) Then
		If Len(aryAttributeInfo) = 0 Then
			setAttributeDetail = -1
			Exit Function
		End If
		pstrName = aryAttributeInfo
	Else
		If UBound(aryAttributeInfo) >= 0 Then
			pstrName = aryAttributeInfo(0)
		Else
			Exit Function
		End If
	End If
	
	pstrPrice = getArrayValue(aryAttributeInfo, 1, "0")
	pstrWeight = getArrayValue(aryAttributeInfo, 2, "0")
	pbytPriceType = getArrayValue(aryAttributeInfo, 3, 0)
	pbytWeightType = getArrayValue(aryAttributeInfo, 4, 0)
	plngAttributeOrder = getArrayValue(aryAttributeInfo, 5, 0)
	pstrSmallImage = getArrayValue(aryAttributeInfo, 6, "")
	pstrLargeImage = getArrayValue(aryAttributeInfo, 7, "")
	pstrFileLocation = getArrayValue(aryAttributeInfo, 8, "")
	plngAttributeID = getArrayValue(aryAttributeInfo, 9, 0)
	
	If Len(CStr(plngAttributeID)) = 0 Then plngAttributeID = -1

	If Len(pstrPrice) = 0 Then
		pstrPrice = 0
		pbytPriceType = 0
	Else
		pstrPrice = CDbl(pstrPrice)
		pbytPriceType = CLng(pbytPriceType)
		If pstrPrice = 0 Then
			pbytPriceType = 0
		ElseIf pstrPrice > 0 Then
			If pbytPriceType <> 2 Then pbytPriceType = 1	'price may come in correctly as a positive value with the type set to be negative
		Else
			pbytPriceType = 2
			pstrPrice = Abs(pstrPrice)
		End If
	End If

	If plngAttributeID < 1 Then
	pstrSQL = "SELECT uid, AttributeID, Name, Price, Weight, PriceType, WeightType, AttributeOrder, SmallImage, LargeImage, FileLocation" _
			& " FROM AttributeDetail" _
			& " WHERE AttributeID=" & wrapSQLValue(lngAttributeID, False, enDatatype_number) _
			& "       AND Name = " & wrapSQLValue(pstrName, False, enDatatype_string)
	Else
		pstrSQL = "SELECT uid, AttributeID, Name, Price, Weight, PriceType, WeightType, AttributeOrder, SmallImage, LargeImage, FileLocation" _
				& " FROM AttributeDetail" _
				& " WHERE uid=" & wrapSQLValue(plngAttributeID, False, enDatatype_number)
	End If

	Set pobjRS = GetRS(pstrSQL)

	If pobjRS.EOF Then
		pstrSQL = "Insert Into AttributeDetail (AttributeID, Name, Price, Weight, PriceType, WeightType, AttributeOrder, SmallImage, LargeImage, FileLocation) Values (" _
				& wrapSQLValue(lngAttributeID, False, enDatatype_number) & ", " _
				& wrapSQLValue(pstrName, False, enDatatype_string) & ", " _
				& wrapSQLValue(pstrPrice, False, enDatatype_string) & ", " _
				& wrapSQLValue(pstrWeight, False, enDatatype_string) & ", " _
				& wrapSQLValue(pbytPriceType, False, enDatatype_number) & ", " _
				& wrapSQLValue(pbytWeightType, False, enDatatype_number) & ", " _
				& wrapSQLValue(plngAttributeOrder, False, enDatatype_number) & ", " _
				& wrapSQLValue(pstrSmallImage, False, enDatatype_string) & ", " _
				& wrapSQLValue(pstrLargeImage, False, enDatatype_string) & ", " _
				& wrapSQLValue(pstrFileLocation, False, enDatatype_string) & ")"
		cnn.Execute pstrSQL,,128
		
		pobjRS.Close
		pstrSQL = "SELECT uid" _
				& " FROM AttributeDetail" _
				& " WHERE AttributeID=" & wrapSQLValue(lngAttributeID, False, enDatatype_number) _
				& "       AND Name = " & wrapSQLValue(pstrName, False, enDatatype_string)
		Set pobjRS = GetRS(pstrSQL)
	Else
		pstrSQL = "Update AttributeDetail Set" _
				& " Name = " & wrapSQLValue(pstrName, False, enDatatype_string) & ", " _
				& " Price = " & wrapSQLValue(pstrPrice, False, enDatatype_string) & ", " _
				& " Weight = " & wrapSQLValue(pstrWeight, False, enDatatype_string) & ", " _
				& " PriceType = " & wrapSQLValue(pbytPriceType, False, enDatatype_number) & ", " _
				& " WeightType = " & wrapSQLValue(pbytWeightType, False, enDatatype_number) & ", " _
				& " AttributeOrder = " & wrapSQLValue(plngAttributeOrder, False, enDatatype_number) & ", " _
				& " SmallImage = " & wrapSQLValue(pstrSmallImage, False, enDatatype_string) & ", " _
				& " LargeImage = " & wrapSQLValue(pstrLargeImage, False, enDatatype_string) & ", " _
				& " FileLocation = " & wrapSQLValue(pstrFileLocation, False, enDatatype_string) & " " _
				& " Where uid=" & wrapSQLValue(pobjRS.Fields("uid").Value, False, enDatatype_number)
		cnn.Execute pstrSQL,,128
	End If
	
	If pobjRS.EOF Then
		setAttributeDetail = -1
	Else
		setAttributeDetail = pobjRS.Fields("uid").Value
	End If

End Function	'setAttributeDetail

'****************************************************************************************************************************************************************

Function setAttributeDetailSortOrder(byVal arySortOrder)

Dim i
Dim pstrSQL

    For i=0 to ubound(arySortOrder)-1
		pstrSQL = "Update AttributeDetail Set AttributeOrder=" & i & " where uid=" & arySortOrder(i)
		cnn.Execute pstrSQL,,128
    Next

End Function	'setAttributeDetailSortOrder

'****************************************************************************************************************************************************************

Function setGiftWrap(byVal strProductID, byVal bytActivate, byVal strPrice)

Dim pstrSQL
Dim pobjRS
Dim pbytIsActive
Dim pstrPrice

	'Check Data
	
	'IsActive must be 0 or 1 or True
	If CStr(bytActivate) = "1" Or CStr(bytActivate) = "-1" Or LCase(CStr(bytActivate)) = "true" Then
		pbytIsActive = 1
	Else
		pbytIsActive = 0
	End If

	'Price must be a number
	If Len(strPrice) > 0 AND isNumeric(strPrice) Then
		pstrPrice = strPrice
	Else
		pstrPrice = 0
	End If

	pstrSQL = "Select gwProdID From sfGiftWraps Where gwProdID=" & wrapSQLValue(strProductID, False, enDatatype_string)
	Set pobjRS = GetRS(pstrSQL)
	If pobjRS.EOF Then
		pstrSQL = "Insert Into sfGiftWraps (gwProdID, gwActivate, gwPrice) Values (" _
				& wrapSQLValue(strProductID, False, enDatatype_string) & ", " _
				& wrapSQLValue(pbytIsActive, False, enDatatype_boolean) & ", " _
				& wrapSQLValue(pstrPrice, False, enDatatype_string) & ")"
	Else
		pstrSQL = "Update sfGiftWraps Set" _
				& " gwActivate = " & wrapSQLValue(pbytIsActive, False, enDatatype_boolean) & ", " _
				& " gwPrice = " & wrapSQLValue(pstrPrice, False, enDatatype_string) & " " _
				& " Where gwProdID=" & wrapSQLValue(strProductID, False, enDatatype_string)	
	End If
	pobjRS.Close
	Set pobjRS = Nothing
	
	On Error Resume Next
	cnn.Execute pstrSQL,,128
	
	setGiftWrap = CBool(Err.number = 0)
	If Err.number <> 0 Then Err.Clear
		
End Function	'setGiftWrap

'****************************************************************************************************************************************************************

Function getManufacturerNameByID(byVal plngID)

Dim pobjRS

	Set pobjRS = GetRS("SELECT mfgName from sfManufacturers Where mfgID=" & wrapSQLValue(plngID, False, enDatatype_number))
	If Not pobjRS.EOF Then getManufacturerNameByID = Trim(pobjRS.Fields("mfgName").Value & "")
	Call ReleaseObject(pobjRS)
					
End Function	'getManufacturerNameByID

'****************************************************************************************************************************************************************

Function getManufacturerByName(byVal pstrName, byVal blnCreate)
'objrsNewProducts comes in as just the Manufacturer field

Dim pstrSQL
Dim pobjRS
Dim plngTempID

	pstrSQL = "SELECT mfgID from sfManufacturers Where mfgName=" & wrapSQLValue(pstrName, False, enDatatype_string)
	Set pobjRS = GetRS(pstrSQL)

	If pobjRS.State = 1 Then
		If pobjRS.EOF Then
			If blnCreate Then
				cnn.Execute "Insert Into sfManufacturers (mfgName) Values (" & wrapSQLValue(pstrName, False, enDatatype_string) & ")",,128
				Call ReleaseObject(pobjRS)
				
				Set pobjRS = GetRS(pstrSQL)
				If pobjRS.State = 1 Then
					If pobjRS.EOF Then
						getManufacturerByName = -1
					Else
						getManufacturerByName = pobjRS.Fields("mfgID").Value
					End If
				Else
					getManufacturerByName = -1
				End If
			Else
				getManufacturerByName = -1
			End If
		Else
			getManufacturerByName = pobjRS.Fields("mfgID").Value
		End If
	Else
		getManufacturerByName = -1
	End If
	Call ReleaseObject(pobjRS)
					
End Function	'getManufacturerByName

'****************************************************************************************************************************************************************

Function getVendorNameByID(byVal plngID)

Dim pobjRS

	Set pobjRS = GetRS("SELECT vendName from sfVendors Where vendID=" & wrapSQLValue(plngID, False, enDatatype_number))
	If Not pobjRS.EOF Then getVendorNameByID = Trim(pobjRS.Fields("vendName").Value & "")
	Call ReleaseObject(pobjRS)
					
End Function	'getVendorNameByID

'****************************************************************************************************************************************************************

Function getVendorByName(byVal pstrName, byVal blnCreate)
'objrsNewProducts comes in as just the Manufacturer field

Dim pstrSQL
Dim pobjRS
Dim plngTempID

	pstrSQL = "SELECT vendID from sfVendors Where vendName=" & wrapSQLValue(pstrName, False, enDatatype_string)
	Set pobjRS = GetRS(pstrSQL)

	If pobjRS.State = 1 Then
		If pobjRS.EOF Then
			If blnCreate Then
				cnn.Execute "Insert Into sfVendors (vendName) Values (" & wrapSQLValue(pstrName, False, enDatatype_string) & ")",,128
				Call ReleaseObject(pobjRS)
				
				Set pobjRS = GetRS(pstrSQL)
				If pobjRS.State = 1 Then
					If pobjRS.EOF Then
						getVendorByName = -1
					Else
						getVendorByName = pobjRS.Fields("vendID").Value
					End If
				Else
					getVendorByName = -1
				End If
			Else
				getVendorByName = -1
			End If
		Else
			getVendorByName = pobjRS.Fields("vendID").Value
		End If
	Else
		getVendorByName = -1
	End If
	Call ReleaseObject(pobjRS)
					
End Function	'getVendorByName

'****************************************************************************************************************************************************************

Function createProductCategoryAssignments(byVal strProductID, byVal aryCategoryUIDs)

Dim i

	Call deleteProductCategoryAssignments(strProductID)
	For i = 0 To UBound(aryCategoryUIDs)
		Call createProductCategoryAssignment(strProductID, aryCategoryUIDs(i))
	Next 'i

	createProductCategoryAssignments = True
	
End Function	'createProductCategoryAssignments

'****************************************************************************************************************************************************************

Function createProductCategoryAssignment(byVal strProductID, byVal lngCategoryID)

Dim pstrSQL
Dim pobjRSProductCategory
Dim pblnAdded

	On Error Resume Next
	
	If cblnSF5AE Then
		pblnAdded = False
		pstrSQL = "Select subcatDetailID From sfSubCatDetail" _
				& " Where ProdID=" & wrapSQLValue(strProductID, False, enDatatype_string) & " And subcatCategoryId=" & wrapSQLValue(lngCategoryID, False, enDatatype_number)
		Set pobjRSProductCategory = Server.CreateObject("ADODB.Recordset")
		With pobjRSProductCategory
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			If .EOF Then
				pstrSQL = "Insert Into sfSubCatDetail (ProdID,subcatCategoryId) Values (" & wrapSQLValue(strProductID, False, enDatatype_string) & ", " & wrapSQLValue(lngCategoryID, False, enDatatype_number) & ")"
				cnn.Execute pstrSQL,,128
				pblnAdded = True
			End If
			.Close
		End With
		Set pobjRSProductCategory = Nothing
	Else
		pstrSQL = "Update sfProducts Set prodCategoryId=" & wrapSQLValue(lngCategoryID, False, enDatatype_number) & " Where prodID=" & wrapSQLValue(strProductID, False, enDatatype_string)
		cnn.Execute pstrSQL,,128
		pblnAdded = True
	End If

	If Err.number <> 0 Then
		pblnAdded = False
		Err.Clear
	End If
	
	createProductCategoryAssignment = pblnAdded
	
End Function	'createProductCategoryAssignment

'****************************************************************************************************************************************************************

Function deleteProductCategoryAssignments(byVal strProductID)

Dim pstrSQL

	On Error Resume Next
	
	pstrSQL = "Delete From sfSubCatDetail Where ProdID=" & wrapSQLValue(strProductID, False, enDatatype_string)
	cnn.Execute pstrSQL,,128
	If Err.number <> 0 Then
		Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><BR>" & vbcrlf
		Response.Write "<font color=red>deletesfSubCatDetailAssignments - pstrSQL: " & pstrSQL & "</font><BR>" & vbcrlf
	End If

	deleteProductCategoryAssignments = CBool(Err.number = 0)
	
End Function	'deleteProductCategoryAssignments

'***********************************************************************************************

Function DeleteProduct(byVal lngUID)

Dim pstrSQL

'On Error Resume Next

	If len(lngUID) = 0 Then Exit Function
	
	'Delete attribute details
	pstrSQL = "Delete From sfAttributeDetail Where uid IN" _
			& " (" _
			& " SELECT sfAttributeDetail.uid" _
			& " FROM sfAttributes LEFT JOIN sfAttributeDetail ON sfAttributes.uid = sfAttributeDetail.AttributeID" _
			& " WHERE sfAttributes.attrProdId=" & wrapSQLValue(lngProductUID, True, enDatatype_string) _
			& " )"
	cnn.Execute pstrSQL,,128

	'Delete attributes
	pstrSQL = "Delete From sfAttributes WHERE attrProdId=" & wrapSQLValue(lngProductUID, True, enDatatype_string)
	cnn.Execute pstrSQL,,128

	'Delete categories
	pstrSQL = "Delete From sfSubCatDetail WHERE ProdID=" & wrapSQLValue(lngProductUID, True, enDatatype_string)
	cnn.Execute pstrSQL,,128

	'Delete GiftWrap
	pstrSQL = "Delete From sfGiftWraps Where gwProdID=" & wrapSQLValue(lngProductUID, True, enDatatype_string)
	cnn.Execute pstrSQL,,128

	'AE Specific
	If cblnSF5AE Then

		'Delete VolumePricing
		pstrSQL = "Delete From sfMTPrices Where mtProdID=" & wrapSQLValue(lngProductUID, True, enDatatype_string)
		cnn.Execute pstrSQL,,128
	
		'Delete Inventory
		pstrSQL = "Delete From sfInventory Where invenProdId=" & wrapSQLValue(lngProductUID, True, enDatatype_string)
		cnn.Execute pstrSQL,,128

		'Delete InventoryInfo
		pstrSQL = "Delete From sfInventoryInfo Where invenProdId=" & wrapSQLValue(lngProductUID, True, enDatatype_string)
		cnn.Execute pstrSQL,,128

	End If

	'Delete Product
	pstrSQL = "Delete from sfProducts where prodID=" & wrapSQLValue(lngProductUID, True, enDatatype_string)
	cnn.Execute pstrSQL, , 128
	
    If (Err.Number = 0) Then
        Call addMessageItem("The product was successfully deleted.", False)
        DeleteProduct = True
    Else
        Call addMessageItem("Error deleting product " & lngProdUID & ": " & Err.Description, True)
        DeleteProduct = False
    End If
    
End Function    'DeleteProduct

'***********************************************************************************************

Function DeleteAllProducts()

Dim pstrSQL
Dim pblnResult

'On Error Resume Next

	pblnResult = True
	
	'Delete attribute details
	pstrSQL = "Delete From sfAttributeDetail"
	cnn.Execute pstrSQL,,128

	'Delete attributes
	pstrSQL = "Delete From sfAttributes"
	cnn.Execute pstrSQL,,128

	'Delete categories
	pstrSQL = "Delete From sfSubCatDetail"
	cnn.Execute pstrSQL,,128

	'Delete GiftWrap
	pstrSQL = "Delete From sfGiftWraps"
	cnn.Execute pstrSQL,,128

	'AE Specific
	If cblnSF5AE Then

		'Delete VolumePricing
		pstrSQL = "Delete From sfMTPrices"
		cnn.Execute pstrSQL,,128
	
		'Delete Inventory
		pstrSQL = "Delete From sfInventory"
		cnn.Execute pstrSQL,,128

		'Delete InventoryInfo
		pstrSQL = "Delete From sfInventoryInfo"
		cnn.Execute pstrSQL,,128

	End If

	'Delete Products
	pstrSQL = "Delete from sfProducts"
	cnn.Execute pstrSQL, , 128
	
    If (Err.Number = 0) Then
        Call addMessageItem("The product catalog was successfully deleted.", False)
    Else
        Call addMessageItem("Error deleting the product catalog: " & Err.Description, True)
        pblnResult = False
    End If
    
	DeleteAllProducts = pblnResult	

End Function    'DeleteAllProducts

'***********************************************************************************************

Function DeleteAttributeCategory(byVal lngUID, byVal lngProductUID)

Dim pstrSQL
Dim pstrAttributeCategoryNames
Dim pobjRS
Dim pstrMessageOut
Dim pblnError

	pblnError = False

'On Error Resume Next

	If Len(lngUID) > 0 Then
		'Get attribute category name
		pstrSQL = "Select [Name] from Attributes where uid=" & lngUID
		Set pobjRS = GetRS(pstrSQL)
		Do While Not pobjRS.EOF
			If Len(pstrAttributeCategoryNames) = 0 Then
				pstrAttributeCategoryNames = pobjrs.Fields("Name").Value
			Else
				pstrAttributeCategoryNames = pstrAttributeCategoryNames & ", " & pobjrs.Fields("Name").Value
			End If
			pobjRS.MoveNext
		Loop
		Call ReleaseObject(pobjRS)
	
		'Delete attribute details
		pstrSQL = "Delete From AttributeDetail Where AttributeID=" & lngUID
		cnn.Execute pstrSQL,,128

		'Delete attributes
		pstrSQL = "Delete From Attributes WHERE uid=" & lngUID
		cnn.Execute pstrSQL,,128

		If (Err.Number = 0) Then
			pstrMessageOut = "The attribute category was successfully deleted."
		Else
			pstrMessageOut = "Error deleting attribute category " & lngUID & ": " & Err.Description
			pblnError = True
		End If

	ElseIf Len(lngProductUID) > 0 Then
		'Get attribute category name
		pstrSQL = "Select [Name] from Attributes where Attributes.ProductID=" & lngProductUID
		Set pobjRS = GetRS(pstrSQL)
		Do While Not pobjRS.EOF
			If Len(pstrAttributeCategoryNames) = 0 Then
				pstrAttributeCategoryNames = pobjrs.Fields("Name").Value
			Else
				pstrAttributeCategoryNames = pstrAttributeCategoryNames & ", " & pobjrs.Fields("Name").Value
			End If
			pobjRS.MoveNext
		Loop
		pobjRS.Close
		Set pobjRS = Nothing
	
		'Delete attribute details
		pstrSQL = "Delete From AttributeDetail Where uid IN" _
				& " (" _
				& " SELECT AttributeDetail.uid" _
				& " FROM Attributes LEFT JOIN AttributeDetail ON Attributes.uid = AttributeDetail.AttributeID" _
				& " WHERE Attributes.ProductID=" & lngProductUID _
				& " )"
		cnn.Execute pstrSQL,,128

		'Delete attributes
		pstrSQL = "Delete From Attributes WHERE ProductID=" & lngProductUID
		cnn.Execute pstrSQL,,128

		If (Err.Number = 0) Then
			pstrMessageOut = "The attribute category was successfully deleted."
		Else
			pstrMessageOut = "Error deleting attribute category " & lngUID & ": " & Err.Description
			pblnError = True
		End If

	Else
		pstrMessageOut = "No Attribute ID nor Product ID specified"
		pblnError = True
	End If

	Call addMessageItem(pstrMessageOut, pblnError)
	DeleteAttributeCategory = Not pblnError
    
End Function    'DeleteAttributeCategory

'***********************************************************************************************

Function DeleteAttributeDetail(byVal lngUID)

Dim pstrSQL
Dim pobjRS
Dim pstrAttributeCategoryNames

'On Error Resume Next

	If len(lngUID) > 0 Then
		'Get attribute name
		pstrSQL = "Select [Name] from AttributeDetail where uid=" & lngUID
		Set pobjRS = GetRS(pstrSQL)
		Do While Not pobjRS.EOF
			If Len(pstrAttributeCategoryNames) = 0 Then
				pstrAttributeCategoryNames = pobjrs.Fields("Name").Value
			Else
				pstrAttributeCategoryNames = pstrAttributeCategoryNames & ", " & pobjrs.Fields("Name").Value
			End If
			pobjRS.MoveNext
		Loop
		Call ReleaseObject(pobjRS)

		pstrSQL = "Delete from AttributeDetail where uid = " & lngUID
		cnn.Execute pstrSQL, , 128
		
		If (Err.Number = 0) Then
			Call addMessageItem("The attribute detail " & pstrAttributeCategoryNames & " was successfully deleted.", False)
			DeleteAttributeDetail = True
		Else
			Call addMessageItem("Error deleting attribute detail " & pstrAttributeCategoryNames & ": " & Err.Description, True)
			DeleteAttributeDetail = False
		End If
		
		If cblnSF5AE Then Call DeleteInventoryByAttributeDetail(lngUID)

	Else
		
    End If	'len(lngUID) > 0

    
End Function    'DeleteAttributeDetail

'***********************************************************************************************

Function DeleteInventoryByAttributeDetail(byVal lngUID)

Dim pstrSQL
Dim pobjRS
Dim pstrAttributeCategoryNames
Dim paryAttDetailID
Dim i

'On Error Resume Next

	If len(lngUID) > 0 And cblnSF5AE Then

		pstrSQL = "Select uid, AttributeDetailID From Inventory Where AttributeDetailID Like '" & lngUID & "'"
		Set pobjRS = GetRS(pstrSQL)
		With pobjRS
			Do While Not .EOF
				paryAttDetailID = Split(.Fields("AttributeDetailID").Value,",")
				For i=0 To uBound(paryAttDetailID)
					'check to make sure actual ids returned since 3 would return 31, 13, etc.
					If cLng(paryAttDetailID(i)) = cLng(lngUID) Then
						pstrSQL = "Delete From Inventory Where uid=" & .Fields("uid").Value
						cnn.Execute pstrSQL,,128
						Exit For
					End If
				Next
				.MoveNext
			Loop
		End With
		Call ReleaseObject(pobjRS)
		
    End If	'len(lngUID) > 0

End Function    'DeleteInventoryByAttributeDetail

'****************************************************************************************************************************************************************

Function updateVolumePricing(byVal strProductID, byVal strBreakLevel, byVal strAmount, byVal strDollarOrPercent)

Dim pstrSQL
Dim aBreakLevel
Dim aAmount
Dim aDollarOrPercent
Dim i,j
Dim paryTemp(2)
Dim plngCount
Dim paryMTP
Dim pblnError

	pblnError = False

	'clear out the existing pricing levels
	pstrSQL = "Delete From sfMTPrices Where mtProdID=" & wrapSQLValue(strProductID, True, enDatatype_string)
	cnn.Execute pstrSQL,,128
		
	If len(strBreakLevel) > 0 Then
		aBreakLevel = Split(strBreakLevel,",")
		aAmount = Split(strAmount,",")
		aDollarOrPercent = Split(strDollarOrPercent,",")
		
		plngCount = ubound(aBreakLevel)
		ReDim paryMTP(plngCount,2)
		For i=0 to plngCount
			paryMTP(i,0) = cLng(Trim(aBreakLevel(i)))
			paryMTP(i,1) = cDbl(Trim(aAmount(i)))
			paryMTP(i,2) = Trim(aDollarOrPercent(i))
		Next

		'Sort them in order by price
		For j=0 to plngCount
			For i=1 to plngCount
				If paryMTP(i-1,0) > paryMTP(i,0) Then
					paryTemp(0) = paryMTP(i,0)
					paryTemp(1) = paryMTP(i,1)
					paryTemp(2) = paryMTP(i,2)
					
					paryMTP(i,0) = paryMTP(i-1,0)
					paryMTP(i,1) = paryMTP(i-1,1)
					paryMTP(i,2) = paryMTP(i-1,2)

					paryMTP(i-1,0) = paryTemp(0)
					paryMTP(i-1,1) = paryTemp(1)
					paryMTP(i-1,2) = paryTemp(2)
				End If
			Next
		Next

		for i=0 to plngCount
			pstrSQL = "Insert Into sfMTPrices (ProductID,BreakLevel,Amount,DollarOrPercent) Values " _
					& "(" & strProductID & "," & paryMTP(i,0) & "," & paryMTP(i,1) & ",'" & paryMTP(i,2) & "')"
			cnn.Execute pstrSQL,,128
		next
	End If
	
	updateVolumePricing = Not pblnError
	
End Function	'updateVolumePricing

'****************************************************************************************************************************************************************

Function createVolumePriceBreak(byVal lngProductUID, byVal bytIndex, byVal lngBreakLevel, byVal strAmount, byVal bytBreakType)

Dim pblnSuccess
Dim pstrBreakType
Dim pstrLocalError
Dim pstrSQL

	pblnSuccess = False
	
	If bytBreakType = 0 Then
		pstrBreakType = "Amount"
	Else
		pstrBreakType = "Percent"
	End If

	pstrSQL = "Delete From sfMTPrices Where mtProdID=" & wrapSQLValue(lngProductUID, True, enDatatype_string) & " AND mtQuantity=" & lngBreakLevel
	pblnSuccess = Execute_NoReturn(pstrSQL, pstrLocalError)

	pstrSQL = "Insert Into sfMTPrices (mtIndex, mtProdID, mtType, mtQuantity, mtValue) " _
			& " Values (" _
			& wrapSQLValue(bytIndex, True, enDatatype_number) & ", " _
			& wrapSQLValue(lngProductUID, True, enDatatype_string) & ", " _
			& wrapSQLValue(pstrBreakType, True, enDatatype_string) & ", " _
			& wrapSQLValue(lngBreakLevel, True, enDatatype_number) & ", " _
			& wrapSQLValue(strAmount, True, enDatatype_string) & " " _
			& ")"

	If Execute_NoReturn(pstrSQL, pstrLocalError) Then
		pblnSuccess = True
	Else
		pblnSuccess = False
		WriteOutput "<b><font color=red>Error setting volume pricing</font></b><br>" & pstrLocalError & vbcrlf
	End If

	createVolumePriceBreak = pblnSuccess
	
End Function	'createVolumePriceBreak

'******************************************************************************************************************************************************************

Function SetImagePath(byVal strImage)

	If len(trim(strImage & "")) > 0 Then
		SetImagePath = SiteURL & strImage
	Else
		SetImagePath = "images/NoImage.gif"
	End If

End Function	'SetImagePath

'**********************************************************
'*	Initialization
'**********************************************************

mlngMaxRecords = LoadRequestValue("PageSize")
If len(mlngMaxRecords) = 0 Then mlngMaxRecords = clngDefaultMaxRecords
mbytSummaryTableHeight = 550


'****************************************************************************************************************************************************************

Function createSFCategory(byVal strCategoryName)

Dim plngCatID
Dim pobjRS
Dim pstrSQL
Dim pstrTempCategory

	plngCatID = -1
	
	pstrTempCategory = CStr("newCategory" & Session.SessionID)

	pstrSQL = "Insert Into sfCategories (catName, catHasSubCategory, catIsActive) Values (" & wrapSQLValue(pstrTempCategory, False, enDatatype_string) & ", 0, 1)"
	cnn.Execute pstrSQL,,128
	
	pstrSQL = "Select catID from sfCategories Where catName=" & wrapSQLValue(pstrTempCategory, False, enDatatype_string)
	Set pobjRS = GetRS(pstrSQL)
	If pobjRS.EOF Then
		debugprint "this shouldn't happen!", ""
	Else
		plngCatID = pobjRS.Fields("catID").Value
	End If
	pobjRS.Close
	Set pobjRS = Nothing

	pstrSQL = "Update sfCategories Set catName=" & wrapSQLValue(strCategoryName, False, enDatatype_string) & " Where catID=" & plngCatID
	cnn.Execute pstrSQL,,128
	
	createSFCategory = plngCatID

End Function	'createSFCategory

'****************************************************************************************************************************************************************

Function createCategory(byRef lngUID, byVal strCategoryName, byVal lngParentLevel, byVal ParentID, byVal bytIsActive)
'This function creates a category by name. It assumes a duplicate name may already exist in the database
'it inserts a unique id, retrieves the UID, and then updates the appropriate record with the actual category name
'Function returns the true if successful, false if unsuccessful
'lngUID is updated to UID if successful, -1 if not

Dim pblnSuccess
Dim pblnAdded
Dim pobjRS
Dim plngCatID
Dim pstrLocalError
Dim pstrSQL
Dim pstrTempCategory

	pblnSuccess = True
	'insert check to make sure empty name is not inserted
	If Len(strCategoryName) = 0 Then
		createCategory = False
		Exit Function
	End If

	pblnAdded = False

	'First need to check and see if a category matching the name, parentLevel, and ParentID match.
	'In theory the parentLevel is superfluous
	If cblnSF5AE Then
		pstrSQL = "Select subcatID As uid" _
				& " From sfSub_Categories" _
				& " Where subcatName=" & wrapSQLValue(strCategoryName, False, enDatatype_string) _
				& "	  AND Depth=" & wrapSQLValue(lngParentLevel, False, enDatatype_number) _
				& "   AND subcatCategoryId=" & wrapSQLValue(ParentID, False, enDatatype_number)
	Else
		pstrSQL = "Select catID as uid, catName from sfCategories Order By catName"
	End If

	'Response.Write "createCategory - SQL: " & pstrSQL & "<BR>"
    Set pobjRS = Server.CreateObject("ADODB.Recordset")
    With pobjRS
        .CursorLocation = 2 'adUseClient
        
        On Error Resume Next
        If Err.number <> 0 Then Err.Clear
        
		.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		If Err.number = 0 Then
			If .EOF Then
				plngCatID = -1
			Else
				plngCatID = .Fields("uid").Value
			End If
		Else
			WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><BR>" & vbcrlf
			WriteOutput "Error finding category sql: " & pstrSQL & "<BR>" & vbcrlf
			WriteOutput "This will result in a new category being created.<BR>" & vbcrlf
			plngCatID = -1
			Err.Clear
		End If
		.Close
	End With

	If plngCatID = -1 Then	
		'Insert the category placeholder and retrieve the UID
		
		If cblnSF5AE Then
			If lngParentLevel = 0 Then
				plngCatID = createSFCategory(strCategoryName)
				plngCatID = CreatePrimarySubCategory(plngCatID, strCategoryName)
			Else
			
			End If
		Else
			plngCatID = createSFCategory(strCategoryName)
		End If
		
		pblnAdded = True
		
	End If	'plngCatID = -1

	Set pobjRS = Nothing
	
	lngUID = plngCatID
	
	createCategory = pblnAdded

End Function	'createCategory

'********************************************************************************************

Function CreatePrimarySubCategory(byVal lngCatID, byVal strCatName)
'Purpose: Creates Initital 

Dim pstrSQL
Dim pobjRS

	pstrSQL = "INSERT INTO sfSub_Categories (subcatCategoryId, subcatName, Depth, HasProds, bottom, CatHierarchy)" _
			& " VALUES (" _
			& wrapSQLValue(lngCatID, True, enDatatype_string) & "," _
			& wrapSQLValue(strCatName, True, enDatatype_string) & "," _
			& "0," _
			& "0," _
			& "1," _
			& wrapSQLValue(Session.SessionID, False, enDatatype_string) _
			& ")"
	'debugprint "pstrSQL", pstrSQL
	cnn.Execute pstrSQL,,128
	
	pstrSQL = "Select subcatID From sfSub_Categories Where CatHierarchy=" & wrapSQLValue(Session.SessionID, False, enDatatype_string)
	'debugprint "pstrSQL", pstrSQL
	Set pobjRS = GetRS(pstrSQL)
	
	If Not pobjRS.EOF Then
		pstrSQL = "Update sfSub_Categories Set CatHierarchy=" & wrapSQLValue("none-" & pobjRS.Fields("subcatID").Value, False, enDatatype_string) & " Where subCatID=" & wrapSQLValue(pobjRS.Fields("subcatID").Value, False, enDatatype_number)
		'debugprint "pstrSQL", pstrSQL
		cnn.Execute pstrSQL,,128
	End If	'Not pobjRS.EOF
	Set pobjRS = Nothing
			
End Function	'CreatePrimarySubCategory

'****************************************************************************************************************************************************************

Function inventoryLowQty(byVal strProductUID, byVal strAttributeIDs)

Dim pstrSQL
Dim pobjRS

	pstrSQL = "Select invenLowFlag FROM sfInventory" _
			& " WHERE invenProdID= " & wrapSQLValue(strProductUID, False, enDatatype_string) _
			& " AND  invenAttDetailID=" & wrapSQLValue(strAttributeIDs, False, enDatatype_string)
	Set pobjRS = GetRS(pstrSQL)

	If pobjRS.EOF Then
		inventoryLowQty = -999999
	Else
		inventoryLowQty = pobjRS.Fields("invenLowFlag").Value
	End If
	Call ReleaseObject(pobjRS)
		
End Function	'inventoryLowQty

'****************************************************************************************************************************************************************

Function inventoryQty(byVal strProductUID, byVal strAttributeIDs)

Dim pstrSQL
Dim pobjRS

	If Len(Trim(strAttributeIDs) & "") = 0 Then strAttributeIDs = "0"
	pstrSQL = "Select invenInstock FROM sfInventory" _
			& " WHERE invenProdID= " & wrapSQLValue(strProductUID, False, enDatatype_string) _
			& " AND  invenAttDetailID=" & wrapSQLValue(strAttributeIDs, False, enDatatype_string)
	Set pobjRS = GetRS(pstrSQL)

	If pobjRS.EOF Then
		inventoryQty = -999999
	Else
		inventoryQty = pobjRS.Fields("invenInstock").Value
	End If
	Call ReleaseObject(pobjRS)
		
End Function	'inventoryQty

'****************************************************************************************************************************************************************

Function isInventoryTracked(byVal strProductUID)

Dim pstrSQL
Dim pobjRS

	pstrSQL = "Select invenbTracked FROM sfInventoryInfo WHERE invenProdID=" & wrapSQLValue(strProductUID, False, enDatatype_string)
	Set pobjRS = GetRS(pstrSQL)

	If pobjRS.EOF Then
		isInventoryTracked = False
	ElseIf pobjRS.Fields("invenbTracked").Value = 1 Then
		isInventoryTracked = True
	Else
		isInventoryTracked = False
	End If
	Call ReleaseObject(pobjRS)
		
End Function	'isInventoryTracked

'****************************************************************************************************************************************************************

Sub updateInventoryQtyDelta(byVal strProductUID, byVal strAttributeIDs, byVal qtyDelta)

Dim pstrSQL

	If Len(Trim(strAttributeIDs) & "") = 0 Then strAttributeIDs = "0"
	pstrSQL = "UPDATE sfInventory SET sfInventory.invenInStock = sfInventory.invenInStock+" & qtyDelta _
			& " WHERE invenProdID= " & wrapSQLValue(strProductUID, False, enDatatype_string) _
			& " AND  invenAttDetailID=" & wrapSQLValue(strAttributeIDs, False, enDatatype_string)
	cnn.Execute pstrSQL,,128
	
End Sub	'updateInventoryQtyDelta

'****************************************************************************************************************************************************************

Function updateCategoryHasProductsStatus(byVal strSubCatID)

Dim pstrSQL

	If Len(strSubCatID) > 0 Then
		pstrSQL = "Update sfSub_Categories Set HasProds=1 Where subcatID=" & strSubCatID
	
	Else
		pstrSQL = "Update sfSub_Categories Set HasProds=0 Where" _
				& " subcatID In " _
				& "  (SELECT sfSub_Categories.subcatID" _
				& "   FROM sfSubCatDetail RIGHT JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
				& "   WHERE sfSubCatDetail.ProdID Is Null" _
				& "  )"
		cnn.Execute pstrSQL,,128
		
		pstrSQL = "Update sfSub_Categories Set HasProds=1 Where" _
				& " subcatID In " _
				& "  (SELECT sfSub_Categories.subcatID" _
				& "   FROM sfSubCatDetail RIGHT JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
				& "   WHERE sfSubCatDetail.ProdID Is Not Null" _
				& "  )"
		cnn.Execute pstrSQL,,128
	End If
	
End Function	'updateCategoryHasProductsStatus

'****************************************************************************************************************************************************************

Function updateSubCategoryHasProductsStatus(byVal strSubCatID, byVal strHasProds)

Dim pstrSQL
Dim pobjRS
Dim pblnHasProds
Dim plngSubCatID
Dim plngDepth
Dim plngParentID
Dim paryCatHeir
Dim pblnResult

	pblnResult = False
	
	If Len(strSubCatID) > 0 Then
		pstrSQL = "Select subcatCategoryId, CatHierarchy, Depth, HasProds, ProdID" _
				& "   FROM sfSubCatDetail RIGHT JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
				& "   WHERE Depth > 1 And subcatID=" & strSubCatID
		Set pobjRS = GetRS(pstrSQL)
		
		If pobjRS.EOF Then
			pblnResult = True
		Else
			If Len(CStr(strHasProds)) = 0 Then
				pblnHasProds = Not isNull(pobjRS.Fields("ProdID").Value)
			Else
				pblnHasProds = ConvertToBoolean(strHasProds, False)
			End If
			
			plngDepth = pobjRS.Fields("Depth").Value - 1	'adjust for 0 based array
			paryCatHeir = Split(Trim(pobjRS.Fields("CatHierarchy").Value), "-")
			plngParentID = paryCatHeir(plngDepth - 1)
			
			If pobjRS.Fields("HasProds").Value <> Abs(pblnHasProds * -1) Then
				pstrSQL = "Update sfSub_Categories Set HasProds=" & Abs(pblnHasProds * -1) _
						& " subcatID=" & strSubCatID
				cnn.Execute pstrSQL,,128
			End If
			
			updateSubCategoryHasProductsStatus = updateSubCategoryHasProductsStatus(plngParentID, pblnHasProds)
		
		End If	'pobjRS.EOF
	
	End If	'Len(strSubCatID) > 0
	
	updateCategoryHasProductsStatus = pblnResult
	
End Function	'updateCategoryHasProductsStatus

'********************************************************************************************

'***************************************************************************************************************************************************************************************
' From Product Reviews
'***************************************************************************************************************************************************************************************

Dim mlngNumReviews
Dim mlngAvgReviewScore

'***************************************************************************************************************************************************************************************

Function getFoundUsefulVotes(byVal lngcontentID, byRef aryVotes)

Dim pblnSuccess
Dim pobjCmd
Dim pobjRS

	pblnSuccess = False
	
	If Len(lngcontentID) = 0 Or Not isNumeric(lngcontentID) Then
		getFoundUsefulVotes = pblnSuccess
		Exit Function
	End If

	Set pobjCmd  = Server.CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select contentFoundUsefulScore From contentFoundUseful Where contentFoundUsefulContentID=?"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("contentFoundUsefulContentID", adInteger, adParamInput, 4, lngcontentID)
		Set pobjRS = .Execute
		
		If Not pobjRS.EOF Then
			aryVotes = pobjRS.GetRows()
			pblnSuccess = True
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing
	
	getFoundUsefulVotes = pblnSuccess

End Function	'getFoundUsefulVotes

'***************************************************************************************************************************************************************************************

Function loadProductReviews(byVal strProductID, byRef aryProductReviews)

Dim pblnResult
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim paryVotes
Dim i,j

	mlngNumReviews = 0
	mlngAvgReviewScore = 0
	pblnResult = False

	If Len(strProductID) = 0 Or Len(strProductID) > 50 Then
		loadProductReviews = pblnResult
		Exit Function
	End If

	pstrSQL = "SELECT contentID, contentContent, contentAuthorName, contentAuthorEmail, contentAuthorShowEmail, contentAuthorRating, contentDateCreated, 0 as foundUseful, 0 as foundWorthless" _
			& " FROM sfProducts RIGHT JOIN content ON sfProducts.sfProductID = content.contentReferenceID" _
			& " WHERE ((content.contentApprovedForDisplay=1 Or content.contentApprovedForDisplay=-1) AND (content.contentContentType=5) AND (sfProducts.prodID=?))" _
			& " ORDER BY contentDateCreated DESC"

	Set pobjCmd  = Server.CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, strProductID)
		Set pobjRS = .Execute
		
		If Not pobjRS.EOF Then
			aryProductReviews = pobjRS.GetRows()

			mlngNumReviews = UBound(aryProductReviews, 2) + 1
			For i = 0 To mlngNumReviews - 1
				mlngAvgReviewScore = mlngAvgReviewScore + aryProductReviews(5, i)
				If getFoundUsefulVotes(aryProductReviews(0, i), paryVotes) Then
					For j = 0 To UBound(paryVotes, 2)
						If paryVotes(0, j) = 0 Then
							aryProductReviews(8, i) = aryProductReviews(8, i) + 1
						Else
							aryProductReviews(7, i) = aryProductReviews(7, i) + 1
						End If
					Next 'j
				End If
			Next 'i
			If mlngNumReviews <> 0 Then mlngAvgReviewScore = mlngAvgReviewScore / mlngNumReviews

			pblnResult = True
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
		
	End With	'pobjCmd
	Set pobjCmd = Nothing
	
	loadProductReviews = pblnResult
	
End Function	'loadProductReviews

'***************************************************************************************************************************************************************************************

Function ratingDisplayImage(byVal dblRating)

	If dblRating >= 4.5 Then
		ratingDisplayImage = "<img src=""../../../images/stars5.jpg"">"
	ElseIf dblRating >= 3.5 Then
		ratingDisplayImage = "<img src=""../../../images/stars4.jpg"">"
	ElseIf dblRating >= 2.5 Then
		ratingDisplayImage = "<img src=""../../../images/stars3.jpg"">"
	ElseIf dblRating >= 1.5 Then
		ratingDisplayImage = "<img src=""../../../images/stars2.jpg"">"
	ElseIf dblRating >= 0.5 Then
		ratingDisplayImage = "<img src=""../../../images/stars1.jpg"">"
	Else
		ratingDisplayImage = "<img src=""../../../images/stars0.jpg"">"
	End If
	
End Function	'ratingDisplayImage

%>
