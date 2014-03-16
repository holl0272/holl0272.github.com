<%
'********************************************************************************
'*   Product Manager Version SF 5.0 					                        *
'*   Release Version:	2.00.002		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		April 18, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   2.00.004 (February 10, 2005)												*
'*   - Updated product summary selection to avoid duplicate display 			*
'*     due to multiple category assignments									    *
'*                                                                              *
'*   2.00.003 (April 18, 2004)													*
'*   - Miscellaneous enhancements									            *
'*                                                                              *
'*   2.00.002 (September 11, 2003)                                              *
'*   - Added support for sites which must use https - ignore warning            *
'*   - Added support for Attribute Extender custom types				        *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Class clsProduct
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private prsProducts
Private prsAttributes
Private prsAttributeDetails
Private pblnError
Private plngRecordCount

Private prsInventory
Private prsInventoryInfo
Private prsMTP
Private prsCategories

'database variables
Private pstrprodID

Private plngprodAttrNum
Private plngprodCategoryId
Private pblnprodCountryTaxIsActive
Private pdtprodDateAdded
Private pdtprodDateModified
Private pstrprodDescription
Private pblnprodEnabledIsActive
Private pblnprodLimitQtyToMTP
Private pstrprodFileName
Private pstrUpgradeVersion
Private pstrPackageCodes
Private pstrVersion
Private pstrReleaseDate
Private pdblInstallationHours
Private pblnMyProduct
Private pblnInstallationRequired
Private pblnIncludeInSearch
Private pblnIncludeInRandomProduct
Private pstrprodMaxDownloads
Private pstrprodDownloadValidFor
Private pdblprodHeight
Private pstrprodImageSmallPath
Private pstrprodImageLargePath
Private pdblprodLength
Private pstrprodLink
Private plngprodManufacturerId
Private pstrprodMessage
Private pstrprodName
Private pstrprodNamePlural
Private pstrprodPrice
Private pblnprodSaleIsActive
Private pstrprodSalePrice
Private pstrprodShip
Private pdblprodFixedShippingCharge
Private pstrprodSpecialShippingMethods
Private pblnprodShipIsActive
Private pstrprodShortDescription
Private pblnprodStateTaxIsActive
Private plngprodVendorId
Private pdblprodWeight
Private pdblprodWidth
Private pblnBuyersClubIsPercentage
Private pdblBuyersClubPointValue

Private plngattrID
Private pstrattrProdId
Private pstrattrName
Private pbytattrDisplayOrder
Private pbytattrDisplayStyle
Private pstrattrURL_Field
Private pstrattrURL
Private pstrattrDisplay_Field
Private pstrattrDisplay
Private pstrattrExtra_Field
Private pstrattrExtra
Private pstrattrImage_Field
Private pstrattrImage
Private pstrattrSKU_Field
Private pstrattrSKU

Private plngattrdtID
Private plngattrdtAttributeId
Private pstrattrdtName
Private pstrattrdtPrice
Private pintattrdtType
Private pintattrdtOrder
Private pstrattrdtImage
Private pstrattrdtImage_Field
Private pdblattrdtWeight
Private pstrattrdtWeight_Field
Private pstrattrdtURL
Private pstrattrdtURL_Field
Private pstrattrdtDisplay
Private pstrattrdtDisplay_Field
Private pstrattrdtExtra
Private pstrattrdtExtra_Field
Private pstrattrdtExtra1
Private pstrattrdtExtra1_Field
Private pstrattrdtFileName
Private pstrattrdtFileName_Field
Private pstrattrdtSKU
Private pstrattrdtSKU_Field
Private pbytattrdtDefault
Private pstrattrdtDefault_Field

Private pblngwActivate
Private pstrgwPrice

Private pstrprodHandlingFee
Private pstrprodSetupFee
Private pstrprodSetupFeeOneTime
Private pstrprodMinQty
Private pstrprodIncrement

Private pstrCategoryList
Private pstrCategoryFilterList

Private pstrProdDictionaryList
Private pstrCategoryArray

Private pblnAEInventoryChanged
Private pblnAEMTPChanged
Private pblnAECategoryChanged

Private pblnCustomMTP
Private pblnTextBasedAttribute
Private pblnPricingLevel
Private pblnAttributeCategoryOrderable

Private pstrFieldAttributeCategoryOrderable

'added for pricing levels
Private pstrprodPLPrice
Private pstrprodPLSalePrice
Private pstrattrdtPLPrice
Private pstrgwPLPrice

'added for SEO
Private pstrpageName
Private pstrmetaTitle
Private pstrmetaDescription
Private pstrmetaKeywords

Private pstrRelatedProducts

Private pstrprodAdditionalImages
Private plngNumAdditionalImages
Private paryAdditionalImageText
Private paryAdditionalImage
Private paryAdditionalImageDesc
Private pbytprodDisplayAdditionalImagesInWindow

Private pblnprodEnableAlsoBought
Private pblnprodEnableReviews

Private paryCustomValues

Public Property Get TextBasedAttribute()
    TextBasedAttribute = pblnTextBasedAttribute
End Property
Public Property Get AttributeCategoryOrderable()
    AttributeCategoryOrderable = pblnAttributeCategoryOrderable
End Property
Public Property Get PricingLevel()
    PricingLevel = pblnPricingLevel
End Property

'***********************************************************************************************

Private Sub class_Initialize()
	
	
	pstrFieldAttributeCategoryOrderable = "attrDisplayOrder"	'you must update the sfAttributes table to use this feature
	'pstrFieldAttributeCategoryOrderable = ""

    cstrDelimeter  = ";"
	pblnPricingLevel =  True
	pblnTextBasedAttribute = True
	pblnAttributeCategoryOrderable = CBool(Len(pstrFieldAttributeCategoryOrderable) > 0)
	pblnCustomMTP = False
	Call InitializeCustomValues(paryCustomValues)

	pstrattrDisplay_Field = "attrDisplay"
	pstrattrExtra_Field = "attrExtra"
	pstrattrImage_Field = "attrImage"
	pstrattrSKU_Field = "attrSKU"
	pstrattrURL_Field = "attrURL"
	
	pstrattrdtDefault_Field = "attrdtDefault"
	pstrattrdtDisplay_Field = "attrdtDisplay"
	pstrattrdtExtra_Field = "attrdtExtra"
	pstrattrdtExtra1_Field = "attrdtExtra1"
	pstrattrdtFileName_Field = "attrdtFileName"
	pstrattrdtImage_Field = "attrdtImage"
	pstrattrdtSKU_Field = "attrdtSKU"
	pstrattrdtURL_Field = "attrdtURL"
	pstrattrdtWeight_Field = "attrdtWeight"
	
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	Call ReleaseObject(prsProducts)
	Call ReleaseObject(prsAttributes)
	Call ReleaseObject(prsAttributeDetails)

	Call ReleaseObject(prsInventory)
	Call ReleaseObject(prsInventoryInfo)
	Call ReleaseObject(prsMTP)
	Call ReleaseObject(prsCategories)

End Sub

'***********************************************************************************************

Private Sub LoadCustomValues(objRS)

Dim i

	If pblnCustomMTP Then
		'If len(pstrprodID) > 0 Then	Set prsMTP = GetRS("Select * from ssPricingLevels Where prodID='" & SQLSafe(pstrprodID) & "' Order By PricingLevel Asc")
		If len(pstrprodID) > 0 Then	Set prsMTP = GetRS("Select * from PricingLevels Where prodID='" & SQLSafe(pstrprodID) & "' Order By PricingLevel Asc")
	End If

	If Not isArray(paryCustomValues) Then Exit Sub
	If objRS.EOF Then Exit Sub
	
	For i = 0 To UBound(paryCustomValues)
		paryCustomValues(i)(2) = objRS.Fields(paryCustomValues(i)(1)).Value
	Next 'i
	
End Sub	'LoadCustomValues

Private Sub UpdateCustomValues(objRS)

Dim i

On Error Resume Next

	If Not isArray(paryCustomValues) Then Exit Sub

	For i = 0 To UBound(paryCustomValues)
		If paryCustomValues(i)(3) = "checkbox" Then
			If Len(paryCustomValues(i)(2)) > 0 Then
				objRS.Fields(paryCustomValues(i)(1)).Value = paryCustomValues(i)(2)
			Else
				objRS.Fields(paryCustomValues(i)(1)).Value = CBool(Len(paryCustomValues(i)(2)) > 0)
			End If
		Else
			objRS.Fields(paryCustomValues(i)(1)).Value = paryCustomValues(i)(2)
			
			If Err.number <> 0 Then
				objRS.Fields(paryCustomValues(i)(1)).Value = Null
				paryCustomValues(i)(2) = ""
				Err.Clear
			End If
		End If
	Next 'i

	If Err.number <> 0 Then
		Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
		Err.Clear
	End If

	If pblnCustomMTP Then Call UpdateCustomMTP
	
End Sub	'UpdateCustomValues

Private Sub LoadCustomValuesFromRequest()

Dim i

	If Not isArray(paryCustomValues) Then Exit Sub
	For i = 0 To UBound(paryCustomValues)
		paryCustomValues(i)(2) = Trim(Request.Form(paryCustomValues(i)(1)))
	Next 'i
	
End Sub	'LoadCustomValuesFromRequest

'***********************************************************************************************

Public Property Get rsPricingLevels()
	If isObject(prsMTP) Then Set rsPricingLevels = prsMTP
End Property

Public Function UpdateCustomMTP()

Dim rs
Dim sql

	If len(pstrprodID) > 0 Then
		
			sql = "Delete From ssPricingLevels Where prodID='" & SQLSafe(pstrprodID) & "'"
			cnn.Execute sql,,128
		
			Dim pstrPricingLevel, aPricingLevel
			Dim pstrPricingAmount, aPricingAmount
			Dim i,j
			Dim paryTemp(1)
			Dim plngCount
			Dim paryPricingLevel
		
			pstrPricingLevel = Request.Form("PricingLevel")
			pstrPricingAmount = Request.Form("PricingAmount")

			If len(pstrPricingLevel) > 0 Then
				aPricingLevel = Split(pstrPricingLevel,",")
				aPricingAmount = Split(pstrPricingAmount,",")
				
				plngCount = ubound(aPricingLevel)
				ReDim paryPricingLevel(plngCount,2)
				For i=0 to plngCount
					paryPricingLevel(i,0) = cLng(Trim(aPricingLevel(i)))
					paryPricingLevel(i,1) = cDbl(Trim(aPricingAmount(i)))
				Next
		
				For j=0 to plngCount
					For i=1 to plngCount
						If paryPricingLevel(i-1,0) > paryPricingLevel(i,0) Then
							paryTemp(0) = paryPricingLevel(i,0)
							paryTemp(1) = paryPricingLevel(i,1)

							paryPricingLevel(i,0) = paryPricingLevel(i-1,0)
							paryPricingLevel(i,1) = paryPricingLevel(i-1,1)

							paryPricingLevel(i-1,0) = paryTemp(0)
							paryPricingLevel(i-1,1) = paryTemp(1)
						End If
					Next
				Next
		
				On Error Resume Next
				for i=0 to plngCount
					sql = "Insert Into ssPricingLevels (prodID,PricingLevel,PricingAmount) Values " _
						& "('" & SQLSafe(pstrprodID) & "'," & paryPricingLevel(i,0) & "," & paryPricingLevel(i,1) & ")"
					cnn.Execute sql,,128

					If Err.number = -2147217900 Then
						Response.Write "<font color=red>You've entered a Pricing Level for Qty-" & paryPricingLevel(i,0) & " that already exists for this Product</font>"
						Err.Clear
					ElseIf Err.number > 0 Then
						debugprint "Error: " & Err.number,Err.Description
						Err.Clear
					End If
					next
			End If
	End If

End Function	'UpdateCustomMTP

'***********************************************************************************************

Public Property Get Message()
    Message = pstrMessage
End Property

Public Sub OutputMessage()

Dim i
Dim aError
	
	If Len(pstrMessage) > 0 Then Response.Write "<fieldset><legend>Result</legend>"

    aError = Split(pstrMessage, cstrDelimeter)
    For i = 0 To UBound(aError)
        If pblnError Then
            Response.Write "<FONT color=Red>" & aError(i) & "</FONT><br />"
        Else
            Response.Write "" & aError(i) & "<br />"
        End If
    Next 'i

	If Len(pstrMessage) > 0 Then Response.Write "</fieldset>"

End Sub 'OutputMessage

'***********************************************************************************************

Public Property Get CustomValues()
    CustomValues = paryCustomValues
End Property

Public Property Get CustomMTP()
    CustomMTP = pblnCustomMTP
End Property

Public Property Get prodAttrNum()
    prodAttrNum = plngprodAttrNum
End Property

Public Property Get prodCategoryId()
    prodCategoryId = plngprodCategoryId
End Property

Public Property Get prodCountryTaxIsActive()
    prodCountryTaxIsActive = pblnprodCountryTaxIsActive
End Property

Public Property Get prodDateAdded()
    prodDateAdded = pdtprodDateAdded
End Property

Public Property Get prodDateModified()
    prodDateModified = pdtprodDateModified
End Property

Public Property Get prodDescription()
    prodDescription = pstrprodDescription
End Property

Public Property Get prodEnabledIsActive()
    prodEnabledIsActive = pblnprodEnabledIsActive
End Property

Public Property Get prodLimitQtyToMTP()
    prodLimitQtyToMTP = pblnprodLimitQtyToMTP
End Property

Public Property Get prodFileName()
    prodFileName = pstrprodFileName
End Property

Public Property Get UpgradeVersion()
    UpgradeVersion = pstrUpgradeVersion
End Property

Public Property Get packageCodes()
    packageCodes = pstrPackageCodes
End Property

Public Property Get Version()
    Version = pstrVersion
End Property

Public Property Get ReleaseDate()
    ReleaseDate = pstrReleaseDate
End Property

Public Property Get InstallationHours()
    InstallationHours = pdblInstallationHours
End Property

Public Property Get MyProduct()
    MyProduct = pblnMyProduct
End Property

Public Property Get InstallationRequired()
    InstallationRequired = pblnInstallationRequired
End Property

Public Property Get IncludeInSearch()
    IncludeInSearch = pblnIncludeInSearch
End Property

Public Property Get IncludeInRandomProduct()
    IncludeInRandomProduct = pblnIncludeInRandomProduct
End Property

Public Property Get prodMaxDownloads()
    prodMaxDownloads = pstrprodMaxDownloads
End Property

Public Property Get prodDownloadValidFor()
    prodDownloadValidFor = pstrprodDownloadValidFor
End Property

Public Property Get prodHeight()
    prodHeight = pdblprodHeight
End Property

Public Property Get prodID()
    prodID = pstrprodID
End Property
Public Property Let prodID(strProdID)
    pstrprodID = strprodID
End Property

Public Property Get prodImageLargePath()
    prodImageLargePath = pstrprodImageLargePath
End Property

Public Property Get prodImageSmallPath()
    prodImageSmallPath = pstrprodImageSmallPath
End Property

Public Property Get prodLength()
    prodLength = pdblprodLength
End Property

Public Property Get prodLink()
    prodLink = pstrprodLink
End Property

Public Property Get prodManufacturerId()
    prodManufacturerId = plngprodManufacturerId
End Property

Public Property Get prodMessage()
    prodMessage = pstrprodMessage
End Property

Public Property Get prodName()
    prodName = pstrprodName
End Property

Public Property Get prodNamePlural()
    prodNamePlural = pstrprodNamePlural
End Property

Public Property Get prodPrice()
    prodPrice = pstrprodPrice
End Property

Public Property Get prodSaleIsActive()
    prodSaleIsActive = pblnprodSaleIsActive
End Property

Public Property Get prodSalePrice()
    prodSalePrice = pstrprodSalePrice
End Property

Public Property Get buyersClubIsPercentage()
    buyersClubIsPercentage = pblnBuyersClubIsPercentage
End Property

Public Property Get buyersClubPointValue()
    buyersClubPointValue = pdblBuyersClubPointValue
End Property

Public Property Get prodShip()
    prodShip = pstrprodShip
End Property

Public Property Get prodFixedShippingCharge()
    prodFixedShippingCharge = pdblprodFixedShippingCharge
End Property

Public Property Get prodSpecialShippingMethods()
    prodSpecialShippingMethods = pstrprodSpecialShippingMethods
End Property

Public Property Get prodShipIsActive()
    prodShipIsActive = pblnprodShipIsActive
End Property

Public Property Get prodShortDescription()
    prodShortDescription = pstrprodShortDescription
End Property

Public Property Get prodStateTaxIsActive()
    prodStateTaxIsActive = pblnprodStateTaxIsActive
End Property

Public Property Get prodVendorId()
    prodVendorId = plngprodVendorId
End Property

Public Property Get prodWeight()
    prodWeight = pdblprodWeight
End Property

Public Property Get prodWidth()
    prodWidth = pdblprodWidth
End Property

Public Property Get attrID()
    attrID = plngattrID
End Property

'----------------------------------------------------------

Public Property Get attrName()
    attrName = pstrattrName
End Property

Public Property Get attrURL()
    attrURL = pstrattrURL
End Property

Public Property Get attrURL_Field()
    attrURL_Field = pstrattrURL_Field
End Property

Public Property Get attrDisplay_Field()
    attrDisplay_Field = pstrattrDisplay_Field
End Property

Public Property Get attrDisplay()
    attrDisplay = pstrattrDisplay
End Property

Public Property Get attrExtra_Field()
    attrExtra_Field = pstrattrExtra_Field
End Property

Public Property Get attrExtra()
    attrExtra = pstrattrExtra
End Property

Public Property Get attrImage_Field()
    attrImage_Field = pstrattrImage_Field
End Property

Public Property Get attrImage()
    attrImage = pstrattrImage
End Property

Public Property Get attrSKU_Field()
    attrSKU_Field = pstrattrSKU_Field
End Property

Public Property Get attrSKU()
    attrSKU = pstrattrSKU
End Property

Public Property Get attrProdId()
    attrProdId = pstrattrProdId
End Property

Public Property Get attrDisplayStyle()
	attrDisplayStyle = pbytattrDisplayStyle
End Property

Public Property Get attrDisplayOrder()
	attrDisplayOrder = pbytattrDisplayOrder
End Property

'----------------------------------------------------------

Public Property Get prodEnableAlsoBought()
	prodEnableAlsoBought = pblnprodEnableAlsoBought
End Property

Public Property Get prodEnableReviews()
	prodEnableReviews = pblnprodEnableReviews
End Property

Public Property Get prodDisplayAdditionalImagesInWindow()
	prodDisplayAdditionalImagesInWindow = pbytprodDisplayAdditionalImagesInWindow
End Property
'----------------------------------------------------------

Public Property Get pageName()
	pageName = pstrpageName
End Property

Public Property Get metaTitle()
	metaTitle = pstrmetaTitle
End Property

Public Property Get metaDescription()
	metaDescription = pstrmetaDescription
End Property

Public Property Get metaKeywords()
	metaKeywords = pstrmetaKeywords
End Property

'----------------------------------------------------------

Public Property Get NumAdditionalImages()
	NumAdditionalImages = plngNumAdditionalImages
End Property

Public Property Get AdditionalImageText()
	AdditionalImageText = paryAdditionalImageText
End Property

Public Property Get AdditionalImage()
	AdditionalImage = paryAdditionalImage
End Property

Public Property Get AdditionalImageDesc()
	AdditionalImageDesc = paryAdditionalImageDesc
End Property

'----------------------------------------------------------

Public Property Get attrdtAttributeId()
    attrdtAttributeId = plngattrdtAttributeId
End Property

Public Property Get attrdtID()
    attrdtID = plngattrdtID
End Property

Public Property Get attrdtName()
    attrdtName = pstrattrdtName
End Property

Public Property Get attrdtOrder()
    attrdtOrder = pintattrdtOrder
End Property

Public Property Get attrdtPrice()
    attrdtPrice = pstrattrdtPrice
End Property

Public Property Get attrdtType()
    attrdtType = pintattrdtType
End Property

Public Property Get attrdtImage()
    attrdtImage = pstrattrdtImage
End Property

Public Property Get attrdtURL()
    attrdtURL = pstrattrdtURL
End Property

Public Property Get attrdtDisplay()
    attrdtDisplay = pstrattrdtDisplay
End Property

Public Property Get attrdtExtra()
    attrdtExtra = pstrattrdtExtra
End Property

Public Property Get attrdtExtra1()
    attrdtExtra1 = pstrattrdtExtra1
End Property

Public Property Get attrdtFileName()
    attrdtFileName = pstrattrdtFileName
End Property

Public Property Get attrdtSKU()
    attrdtSKU = pstrattrdtSKU
End Property

Public Property Get attrdtDefault()
    attrdtDefault = pbytattrdtDefault
End Property

Public Property Get attrdtWeight()
    attrdtWeight = pdblattrdtWeight
End Property

Public Property Get attrdtImage_Field()
    attrdtImage_Field = pstrattrdtImage_Field
End Property

Public Property Get attrdtURL_Field()
    attrdtURL_Field = pstrattrdtURL_Field
End Property

Public Property Get attrdtDisplay_Field()
    attrdtDisplay_Field = pstrattrdtDisplay_Field
End Property

Public Property Get attrdtExtra_Field()
    attrdtExtra_Field = pstrattrdtExtra_Field
End Property

Public Property Get attrdtExtra1_Field()
    attrdtExtra1_Field = pstrattrdtExtra1_Field
End Property

Public Property Get attrdtFileName_Field()
    attrdtFileName_Field = pstrattrdtFileName_Field
End Property

Public Property Get attrdtSKU_Field()
    attrdtSKU_Field = pstrattrdtSKU_Field
End Property

Public Property Get attrdtDefault_Field()
    attrdtDefault_Field = pstrattrdtDefault_Field
End Property

Public Property Get attrdtWeight_Field()
    attrdtWeight_Field = pstrattrdtWeight_Field
End Property

'----------------------------------------------------------

Public Property Get rsProducts()
	Set rsProducts = prsProducts
End Property
Public Property Get rsAttributes()
	If isObject(prsAttributes) Then Set rsAttributes = prsAttributes
End Property
Public Property Get rsAttributeDetails()
	Set rsAttributeDetails = prsAttributeDetails
End Property

'added for pricing levels
Public Property Get prodPLPrice()
    prodPLPrice = pstrprodPLPrice
End Property
Public Property Get prodPLSalePrice()
    prodPLSalePrice = pstrprodPLSalePrice
End Property
Public Property Get attrdtPLPrice()
    attrdtPLPrice = pstrattrdtPLPrice
End Property
Public Property Get gwPLPrice
	gwPLPrice = pstrgwPLPrice
End Property

Public Property Get prodSetupFee
	prodSetupFee = pstrprodSetupFee
End Property

Public Property Get prodSetupFeeOneTime
	prodSetupFeeOneTime = pstrprodSetupFeeOneTime
End Property

Public Property Get prodHandlingFee
	prodHandlingFee = pstrprodHandlingFee
End Property

Public Property Get prodMinQty
	prodMinQty = pstrprodMinQty
End Property

Public Property Get prodIncrement
	prodIncrement = pstrprodIncrement
End Property

'----------------------------------------------------------

Public Property Get relatedProducts()
	relatedProducts = pstrRelatedProducts
End Property

'***********************************************************************************************

Private Sub ClearValues()


End Sub 'ClearValues

'***********************************************************************************************

Public Sub LoadProductValues(objRS)

Dim i

	With objRS
		If Not .EOF Then
			plngprodAttrNum = trim(.Fields("prodAttrNum").Value)
			plngprodCategoryId = trim(.Fields("prodCategoryId").Value)
			pblnprodCountryTaxIsActive = ConvertBoolean(.Fields("prodCountryTaxIsActive").Value)
			pdtprodDateAdded = trim(.Fields("prodDateAdded").Value)
			pdtprodDateModified = trim(.Fields("prodDateModified").Value)
			pstrprodDescription = trim(.Fields("prodDescription").Value)
			pblnprodEnabledIsActive = ConvertBoolean(.Fields("prodEnabledIsActive").Value)
			pblnprodLimitQtyToMTP = ConvertBoolean(.Fields("prodLimitQtyToMTP").Value)
			pdblprodHeight = trim(.Fields("prodHeight").Value)
			pstrprodID = trim(.Fields("prodID").Value)
			pstrprodImageLargePath = trim(.Fields("prodImageLargePath").Value)
			pstrprodImageSmallPath = trim(.Fields("prodImageSmallPath").Value)
			pdblprodLength = trim(.Fields("prodLength").Value)
			pstrprodLink = trim(.Fields("prodLink").Value)
			plngprodManufacturerId = trim(.Fields("prodManufacturerId").Value)
			pstrprodMessage = trim(.Fields("prodMessage").Value)
			pstrprodName = trim(.Fields("prodName").Value)
			pstrprodNamePlural = trim(.Fields("prodNamePlural").Value)
			pstrprodPrice = trim(.Fields("prodPrice").Value)
			pblnprodSaleIsActive = ConvertBoolean(.Fields("prodSaleIsActive").Value)
			pstrprodSalePrice = trim(.Fields("prodSalePrice").Value)
			pstrprodShip = trim(.Fields("prodShip").Value)
			pdblprodFixedShippingCharge = trim(.Fields("prodFixedShippingCharge").Value)
			pstrprodSpecialShippingMethods = trim(.Fields("prodSpecialShippingMethods").Value)
			pblnprodShipIsActive = ConvertBoolean(.Fields("prodShipIsActive").Value)
			pstrprodShortDescription = trim(.Fields("prodShortDescription").Value)
			pblnprodStateTaxIsActive = ConvertBoolean(.Fields("prodStateTaxIsActive").Value)
			plngprodVendorId = trim(.Fields("prodVendorId").Value)
			pdblprodWeight = trim(.Fields("prodWeight").Value)
			pdblprodWidth = trim(.Fields("prodWidth").Value)
			pblnBuyersClubIsPercentage = ConvertBoolean(.Fields("BuyersClubIsPercentage").Value)
			pdblBuyersClubPointValue = trim(.Fields("BuyersClubPointValue").Value)
		    
			pstrprodHandlingFee = trim(.Fields("prodHandlingFee").Value)
			pstrprodSetupFee = trim(.Fields("prodSetupFee").Value)
			pstrprodSetupFeeOneTime = trim(.Fields("prodSetupFeeOneTime").Value)

			pstrprodMinQty = trim(.Fields("prodMinQty").Value)
			pstrprodIncrement = trim(.Fields("prodIncrement").Value)

		    For i = 0 To UBound(maryImageFields)
				maryImageFields(i)(2) = getRSFieldValue(objRS, maryImageFields(i)(1))
		    Next 'i
		    
			If pblnPricingLevel Then
				pstrprodPLPrice = getRSFieldValue(objRS, "prodPLPrice")
				pstrprodPLSalePrice = getRSFieldValue(objRS, "prodPLSalePrice")
			End If
			pstrprodFileName = getRSFieldValue(objRS, "prodFileName")
			pstrUpgradeVersion = getRSFieldValue(objRS, "UpgradeVersion")
			pstrPackageCodes = getRSFieldValue(objRS, "packageCodes")
			pstrVersion = getRSFieldValue(objRS, "Version")
			pstrReleaseDate = getRSFieldValue(objRS, "ReleaseDate")
			pdblInstallationHours = getRSFieldValue(objRS, "InstallationHours")
			pblnMyProduct = getRSFieldValue(objRS, "MyProduct")
			pblnInstallationRequired = getRSFieldValue(objRS, "InstallationRequired")
			pblnIncludeInSearch = getRSFieldValue(objRS, "IncludeInSearch")
			pblnIncludeInRandomProduct = getRSFieldValue(objRS, "IncludeInRandomProduct")
			pstrprodMaxDownloads = getRSFieldValue(objRS, "prodMaxDownloads")
			pstrprodDownloadValidFor = getRSFieldValue(objRS, "prodDownloadValidFor")
			pblnprodEnableAlsoBought = getRSFieldValue(objRS, "prodEnableAlsoBought")
			pblnprodEnableReviews = getRSFieldValue(objRS, "prodEnableReviews")
			If cblnAddon_DynamicProductDisplay Then pstrRelatedProducts = getRSFieldValue(objRS, "relatedProducts")
			Call LoadAdditionalImageValues(objRS)

			pstrpageName = getRSFieldValue(objRS, "pageName")
			pstrmetaTitle = getRSFieldValue(objRS, "metaTitle")
			pstrmetaDescription = getRSFieldValue(objRS, "metaDescription")
			pstrmetaKeywords = getRSFieldValue(objRS, "metaKeywords")
		End If
	End With
	
    Call LoadCustomValues(objRS)

End Sub 'LoadProductValues

'***********************************************************************************************

Public Sub LoadIndividualProduct(strProductID)
'this sub added for cases when product added does not appear in summary

Dim pobjRS_Product
Dim pstrSQL

	pstrSQL = "Select * from sfProducts where prodID = '" & SQLSafe(strProductID) & "'"
	Set pobjRS_Product = GetRS(pstrSQL)
	Call LoadProductAttributes(strProductID)
	Call LoadProductValues(pobjRS_Product)
	pobjRS_Product.Close
	Set pobjRS_Product = Nothing

End Sub 'LoadIndividualProduct

'***********************************************************************************************

Private Sub LoadAttributeValues

On Error Resume Next

	If prsAttributes.EOF Then Exit Sub
    plngattrID = trim(prsAttributes.Fields("attrID").Value)
    pstrattrName = trim(prsAttributes.Fields("attrName").Value)
    If Len(pstrattrURL_Field) > 0 Then pstrattrURL = trim(prsAttributes.Fields(pstrattrURL_Field).Value)
    If Len(pstrattrDisplay_Field) > 0 Then pstrattrDisplay = trim(prsAttributes.Fields(pstrattrDisplay_Field).Value)
    If Len(pstrattrExtra_Field) > 0 Then pstrattrExtra = trim(prsAttributes.Fields(pstrattrExtra_Field).Value)
    If Len(pstrattrImage_Field) > 0 Then pstrattrImage = trim(prsAttributes.Fields(pstrattrImage_Field).Value)
    If Len(pstrattrSKU_Field) > 0 Then pstrattrSKU = trim(prsAttributes.Fields(pstrattrSKU_Field).Value)
    pstrattrProdId = trim(prsAttributes.Fields("attrProdId").Value)
	
    If pblnTextBasedAttribute Then
		pbytattrDisplayStyle = trim(prsAttributes.Fields("attrDisplayStyle").Value)
		If Len(pbytattrDisplayStyle) = 0 Then pbytattrDisplayStyle = 0
		If Err.number = 3265 Then
			pbytattrDisplayStyle = False
			Err.Clear
		ElseIf Err.number > 0 Then
			debugprint "Error:" & Err.number,Err.Description
			Err.Clear
		End If
	End If
    If pblnAttributeCategoryOrderable Then
		pbytattrDisplayOrder = trim(prsAttributes.Fields(pstrFieldAttributeCategoryOrderable).Value)
		If Len(pbytattrDisplayOrder) = 0 Then pbytattrDisplayOrder = 0
		If Err.number = 3265 Then
			pblnAttributeCategoryOrderable = False
			Err.Clear
		ElseIf Err.number > 0 Then
			debugprint "Error:" & Err.number,Err.Description
			Err.Clear
		End If
	End If

End Sub 'LoadAttributeValues

Private Sub LoadAttrDetailValues

	If prsAttributeDetails.EOF Then Exit Sub
    plngattrdtID = trim(prsAttributeDetails.Fields("attrdtID").Value)
    plngattrdtAttributeId = trim(prsAttributeDetails.Fields("attrdtAttributeId").Value)
    pstrattrdtName = trim(prsAttributeDetails.Fields("attrdtName").Value)
    pstrattrdtPrice = trim(prsAttributeDetails.Fields("attrdtPrice").Value)
    
    pintattrdtType = trim(prsAttributeDetails.Fields("attrdtType").Value)
    pintattrdtOrder = trim(prsAttributeDetails.Fields("attrdtOrder").Value)

    If Len(pstrattrdtImage_Field) > 0 Then pstrattrdtImage = trim(prsAttributeDetails.Fields(pstrattrdtImage_Field).Value)
    If Len(pstrattrdtURL_Field) > 0 Then pstrattrdtURL = trim(prsAttributeDetails.Fields(pstrattrdtURL_Field).Value)
    If Len(pstrattrdtDisplay_Field) > 0 Then pstrattrdtDisplay = trim(prsAttributeDetails.Fields(pstrattrdtDisplay_Field).Value)
    If Len(pstrattrdtExtra_Field) > 0 Then pstrattrdtExtra = trim(prsAttributeDetails.Fields(pstrattrdtExtra_Field).Value)
    If Len(pstrattrdtExtra1_Field) > 0 Then pstrattrdtExtra1 = trim(prsAttributeDetails.Fields(pstrattrdtExtra1_Field).Value)
    If Len(pstrattrdtFileName_Field) > 0 Then pstrattrdtFileName = trim(prsAttributeDetails.Fields(pstrattrdtFileName_Field).Value)
    If Len(pstrattrdtSKU_Field) > 0 Then pstrattrdtSKU = trim(prsAttributeDetails.Fields(pstrattrdtSKU_Field).Value)
    If Len(pstrattrdtDefault_Field) > 0 Then pbytattrdtDefault = trim(prsAttributeDetails.Fields(pstrattrdtDefault_Field).Value)
    If Len(pstrattrdtWeight_Field) > 0 Then pdblattrdtWeight = trim(prsAttributeDetails.Fields(pstrattrdtWeight_Field).Value)

    If pblnPricingLevel Then
		pstrattrdtPLPrice = trim(prsAttributeDetails("attrdtPLPrice").Value)
	End If

End Sub 'LoadAttrDetailValues

Private Sub LoadAdditionalImageValues(byRef objRS)

Dim i
Dim paryTemp
Dim paryTempImage

	plngNumAdditionalImages = 0
	pstrprodAdditionalImages = trim(objRS.Fields("prodAdditionalImages").Value)
	If Len(pstrprodAdditionalImages) > 0 Then
		paryTemp = Split(pstrprodAdditionalImages, "|")
		plngNumAdditionalImages = UBound(paryTemp)
		ReDim paryAdditionalImageText(plngNumAdditionalImages)
		ReDim paryAdditionalImage(plngNumAdditionalImages)
		ReDim paryAdditionalImageDesc(plngNumAdditionalImages)
		For i = 0 To plngNumAdditionalImages
			paryTempImage = Split(paryTemp(i), ";")
			paryAdditionalImageText(i) = paryTempImage(0)
			paryAdditionalImage(i) = paryTempImage(1)
			paryAdditionalImageDesc(i) = paryTempImage(2)
		Next 'i
		plngNumAdditionalImages = plngNumAdditionalImages + 1
	End If

End Sub 'LoadAdditionalImageValues

Private Sub LoadAdditionalImagesFromRequest

Dim i
Dim pblnDone
Dim pstradditionalImageText
Dim pstradditionalImage
Dim pstradditionalImageDesc
Dim pstrTemp

	i = -1
	pblnDone = False
	
	Do While Not pblnDone
		i = i + 1
		pstradditionalImageText = Trim(Request.Form("additionalImageText" & i))
		pstradditionalImage = Trim(Request.Form("additionalImage" & i))
		pstradditionalImageDesc = Trim(Request.Form("additionalImageDesc" & i))

		pblnDone = Len(pstradditionalImageText & pstradditionalImage & pstradditionalImageDesc) = 0
		If Not pblnDone Then
			pstrTemp = pstradditionalImageText & ";" & pstradditionalImage & ";" & pstradditionalImageDesc
			If Len(pstrprodAdditionalImages) = 0 Then
				pstrprodAdditionalImages = pstrTemp
			Else
				pstrprodAdditionalImages = pstrprodAdditionalImages & "|" & pstrTemp
			End If
		End If
		
	Loop
	
End Sub

Private Sub LoadFromRequest

Dim i

On Error Goto 0

    With Request.Form
        plngprodAttrNum = Trim(.Item("prodAttrNum"))
        plngprodCategoryId = Trim(.Item("prodCategoryId"))
        pblnprodCountryTaxIsActive = (lCase(.Item("prodCountryTaxIsActive")) = "on")
        pstrprodDescription = Trim(.Item("prodDescription"))
        pblnprodEnabledIsActive = (lCase(.Item("prodEnabledIsActive")) = "on")
        pblnprodLimitQtyToMTP = (lCase(.Item("prodLimitQtyToMTP")) = "on")
        pstrprodFileName = Trim(.Item("prodFileName"))
        pstrUpgradeVersion = Trim(.Item("UpgradeVersion"))
        pstrPackageCodes = Trim(.Item("packageCodes"))
		pstrVersion = Trim(.Item("Version"))
		pstrReleaseDate = Trim(.Item("ReleaseDate"))
		pdblInstallationHours = Trim(.Item("InstallationHours"))
		pblnMyProduct = Trim(.Item("MyProduct"))
		pblnInstallationRequired = Trim(.Item("InstallationRequired"))
		pblnIncludeInSearch = Trim(.Item("IncludeInSearch"))
		pblnIncludeInRandomProduct = Trim(.Item("IncludeInRandomProduct"))
        pstrprodMaxDownloads = Trim(.Item("prodMaxDownloads"))
        pstrprodDownloadValidFor = Trim(.Item("prodDownloadValidFor"))
        pdblprodHeight = Trim(.Item("prodHeight"))
        pstrprodID = Trim(.Item("prodID"))
        pstrprodImageLargePath = Trim(.Item("prodImageLargePath"))
        pstrprodImageSmallPath = Trim(.Item("prodImageSmallPath"))
        pdblprodLength = Trim(.Item("prodLength"))
        pstrprodLink = Trim(.Item("prodLink"))
        plngprodManufacturerId = Trim(.Item("prodManufacturerId"))
        pstrprodMessage = Trim(.Item("prodMessage"))
        pstrprodName = Trim(.Item("prodName"))
        pstrprodNamePlural = Trim(.Item("prodNamePlural"))
        pstrprodPrice = Trim(.Item("prodPrice"))
        pblnprodSaleIsActive = (lCase(.Item("prodSaleIsActive")) = "on")
        pstrprodSalePrice = Trim(.Item("prodSalePrice"))
        pstrprodShip = Trim(.Item("prodShip"))
        pdblprodFixedShippingCharge = Trim(.Item("prodFixedShippingCharge"))
        pstrprodSpecialShippingMethods = Trim(.Item("prodSpecialShippingMethods"))
        pblnprodShipIsActive = (lCase(.Item("prodShipIsActive")) = "on")
        pblnBuyersClubIsPercentage = (lCase(.Item("BuyersClubIsPercentage")) = "on")
        pdblBuyersClubPointValue = Trim(.Item("BuyersClubPointValue"))

        pbytprodDisplayAdditionalImagesInWindow = Trim(.Item("prodDisplayAdditionalImagesInWindow"))
        If Len(prodDisplayAdditionalImagesInWindow) = 0 Then prodDisplayAdditionalImagesInWindow = 0
        
        pblnprodEnableAlsoBought = (lCase(.Item("prodEnableAlsoBought")) = "on")
        pblnprodEnableReviews = (lCase(.Item("prodEnableReviews")) = "on")

        pstrpageName = Trim(.Item("pageName"))
        pstrmetaTitle = Trim(.Item("metaTitle"))
        pstrmetaDescription = Trim(.Item("metaDescription"))
        pstrmetaKeywords = Trim(.Item("metaKeywords"))

        pstrprodHandlingFee = Trim(.Item("prodHandlingFee"))
        pstrprodSetupFee = Trim(.Item("prodSetupFee"))
        pstrprodSetupFeeOneTime = Trim(.Item("prodSetupFeeOneTime"))
		If Len(pstrprodHandlingFee) = 0 Then pstrprodHandlingFee = 0
		If Len(pstrprodSetupFee) = 0 Then pstrprodSetupFee = 0
		If Len(pstrprodSetupFeeOneTime) = 0 Then pstrprodSetupFeeOneTime = 0
		
        pstrprodMinQty = Trim(.Item("prodMinQty"))
        pstrprodIncrement = Trim(.Item("prodIncrement"))
		If Len(pstrprodMinQty) = 0 Then pstrprodMinQty = 0
		If Len(pstrprodIncrement) = 0 Then pstrprodIncrement = 0
		
		For i = 0 To UBound(maryImageFields)
			maryImageFields(i)(2) = Trim(.Item(maryImageFields(i)(1)))
		Next 'i

        pstrprodShortDescription = Trim(.Item("prodShortDescription"))
        If Len(mlngShortDescriptionLength) > 0 Then
			If len(pstrprodShortDescription) > CLng(mlngShortDescriptionLength) Then pstrprodShortDescription = Left(pstrprodShortDescription, CLng(mlngShortDescriptionLength))
		End If
        
        pblnprodStateTaxIsActive = (lCase(.Item("prodStateTaxIsActive")) = "on")
        plngprodVendorId = Trim(.Item("prodVendorId"))
        pdblprodWeight = Trim(.Item("prodWeight"))
        pdblprodWidth = Trim(.Item("prodWidth"))
        pdtprodDateAdded = Trim(.Item("prodDateAdded"))
        pdtprodDateModified = Trim(.Item("prodDateModified"))

        plngattrID = Trim(.Item("attrID"))
        pstrattrName = Trim(.Item("attrName"))
        pstrattrURL = Trim(.Item("attrURL"))
        pstrattrDisplay = Trim(.Item("attrDisplay"))
        pstrattrExtra = Trim(.Item("attrExtra"))
        pstrattrImage = Trim(.Item("attrImage"))
        pstrattrSKU = Trim(.Item("attrSKU"))
        pstrattrProdId = Trim(.Item("attrProdId"))
		pbytattrDisplayStyle = trim(.Item("attrDisplayStyle"))
		pbytattrDisplayOrder = trim(.Item(pstrFieldAttributeCategoryOrderable))

        plngattrdtAttributeId = Trim(.Item("attrdtAttributeId"))
        plngattrdtID = Trim(.Item("attrdtID"))
        pstrattrdtName = Trim(.Item("attrdtName"))
        pintattrdtOrder = Trim(.Item("attrdtOrder"))
        pstrattrdtPrice = Trim(.Item("attrdtPrice"))
        pintattrdtType = Trim(.Item("attrdtType"))
        pstrattrdtImage = Trim(.Item("attrdtImage"))
        pdblattrdtWeight = Trim(.Item("attrdtWeight"))
        pstrattrdtURL = Trim(.Item("attrdtURL"))
        pstrattrdtDisplay = Trim(.Item("attrdtDisplay"))
        pstrattrdtExtra = Trim(.Item("attrdtExtra"))
        pstrattrdtExtra1 = Trim(.Item("attrdtExtra1"))
        pstrattrdtFileName = Trim(.Item("attrdtFileName"))
        pstrattrdtSKU = Trim(.Item("attrdtSKU"))
        pbytattrdtDefault = Trim(.Item("attrdtDefault"))
		pblnAEInventoryChanged = ConvertBoolean(.Item("ChangeInventory"))
		pblnAEMTPChanged = ConvertBoolean(.Item("ChangeMTP"))
		pblnAECategoryChanged = ConvertBoolean(.Item("ChangeCategory"))
		
		If pblnPricingLevel Then
			pstrprodPLPrice = splitPrices(.Item("prodPLPrice"))
			pstrprodPLSalePrice = splitPrices(.Item("prodPLSalePrice"))
			pstrattrdtPLPrice = splitPrices(.Item("attrdtPLPrice"))
		End If

        If cblnAddon_DynamicProductDisplay Then pstrRelatedProducts = Replace(Trim(.Item("relatedProducts")), ", ", ";")
        Call LoadAdditionalImagesFromRequest
		Call LoadCustomValuesFromRequest
		
		Dim pstrTemp
		pstrTemp = Trim(.Item("ManufacturerNew"))
		If Len(pstrTemp) > 0 Then
			plngprodManufacturerId = getManufacturerByName(pstrTemp, True)
			If plngprodManufacturerId = -1 Then
				plngprodManufacturerId = 1	'Reset to No Manufacturer if an error resulted
			Else
				Call addMessageItem("Manufacturer <em>" & pstrTemp & "</em> created.", False)
				Call resetCombo_Saved("manufacturer")
			End If
		End If

		pstrTemp = Trim(.Item("VendorNew"))
		If Len(pstrTemp) > 0 Then
			plngprodVendorId = getVendorByName(pstrTemp, True)
			If plngprodVendorId = -1 Then
				plngprodVendorId = 1	'Reset to No Vendor if an error resulted
			Else
				Call addMessageItem("Vendor <em>" & pstrTemp & "</em> created.", False)
				Call resetCombo_Saved("vendor")
			End If
		End If

    End With

End Sub 'LoadFromRequest

'added because of Euros
Function isNumber(strQuestion)

	If isNumeric(strQuestion) And Len(Trim(strQuestion)) > 0 Then
		isNumber = True
	Else
		isNumber = False
	End If

End Function	'isNumber

Function splitPrices(strPrice)

Dim pstrTemp
Dim i
Dim plngLength
Dim pstrChar
Dim pstrPrevChar
Dim pstrNextChar
Dim pblnIsNumber

	pstrTemp = (Trim(strPrice))
	'pstrTemp = Replace(pstrTemp, " ", "")
	If Len(pstrTemp) = 0 Then Exit Function

	plngLength = Len(pstrTemp)
	
	pstrPrevChar = Mid(pstrTemp,1,1)
	For i = 2 To (plngLength - 1)
		pstrChar = Mid(pstrTemp,i,1)
		pstrNextChar = Mid(pstrTemp, i+1, 1)
		If isNumber(pstrPrevChar) And isNumber(pstrNextChar) And pstrChar="," Then
			pstrTemp = Left(pstrTemp,i-1) & "|" & Right(pstrTemp,plngLength - i)
		End If

		pstrPrevChar = pstrChar
	Next 'i
	
	pstrTemp = Replace(pstrTemp, ",", ";")
	pstrTemp = Replace(pstrTemp, "|", ",")

	splitPrices = pstrTemp

End Function	'splitPrices

'***********************************************************************************************

Public Function FindProduct(strprodID)

'On Error Resume Next

    With prsProducts
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(strprodID) <> 0 Then
                .Find "prodID='" & SQLSafe(strprodID) & "'"
            Else
                .MoveLast
            End If
            If .EOF Then 
				Call LoadIndividualProduct(strprodID)
            Else
				Call LoadProductValues(prsProducts)
				FindProduct = True
			End If
        Else
			FindProduct = False
        End If
    End With

End Function    'FindProduct

Public Function FindAttribute(lngID)

'On Error Resume Next

	If Not isObject(prsAttributes) Then Exit Function

    With prsAttributes
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "attrID=" & lngID
            Else
                .MoveLast
            End If
            If Not .EOF Then 
				FindProduct prsAttributes("attrProdId")
				Call LoadAttributeValues
			End If
        End If
    End With

End Function    'FindsubProduct

Public Function FindAttrDetail(lngID)

'On Error Resume Next

	If Not isObject(prsAttributeDetails) Then Exit Function

	If isObject(prsAttributeDetails) Then
		With prsAttributeDetails
			If .RecordCount > 0 Then
				.MoveFirst
				If Len(lngID) <> 0 Then
					.Find "attrdtID=" & lngID
				Else
					.MoveLast
				End If
				If Not .EOF Then 
					FindAttribute prsAttributeDetails("attrdtAttributeId")
					Call LoadAttrDetailValues
				End If
			End If
		End With
	End If

End Function    'FindsubProduct

'***********************************************************************************************

Public Property Get gwActivate
	gwActivate = pblngwActivate
End Property
Public Property Get gwPrice
	gwPrice = pstrgwPrice
End Property

Public Function LoadAE()

Dim prs

	If len(pstrprodID) > 0 Then
		Set prsInventory = GetRS("Select * from sfInventory Where invenProdId='" & SQLSafe(pstrprodID) & "'")
		Set prsInventoryInfo = GetRS("Select * from sfInventoryInfo Where invenProdID='" & SQLSafe(pstrprodID) & "'")
		Set prsMTP = GetRS("Select * from sfMTPrices Where mtProdID='" & SQLSafe(pstrprodID) & "'")
		Set prsCategories = GetRS("Select * from sfSub_Categories")

		Set prs = GetRS("Select * from sfGiftWraps Where gwProdID='" & SQLSafe(pstrprodID) & "'")
		If Not prs.EOF Then
			pblngwActivate = ConvertBoolean(prs.Fields("gwActivate").value)
			pstrgwPrice = prs.Fields("gwPrice").value
		End If
		prs.Close
		Set prs = Nothing
		
	End If
	Call AEProductCategory

End Function	'LoadAE

'***********************************************************************************************

Public Function UpdateAE()

Dim rs
Dim sql
Dim p_blnTrackInventory

	If len(pstrprodID) > 0 Then

        Set rs = server.CreateObject("adodb.Recordset")
		sql = "Select * from sfInventoryInfo Where invenProdID='" & SQLSafe(pstrprodID) & "'"
		With rs
			.open sql, cnn, 1, 3
			If .EOF Then 
				.AddNew
				.Fields("invenProdId").Value = pstrprodID
			End If

			.Fields("invenbBackOrder").Value = ConvertBoolean(lcase(Request.Form("invenbBackOrder"))="on")*-1
			p_blnTrackInventory = ConvertBoolean(lcase(Request.Form("invenbTracked"))="on")
			If (.Fields("invenbTracked").Value = 0) And p_blnTrackInventory Then pblnAEInventoryChanged = True

			.Fields("invenbTracked").Value = p_blnTrackInventory*-1
			.Fields("invenbStatus").Value = ConvertBoolean(lcase(Request.Form("invenbStatus"))="on")*-1
			.Fields("invenbNotify").Value = ConvertBoolean(lcase(Request.Form("invenbNotify"))="on")*-1
			If Len(Trim(Request.Form("invenInStockDEF"))) > 0 Then 
				.Fields("invenInStockDEF").Value = Request.Form("invenInStockDEF")
			Else
				.Fields("invenInStockDEF").Value = 0
			End If
			If Len(Trim(Request.Form("invenLowFlagDEF"))) > 0 Then 
				.Fields("invenLowFlagDEF").Value = Request.Form("invenLowFlagDEF")
			Else
				.Fields("invenLowFlagDEF").Value = 0
			End If

			.Update
			.Close
		End With
		Set rs = Nothing

        Set rs = server.CreateObject("adodb.Recordset")
		sql = "Select * from sfGiftWraps Where gwProdID='" & SQLSafe(pstrprodID) & "'"
		With rs
			.open sql, cnn, 1, 3
			If .EOF Then 
				.AddNew
				.Fields("gwProdID").Value = pstrprodID
			End If

			.Fields("gwActivate").Value = ConvertBoolean(lcase(Request.Form("gwActivate"))="on")*-1
			If Len(Trim(Request.Form("gwPrice"))) > 0 Then 
				.Fields("gwPrice").Value = Request.Form("gwPrice")
			Else
				.Fields("gwPrice").Value = 0
			End If

			.Update
			.Close
		End With
		Set rs = Nothing
		
		'Get MTP
		If pblnAEMTPChanged Then
			sql = "Delete From sfMTPrices Where mtProdID='" & SQLSafe(pstrprodID) & "'"
			cnn.Execute sql,,128
		
			Dim pstrmtQuantity, amtQuantity
			Dim pstrmtValue, amtValue
			Dim pstrmtType, amtType
			Dim i,j
			Dim paryTemp
			Dim plngCount
			Dim paryMTP
		

			If pblnPricingLevel Then
				ReDim paryTemp(3)
			Else
				ReDim paryTemp(2)
			End If

			pstrmtQuantity = Request.Form("mtQuantity")
			pstrmtValue = Request.Form("mtValue")
			pstrmtType = Request.Form("mtType")
			'debugprint "pstrmtQuantity",pstrmtQuantity
			'debugprint "pstrmtValue",pstrmtValue
			'debugprint "pstrmtType",pstrmtType

			If len(pstrmtQuantity) > 0 Then
				amtQuantity = Split(pstrmtQuantity,",")
				amtValue = Split(pstrmtValue,",")
				amtType = Split(pstrmtType,",")
				
				plngCount = ubound(amtQuantity)
				If pblnPricingLevel Then
					ReDim paryMTP(plngCount,3)
				Else
					ReDim paryMTP(plngCount,2)
				End If
				For i=0 to plngCount
					paryMTP(i,0) = cLng(Trim(amtQuantity(i)))
					paryMTP(i,1) = cDbl(Trim(amtValue(i)))
					paryMTP(i,2) = Trim(amtType(i))
					If pblnPricingLevel Then paryMTP(i,3) = Replace(Replace(Trim(Request.Form("mtPLValue" & i + 1)),",",";")," ","")
				Next
		
				For j=0 to plngCount
					For i=1 to plngCount
						If paryMTP(i-1,0) > paryMTP(i,0) Then
							paryTemp(0) = paryMTP(i,0)
							paryTemp(1) = paryMTP(i,1)
							paryTemp(2) = paryMTP(i,2)
							If pblnPricingLevel Then paryTemp(3) = paryMTP(i,3)
							
							paryMTP(i,0) = paryMTP(i-1,0)
							paryMTP(i,1) = paryMTP(i-1,1)
							paryMTP(i,2) = paryMTP(i-1,2)
							If pblnPricingLevel Then paryMTP(i,3) = paryMTP(i-1,3)

							paryMTP(i-1,0) = paryTemp(0)
							paryMTP(i-1,1) = paryTemp(1)
							paryMTP(i-1,2) = paryTemp(2)
							If pblnPricingLevel Then paryMTP(i-1,3) = paryTemp(3)
						End If
					Next
				Next
		
				for i=0 to plngCount
					If pblnPricingLevel Then
						If Len(paryMTP(i,3)) > 0 Then 
							sql = "Insert Into sfMTPrices (mtProdID,mtIndex,mtQuantity,mtValue,mtType,mtPLValue) Values " _
								& "('" & SQLSafe(pstrprodID) & "'," & i & "," & paryMTP(i,0) & "," & paryMTP(i,1) & ",'" & paryMTP(i,2) & "','" & paryMTP(i,3) & "')"
						Else
							sql = "Insert Into sfMTPrices (mtProdID,mtIndex,mtQuantity,mtValue,mtType) Values " _
								& "('" & SQLSafe(pstrprodID) & "'," & i & "," & paryMTP(i,0) & "," & paryMTP(i,1) & ",'" & paryMTP(i,2) & "')"
						End If
					Else
						sql = "Insert Into sfMTPrices (mtProdID,mtIndex,mtQuantity,mtValue,mtType) Values " _
							& "('" & SQLSafe(pstrprodID) & "'," & i & "," & paryMTP(i,0) & "," & paryMTP(i,1) & ",'" & paryMTP(i,2) & "')"
					End If
					cnn.Execute sql,,128
				next
			End If
		End If		

		If Not p_blnTrackInventory Then
			DeleteInventoryFields(pstrprodID)
			pblnAEInventoryChanged = False
		End If

		If pblnAEInventoryChanged Then
		
			Dim pstrinvenId, paryinvenId
			Dim pstrinvenInStock, paryinvenInStock
			Dim pstrinvenLowFlag, paryinvenLowFlag
		
			pstrinvenId = Request.Form("invenId")
			pstrinvenInStock = Request.Form("invenInStock")
			pstrinvenLowFlag = Request.Form("invenLowFlag")
			
			If len(pstrinvenId) > 0 Then
				paryinvenId = Split(pstrinvenId,",")
				paryinvenInStock = Split(pstrinvenInStock,",")
				paryinvenLowFlag = Split(pstrinvenLowFlag,",")
				
				plngCount = ubound(paryinvenId)
				for i=0 to plngCount
				sql = "Update sfInventory Set invenInStock=" & paryinvenInStock(i) & ", invenLowFlag=" & paryinvenLowFlag(i) & " Where invenId=" & SQLSafe(paryinvenId(i))
				cnn.Execute sql,,128
				next

			End If		

			Call UpdateInventoryFields
		End If
		
		If pblnAECategoryChanged Then Call UpdateProductsInCategory
		
	End If

End Function	'UpdateAE

Public Function DeleteInventoryFields(strprodID)
	If len(strprodID) > 0 Then cnn.Execute "Delete from sfInventory Where invenProdId='" & SQLSafe(strprodID) & "'",,128
End Function

Public Function UpdateInventoryFields

Dim p_rsAttributeCategories
Dim p_rsAttributes
Dim p_rsInventory
Dim sql
Dim i,j,k
Dim plngPos
Dim plngCount
Dim paryAttr()
Dim paryAttCats()

	If Not cblnSF5AE Then Exit Function
	If Not ConvertBoolean(lcase(Request.Form("invenbTracked"))="on") Then
		DeleteInventoryFields(pstrprodID)
		Exit Function
	End If

	Set p_rsAttributeCategories = GetRS("Select attrID from sfAttributes Where attrProdId='" & SQLSafe(pstrprodID) & "'")
	If Not p_rsAttributeCategories.EOF Then
		ReDim paryAttCats(p_rsAttributeCategories.RecordCount,3)
		sql = "SELECT sfAttributeDetail.attrdtID, sfAttributeDetail.attrdtAttributeId, sfAttributeDetail.attrdtName" _
			& " FROM sfAttributes LEFT JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId" _
			& " WHERE sfAttributes.attrProdId='" & SQLSafe(pstrprodID) & "'" _
			& " ORDER BY sfAttributeDetail.attrdtID"
		Set p_rsAttributes = GetRS(sql)
		plngCount = 1
		For i=1 to p_rsAttributeCategories.RecordCount
			paryAttCats(i,0) = p_rsAttributeCategories.Fields("attrID").Value
			p_rsAttributes.Filter = "attrdtAttributeId='" & paryAttCats(i,0) & "'"
			paryAttCats(i,1) = p_rsAttributes.RecordCount
			plngCount =  plngCount * p_rsAttributes.RecordCount
			paryAttCats(i,2) = plngCount
			p_rsAttributeCategories.MoveNext
		Next
		paryAttCats(p_rsAttributeCategories.RecordCount,3) = 1
		For i=p_rsAttributeCategories.RecordCount-1 to 1 step -1
		paryAttCats(i,3) = paryAttCats(i+1,3) * paryAttCats(i+1,1)
		Next

		ReDim paryAttr(plngCount,1)
		For j=1 to ubound(paryAttCats)
			p_rsAttributes.Filter = "attrdtAttributeId='" & paryAttCats(j,0) & "'"
			If p_rsAttributes.RecordCount > 0 Then
				p_rsAttributes.MoveFirst
				k=1
				For i=1 to plngCount
					If len(paryAttr(i,0)) = 0 Then
						paryAttr(i,0) = p_rsAttributes.Fields("attrdtID").Value
					Else
						paryAttr(i,0) = paryAttr(i,0) & "," & p_rsAttributes.Fields("attrdtID").Value
					End If
					If len(paryAttr(i,1)) = 0 Then
						paryAttr(i,1) = p_rsAttributes.Fields("attrdtName").Value
					Else
						paryAttr(i,1) = paryAttr(i,1) & " " & p_rsAttributes.Fields("attrdtName").Value
					End If
					If p_rsAttributes.EOF Then p_rsAttributes.MoveFirst
					If k=paryAttCats(j,3) Then
						p_rsAttributes.MoveNext
						If p_rsAttributes.EOF Then p_rsAttributes.MoveFirst
						k=0
					End If
					k=k+1
				Next
			End If
		Next 'j
		
		'For i=1 to plngCount
		'debugprint i & ": " & paryAttr(i,0),paryAttr(i,1)
		'Next

		'Clean out extra rows
		Dim plnginvenInStockDEF, plnginvenLowFlagDEF
		Set p_rsInventory = GetRS("Select invenInStockDEF, invenLowFlagDEF from sfInventoryInfo Where invenProdId='" & SQLSafe(pstrprodID) & "'")
		
		If Not p_rsInventory.EOF Then
		
			plnginvenInStockDEF = p_rsInventory.Fields("invenInStockDEF").Value
			plnginvenLowFlagDEF = p_rsInventory.Fields("invenLowFlagDEF").Value
			Set p_rsInventory = GetRS("Select invenId,invenAttDetailID from sfInventory Where invenProdId='" & SQLSafe(pstrprodID) & "'")

			Dim pobjDic
			set pobjDic = Server.CreateObject("Scripting.Dictionary")
			For i=1 to plngCount
				'debugprint paryAttr(i,0),paryAttr(i,1)
				pobjDic.Add "ss|" & Trim(paryAttr(i,0)),i
				p_rsInventory.Filter = "invenAttDetailID='" & paryAttr(i,0) & "'"
				If p_rsInventory.EOF Then
					sql = "Insert Into sfInventory (invenProdId,invenAttDetailID,invenAttName,invenInStock,invenLowFlag)" _
						& " Values ('" & SQLSafe(pstrprodID) & "','" & paryAttr(i,0) & "','" & SQLSafe(paryAttr(i,1)) & "'," & plnginvenInStockDEF & "," & plnginvenLowFlagDEF & ")"
					cnn.Execute sql,,128
				Else
					sql = "Update sfInventory Set invenAttName='" & SQLSafe(paryAttr(i,1)) & "'" _
						& " Where invenProdId='" & SQLSafe(pstrprodID) & "' AND invenAttDetailID='" & SQLSafe(paryAttr(i,0)) & "'"
					cnn.Execute sql,,128
				End If
			Next
			p_rsInventory.Filter = ""
			
			Do While Not p_rsInventory.EOF
				If Not pobjDic.Exists("ss|" & Trim(p_rsInventory.Fields("invenAttDetailID").Value)) Then
					sql = "Delete from sfInventory Where invenId=" & p_rsInventory.Fields("invenId").Value
					cnn.Execute sql,,128
				End If
				p_rsInventory.MoveNext
			Loop
		End If
	Else	'No attributes
	
		Set p_rsInventory = GetRS("Select invenInStockDEF, invenLowFlagDEF from sfInventoryInfo Where invenProdId='" & SQLSafe(pstrprodID) & "'")
		If Not p_rsInventory.EOF Then
			plnginvenInStockDEF = p_rsInventory.Fields("invenInStockDEF").Value
			plnginvenLowFlagDEF = p_rsInventory.Fields("invenLowFlagDEF").Value
			Set p_rsInventory = GetRS("Select invenId,invenAttDetailID from sfInventory Where invenProdId='" & SQLSafe(pstrprodID) & "'")
			If p_rsInventory.EOF Then
				sql = "Insert Into sfInventory (invenProdId,invenAttDetailID,invenAttName,invenInStock,invenLowFlag)" _
					& " Values ('" & SQLSafe(pstrprodID) & "','0',''," & plnginvenInStockDEF & "," & plnginvenLowFlagDEF & ")"
				cnn.Execute sql,,128
			End If
		End If

	End If
	
	'Now issue the updates
	Dim paryinvenId
	Dim paryinvenInStock
	Dim paryinvenLowFlag
	
	paryinvenId = Split(Request.Form("invenId"),",")
	paryinvenInStock = Split(Request.Form("invenInStock"),",")
	paryinvenLowFlag = Split(Request.Form("invenLowFlag"),",")

	For i=0 to uBound(paryinvenId)
		sql = "Update sfInventory Set invenInStock=" & paryinvenInStock(i) & ", invenLowFlag=" & paryinvenLowFlag(i) & " Where invenId=" & paryinvenId(i)
		cnn.Execute sql,,128
	Next
	
	Set pobjDic = Nothing
	ReleaseObject(p_rsInventory)
	ReleaseObject(p_rsAttributes)

End Function	'UpdateInventoryFields


Private Function UpdateProductsInCategory

Dim plngPos
Dim prsAEProductCategories
Dim prsCatTest
Dim pstrCategories
Dim paryCategories
Dim pstrCats
Dim i
Dim sql

	pstrCategories = Trim(Request.Form("Categories"))
	paryCategories = Split(pstrCategories,",")
	pstrCats = "|" & Replace(pstrCategories,",","|") & "|"
	pstrCats = Replace(pstrCats," ","")
	Set prsAEProductCategories = GetRS("Select subcatDetailID, subcatCategoryId from sfSubCatDetail Where ProdID='" & SQLSafe(pstrprodID) & "'")

	'add the new categories
	For i=0 to uBound(paryCategories)
		plngPos = Instr(1,paryCategories(i),"none-")
		If plngPos > 0 Then
			Dim prsCatInfo
			Dim prsNewSubCat
			Dim p_lngCatID
			
			p_lngCatID = Right(paryCategories(i),Len(paryCategories(i))-plngPos-4)
			Set prsCatInfo = GetRS("Select catName from sfCategories Where catID=" & p_lngCatID)
			Set prsNewSubCat = server.CreateObject("adodb.Recordset")
			prsNewSubCat.CursorLocation = 3
			prsNewSubCat.CacheSize = 1
			prsNewSubCat.open "Select * from sfSub_Categories", cnn, 1, 3
		    prsNewSubCat.AddNew
			prsNewSubCat.Fields("subcatCategoryId").Value = p_lngCatID
			prsNewSubCat.Fields("subcatName").Value = prsCatInfo.Fields("catName").value
			prsNewSubCat.Fields("subcatIsActive").Value = 1
			prsNewSubCat.Fields("Depth").Value = 0
			prsNewSubCat.Fields("HasProds").Value = 1
			prsNewSubCat.Fields("bottom").Value = 1
				
			prsNewSubCat.Update
			paryCategories(i) = prsNewSubCat.Fields("subcatID").Value
			prsNewSubCat.Fields("CatHierarchy").Value = "none-" & paryCategories(i)
			prsNewSubCat.Update

			prsNewSubCat.Close
			Set prsNewSubCat = Nothing		

			prsCatInfo.Close
			Set prsCatInfo = Nothing		
		
		End If

		prsAEProductCategories.Filter = "subcatCategoryId=" & paryCategories(i)
		If prsAEProductCategories.EOF Then
			sql = "Insert Into sfSubCatDetail (subcatCategoryId,ProdID,ProdName) Values (" & paryCategories(i) & ",'" & SQLSafe(pstrprodID) & "','" & SQLSafe(pstrProdName) & "')"
			cnn.Execute sql,,128
			sql = "Update sfSub_Categories Set HasProds=1 Where subcatCategoryId=" & paryCategories(i)
			cnn.Execute sql,,128
		End If
	Next
	
	'delete the extra
	prsAEProductCategories.Filter = ""
	For i=1 to prsAEProductCategories.RecordCount
		If Instr(1,pstrCats,"|" & Trim(prsAEProductCategories.Fields("subcatCategoryId").Value) & "|") = 0 Then
			sql = "Delete from sfSubCatDetail Where subcatCategoryId=" & prsAEProductCategories.Fields("subcatCategoryId").Value & " AND ProdID='" & SQLSafe(pstrprodID) & "'"
			cnn.Execute sql,,128
			
			'check to see if this was the last product in the sub-category
			sql = "Select subcatDetailID from sfSubCatDetail where subcatCategoryId=" & prsAEProductCategories.Fields("subcatCategoryId").Value
			Set prsCatTest = GetRS(sql)
			If prsCatTest.EOF Then
				sql = "Update sfSub_Categories Set HasProds=0 Where subcatCategoryId=" & prsAEProductCategories.Fields("subcatCategoryId").Value
				cnn.Execute sql,,128
			End If
		End If
		prsAEProductCategories.MoveNext
	Next
	
	ReleaseObject(prsCatTest)
	ReleaseObject(prsAEProductCategories)

End Function	'UpdateProductsInCategory

Public Property Get rsInventory()
	If isObject(prsInventory) Then Set rsInventory = prsInventory
End Property
Public Property Get rsInventoryInfo()
    Set rsInventoryInfo = prsInventoryInfo
End Property
Public Property Get rsMTP()
	If isObject(prsMTP) Then Set rsMTP = prsMTP
End Property
Public Property Get Categories()
    Set Categories = pCategories
End Property

Public Property Get CategoryFilterList()
    CategoryFilterList = pstrCategoryFilterList
End Property

Public Property Get CategoryList()
    CategoryList = pstrCategoryList
End Property
Public Property Get ProdDictionaryList()
    ProdDictionaryList = pstrProdDictionaryList
End Property
Public Property Get CategoryArray()
    CategoryArray = pstrCategoryArray
End Property

Private Sub AEProductCategory

Dim paryCategories_All
Dim prsAEProductCategories
Dim i,j,k
Dim paryCategories()
Dim plngCatCount
Dim pblnBottom
Dim pdicCategories
Dim pstrTemp
Dim pstrCatValue
Dim pstrCatName
Dim pbytCatDepth
Dim pstrsubCatValue
Dim pstrsubCatName
Dim pstrSelected
Dim pstrSQL

	pstrSQL = "SELECT sfCategories.catID, sfSub_Categories.subcatID, sfCategories.catName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName, sfSub_Categories.Depth, sfSub_Categories.bottom" _
			& " FROM sfSub_Categories RIGHT JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
			& " ORDER BY sfCategories.catName, sfSub_Categories.Depth, sfSub_Categories.subcatName"

	Set prsAEProductCategories = GetRS(pstrSQL)
	With prsAEProductCategories
		paryCategories_All = .getRows()

		'Determine how many leaf nodes
		.Filter = "bottom=1 or bottom=Null"
		plngCatCount = .RecordCount - 1

		k = 0
		If plngCatCount > 0 Then
			redim paryCategories(plngCatCount,7)
			Set pdicCategories = Server.CreateObject("SCRIPTING.DICTIONARY")
			For i = 0 to UBound(paryCategories_All, 2)
			
				'do a little cleanup
				paryCategories_All(1, i) = Trim(paryCategories_All(1, i) & "")
				paryCategories_All(2, i) = Trim(paryCategories_All(2, i) & "")
				paryCategories_All(3, i) = Trim(paryCategories_All(3, i) & "")
				paryCategories_All(4, i) = Trim(paryCategories_All(4, i) & "")
				
				If Len(paryCategories_All(1, i)) > 0 Then pdicCategories.Add CStr(paryCategories_All(1, i)), paryCategories_All(4, i)
				pblnBottom = paryCategories_All(6, i) = 1 OR isNull(paryCategories_All(6, i))
'				pblnBottom = .Fields("bottom").value = 1
				If pblnBottom Then
					paryCategories(k,0) = paryCategories_All(0, i)	'.Fields("catID").value
					paryCategories(k,1) = paryCategories_All(1, i)	'.Fields("subcatID").value
					paryCategories(k,2) = paryCategories_All(2, i)	'.Fields("catName").value
					paryCategories(k,3) = paryCategories_All(4, i)	'.Fields("subcatName").value
					paryCategories(k,4) = Trim(paryCategories_All(3, i) & "")	'Trim(.Fields("CatHierarchy").value & "")
					paryCategories(k,5) = paryCategories_All(5, i)	'.Fields("Depth").value
					paryCategories(k,6) = pblnBottom
					k = k + 1
				End If
			Next

			'this creates the first dropdown
			If Len(mbytCategoryFilter) = 0 And Len(mbytsubCategoryFilter) = 0 Then pstrSelected = "selected"
			pstrTemp = "<select id='CategoryFilter' name='CategoryFilter' size='10' multiple>" & vbcrlf _
					 & "  <option value='' " & pstrSelected & ">- All -</Option>" & vbcrlf

			For i = 0 to UBound(paryCategories_All, 2)
			
				If len(paryCategories_All(5, i) & "") > 0 Then 'Then No sub-categories so don't display
					pbytCatDepth = paryCategories_All(5, i)
					
					If pstrCatName <> paryCategories_All(2, i) Then
						pstrCatValue = paryCategories_All(0, i) & "."
						pstrCatName = paryCategories_All(2, i)
						pstrSelected = ""
						If pstrCatValue = CStr(mbytCategoryFilter & "." & mbytsubCategoryFilter) Then pstrSelected = "selected"
						pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrCatValue & chr(34) & pstrSelected & ">" & pstrCatName & "</option>" & vbcrlf 
						
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & paryCategories_All(1, i)
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & paryCategories_All(4, i)
							pstrSelected = ""
							If pstrsubCatValue = CStr(mbytCategoryFilter & "." & mbytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrsubCatValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					Else
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & paryCategories_All(1, i)
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & paryCategories_All(4, i)
							pstrSelected = ""
							If pstrsubCatValue = CStr(mbytCategoryFilter & "." & mbytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrsubCatValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					End If
				End If

			Next

			pstrTemp = pstrTemp & "</select>" & vbcrlf

			Dim maryCatHeir
			Const cstrSpacer = " --> "
			For i=0 to plngCatCount
				If isNull(paryCategories(i,1)) Then
					paryCategories(i,7) = paryCategories(i,2)
					paryCategories(i,1) = "none-" & paryCategories(i,0)
				Else
					If InStr(1,paryCategories(i,4),"none") > 0 Then
						paryCategories(i,7) = paryCategories(i,2)
					Else
						maryCatHeir = Split(paryCategories(i,4),"-")
						paryCategories(i,7) = paryCategories(i,2)
						If isArray(maryCatHeir) Then
	'						paryCategories(i,7) = paryCategories(i,7) & cstrSpacer & paryCategories(i,3)
							For j=0 to uBound(maryCatHeir)
								If maryCatHeir(j) <> "none" Then
'									.Filter = "subcatID=" & maryCatHeir(j)
'									paryCategories(i,7) = paryCategories(i,7) & cstrSpacer & .Fields("subcatName").value
									paryCategories(i,7) = paryCategories(i,7) & cstrSpacer & pdicCategories.Item(cStr(maryCatHeir(j)))
								End If
							Next
						End If
					End If
				End If
			Next

			pstrCategoryList = "<select id=CatSource name=CatSource size=10 multiple>"
			pstrCategoryArray = "var maryCategories = new Array();" & vbcrlf

			For i=0 to plngCatCount
				pstrCategoryArray = pstrCategoryArray & "maryCategories[" & paryCategories(i,1) & "] = " & chr(34) & paryCategories(i,7) & chr(34) & ";" & vbcrlf
				pstrCategoryList = pstrCategoryList & "<option value='" & paryCategories(i,1) & "'>" & paryCategories(i,7) & "</option>"
			Next
			pstrCategoryList = pstrCategoryList & "</select>"
			
			pstrCategoryFilterList = pstrTemp
			
			Set pdicCategories = Nothing
		Else
			pstrCategoryList = "<font color=red>No Categories</font><br>"
			pstrCategoryList = pstrCategoryList & "<select id=CatSource name=CatSource size=10 multiple>"
			pstrCategoryList = pstrCategoryList & "</select>"
		End If
	End With

	'Now create the dictionary which will populate the selected categories
	If len(pstrprodID) > 0 Then
		pstrTemp = ""
		Set prsAEProductCategories = GetRS("Select subcatCategoryId from sfSubCatDetail Where ProdID='" & SQLSafe(pstrprodID) & "' Order By subcatCategoryId")
		With prsAEProductCategories
			For i=1 to .RecordCount
				If pstrTemp <> .Fields("subcatCategoryId").Value Then
					pstrTemp = .Fields("subcatCategoryId").Value
					pstrProdDictionaryList = pstrProdDictionaryList & "mdicCategory.add (" & chr(34) & "" & pstrTemp & chr(34) & "," & chr(34) & "" & chr(34) & ");" & vbcrlf
				End If
				.MoveNext
			Next
		End With
	End If
	ReleaseObject(prsAEProductCategories) 
	
End Sub	'AEProductCategory

'***********************************************************************************************

Public Function Load(byVal strProdID)

dim pstrSQL
dim p_strWhere
dim i
dim sql

'On Error Resume Next

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	
	set	prsProducts = server.CreateObject("adodb.recordset")
	with prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
        pstrSQL = "SELECT sfProducts.*, [sfGiftWraps].[gwActivate], [sfGiftWraps].[gwPrice], [sfInventory].[invenId], [sfInventory].[invenAttDetailID], [sfInventory].[invenAttName], [sfInventory].[invenInStock], [sfInventory].[invenLowFlag], [sfInventoryInfo].[invenbBackOrder], [sfInventoryInfo].[invenbTracked], [sfInventoryInfo].[invenbStatus], [sfInventoryInfo].[invenbNotify], [sfInventoryInfo].[invenInStockDEF], [sfInventoryInfo].[invenLowFlagDEF], [sfMTPrices].[mtIndex], [sfMTPrices].[mtQuantity], [sfMTPrices].[mtValue], [sfMTPrices].[mtType]" _
				& "FROM ((sfMTPrices RIGHT JOIN (sfProducts LEFT JOIN sfGiftWraps ON [sfProducts].[prodID]=[sfGiftWraps].[gwProdID]) ON [sfMTPrices].[mtProdID]=[sfProducts].[prodID]) LEFT JOIN sfInventory ON [sfProducts].[prodID]=[sfInventory].[invenProdId]) LEFT JOIN sfInventoryInfo ON [sfProducts].[prodID]=[sfInventoryInfo].[invenProdId]" _
				& "" 'mstrsqlWhere
				
		If (Len(mbytCategoryFilter) > 0 And cblnSF5AE) OR mblnShowUnassignedProducts Then
	        pstrSQL = "SELECT sfProducts.* FROM sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID " & mstrSQLWhere
			If mblnShowUnassignedProducts Then
				pstrSQL = "SELECT sfProducts.*" _
						& " FROM (sfProducts LEFT JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) LEFT JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " & mstrSQLWhere
			Else
				pstrSQL = "SELECT sfProducts.*" _
						& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " & mstrSQLWhere
			End If
	    Else
	        pstrSQL = "Select * from sfProducts " & mstrsqlWhere
	    End If
	    
		If len(mlngMaxRecords) > 0 Then 
			.CacheSize = mlngMaxRecords
			.PageSize = mlngMaxRecords
		End If
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		mlngPageCount = .PageCount
		If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
'debugprint "pstrSQL", pstrSQL		
	end with

	Call LoadProductAttributes(strprodID)

    Load = (Not prsProducts.EOF)

End Function    'Load

'***********************************************************************************************

Public Sub LoadProductAttributes(byVal strProdID)

dim pstrSQL

'On Error Resume Next

	If Not prsProducts.EOF Then
	
		pstrSQL = "Select * from sfAttributes where attrProdId = '" & strprodID & "' Order By attrName"
		If pblnAttributeCategoryOrderable Then
			On Error Resume Next
			Set prsAttributes = GetRS(Replace(pstrSQL, "attrName", pstrFieldAttributeCategoryOrderable))
			If Err.number = -2147217904 Then '-2147217904 = No value given for one or more required parameters
				pblnAttributeCategoryOrderable = False
				Set prsAttributes = GetRS(pstrSQL)
				Err.Clear
			ElseIf Err.number > 0 Then
				debugprint "Error: " & Err.number,Err.Description
			End If
			On Error Goto 0
		Else
			Set prsAttributes = GetRS(pstrSQL)
		End If
		If isObject(prsAttributes) Then Set prsAttributeDetails = GetRS("SELECT sfAttributeDetail.* FROM sfAttributes INNER JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId WHERE attrProdId='" & strprodID & "' Order By attrdtOrder")
	End If	'prsProducts.EOF

End Sub    'LoadProductAttributes

'***********************************************************************************************

Public Function LoadSummary(byRef strProdID)

dim pstrSQL
dim p_strWhere
dim i
dim sql
Dim pstrProductIDList
Dim mblnBadSort

'On Error Resume Next

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	mblnBadSort = False
	
	set	prsProducts = server.CreateObject("adodb.recordset")
	with prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
				
		If Len(mstrsqlWhere) > 0 AND Len(strProdID) > 0 Then mstrsqlWhere = Replace(mstrsqlWhere, "Where ", "Where prodID='" & sqlSafe(strProdID) & "' OR ", 1, 1)
		
		If (Len(mbytCategoryFilter) > 0 And cblnSF5AE) OR mblnShowUnassignedProducts Then
			If mblnShowUnassignedProducts Then
				'Distinct modifier removed because it fails with an order by clause
				'pstrSQL = "Select distinct sfProducts.prodID" _
				pstrSQL = "Select sfProducts.prodID" _
						& " FROM (sfProducts LEFT JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) LEFT JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " _
						& mstrSQLWhere
			Else
				'Distinct modifier removed because it fails with an order by clause
				'pstrSQL = "Select distinct sfProducts.prodID" _
				pstrSQL = "Select sfProducts.prodID" _
						& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " _
						& mstrSQLWhere
			End If
	    Else
	        pstrSQL = "Select sfProducts.prodID" _
					& " FROM sfProducts " _
					& mstrsqlWhere
	    End If

		If len(mlngMaxRecords) > 0 Then 
			.CacheSize = mlngMaxRecords
			.PageSize = mlngMaxRecords
		End If
		On Error Resume Next
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		If Err.number <> 0 Then
			Call commonError(Err)
			Response.Write "<h3><font color=red>Unable to perform filter. This can be caused by sorting on a price/sale price column if products have no prices assigned</font></h3>"
			Err.Clear
	        pstrSQL = "Select sfProducts.prodID FROM sfProducts"
			.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
			mblnBadSort = True
		End If
		
		mlngPageCount = .PageCount
		If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
		
		If .EOF Then
			plngRecordCount = 0
		Else
			Dim plnglbound
			If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
			If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
			plnglbound = (mlngAbsolutePage - 1) * .PageSize + 1
			.AbsolutePosition = plnglbound

			plngRecordCount = .RecordCount
			If Len(strProdID) = 0 Then strProdID = .Fields("prodID").Value
		
			pstrProductIDList = .getString(adClipString, mlngMaxRecords, "-", "', '", "")
			pstrProductIDList = "'" & Left(pstrProductIDList, Len(pstrProductIDList) - 3)	'remove the trailing comma

			If Len(pstrProductIDList) > 0 Then
				mstrSQLWhere = "Where sfProducts.prodID In (" & pstrProductIDList & ")"
				If Not mblnBadSort Then Call LoadSort
			Else
				mstrSQLWhere = "Where prodEnabledIsActive=1 And prodEnabledIsActive=0"
			End If
			
			pstrSQL = "Select sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDateAdded, sfProducts.prodEnabledIsActive"
				pstrSQL = pstrSQL _
						& " FROM sfProducts " _
						& mstrsqlWhere
			.Close
			.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		End If
		
		'debugprint "pstrSQL", pstrSQL		
	end with

    LoadSummary = (Not prsProducts.EOF)

End Function    'LoadSummary

'***********************************************************************************************

Public Function LoadSummary_old(byRef strProdID)

dim pstrSQL
dim p_strWhere
dim i
dim sql

'On Error Resume Next

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	
	set	prsProducts = server.CreateObject("adodb.recordset")
	with prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
				
		If Len(mstrsqlWhere) > 0 AND Len(strProdID) > 0 Then mstrsqlWhere = Replace(mstrsqlWhere, "Where ", "Where prodID='" & sqlSafe(strProdID) & "' OR ", 1, 1)
		
	    pstrSQL = "Select sfProducts.prodID, sfProducts.prodName, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDateAdded, sfProducts.prodEnabledIsActive"
		If (Len(mbytCategoryFilter) > 0 And cblnSF5AE) OR mblnShowUnassignedProducts Then
			If mblnShowUnassignedProducts Then
				pstrSQL = pstrSQL _
						& " FROM (sfProducts LEFT JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) LEFT JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " _
						& mstrSQLWhere
			Else
				pstrSQL = pstrSQL _
						& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " _
						& mstrSQLWhere
			End If
	    Else
	        pstrSQL = pstrSQL _
					& " FROM sfProducts " _
					& mstrsqlWhere
	    End If
	    
		If len(mlngMaxRecords) > 0 Then 
			.CacheSize = mlngMaxRecords
			.PageSize = mlngMaxRecords
		End If
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		mlngPageCount = .PageCount
		If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
		
		If Not .EOF And Len(strProdID) = 0 Then strProdID = .Fields("prodID").Value
		'debugprint "pstrSQL", pstrSQL		
	end with

    LoadSummary = (Not prsProducts.EOF)

End Function    'LoadSummary_old

'***********************************************************************************************

Public Function Activate(byVal strProdID, byVal blnActivate)


Dim pstrSQL
Dim pstrProdName

'On Error Resume Next

	Call getProductInfo(strProdID, pstrProdName)

	if blnActivate then
		pstrSQL = "Update sfProducts Set prodEnabledIsActive=1 where prodID='" & strProdID & "'"
		cnn.Execute pstrSQL, , 128
		pstrMessage = pstrMessage & cstrdelimeter & pstrProdName & " (" & strProdID & ") successfully activated."
    else
		pstrSQL = "Update sfProducts Set prodEnabledIsActive=0 where prodID='" & strProdID & "'"
		cnn.Execute pstrSQL, , 128
		pstrMessage = pstrMessage & cstrdelimeter & pstrProdName & " (" & strProdID & ") successfully deactivated."
	end if
    
    If (Err.Number = 0) Then
        Activate = True
    Else
		pstrMessage = pstrMessage & cstrdelimeter & "Error activating product " & pstrProdName & "(" & strProdID & "): " & Err.Description
		pblnError = True
        Activate = False
    End If

End Function    'Activate

'***********************************************************************************************

Public Function DeleteProduct(strProdID)

Dim sql
Dim p_rsAttr
Dim i

'On Error Resume Next

	If len(strProdID) = 0 Then Exit Function
	sql = "Select attrID from sfAttributes where attrProdId='" & SQLSafe(strProdID) & "'"
	set p_rsAttr = GetRS(sql)
	
	for i = 1 to p_rsAttr.RecordCount
		Call DeleteAttribute(p_rsAttr("attrID"))
		p_rsAttr.MoveNext
	next
	
	sql = "Delete from sfProducts where ProdId='" & SQLSafe(strProdID) & "'"
	cnn.Execute sql, , 128
	
    If (Err.Number = 0) Then
        pstrMessage = "The product was successfully deleted."
        DeleteProduct = True
    Else
        pstrMessage = Err.Description
        DeleteProduct = False
    End If
    
    p_rsAttr.Close
    Set p_rsAttr = Nothing
    
	If cblnSF5AE Then

		'AE Specific
		sql = "Delete From sfGiftWraps Where gwProdID='" & SQLSafe(strProdID) & "'"
		cnn.Execute sql,,128

		sql = "Delete From sfMTPrices Where mtProdID='" & SQLSafe(strProdID) & "'"
		cnn.Execute sql,,128

		sql = "Delete From sfInventory Where invenProdId='" & SQLSafe(strProdID) & "'"
		cnn.Execute sql,,128

		sql = "Delete From sfInventoryInfo Where invenProdId='" & SQLSafe(strProdID) & "'"
		cnn.Execute sql,,128

		sql = "Delete From sfSubCatDetail Where ProdID='" & SQLSafe(strProdID) & "'"
		cnn.Execute sql,,128

	End If

End Function    'DeleteProduct

'***********************************************************************************************

Public Function DeleteAllAttributes(strProdID)

Dim sql
Dim pobjRS

'On Error Resume Next

	sql = "Select attrID from sfAttributes where attrProdId='" & SQLSafe(strProdID) & "'"
	Set pobjRS = GetRS(sql)
	Do While Not pobjRS.EOF
		Call DeleteAttribute(pobjRS.Fields("attrID").Value)
		pobjRS.MoveNext
	Loop
	pobjRS.Close
	Set pobjRS = Nothing
    
   If (Err.Number = 0) Then
        pstrMessage =  "The product attributes were successfully deleted.<br>" & pstrMessage
        DeleteAllAttributes = True
    Else
        pstrMessage = Err.Description
        DeleteAllAttributes = False
    End If

End Function    'DeleteAllAttributes

'***********************************************************************************************

Public Function DeleteAttribute(lngID)

Dim sql
Dim pobjRS
Dim p_strID, p_strName
Dim p_intCount

'On Error Resume Next

	sql = "Select attrProdId,attrName from sfAttributes where attrID=" & lngID
	Set pobjRS = GetRS(sql)
	if not pobjRS.eof Then 
		p_strID = pobjRS("attrProdId").value
		p_strName = pobjRS("attrName").value
	end if
	pobjRS.Close
    Set pobjRS = Nothing
    
	If len(lngID) = 0 Then Exit Function
    sql = "Delete from sfAttributeDetail where attrdtAttributeId=" & lngID
    cnn.Execute sql, , 128

    sql = "Delete from sfAttributes where attrID=" & lngID
    cnn.Execute sql, , 128

	sql = "Select * from sfAttributes where attrProdId='" & sqlSafe(p_strID) & "'"
	Set pobjRS = GetRS(sql)
	If pobjRS.EOF Then
		sql = "Update sfProducts set prodAttrNum=0 where prodID='" & sqlSafe(p_strID) & "'"
		cnn.Execute sql,,128
	Else
		sql = "Update sfProducts set prodAttrNum=" & pobjRS.RecordCount & " where prodID='" & sqlSafe(p_strID) & "'"
		cnn.Execute sql,,128
	End If
	pobjRS.Close
    Set pobjRS = Nothing
    
   If (Err.Number = 0) Then
        pstrMessage = pstrMessage & "<BR>The product attribute " & p_strName & " was successfully deleted."
		If cblnSF5AE Then Call UpdateInventoryFields
        DeleteAttribute = True
    Else
        pstrMessage = Err.Description
        DeleteAttribute = False
    End If

End Function    'DeleteAttribute

'***********************************************************************************************

Public Function DeleteAttributeDetail(lngID)

Dim sql

'On Error Resume Next

	If len(lngID) = 0 Then Exit Function
    sql = "Delete from sfAttributeDetail where attrdtID = " & lngID
    cnn.Execute sql, , 128

    If (Err.Number = 0) Then
        pstrMessage = "The product attribute detail was successfully deleted."
        DeleteAttributeDetail = True
    Else
        pstrMessage = Err.Description
        DeleteAttributeDetail = False
    End If
    
	If cblnSF5AE Then

		'Delete AE Inventory
		Dim prs
		Dim paryAttDetailID
		Dim i
    
		sql = "Select invenId,invenAttDetailID From sfInventory Where invenAttDetailID Like '" & lngID & "'"
		Set prs = GetRS(sql)
		With prs
			Do While Not .EOF
				paryAttDetailID = Split(.Fields("invenAttDetailID").Value,",")
				For i=0 To uBound(paryAttDetailID)
					If cLng(paryAttDetailID(i)) = cLng(lngID) Then
						sql = "Delete From sfInventory Where invenId=" & .Fields("invenId").Value
						cnn.Execute sql,,128
						Exit For
					End If
				Next
				.MoveNext
			Loop
		End With
		
	End If
	
End Function    'DeleteAttributeDetail

'***********************************************************************************************

Public Function CopyAE(strProductID,strNewProductID)

Dim sql
Dim prsSource, prsTarget
Dim i,j
Dim pstrField
Dim plngNewID

'On Error Resume Next

	If len(strProductID) = 0 Then Exit Function
	
	'Update GiftWraps
	On Error Resume Next
	sql = "Select gwActivate,gwPrice from sfGiftWraps where gwProdID = '" & SQLSafe(strProductID) & "'"
	Set prsSource = GetRS(sql)
	For i=1 to prsSource.RecordCount
		sql = "Insert Into sfGiftWraps (gwProdID,gwActivate,gwPrice) Values (" _
			& "'" & SQLSafe(strNewProductID) & "'," _
			& prsSource.Fields("gwActivate").Value & ",'" _
			& prsSource.Fields("gwPrice").Value & "')"
		cnn.Execute sql,,128
		prsSource.MoveNext
	Next
	On Error Goto 0
	
	'Update MTP
	On Error Resume Next
	sql = "Select mtIndex,mtQuantity,mtValue,mtType from sfMTPrices where mtProdID = '" & SQLSafe(strProductID) & "'"
	Set prsSource = GetRS(sql)
	For i=1 to prsSource.RecordCount
		sql = "Insert Into sfMTPrices (mtProdID,mtIndex,mtQuantity,mtValue,mtType) Values (" _
			& "'" & SQLSafe(strNewProductID) & "'," _
			& prsSource.Fields("mtIndex").Value & "," _
			& prsSource.Fields("mtQuantity").Value & "," _
			& prsSource.Fields("mtValue").Value & ",'" _
			& prsSource.Fields("mtType").Value & "')"
		cnn.Execute sql,,128
		prsSource.MoveNext
	Next
	On Error Goto 0
	
	'Update InventoryInfo
	On Error Resume Next
	sql = "Select invenbBackOrder,invenbTracked,invenbStatus,invenbNotify,invenInStockDEF,invenLowFlagDEF from sfInventoryInfo where invenProdId = '" & SQLSafe(strProductID) & "'"
	Set prsSource = GetRS(sql)
	For i=1 to prsSource.RecordCount
		sql = "Insert Into sfInventoryInfo (invenProdId,invenbBackOrder,invenbTracked,invenbStatus,invenbNotify,invenInStockDEF,invenLowFlagDEF) Values (" _
			& "'" & SQLSafe(strNewProductID) & "'," _
			& prsSource.Fields("invenbBackOrder").Value & "," _
			& prsSource.Fields("invenbTracked").Value & "," _
			& prsSource.Fields("invenbStatus").Value & "," _
			& prsSource.Fields("invenbNotify").Value & "," _
			& prsSource.Fields("invenInStockDEF").Value & "," _
			& prsSource.Fields("invenLowFlagDEF").Value & ")"
		cnn.Execute sql,,128
		prsSource.MoveNext
	Next
	On Error Goto 0
	
	'Update Categories
'	sql = "Select subcatCategoryId,ProdName from sfSub_Categories where subcatID = '" & SQLSafe(strProductID) & "'"
'	Set prsSource = GetRS(sql)
'	For i=1 to prsSource.RecordCount
'		sql = "Insert Into sfInventoryInfo (invenProdId,invenbBackOrder,invenbTracked,invenbStatus,invenbNotify,invenInStockDEF,invenLowFlagDEF) Values (" _
'			& "'" & SQLSafe(strNewProductID) & "'," _
'			& prsSource.Fields("invenbBackOrder").Value & "," _
'			& prsSource.Fields("invenbTracked").Value & "," _
'			& prsSource.Fields("invenbStatus").Value & "," _
'			& prsSource.Fields("invenbNotify").Value & "," _
'			& prsSource.Fields("invenInStockDEF").Value & "," _
'			& prsSource.Fields("invenLowFlagDEF").Value & ")"
'		cnn.Execute sql,,128
'		prsSource.MoveNext
'	Next

	'Inserted to keep inventory updated
	pstrProdID = strNewProductID
	Call UpdateInventoryFields

End Function	'CopyAE

'***********************************************************************************************

Public Function CopyProductAttributesToExistingProduct(strProductID,strNewProductID,lngAttrID)

Dim sql
Dim prsSource
Dim rsAttribute
Dim p_strAttName
Dim i
Dim pstrTempMessage

'added to do unlimited attribute copying
Dim plngProductCounter
Dim pstrTempProductID
Dim paryProductID

'On Error Resume Next

	If len(strProductID) = 0 Then Exit Function
	paryProductID = Split(strNewProductID,",")
	
	For plngProductCounter = 0 To UBound(paryProductID)
		pstrTempProductID = Trim(paryProductID(plngProductCounter))
	
		sql = "Select prodID from sfProducts where prodID = '" & SQLSafe(pstrTempProductID) & "'"
		Set prsSource = GetRS(sql)

		If prsSource.EOF Then
			pstrTempMessage = pstrTempMessage & "Product " & pstrTempProductID & " does not exist.<br>"
			CopyProductAttributesToExistingProduct = False
		Else
			
			If Len(lngAttrID) = 0 Then
				sql = "Select * from sfAttributes where attrProdId='" & strProductID & "'"
				Set rsAttribute = GetRS(sql)
				For i = 1 to rsAttribute.RecordCount
					Call CopyAttribute(rsAttribute.Fields("attrID").Value, pstrTempProductID, "")
					rsAttribute.MoveNext
				Next
			Else
				sql = "Select * from sfAttributes where attrID=" & lngAttrID
				Set rsAttribute = GetRS(sql)
				If rsAttribute.EOF Then
					pstrTempMessage = pstrTempMessage & "Error Locating Attribute.<br>"
					CopyProductAttributesToExistingProduct = False
				Else
					p_strAttName = rsAttribute.Fields("attrName").Value
					Call CopyAttribute(rsAttribute.Fields("attrID").Value, pstrTempProductID, "")
				End If
			End If
			
			rsAttribute.Close
			Set rsAttribute = Nothing
			
			If (Err.Number = 0) Then
				If Len(lngAttrID) > 0 Then
					pstrTempMessage = pstrTempMessage & "Attribute " & p_strAttName & " was successfully copied to Product " & pstrTempProductID & ".<br>"
				Else
					pstrTempMessage = pstrTempMessage & "Product " & strProductID & "'s attributes were successfully copied to Product " & pstrTempProductID & ".<br>"
				End If
				CopyProductAttributesToExistingProduct = True
			ElseIf Err.number <> 0 Then
				pstrTempMessage = Err.Description
				CopyProductAttributesToExistingProduct = False
			End If
			
		End If
		
		Call UpdateAttributeCount(pstrTempProductID)
		
	Next 'plngProductCounter

	pstrMessage = pstrTempMessage
	prsSource.Close
	set prsSource = Nothing
	
End Function    'CopyProductAttributesToExistingProduct

'***********************************************************************************************

Public Sub UpdateAttributeCount(strProductID)

Dim sql
Dim rs

'On Error Resume Next

		Set rs = Server.CreateObject("ADODB.RECORDSET")
		sql = "Select * from sfAttributes where attrProdId='" & sqlSafe(strProductID) & "'"
		rs.open sql, cnn, 1, 3
		If Not rs.EOF Then
 			sql = "Update sfProducts set prodAttrNum=" & rs.RecordCount & " where prodID='" & sqlSafe(strProductID) & "'"
			cnn.Execute sql,,128
		End If
		rs.Close
        Set rs = Nothing

End Sub    'UpdateAttributeCount

'***********************************************************************************************

Private Function replaceProductIDWithNew(strProductID,strNewProductID, objField)

Dim pstrFieldName
Dim pstrTempValue

	pstrFieldName = objField.Name
	
	Select Case pstrFieldName
		Case "prodImageSmallPath", "prodImageLargePath", "prodLink"
			pstrTempValue = Replace(Trim(objField.Value & ""), strProductID, strNewProductID)
		Case Else
			pstrTempValue = objField.Value
			For i = 0 To UBound(maryImageFields)
				If maryImageFields(i)(1) = pstrFieldName Then
					pstrTempValue = Replace(Trim(objField.Value & ""), strProductID, strNewProductID)
				End If
			Next 'i
	End Select
	
	replaceProductIDWithNew = pstrTempValue

End Function	'replaceProductIDWithNew

'***********************************************************************************************

Public Function DuplicateProduct(strProductID,strNewProductID)

Dim sql
Dim prsSource, prsTarget
Dim i,j
Dim pstrField
Dim plngNewID

'On Error Resume Next

	If len(strProductID) = 0 Then Exit Function
	
	sql = "Select * from sfProducts where prodID = '" & SQLSafe(strProductID) & "'"
	Set prsSource = GetRS(sql)

	If Not prsSource.EOF Then
	
		sql = "Select * from sfProducts where prodID = '" & SQLSafe(strNewProductID) & "'"
		Set prsTarget = server.CreateObject("adodb.Recordset")
		prsTarget.open sql, cnn, 1, 3
		If Not prsTarget.EOF Then
			prsTarget.Close
			Set prsTarget = Nothing
			prsSource.Close
			set prsSource = Nothing
			pstrMessage = "The ProductID " & strNewProductID & " already exists."
			DuplicateProduct = False
			Exit Function
		End If
		prsTarget.AddNew

		For j=0 to prsSource.fields.count-1
			pstrField = prsTarget.fields(j).name
			'skip autonumber and timestamps
			If LCase(pstrField) <> "sfproductid" And prsSource(pstrField).Type <> 128 Then
				On Error Resume Next
				prsTarget(pstrField).value = replaceProductIDWithNew(strProductID, strNewProductID, prsSource(pstrField))
				If Err.number <> 0 Then
					Response.Write "<fieldset><legend>Error</legend>" _
								 & "Error " & Err.number & ": " & Err.Description & "<br>" _
								 & "Field: " & prsSource(pstrField).Name & "<br>" _
								 & "Type : " & prsSource(pstrField).Type & "<br>" _
								 & "prsSource(pstrField): " & prsSource(pstrField) & "<br>" _
								 & "strProductID: " & strProductID & "<br>" _
								 & "strNewProductID: " & strNewProductID & "<br>" _
								 & "replaceProductIDWithNew: " & replaceProductIDWithNew(strProductID, strNewProductID, prsSource(pstrField))& "<br>" _
								 & "</fieldset>"
					Err.Clear
				End If
			Else
				'Response.Write "<font color=red>" & pstrField & " was skipped</font><br />"
			End If
		Next
		prsTarget("prodID") = strNewProductID

		prsTarget.Update
		prsTarget.Close
		Set prsTarget = Nothing
		
		prsSource.Close
		sql = "Select attrID from sfAttributes where attrProdId='" & strProductID & "'"
		Set prsSource = GetRS(sql)

		For i = 1 to prsSource.RecordCount
			Call CopyAttribute(prsSource("attrID"),strNewProductID,"")
			prsSource.MoveNext
		Next
		
	End If
	
	prsSource.Close
	set prsSource = Nothing
	
	If cblnSF5AE Then Call CopyAE(strProductID,strNewProductID)
	
	If (Err.Number = 0) Then
	    pstrMessage = strNewProductID & " was successfully created."
	    DuplicateProduct = True
	Else
	    pstrMessage = Err.Description
	    DuplicateProduct = False
	End If

End Function    'DuplicateProduct

'***********************************************************************************************

Public Function CopyProductCategories(strProductID, strNewProductID)

Dim sql
Dim prsSource
Dim pstrNewProductName

'On Error Resume Next

	If len(strProductID) = 0 Then Exit Function
	If Not cblnSF5AE Then Exit Function
	
	Call CopyAE(strProductID,strNewProductID)
	
	'Get prodName for new product ID
	sql = "Select prodName From sfProducts where prodID = '" & SQLSafe(strProductID) & "'"
	Set prsSource = GetRS(sql)
	If Not prsSource.EOF Then pstrNewProductName = Trim(prsSource.Fields("prodName").Value & "")
	prsSource.Close
	set prsSource = Nothing

	'Get category assignments from source product
	sql = "Select subcatCategoryId From sfSubCatDetail where prodID = '" & SQLSafe(strProductID) & "'"
	Set prsSource = GetRS(sql)
	Do While Not prsSource.EOF
		sql = "Insert Into sfSubCatDetail (subcatCategoryId, prodID, ProdName) Values (" & prsSource.Fields("subcatCategoryId").Value & ", '" & SQLSafe(strNewProductID) & "', '" & SQLSafe(pstrNewProductName) & "')"
		cnn.Execute sql,,128
		prsSource.MoveNext
	Loop
	prsSource.Close
	set prsSource = Nothing
	
	If (Err.Number = 0) Then
	    pstrMessage = pstrMessage & "<BR>Categories successfully assigned to " & strNewProductID & "."
	    CopyProductCategories = True
	Else
	    pstrMessage = Err.Description
	    CopyProductCategories = False
	End If

End Function    'CopyProductCategories

'***********************************************************************************************

Public Function CopyAttribute(lngattrID,strProductID,strNewName)

Dim sql
Dim prsSource, prsTarget
Dim i
Dim plngNewID
Dim pBookmark

'On Error Resume Next

	sql = "Select * from sfAttributes where attrID=" & lngattrID
	Set prsSource = GetRS(sql)

	If Not prsSource.EOF Then
	
		sql = "Select * from sfAttributes where attrID=0"
		sql = "sfAttributes Order By attrID"
		Set prsTarget = server.CreateObject("adodb.Recordset")
		
		prsTarget.CursorLocation = 3			'adUseClient
		prsTarget.open sql, cnn, 2, 3,&H0002	'adOpenDynamic, adLockOptimistic, adCmdTable
		prsTarget.AddNew

		For i = 0 to prsTarget.Fields.Count - 1
			Select Case prsTarget.Fields(i).name
				Case "attrProdId"
					If len(strProductID) > 0 Then
						prsTarget("attrProdId") = strProductID
					Else
						prsTarget("attrProdId") = prsSource("attrProdId")
					End If
				Case "attrName"
					If len(strNewName) > 0 Then
						prsTarget("attrName") = strNewName
					Else
						prsTarget("attrName") = prsSource("attrName")
					End If
				Case "attrID"
					'do nothing
				Case Else
					If prsSource.Fields(i).Type <> 128 Then prsTarget.Fields(i).Value = prsSource.Fields(i).Value
			End Select
		Next
		
		prsTarget.Update
		plngNewID = prsTarget.Fields("attrID").Value

		If plngNewID = 0 Then
			pBookmark = prsTarget.AbsolutePosition 
			prsTarget.Requery 
			prsTarget.AbsolutePosition = pBookmark
			plngNewID = prsTarget.Fields("attrID").Value
			If plngNewID = 0 Then pstrMessage = "<font color=red>Error creating " & strNewName & " - could not retrieve attributeID</font><br>"
		End If

		prsTarget.Close
		Set prsTarget = Nothing
		sql = "Select attrdtID from sfAttributeDetail where attrdtAttributeId=" & lngattrID

		prsSource.Close
		
		Set prsSource = GetRS(sql)
		For i = 1 to prsSource.RecordCount
			Call CopyAttributeDetail(prsSource("attrdtID"),plngNewID,"")
			prsSource.MoveNext
		Next
	End If
	
	prsSource.Close
	set prsSource = Nothing

	If (Err.Number = 0) Then
	    pstrMessage = strNewName & " was successfully created."
	    CopyAttribute = True
	Else
	    pstrMessage = Err.Description
	    CopyAttribute = False
	End If

End Function    'CopyAttribute

'***********************************************************************************************

Public Function CopyAttributeDetail(lngattrdtID,lngattrID,strNewName)

Dim sql
Dim prsSource, prsTarget
Dim i

'On Error Resume Next

	sql = "Select * from sfAttributeDetail where attrdtID = " & lngattrdtID
	Set prsSource = GetRS(sql)

	If Not prsSource.EOF Then
	
		sql = "Select * from sfAttributeDetail where attrdtID = 0"
		Set prsTarget = server.CreateObject("adodb.Recordset")
		prsTarget.open sql, cnn, 1, 3
		prsTarget.AddNew
	
		For i = 0 to prsTarget.Fields.Count - 1
			Select Case prsTarget.Fields(i).name
				Case "attrdtAttributeId"
					If len(lngattrID) > 0 Then 
						prsTarget("attrdtAttributeId") = lngattrID
					Else
						prsTarget("attrdtAttributeId") = prsSource("attrdtAttributeId")
					End If
				Case "attrdtName"
					If len(strNewName) > 0 Then 
						prsTarget("attrdtName") = pstrNewName
					Else
						prsTarget("attrdtName") = prsSource("attrdtName")
					End If
				Case "attrdtID"
					'do nothing
				Case Else
					If prsSource.Fields(i).Type <> 128 Then prsTarget.Fields(i).Value = prsSource.Fields(i).Value
			End Select
		Next

		prsTarget.Update
		prsTarget.Close
		Set prsTarget = Nothing
	End If
	
	prsSource.Close
	Set prsSource = Nothing
	
	If (Err.Number = 0) Then
	    pstrMessage = strNewName & " was successfully created."
	    CopyAttributeDetail = True
	Else
	    pstrMessage = Err.Description
	    CopyAttributeDetail = False
	End If

End Function    'CopyAttributeDetail

'***********************************************************************************************

Public Function Update()

Dim sql
Dim rs
Dim strErrorMessage
Dim blnAdd
Dim pstrOrigprodID
Dim p_strTableName, p_strFieldName

'On Error Resume Next

    pblnError = False
    Call LoadFromRequest

    strErrorMessage = ValidateValues
    If ValidateValues Then
    
        If Len(pstrprodID) = 0 Then pstrprodID = "0"
		pstrOrigprodID = trim(Request.Form("OrigprodID"))

		If ((Len(pstrOrigprodID) > 0) and (pstrOrigprodID <> pstrprodID)) Then
			p_strTableName = "sfProducts"
			p_strFieldName = "prodID"
			sql = "Update " & p_strTableName & " set " & p_strFieldName & "='" & sqlSafe(pstrprodID) & "' where " & p_strFieldName & "='" & sqlSafe(pstrOrigprodID) & "'"
			cnn.Execute sql,,128
			If Err.Number <> 0 Then
			    pstrMessage = pstrMessage & cstrdelimeter & "Error updating the table " & p_strTableName & ": " & Err.Number & " - " & Err.Description & "<BR>"
				pblnError = True
			    Err.Clear
			End If
			
			p_strTableName = "sfAttributes"
			p_strFieldName = "attrProdId"
			sql = "Update " & p_strTableName & " set " & p_strFieldName & "='" & sqlSafe(pstrprodID) & "' where " & p_strFieldName & "='" & sqlSafe(pstrOrigprodID) & "'"
			cnn.Execute sql,,128
			If Err.Number <> 0 Then
			    pstrMessage = pstrMessage & cstrdelimeter & "Error updating the table " & p_strTableName & ": " & Err.Number & " - " & Err.Description & "<BR>"
				pblnError = True
			    Err.Clear
			End If
			
			If cblnSF5AE Then 

				'added for AE
				p_strTableName = "sfGiftWraps"
				p_strFieldName = "gwProdID"
				sql = "Update " & p_strTableName & " set " & p_strFieldName & "='" & sqlSafe(pstrprodID) & "' where " & p_strFieldName & "='" & sqlSafe(pstrOrigprodID) & "'"
				cnn.Execute sql,,128
				If Err.Number <> 0 Then
				    pstrMessage = pstrMessage & cstrdelimeter & "Error updating the table " & p_strTableName & ": " & Err.Number & " - " & Err.Description & "<BR>"
					pblnError = True
				    Err.Clear
				End If
			
				p_strTableName = "sfInventory"
				p_strFieldName = "invenProdId"
				sql = "Update " & p_strTableName & " set " & p_strFieldName & "='" & sqlSafe(pstrprodID) & "' where " & p_strFieldName & "='" & sqlSafe(pstrOrigprodID) & "'"
				cnn.Execute sql,,128
				If Err.Number <> 0 Then
				    pstrMessage = pstrMessage & cstrdelimeter & "Error updating the table " & p_strTableName & ": " & Err.Number & " - " & Err.Description & "<BR>"
					pblnError = True
				    Err.Clear
				End If
			
				p_strTableName = "sfInventoryInfo"
				p_strFieldName = "invenProdId"
				sql = "Update " & p_strTableName & " set " & p_strFieldName & "='" & sqlSafe(pstrprodID) & "' where " & p_strFieldName & "='" & sqlSafe(pstrOrigprodID) & "'"
				cnn.Execute sql,,128
				If Err.Number <> 0 Then
				    pstrMessage = pstrMessage & cstrdelimeter & "Error updating the table " & p_strTableName & ": " & Err.Number & " - " & Err.Description & "<BR>"
					pblnError = True
				    Err.Clear
				End If
			
				p_strTableName = "sfMTPrices"
				p_strFieldName = "mtProdID"
				sql = "Update " & p_strTableName & " set " & p_strFieldName & "='" & sqlSafe(pstrprodID) & "' where " & p_strFieldName & "='" & sqlSafe(pstrOrigprodID) & "'"
				cnn.Execute sql,,128
				If Err.Number <> 0 Then
				    pstrMessage = pstrMessage & cstrdelimeter & "Error updating the table " & p_strTableName & ": " & Err.Number & " - " & Err.Description & "<BR>"
					pblnError = True
				    Err.Clear
				End If
			
				p_strTableName = "sfSubCatDetail"
				p_strFieldName = "ProdID"
				sql = "Update " & p_strTableName & " set " & p_strFieldName & "='" & sqlSafe(pstrprodID) & "' where " & p_strFieldName & "='" & sqlSafe(pstrOrigprodID) & "'"
				cnn.Execute sql,,128
				If Err.Number <> 0 Then
				    pstrMessage = pstrMessage & cstrdelimeter & "Error updating the table " & p_strTableName & ": " & Err.Number & " - " & Err.Description & "<BR>"
					pblnError = True
				    Err.Clear
				End If
				
			End If
			
		Else
		End If
		sql = "Select * from sfProducts where prodID = '" & pstrprodID & "'"
		
        Set rs = server.CreateObject("adodb.Recordset")
		rs.CursorLocation = 3
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

'       rs("prodAttrNum") = plngprodAttrNum
        rs("prodID") = pstrprodID
        rs("prodCountryTaxIsActive") = pblnprodCountryTaxIsActive * -1
        rs("prodDescription") = pstrprodDescription
        rs("prodEnabledIsActive") = pblnprodEnabledIsActive * -1
        rs("prodLimitQtyToMTP") = pblnprodLimitQtyToMTP * -1
        
		rs.Fields("prodHeight").Value = wrapSQLValue(pdblprodHeight, False, enDatatype_NA)	'NA type used to prevent quote wraps
		rs.Fields("prodWeight").Value = wrapSQLValue(pdblprodWeight, False, enDatatype_NA)	'NA type used to prevent quote wraps
		rs.Fields("prodWidth").Value = wrapSQLValue(pdblprodWidth, False, enDatatype_NA)	'NA type used to prevent quote wraps
		rs.Fields("prodLength").Value = wrapSQLValue(pdblprodLength, False, enDatatype_NA)	'NA type used to prevent quote wraps
	    
		Call setRSFieldValue(rs, "prodEnableAlsoBought", pblnprodEnableAlsoBought * -1)
		Call setRSFieldValue(rs, "prodEnableReviews", pblnprodEnableReviews * -1)

		Call setRSFieldValue(rs, "pageName", pstrpageName)
		Call setRSFieldValue(rs, "metaTitle", pstrmetaTitle)
		Call setRSFieldValue(rs, "metaDescription", pstrmetaDescription)
		Call setRSFieldValue(rs, "metaKeywords", pstrmetaKeywords)

		Call setRSFieldValue(rs, "prodDisplayAdditionalImagesInWindow", pbytprodDisplayAdditionalImagesInWindow)
		Call setRSFieldValue(rs, "prodAdditionalImages", pstrprodAdditionalImages)
		Call setRSFieldValue(rs, "prodImageSmallPath", pstrprodImageSmallPath)
		Call setRSFieldValue(rs, "prodImageLargePath", pstrprodImageLargePath)
		For i = 0 To UBound(maryImageFields)
			Call setRSFieldValue(rs, maryImageFields(i)(1), maryImageFields(i)(2))
		Next 'i

        rs("prodCategoryId") = plngprodCategoryId
        rs("prodManufacturerId") = plngprodManufacturerId
        rs("prodVendorId") = plngprodVendorId
        rs("prodPrice") = pstrprodPrice
        rs("prodSaleIsActive") = pblnprodSaleIsActive * -1
        rs("prodSalePrice") = pstrprodSalePrice
        rs("prodShip") = pstrprodShip
 		Call setRSFieldValue(rs, "prodFixedShippingCharge", pdblprodFixedShippingCharge)
 		Call setRSFieldValue(rs, "prodSpecialShippingMethods", pstrprodSpecialShippingMethods)
        rs("prodShipIsActive") = pblnprodShipIsActive * -1
        rs("prodStateTaxIsActive") = pblnprodStateTaxIsActive * -1
        
 		Call setRSFieldValue(rs, "prodShortDescription", pstrprodShortDescription)
 		Call setRSFieldValue(rs, "prodName", pstrprodName)
 		Call setRSFieldValue(rs, "prodNamePlural", pstrprodNamePlural)
 		Call setRSFieldValue(rs, "prodMessage", pstrprodMessage)
 		Call setRSFieldValue(rs, "prodLink", pstrprodLink)
 		Call setRSFieldValue(rs, "BuyersClubIsPercentage", pblnBuyersClubIsPercentage)
 		Call setRSFieldValue(rs, "BuyersClubPointValue", pdblBuyersClubPointValue)

 		Call setRSFieldValue(rs, "prodHandlingFee", pstrprodHandlingFee)
 		Call setRSFieldValue(rs, "prodSetupFee", pstrprodSetupFee)
 		Call setRSFieldValue(rs, "prodSetupFeeOneTime", pstrprodSetupFeeOneTime)
       
 		Call setRSFieldValue(rs, "prodMinQty", pstrprodMinQty)
 		Call setRSFieldValue(rs, "prodIncrement", pstrprodIncrement)
       
        rs("prodDateModified") =  Now()
        
		If Len(pdtprodDateAdded) = 0 Then
			rs("prodDateAdded") =  Now()
		ElseIf isDate(pdtprodDateAdded) Then
			rs("prodDateAdded") =  pdtprodDateAdded
		End If
		
		If pblnPricingLevel Then
			Call setRSFieldValue(rs, "prodPLPrice", pstrprodPLPrice)
			Call setRSFieldValue(rs, "prodPLSalePrice", pstrprodPLSalePrice)
		End If
		Call setRSFieldValue(rs, "prodFileName", pstrprodFileName)
		Call setRSFieldValue(rs, "UpgradeVersion", pstrUpgradeVersion)
		Call setRSFieldValue(rs, "packageCodes", pstrPackageCodes)
		Call setRSFieldValue(rs, "Version", pstrVersion)
		Call setRSFieldValue(rs, "ReleaseDate", pstrReleaseDate)
		Call setRSFieldValue(rs, "InstallationHours", pdblInstallationHours)
		Call setRSFieldValue(rs, "MyProduct", pblnMyProduct)
		Call setRSFieldValue(rs, "InstallationRequired", pblnInstallationRequired)
		Call setRSFieldValue(rs, "IncludeInSearch", pblnIncludeInSearch)
		Call setRSFieldValue(rs, "IncludeInRandomProduct", pblnIncludeInRandomProduct)
		Call setRSFieldValue(rs, "prodMaxDownloads", pstrprodMaxDownloads)
		Call setRSFieldValue(rs, "prodDownloadValidFor", pstrprodDownloadValidFor)
        If cblnAddon_DynamicProductDisplay Then Call setRSFieldValue(rs, "relatedProducts", wrapSQLValue(pstrRelatedProducts, False, enDatatype_NA))	'NA type used to prevent quote wraps

        Call UpdateCustomValues(rs)

		rs.Update
		
        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = pstrMessage &  cstrdelimeter & "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
        End If
		If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstrMessage & cstrdelimeter & pstrprodID & " was successfully added."
            Else
                pstrMessage = pstrMessage & cstrdelimeter & "The changes to " & pstrprodID & " were successfully saved."
            End If
        Else
            pblnError = True
        End If

        rs.Close

		If len(pstrattrName) > 0 Then
			If len(plngattrID) > 0 Then
				sql = "Select * from sfAttributes where attrID=" & plngattrID
			Else
				sql = "Select * from sfAttributes where attrID=0"
			End If

			rs.CursorLocation = 3
			rs.open sql, cnn, 1, 3
			If rs.EOF Then
			    rs.AddNew
			    blnAdd = True
			    pblnAEInventoryChanged = True
			Else
			    blnAdd = False
			End If
			
			rs.Fields("attrName").Value = pstrattrName
			If Len(pstrattrURL_Field) > 0 Then Call setRSFieldValue(rs, pstrattrURL_Field, Trim(pstrattrURL))
			If Len(pstrattrDisplay_Field) > 0 Then Call setRSFieldValue(rs, pstrattrDisplay_Field, Trim(pstrattrDisplay))
			If Len(pstrattrExtra_Field) > 0 Then Call setRSFieldValue(rs, pstrattrExtra_Field, Trim(pstrattrExtra))
			If Len(pstrattrImage_Field) > 0 Then Call setRSFieldValue(rs, pstrattrImage_Field, Trim(pstrattrImage))
			If Len(pstrattrSKU_Field) > 0 Then Call setRSFieldValue(rs, pstrattrSKU_Field, Trim(pstrattrSKU))
			rs.Fields("attrProdId").Value = pstrProdId
			
			'Response.Write "<h3>pblnTextBasedAttribute: " & pblnTextBasedAttribute & "</h3>"
			If pblnTextBasedAttribute Then 
				rs.Fields("attrDisplayStyle").Value = pbytattrDisplayStyle
				If Err.number <> 0 Then
					If Err.number = 3265 Then
						pblnTextBasedAttribute = False
					Else
						debugprint "Error: " & Err.number,Err.Description
					End If
					Err.Clear
				End If
			End If

			rs.Update
		
			If Err.Number = 0 Then
			    If blnAdd Then
			        pstrMessage = pstrMessage & cstrdelimeter & pstrattrName & " was successfully added."
			    Else
			        pstrMessage = pstrMessage & cstrdelimeter & "The changes to " & pstrattrName & " were successfully saved."
			    End If
			Else
			    pblnError = True
			End If
			
			plngattrID = rs("attrID")
			rs.Close

			If len(pstrattrdtName) > 0 Then

				If len(plngattrdtID) > 0 Then
					sql = "Select * from sfAttributeDetail where attrdtID=" & plngattrdtID
				Else
					sql = "Select * from sfAttributeDetail where attrdtID=0"
				End If
				
				rs.open sql, cnn, 1, 3
				If rs.EOF Then
				    rs.AddNew
				    blnAdd = True
					pblnAEInventoryChanged = True
				Else
				    blnAdd = False
				End If
				
				rs.Fields("attrdtAttributeId").Value = plngattrID
				Call setRSFieldValue(rs, "attrdtName", pstrattrdtName)
				Call setRSFieldValue(rs, "attrdtPrice", pstrattrdtPrice)
				rs.Fields("attrdtType").Value = pintattrdtType
				If pblnPricingLevel Then Call setRSFieldValue(rs, "attrdtPLPrice", pstrattrdtPLPrice)
				If Len(pstrattrdtWeight_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtWeight_Field, wrapSQLValue(pdblattrdtWeight, False, enDatatype_NA))
				If Len(pstrattrdtImage_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtImage_Field, pstrattrdtImage)
				If Len(pstrattrdtDisplay_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtDisplay_Field, pstrattrdtDisplay)
				If Len(pstrattrdtExtra_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtExtra_Field, pstrattrdtExtra)
				If Len(pstrattrdtExtra1_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtExtra1_Field, pstrattrdtExtra1)
				If Len(pstrattrdtFileName_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtFileName_Field, pstrattrdtFileName)
				If Len(pstrattrdtURL_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtURL_Field, pstrattrdtURL)
				If Len(pstrattrdtSKU_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtSKU_Field, pstrattrdtSKU)
				If Len(pstrattrdtDefault_Field) > 0 Then Call setRSFieldValue(rs, pstrattrdtDefault_Field, ConvertToBoolean(pbytattrdtDefault, False))

				rs.Update
				
				If Err.Number = 0 Then
				    If blnAdd Then
				        pstrMessage = pstrMessage & cstrdelimeter & pstrattrdtName & " was successfully added."
				    Else
				        pstrMessage = pstrMessage & cstrdelimeter & "The changes to " & pstrattrdtName & " were successfully saved."
				    End If
				Else
				    pblnError = True
				End If
		
			End If

		End If
 
		Call UpdateAttributeCount(pstrOrigprodID)
        
        Dim parySortOrder,i
        
		'for the attribute categories
		If pblnAttributeCategoryOrderable Then
			parySortOrder = split(pbytattrDisplayOrder,",")
			For i=0 to ubound(parySortOrder)-1
				sql = "Update sfAttributes Set attrDisplayOrder=" & i & " where attrID=" & parySortOrder(i)
				cnn.Execute sql,,128
				If Err.number = -2147217904  Then	'No value given for one or more required parameters
					pblnAttributeCategoryOrderable = False
					Err.Clear
					Exit For
				ElseIf Err.number <> 0 Then
					Debugprint "Error: " & Err.number, Err.Description
					Err.Clear
				End If
			Next
		End If

		'for the attribute details
        parySortOrder = split(pintattrdtOrder,",")
        For i=0 to ubound(parySortOrder)-1
			sql = "Update sfAttributeDetail Set attrdtOrder=" & i & " where attrdtID=" & parySortOrder(i)
			cnn.Execute sql,,128
        Next

		If cblnSF5AE Then Call UpdateAE

    Else
        pblnError = True
    End If

    Update = (not pblnError)

End Function    'Update

'***********************************************************************************************

Public Sub OutputAttrValues()

Dim i,j
Dim pstrReplaceQuote
Dim pbytDisplayStyle

Dim pstrAttributeNames

On Error Resume Next

	If Not isObject(prsAttributes) Then Exit Sub
	
	With prsAttributes
		.filter = "attrProdId='" & pstrProdID & "'"
		if .RecordCount > 0 then 
			.MoveFirst
			pstrAttributeNames = "var aryAttributeNames = new Array("
			pstrAttributeNames = pstrAttributeNames & .Fields("attrID").Value & "," & jsOutputValue(.Fields("attrName").Value)		
			.MoveNext
			For i = 2 to .RecordCount
				pstrAttributeNames = pstrAttributeNames & "," & .Fields("attrID").Value & "," & jsOutputValue(.Fields("attrName").Value)		
				.MoveNext
			Next
			pstrAttributeNames = pstrAttributeNames & ");" & vbcrlf
			Response.Write pstrAttributeNames
			.MoveFirst

			'added for text based attributes
			Dim pblnError
			Dim pstrTextBasedAttributes
			Dim pstrTextAttrURL
			Dim pstrTextattrDisplay
			Dim pstrTextAttrExtra
			Dim pstrTextattrImage
			Dim pstrTextattrSKU
			pblnError = False
			
			If pblnTextBasedAttribute Then
				pstrTextBasedAttributes = "var aryAttributeDisplay = new Array();" & vbcrlf
				For i = 1 to .RecordCount
					If isNull(prsAttributes("attrDisplayStyle")) Then
						pbytDisplayStyle = 0
					Else
						pbytDisplayStyle = prsAttributes("attrDisplayStyle")
					End If
					
					If Len(pstrattrURL_Field) > 0 Then pstrTextAttrURL = jsEncodeApostrophe(Trim(prsAttributes.Fields(pstrattrURL_Field).Value))
					If Len(pstrattrDisplay_Field) > 0 Then pstrTextattrDisplay = jsEncodeApostrophe(Trim(prsAttributes.Fields(pstrattrDisplay_Field).Value))
					If Len(pstrattrExtra_Field) > 0 Then pstrTextAttrExtra = jsEncodeApostrophe(Trim(prsAttributes.Fields(pstrattrExtra_Field).Value))
					If Len(pstrattrImage_Field) > 0 Then pstrTextattrImage = jsEncodeApostrophe(Trim(prsAttributes.Fields(pstrattrImage_Field).Value))
					If Len(pstrattrSKU_Field) > 0 Then pstrTextattrSKU = jsEncodeApostrophe(Trim(prsAttributes.Fields(pstrattrSKU_Field).Value))

					pstrTextBasedAttributes = pstrTextBasedAttributes & "aryAttributeDisplay[" & prsAttributes.Fields("attrID").Value & "] = " _
											& "[" _
											& "'" & pbytDisplayStyle & "'," _
											& "'" & pstrTextattrDisplay & "'," _
											& "'" & pstrTextAttrExtra & "'," _
											& "'" & pstrTextattrImage & "'," _
											& "'" & pstrTextattrSKU & "'," _
											& "'" & pstrTextAttrURL & "'" _
											& "];" & vbcrlf
					
					If Err.number <> 0 Then
						If Err.number = 3265 Then
							pblnTextBasedAttribute = False
						Else
							debugprint "Error: " & Err.number,Err.Description
						End If
						pblnError = True
						Err.Clear
					End If
					.MoveNext
				Next
				.MoveFirst
				If Not pblnError Then Response.Write pstrTextBasedAttributes
			End If
			'end addition for text based attributes
			
			Response.Write "var aryAttributes = new Array();" & vbcrlf
			Response.Write vbcrlf
			For i = 1 to .RecordCount
				prsAttributeDetails.filter = "attrdtAttributeId=" & prsAttributes("attrID")
				If prsAttributeDetails.RecordCount > 0 then 
					prsAttributeDetails.MoveFirst
					If not prsAttributeDetails.EOF Then 
						Response.Write "aryAttributes[" & i & "] = ["
						Response.Write jsOutputValue(prsAttributeDetails("attrdtName")) & "," & prsAttributeDetails("attrdtID")
						prsAttributeDetails.MoveNext
						For j = 2 to prsAttributeDetails.RecordCount
							Response.Write "," & jsOutputValue(prsAttributeDetails("attrdtName")) & "," & prsAttributeDetails("attrdtID")
							prsAttributeDetails.MoveNext
						Next
						Response.Write "]" & vbcrlf
					End If
				End If
  				.MoveNext
			Next
		Else
			Response.Write "var aryAttributeNames = new Array();"
		End If
	End With

End Sub

Public Sub OutputAttrDetailValues()

Dim i,j,k
Dim paryCustomAttrDetailOut(11)

	If Not isObject(prsAttributes) Then Exit Sub

	Response.Write "var aryAttributeDetails = new Array();" & vbcrlf
	If pblnPricingLevel Then Response.Write "var aryAttributePLDetails = new Array();" & vbcrlf
	If prsAttributes.RecordCount > 0 then prsAttributes.MoveFirst
	Response.Write vbcrlf
	For i = 1 to prsAttributes.RecordCount
		prsAttributeDetails.filter = "attrdtAttributeId=" & prsAttributes("attrID")
		If prsAttributeDetails.RecordCount > 0 then 
			prsAttributeDetails.MoveFirst
			For j = 1 to prsAttributeDetails.RecordCount
				paryCustomAttrDetailOut(0) = jsOutputValue(getRSFieldValue(prsAttributeDetails, "attrdtName"))
				paryCustomAttrDetailOut(1) = getRSFieldValue(prsAttributeDetails, "attrdtPrice")
				paryCustomAttrDetailOut(2) = getRSFieldValue(prsAttributeDetails, "attrdtType")
				
				If Len(pstrattrdtWeight_Field) > 0 Then paryCustomAttrDetailOut(3) = jsOutputValue(getRSFieldValue(prsAttributeDetails, pstrattrdtWeight_Field))
				If Len(pstrattrdtImage_Field) > 0 Then paryCustomAttrDetailOut(4) = jsOutputValue(getRSFieldValue(prsAttributeDetails, pstrattrdtImage_Field))
				If Len(pstrattrdtDisplay_Field) > 0 Then paryCustomAttrDetailOut(5) = jsOutputValue(getRSFieldValue(prsAttributeDetails, pstrattrdtDisplay_Field))
				If Len(pstrattrdtFileName_Field) > 0 Then paryCustomAttrDetailOut(6) = jsOutputValue(getRSFieldValue(prsAttributeDetails, pstrattrdtFileName_Field))
				If Len(pstrattrdtSKU_Field) > 0 Then paryCustomAttrDetailOut(7) = jsOutputValue(getRSFieldValue(prsAttributeDetails, pstrattrdtSKU_Field))
				If Len(pstrattrdtURL_Field) > 0 Then paryCustomAttrDetailOut(8) = jsOutputValue(getRSFieldValue(prsAttributeDetails, pstrattrdtURL_Field))
				If Len(pstrattrdtDefault_Field) > 0 Then paryCustomAttrDetailOut(9) = getRSFieldValue(prsAttributeDetails, pstrattrdtDefault_Field)
				paryCustomAttrDetailOut(9) = LCase(CStr(ConvertToBoolean(paryCustomAttrDetailOut(9), False)))
				If Len(pstrattrdtExtra_Field) > 0 Then paryCustomAttrDetailOut(10) = jsOutputValue(getRSFieldValue(prsAttributeDetails, pstrattrdtExtra_Field))
				If Len(pstrattrdtExtra1_Field) > 0 Then paryCustomAttrDetailOut(11) = jsOutputValue(getRSFieldValue(prsAttributeDetails, pstrattrdtExtra1_Field))

				Response.Write "aryAttributeDetails[" & prsAttributeDetails("attrdtID") & "] = ["
				For k = 0 To UBound(paryCustomAttrDetailOut) - 1
					Response.Write paryCustomAttrDetailOut(k) & ","
				Next 'i
				Response.Write paryCustomAttrDetailOut(UBound(paryCustomAttrDetailOut)) & "]" & vbcrlf

				If pblnPricingLevel Then
					Response.Write "aryAttributePLDetails[" & prsAttributeDetails("attrdtID") & "] = [" _
																								   & chr(34) & prsAttributeDetails("attrdtPLPrice") & chr(34) _
																								   & "]" & vbcrlf
				End If
				Response.Write "" & vbcrlf
			prsAttributeDetails.MoveNext
			Next
		End If
		prsAttributes.MoveNext
	Next

End Sub

'***********************************************************************************************

Public Sub OutputSummary()

'On Error Resume Next

Dim i
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr
Dim pstrSelect, pstrHighlight
Dim pstrID
Dim pblnSelected, pblnHasDetails, pblnExpanded
Dim pstrCheckbox
Dim plngNumActiveColumns
Dim paryOutputColumns(9)	'two more than maryDisplayField due to checkboxes and attribute
Dim plngFieldCounter

Dim enSummary_HeaderTitle:	enSummary_HeaderTitle = 0
Dim enSummary_HeaderText:	enSummary_HeaderText = 1
Dim enSummary_HeaderAlign:	enSummary_HeaderAlign = 2
Dim enSummary_HeaderWidth:	enSummary_HeaderWidth = 3
Dim enSummary_HeaderSort:	enSummary_HeaderSort = 4
Dim enSummary_HeaderTitle_Asc:	enSummary_HeaderTitle_Asc = 5
Dim enSummary_HeaderTitle_Desc:	enSummary_HeaderTitle_Desc = 6
Dim enSummary_CellAlign:	enSummary_CellAlign = 7

Dim aSortHeader(7)
	aSortHeader(0) = Array("", "", "center", "", "", "", "", "center")				'Spacer for checkbox
	aSortHeader(1) = Array("", "", "center", "", "", "", "", "center")				'Spacer for attribute +/-
	aSortHeader(2) = Array("", "Product ID", "left", "", "prodID", "Sort Product IDs in descending order", "Sort Product IDs in ascending order", "left")
	aSortHeader(3) = Array("", "Product Name", "left", "", "prodName", "Sort Product Names in descending order", "Sort Product Names in ascending order", "left")
	aSortHeader(4) = Array("", "Price", "center", "", "prodPrice", "Sort Prices in descending order", "Sort Prices in ascending order", "center")
	aSortHeader(5) = Array("", "Sale Price", "center", "", "prodSalePrice", "Sort Sales Prices in descending order", "Sort Sale Prices in ascending order", "center")
	aSortHeader(6) = Array("", "Date Added", "center", "", "prodDateAdded", "Sort by date added in descending order", "Sort by date added in ascending order", "center")
	aSortHeader(7) = Array("", "Active", "center", "", "prodEnabledIsActive", "Sort Active Categories first", "Sort Inactive Categories first", "center")

	plngNumActiveColumns = 2
    For plngFieldCounter = 0 To UBound(maryDisplayField)
		If maryDisplayField(plngFieldCounter)(1) Then plngNumActiveColumns = plngNumActiveColumns + 1
    Next 'plngFieldCounter

	With Response

		.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke' id='tblSummary'>"
		.Write "<colgroup align='center'>"
		.Write "	<tr class='tblhdr'>"

	If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
		For i = 0 To UBound(aSortHeader)
			aSortHeader(i)(enSummary_HeaderTitle) = aSortHeader(i)(enSummary_HeaderTitle_Asc)
		Next 'i
		pstrSortOrder = "ASC"
	Else
		For i = 0 To UBound(aSortHeader)
			aSortHeader(i)(enSummary_HeaderTitle) = aSortHeader(i)(enSummary_HeaderTitle_Desc)
		Next 'i
		pstrSortOrder = "DESC"
	End If
	
	.Write "<TH align=" & aSortHeader(0)(enSummary_HeaderAlign) & "><input type='checkbox' name='chkCheckAll' id='chkCheckAll'  onclick='checkAll(this.form.chkProductID, this.checked);' value=''></TH>"
	.Write "<TH align=" & aSortHeader(1)(enSummary_HeaderAlign) & ">&nbsp;</TH>"
	if len(mstrOrderBy) > 0 Then
		pstrOrderBy = mstrOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
	for i = 2 to UBound(aSortHeader)
		If pstrOrderBy = aSortHeader(i)(enSummary_HeaderSort) Then
			If (pstrSortOrder = "ASC") Then
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn('" & aSortHeader(i)(enSummary_HeaderSort) & "','DESC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i)(enSummary_HeaderTitle) & "' align=" & aSortHeader(i)(enSummary_HeaderAlign) & ">" & aSortHeader(i)(enSummary_HeaderText) & _
								"&nbsp;&nbsp;&nbsp;<img src='images/up.gif' border=0 align=bottom></TH>" & vbCrLf
			Else
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn('" & aSortHeader(i)(enSummary_HeaderSort) & "','ASC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i)(enSummary_HeaderTitle) & "' align=" & aSortHeader(i)(enSummary_HeaderAlign) & ">" & aSortHeader(i)(enSummary_HeaderText) & _
								"&nbsp;&nbsp;&nbsp;<img src='images/down.gif' border=0 align=bottom></TH>" & vbCrLf
			End If
		Else
		    .Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn('" & aSortHeader(i)(enSummary_HeaderSort) & "','" & pstrSortOrder & "');" & chr(34) & _
							" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
							" title='" & aSortHeader(i)(enSummary_HeaderTitle) & "' align=" & aSortHeader(i)(enSummary_HeaderAlign) & ">" & aSortHeader(i)(enSummary_HeaderText) & "</TH>" & vbCrLf
		End If
	next 'i

		.Write "	</tr>" & vbcrlf

    If prsProducts.RecordCount > 0 Then
        prsProducts.MoveFirst

'Need to calculate current recordset page and upper bound to loop through
dim plnguBound, plnglbound, pstrDisplay

        For i = 1 To prsProducts.RecordCount
        
			pstrID = trim(prsProducts("prodID"))
			pstrURL = "sfProductAdmin.asp?Action=View&prodID=" & pstrID
			pstrTitle = "Click to view " & Replace(prsProducts("prodName") & "","'","")
			pstrHighlight = "title='" & pstrTitle & "' " _
						  & "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
						  & "onmouseout='doMouseOutRow(this); ClearTitle();' "
			pstrSelect = "title='" & pstrTitle & "' " _
					   & "onmouseover='DisplayTitle(this);' " _
					   & "onmouseout='ClearTitle();' " _
					   & "onmousedown='ViewProduct(" & chr(34) & pstrID & chr(34) & ");'"
			
			pblnSelected = (pstrID = pstrprodID)
			
			If mblnExpandAttributesAutomatically Then
				pblnExpanded = pblnSelected
			Else
				pblnExpanded = False
			End If

			prsAttributes.Filter = "attrProdId='" & SQLSafe(pstrID) & "'"
			pblnHasDetails = (Not prsAttributes.EOF)

			pstrCheckbox = "<input type=checkbox name=chkProductID id=chkProductID value=" & Chr(34) & Server.HTMLEncode(pstrID) & Chr(34) & ">&nbsp;"

            If pblnSelected Then
                '.Write "<TR class='Selected' onmouseover='doMouseOverRow(this);' onmouseout='doMouseOutRow(this);'>"
				'.Write " <TD>&nbsp;</TD>"
                .Write "<TR class='Selected' " & pstrHighlight & " id='selectedSummaryItem'>"
            Else
				if ConvertBoolean(prsProducts("prodEnabledIsActive")) then
					.Write " <TR class='Active' " & pstrHighlight & ">"
				else
					.Write " <TR class='Inactive' " & pstrHighlight & ">"
        		end if
            End If
            
			paryOutputColumns(0) = "<TD align=" & aSortHeader(0)(enSummary_CellAlign) & ">" & pstrCheckbox & "</TD>"
            
			if pblnHasDetails then
				If pblnExpanded Then
					paryOutputColumns(1) = " <TD onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Hide attributes' onclick='return ExpandProduct(this," & chr(34) & pstrID & chr(34) & ");' align=" & aSortHeader(1)(enSummary_CellAlign) & ">-</TD>"
				Else
					paryOutputColumns(1) = " <TD onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Show attributes' onclick='return ExpandProduct(this," & chr(34) & pstrID & chr(34) & ");' align=" & aSortHeader(1)(enSummary_CellAlign) & ">+</TD>"
				End If
			else
				paryOutputColumns(1) = " <TD " & pstrSelect & ">&nbsp;</TD>"
       		end if
       		
        	paryOutputColumns(GetDisplayFieldIndex("prodID") + 2) = "<TD " & pstrSelect & " align=" & aSortHeader(GetDisplayFieldIndex("prodID") + 2)(enSummary_CellAlign) & "><a onclick='return false;' href='" & pstrURL & "' " & pstrSelect & ">" & prsProducts("prodID") & "</a></TD>"
        	paryOutputColumns(GetDisplayFieldIndex("prodName") + 2) = "<TD " & pstrSelect & " align=" & aSortHeader(GetDisplayFieldIndex("prodName") + 2)(enSummary_CellAlign) & ">" & prsProducts("prodName") & "&nbsp;</TD>"
        	paryOutputColumns(GetDisplayFieldIndex("prodPrice") + 2) = "<TD " & pstrSelect & " align=" & aSortHeader(GetDisplayFieldIndex("prodPrice") + 2)(enSummary_CellAlign) & ">" & WriteCurrency(prsProducts("prodPrice")) & "</TD>"

			if ConvertBoolean(prsProducts("prodSaleIsActive")) then
	       		paryOutputColumns(GetDisplayFieldIndex("prodSalePrice") + 2) = "<TD align=" & aSortHeader(GetDisplayFieldIndex("prodSalePrice") + 2)(enSummary_CellAlign) & ">" & WriteCurrency(prsProducts("prodSalePrice")) & "</TD>"
			else
	       		paryOutputColumns(GetDisplayFieldIndex("prodSalePrice") + 2) = "<TD align=" & aSortHeader(GetDisplayFieldIndex("prodSalePrice") + 2)(enSummary_CellAlign) & ">-</TD>"
	       	end if

        	paryOutputColumns(GetDisplayFieldIndex("prodDateAdded") + 2) = "<TD " & pstrSelect & " align=" & aSortHeader(GetDisplayFieldIndex("prodDateAdded") + 2)(enSummary_CellAlign) & ">" & prsProducts("prodDateAdded") & "&nbsp;</TD>"
            
			if ConvertBoolean(prsProducts("prodEnabledIsActive")) then
	       		paryOutputColumns(GetDisplayFieldIndex("prodEnabledIsActive") + 2) = "<TD align=" & aSortHeader(GetDisplayFieldIndex("prodEnabledIsActive") + 2)(enSummary_CellAlign) & "><a href='sfProductAdmin.asp?Action=Deactivate&prodID=" & prsProducts("prodID") _
																				   & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to deactivate " & pstrTitle & "'>Active</a></TD>"
			else
	       		paryOutputColumns(GetDisplayFieldIndex("prodEnabledIsActive") + 2) = "<TD align=" & aSortHeader(GetDisplayFieldIndex("prodEnabledIsActive") + 2)(enSummary_CellAlign) & "><a href='sfProductAdmin.asp?Action=Activate&prodID=" & prsProducts("prodID") _
																				   & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to activate " & pstrTitle & "'>Inactive</a></TD>"
	       	end if

            For plngFieldCounter = 0 To 1
				Response.Write paryOutputColumns(plngFieldCounter)
            Next 'plngFieldCounter
            
            For plngFieldCounter = 2 To UBound(maryDisplayField) + 2
				If maryDisplayField(CLng(plngFieldCounter-2))(1) Then Response.Write paryOutputColumns(plngFieldCounter)
            Next 'plngFieldCounter
        	.Write "</TR>" & vbcrlf
            
        	.Write "</TR>" & vbcrlf
        	
			If pblnHasDetails Then 
	        	.Write "<TR><td colspan=1></td><TD COLSPAN=" & CStr(plngNumActiveColumns - 1)& ">"
				Call OutputAttr(pstrID,pblnExpanded,pblnSelected)
        		.Write "</TD></TR>" & vbcrlf
			End If

            prsProducts.MoveNext
        Next
    Else
			.Write "<TR><TD align=center COLSPAN=" & CStr(plngNumActiveColumns)& "><h3>There are no Products</h3></TD></TR>"
    End If
    
		.Write "<tr class='tblhdr'><TH COLSPAN=" & CStr(plngNumActiveColumns)& " align=center>"
		If plngRecordCount = 0 Then
			.Write "No Products match your search criteria"
		Elseif plngRecordCount = 1 Then
			.Write "1 product matches your search criteria"
		Else 
			.Write plngRecordCount & " products match your search criteria<br>"

		dim pstrCheck
		pstrCheck = "return isInteger(this, true, ""Please enter a positive integer for the recordset page size."");"
		.Write "Show&nbsp;<input type='Text' id='PageSize' name='PageSize' value='" & prsProducts.PageSize & "' maxlength='4' size='4' style='text-align:center;' onblur='" & pstrCheck & "'>&nbsp;records at a time.&nbsp;&nbsp;"

		If mlngPageCount > 1 Then
			If True Then
				Response.Write "&nbsp;Goto&nbsp;<select name=pageSelect id=pageSelect onchange='return ViewPage(this.selectedIndex+1);'>"
				For i=1 to mlngPageCount
					plnglbound = (i-1) * mlngMaxRecords + 1
					plnguBound = i * mlngMaxRecords
					if plnguBound > plngRecordCount Then plnguBound = plngRecordCount
					'pstrDisplay = plnglbound & " - " & plnguBound & "&nbsp;"
					'If i = cInt(mlngAbsolutePage) Then
					'	Response.Write pstrDisplay
					'Else
					'	Response.Write "<a href='#' onclick='return ViewPage(" & i & ");'>" & pstrDisplay & "</a>&nbsp;"
					'End If
					Response.Write "<option " & isSelected(i = cInt(mlngAbsolutePage)) & ">" & "Page " & i & " (" & plnglbound & " - " & plnguBound & ")</option>"
				Next
				Response.Write "</select>"
			Else
				Response.Write "&nbsp;Page " & mlngAbsolutePage & " of " & mlngPageCount & " pages total. Goto&nbsp;page<input name=pageSelect id=pageSelect size=6 onchange='var plngDesiredPage=this.value; if (plngDesiredPage<=" & mlngPageCount & "){return ViewPage(plngDesiredPage);}else{return false;}'>"
			End If
		End If	'mlngPageCount > 1
		End If
		.Write "</TH></TR>"
		.Write "</TABLE>"
	End With
End Sub      'OutputSummary

Private Sub OutputAttr(prodID,blnExpand,blnSelected)

'On Error Resume Next

Dim i
Dim pstrTitle, pstrURL, pstrAbbr
Dim pstrSelect, pstrHighlight
Dim plngID
Dim pblnHasDetails

	If prsAttributes.RecordCount > 0 Then
	  With Response
		If blnExpand Then
			.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				& "bgcolor='whitesmoke' id='tbl" & prodID & "'>"
		Else
			.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				& "bgcolor='whitesmoke' style='display: none;' id='tbl" & prodID & "'>"
		End If
		.Write "<colgroup align='left' width='3%'>" & vbcrlf
		.Write "<colgroup align='left' width='3%'>" & vbcrlf
		.Write "<colgroup align='left' width='94%'>" & vbcrlf

		For i = 1 to prsAttributes.RecordCount
		
			plngID = prsAttributes("attrID")
			pstrURL = "sfProductAdmin.asp?Action=View&attrID=" & plngID
			pstrTitle = "Click to view " & Replace(prsAttributes("attrName") & "","'","")

			pstrHighlight = "title='" & pstrTitle & "' " _
						  & "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
						  & "onmouseout='doMouseOutRow(this); ClearTitle();'"
			pstrSelect = "title='" & pstrTitle & "' " _
					   & "onmouseover='DisplayTitle(this);' " _
					   & "onmouseout='ClearTitle();' " _
					   & "onmousedown='ViewAttribute(" & plngID & ");'"

			prsAttributeDetails.Filter = "attrdtAttributeId=" & plngID
			pblnHasDetails = (Not prsAttributeDetails.EOF)
			
			If blnSelected Then
                .Write " <TR onmouseover='doMouseOverRow(this);' onmouseout='doMouseOutRow(this);' onmousedown=" & chr(34) & "SelectAttr(" & plngID & ");" & chr(34) & ">"
				.Write " <TD>&nbsp;</TD>" & vbcrlf
			Else
				.Write " <TR " & pstrHighlight & ">" & vbcrlf
				.Write " <TD " & pstrSelect & ">&nbsp;</TD>" & vbcrlf
			End If

			If pblnHasDetails Then
				If blnExpand Then
					.Write " <TD onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Hide attributes' onclick='return ExpandAttr(this," & plngID & ");'>-</TD>" & vbcrlf
				Else
					.Write " <TD onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Show attributes' onclick='return ExpandAttr(this," & plngID & ");'>+</TD>" & vbcrlf
				End If
			Else
				If blnSelected Then
					.Write " <TD>&nbsp;</TD>" & vbcrlf
				Else
					.Write " <TD " & pstrSelect & ">&nbsp;</TD>" & vbcrlf
				End If
			End If
			
			If blnSelected Then
       			.Write "<TD>" & prsAttributes("attrName") & "</span></TD>" & vbcrlf
			Else
       			.Write "<TD " & pstrSelect & "><a href='" & pstrURL & "' " & pstrSelect & ">" & prsAttributes("attrName") & "</a></TD>" & vbcrlf
			End If

			.Write "</TR>" & vbcrlf
			
			If pblnHasDetails Then 
	        	.Write "<TR><TD COLSPAN=3>"
				Call OutputDetail(plngID,blnExpand,blnSelected)
        		.Write "</TD></TR>"
			End If
			prsAttributes.MoveNext
		Next
		.Write "</TABLE>"
	  End With
    End If
	
End Sub      'OutputAttr

Private Sub OutputDetail(attrID,blnExpand,blnSelected)

'On Error Resume Next

Dim i
Dim pstrTitle, pstrURL, pstrAbbr
Dim plngID, pstrSelect, pstrHighlight

	If prsAttributeDetails.RecordCount > 0 Then
	  With Response
		If blnExpand Then
			.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				& "bgcolor='whitesmoke' style='cursor:hand;' id='tblAttDetail" & attrID & "'" _
				& ">"
		Else
			.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				& "bgcolor='whitesmoke' style='display: none; cursor:hand;' id='tblAttDetail" & attrID & "'" _
				& ">"
		End If

		.Write "<colgroup align='left' width='6%'>"
		.Write "<colgroup align='left' width='52%'>"
		.Write "<colgroup align='right' width='13%'>"
		.Write "<colgroup align='left' width='29%'>"

		For i = 1 to prsAttributeDetails.RecordCount
			plngID = prsAttributeDetails("attrdtID")
			pstrURL = "sfProductAdmin.asp?Action=View&attrdtID=" & plngID
			pstrTitle = "Click to view " & prsAttributeDetails("attrdtName")
		
			pstrHighlight = "title='" & pstrTitle & "' " _
						  & "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
						  & "onmouseout='doMouseOutRow(this); ClearTitle();'"
			pstrSelect = "title='" & pstrTitle & "' " _
					   & "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
					   & "onmouseout='doMouseOutRow(this); ClearTitle();'" _
					   & "onmousedown='ViewAttrDetail(" & plngID & ");'"

			If blnSelected Then
				.Write " <TR " & pstrHighlight & " onmousedown=" & chr(34) & "SelectAttrDetail(" & attrID & "," & plngID & ")" & chr(34) & ">"
				.Write " <TD>&nbsp;</TD>"
				.Write " <TD>&nbsp;&nbsp;&nbsp;" & prsAttributeDetails("attrdtName") & "</TD>"
			Else
				.Write " <TR " & pstrSelect & ">"
				.Write " <TD>&nbsp;</TD>"
       			.Write " <TD>&nbsp;&nbsp;&nbsp;<a href='" & pstrURL & "' " & pstrSelect & ">" & prsAttributeDetails("attrdtName") & "&nbsp;</a></TD>"
			End If
			
			Select Case prsAttributeDetails("attrdtType")
				Case 1:	.Write " <TD>+ " & WriteCurrency(prsAttributeDetails("attrdtPrice")) & "&nbsp;</TD>"
				Case 2:	.Write " <TD>- " & WriteCurrency(prsAttributeDetails("attrdtPrice")) & "&nbsp;</TD>"
				Case 0:	.Write " <TD>&nbsp; " & WriteCurrency(prsAttributeDetails("attrdtPrice")) & "&nbsp;</TD>"
			End Select
			.Write " <TD>&nbsp;</TD>"
			.Write " </TR>"
			
			prsAttributeDetails.MoveNext
		Next
		.Write " </TABLE>"
	  End With
	End If
	
End Sub      'OutputDetail

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    If Len(pstrprodID) = 0 Then strError = strError & "Please enter a Product ID." & cstrDelimeter
    If Not IsNumeric(pstrprodPrice) Then strError = strError & "Please enter a number for the product price." & cstrDelimeter
    If Not IsNumeric(pstrprodHandlingFee) Then strError = strError & "Please enter a number for the product Handling Fee." & cstrDelimeter
    If Not IsNumeric(pstrprodSetupFee) Then strError = strError & "Please enter a number for the product Setup Fee." & cstrDelimeter
    If Not IsNumeric(pstrprodSetupFeeOneTime) Then strError = strError & "Please enter a number for the product Setup Fee." & cstrDelimeter

    If Not IsNumeric(pstrprodSalePrice) And Len(pstrprodSalePrice)<>0 Then strError = strError & "Please enter a number for the sale price." & cstrDelimeter
    If Not IsNumeric(pstrprodShip) And Len(pstrprodShip)<>0 Then strError = strError & "Please enter a number for the product based shipping charge." & cstrDelimeter
    If Not IsNumeric(pdblprodFixedShippingCharge) And Len(pdblprodFixedShippingCharge)<>0 Then strError = strError & "Please enter a number for the fixed shipping charge." & cstrDelimeter
    If Not IsNumeric(pdblprodHeight) And Len(pdblprodHeight)<>0 Then strError = strError & "Please enter a number for the product height." & cstrDelimeter
    If Not IsNumeric(pdblprodWeight) And Len(pdblprodWeight)<>0 Then strError = strError & "Please enter a number for the product weight." & cstrDelimeter
    If Not IsNumeric(pdblprodWidth) And Len(pdblprodWidth)<>0 Then strError = strError & "Please enter a number for the product width." & cstrDelimeter
    If Not IsNumeric(pdblprodLength) And Len(pdblprodLength)<>0 Then strError = strError & "Please enter a number for the product length." & cstrDelimeter

	If len(pstrattrName) > 0 Then
		pstrattrProdId = pstrProdID
	End If

	If len(pstrattrdtName) > 0 and len(pstrattrProdId) > 0 Then
		If Not IsNumeric(pstrattrdtPrice) Then pstrattrdtPrice = "0"
		If (pstrattrdtPrice = "0") or (len(pintattrdtType) = 0) Then pintattrdtType = 0
		If len(pintattrdtOrder) = 0 Then pintattrdtOrder = 0
	End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues

'******************************************************************************************************************************************************************

Function ConvertBoolean(vntValue)

	If Len(Trim(vntValue & "")) = 0 Then
		ConvertBoolean = False
	Else
		On Error Resume Next
		ConvertBoolean = cBool(vntValue)
		If Err.number <> 0 Then 
			ConvertBoolean = False
			Err.Clear
		End If
	End If

End Function	'ConvertBoolean

'******************************************************************************************************************************************************************

End Class   'clsProduct
%>