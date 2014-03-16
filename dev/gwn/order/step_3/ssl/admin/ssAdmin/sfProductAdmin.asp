<%Option Explicit
'********************************************************************************
'*   Product Manager Version SF 5.0 					                        *
'*   Release Version   2.0.0.13	Beta                                            *
'*   Release Date      November 2, 2002											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Dim mlngPageCount
Dim mlngAbsolutePage
Dim mlngMaxRecords
Dim mlngShortDescriptionLength
Dim mblnShowAttributes
Const mblnIsAEPM = False

mlngMaxRecords = LoadRequestValue("PageSize")

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Const mblnShowTabs = True			'Set this to true to show tabs
Const mblnAutoShowTable = True		'Set to true to automatically display database summary
Const mbytSummaryTableHeight = 300	'Summary Table Height
mblnAutoShowDetailInWindow = True			'Set this to true to show tabs
mblnShowAttributes = False		'Set to true to show attributes in summary table, false not to. Large databases should set this value to false
Server.ScriptTimeout = 300			'in seconds. Adjust for large databases or if some products have a lot of attributes. Server Default is usually 90 seconds
mlngShortDescriptionLength = ""	'by default this field is limited to 255 characters. If you desire, you can set it higher; "" is unlimited - this will required a change to the sfProducts table

If len(mlngMaxRecords) = 0 Then mlngMaxRecords = 50	'Set your default Maximum Records to show in summary table

'added for pricing levels
Const mblnBasePrice = False			'Show pricing levels for the regular price
Const mblnSalePrice = False			'Show pricing levels for the sale price
Const mblnAttrPrice = False			'Show pricing levels for the attribute price
Const mblnMTPrice = False			'Show pricing levels for the multi-tier price

'/
'/////////////////////////////////////////////////

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
Private pstrprodFileName
Private pdblprodHeight
Private pstrprodImageLargePath
Private pstrprodImageSmallPath
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
Private pblnprodShipIsActive
Private pstrprodShortDescription
Private pblnprodStateTaxIsActive
Private plngprodVendorId
Private pdblprodWeight
Private pdblprodWidth

Private plngattrID
Private pstrattrProdId
Private pstrattrName
Private pbytattrDisplayOrder
Private pbytattrDisplayStyle

Private plngattrdtID
Private plngattrdtAttributeId
Private pstrattrdtName
Private pstrattrdtPrice
Private pintattrdtType
Private pintattrdtOrder

Private pblngwActivate
Private pstrgwPrice

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

'added for pricing levels
Private pstrprodPLPrice
Private pstrprodPLSalePrice
Private pstrattrdtPLPrice
Private pstrgwPLPrice

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
    cstrDelimeter  = ";"
	pblnTextBasedAttribute = True
	pblnPricingLevel = False
	pblnAttributeCategoryOrderable = True	'you must update the sfAttributes table to use this feature
	pblnCustomMTP = False
	Call InitializeCustomValues
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

Private Sub InitializeCustomValues

	Exit Sub
	ReDim paryCustomValues(6)
	
	'format: Display Text, field name, field value(must be ""), DisplayType, DisplayLength, sqlSource
	paryCustomValues(0) = Array("Version","version","","")
	paryCustomValues(1) = Array("Release Date","releaseDate","","")
	paryCustomValues(2) = Array("Installation Hours","InstallationHours","","")
	paryCustomValues(3) = Array("My Product","MyProduct","","checkbox")
	paryCustomValues(4) = Array("Installation Required","InstallationRequired","","checkbox")
	paryCustomValues(5) = Array("Include In Search","IncludeInSearch","","checkbox")
	paryCustomValues(6) = Array("Include In Random Product","IncludeInRandomProduct","","checkbox")
	

End Sub	'InitializeCustomValues

Private Sub LoadCustomValues(objRS)

Dim i

	If pblnCustomMTP Then
		'If len(pstrprodID) > 0 Then	Set prsMTP = GetRS("Select * from ssPricingLevels Where prodID='" & SQLSafe(pstrprodID) & "' Order By PricingLevel Asc")
		If len(pstrprodID) > 0 Then	Set prsMTP = GetRS("Select * from PricingLevels Where prodID='" & SQLSafe(pstrprodID) & "' Order By PricingLevel Asc")
	End If

	If Not isArray(paryCustomValues) Then Exit Sub
	For i = 0 To UBound(paryCustomValues)
		paryCustomValues(i)(2) = objRS.Fields(paryCustomValues(i)(1)).Value
	Next 'i
	
End Sub	'LoadCustomValues

Private Sub UpdateCustomValues(objRS)

Dim i

	If Not isArray(paryCustomValues) Then Exit Sub
	For i = 0 To UBound(paryCustomValues)
		objRS.Fields(paryCustomValues(i)(1)).Value = paryCustomValues(i)(2)
	Next 'i

	If pblnCustomMTP Then Call UpdateCustomMTP
	
End Sub	'LoadCustomValues

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

    aError = Split(pstrMessage, cstrDelimeter)
    For i = 0 To UBound(aError)
        If pblnError Then
            Response.Write "<P align='center'><H4><FONT color=Red>" & aError(i) & "</FONT></H4></P>"
        Else
            Response.Write "<P align='center'><H4>" & aError(i) & "</H4></P>"
        End If
    Next 'i

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

Public Property Get prodFileName()
    prodFileName = pstrprodFileName
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

Public Property Get prodShip()
    prodShip = pstrprodShip
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

'***********************************************************************************************

Private Sub ClearValues()


End Sub 'ClearValues

Public Sub LoadProductValues

	If prsProducts.EOF Then Exit Sub
    plngprodAttrNum = trim(prsProducts("prodAttrNum").Value)
    plngprodCategoryId = trim(prsProducts("prodCategoryId").Value)
    pblnprodCountryTaxIsActive = ConvertBoolean(prsProducts("prodCountryTaxIsActive").Value)
    pdtprodDateAdded = trim(prsProducts("prodDateAdded").Value)
    pdtprodDateModified = trim(prsProducts("prodDateModified").Value)
    pstrprodDescription = trim(prsProducts("prodDescription").Value)
    pblnprodEnabledIsActive = ConvertBoolean(prsProducts("prodEnabledIsActive").Value)
    pstrprodFileName = trim(prsProducts("prodFileName").Value)
    pdblprodHeight = trim(prsProducts("prodHeight").Value)
    pstrprodID = trim(prsProducts("prodID").Value)
    pstrprodImageLargePath = trim(prsProducts("prodImageLargePath").Value)
    pstrprodImageSmallPath = trim(prsProducts("prodImageSmallPath").Value)
    pdblprodLength = trim(prsProducts("prodLength").Value)
    pstrprodLink = trim(prsProducts("prodLink").Value)
    plngprodManufacturerId = trim(prsProducts("prodManufacturerId").Value)
    pstrprodMessage = trim(prsProducts("prodMessage").Value)
    pstrprodName = trim(prsProducts("prodName").Value)
    pstrprodNamePlural = trim(prsProducts("prodNamePlural").Value)
    pstrprodPrice = trim(prsProducts("prodPrice").Value)
    pblnprodSaleIsActive = ConvertBoolean(prsProducts("prodSaleIsActive").Value)
    pstrprodSalePrice = trim(prsProducts("prodSalePrice").Value)
    pstrprodShip = trim(prsProducts("prodShip").Value)
    pblnprodShipIsActive = ConvertBoolean(prsProducts("prodShipIsActive").Value)
    pstrprodShortDescription = trim(prsProducts("prodShortDescription").Value)
    pblnprodStateTaxIsActive = ConvertBoolean(prsProducts("prodStateTaxIsActive").Value)
    plngprodVendorId = trim(prsProducts("prodVendorId").Value)
    pdblprodWeight = trim(prsProducts("prodWeight").Value)
    pdblprodWidth = trim(prsProducts("prodWidth").Value)
    
    If pblnPricingLevel Then
		pstrprodPLPrice = trim(prsProducts("prodPLPrice").Value)
		pstrprodPLSalePrice = trim(prsProducts("prodPLSalePrice").Value)
	End If

    Call LoadCustomValues(prsProducts)

End Sub 'LoadProductValues

Private Sub LoadAttributeValues

On Error Resume Next

	If prsAttributes.EOF Then Exit Sub
    plngattrID = trim(prsAttributes("attrID").Value)
    pstrattrName = trim(prsAttributes("attrName").Value)
    pstrattrProdId = trim(prsAttributes("attrProdId").Value)
	
    If pblnTextBasedAttribute Then
		pbytattrDisplayStyle = trim(prsAttributes("attrDisplayStyle").Value)
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
		pbytattrDisplayOrder = trim(prsAttributes("attrDisplayOrder").Value)
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
    plngattrdtID = trim(prsAttributeDetails("attrdtID").Value)
    plngattrdtAttributeId = trim(prsAttributeDetails("attrdtAttributeId").Value)
    pstrattrdtName = trim(prsAttributeDetails("attrdtName").Value)
    pintattrdtOrder = trim(prsAttributeDetails("attrdtOrder").Value)
    pstrattrdtPrice = trim(prsAttributeDetails("attrdtPrice").Value)
    pintattrdtType = trim(prsAttributeDetails("attrdtType").Value)

    If pblnPricingLevel Then
		pstrattrdtPLPrice = trim(prsAttributeDetails("attrdtPLPrice").Value)
	End If

End Sub 'LoadAttrDetailValues

Private Sub LoadFromRequest

On Error Goto 0

    With Request.Form
        plngprodAttrNum = Trim(.Item("prodAttrNum"))
        plngprodCategoryId = Trim(.Item("prodCategoryId"))
        pblnprodCountryTaxIsActive = (lCase(.Item("prodCountryTaxIsActive")) = "on")
        pstrprodDescription = Trim(.Item("prodDescription"))
        pblnprodEnabledIsActive = (lCase(.Item("prodEnabledIsActive")) = "on")
        pstrprodFileName = Trim(.Item("prodFileName"))
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
        pblnprodShipIsActive = (lCase(.Item("prodShipIsActive")) = "on")
        
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
        pstrattrProdId = Trim(.Item("attrProdId"))
		pbytattrDisplayStyle = trim(.Item("attrDisplayStyle"))
		pbytattrDisplayOrder = trim(.Item("attrDisplayOrder"))

        plngattrdtAttributeId = Trim(.Item("attrdtAttributeId"))
        plngattrdtID = Trim(.Item("attrdtID"))
        pstrattrdtName = Trim(.Item("attrdtName"))
        pintattrdtOrder = Trim(.Item("attrdtOrder"))
        pstrattrdtPrice = Trim(.Item("attrdtPrice"))
        pintattrdtType = Trim(.Item("attrdtType"))
    
		pblnAEInventoryChanged = ConvertBoolean(.Item("ChangeInventory"))
		pblnAEMTPChanged = ConvertBoolean(.Item("ChangeMTP"))
		pblnAECategoryChanged = ConvertBoolean(.Item("ChangeCategory"))
		
		If pblnPricingLevel Then
			pstrprodPLPrice = Replace(Replace(Trim(.Item("prodPLPrice")),",",";")," ","")
			pstrprodPLSalePrice = Replace(Replace(Trim(.Item("prodPLSalePrice")),",",";")," ","")
			pstrattrdtPLPrice = Replace(Replace(Trim(.Item("attrdtPLPrice")),",",";")," ","")
		End If

		Call LoadCustomValuesFromRequest

    End With

End Sub 'LoadFromRequest

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
            If Not .EOF Then LoadProductValues
        End If
    End With

End Function    'FindProduct

Public Function FindAttribute(lngID)

'On Error Resume Next

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
			Dim paryTemp(2)
			Dim plngCount
			Dim paryMTP
		
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

	If Not mblnIsAEPM Then Exit Function
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
	Set prsAEProductCategories = GetRS("Select subcatDetailID,subcatCategoryId from sfSubCatDetail Where ProdID='" & SQLSafe(pstrprodID) & "'")

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

Dim sql
Dim prsAEProductCategories
Dim i,j,k
Dim paryCategories()
Dim plngCatCount
Dim pblnBottom
Dim pdicCategories

	sql = "SELECT sfCategories.catID, sfSub_Categories.subcatID, sfCategories.catName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName, sfSub_Categories.Depth, sfSub_Categories.bottom" _
		& " FROM sfSub_Categories RIGHT JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
		& " ORDER BY sfCategories.catName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName"
	'debugprint "sql",sql
	
	Set prsAEProductCategories = GetRS(sql)
	With prsAEProductCategories
'		.Fields("subcatID").Properties("Optimize") = True
		.Filter = "bottom=1 or bottom=Null"
'		.Filter = "bottom=1"
		plngCatCount = .RecordCount-1
		.Filter = ""
		k = 0
		If plngCatCount > 0 Then
			redim paryCategories(plngCatCount,7)
			Set pdicCategories = Server.CreateObject("SCRIPTING.DICTIONARY")
			For i=1 to .RecordCount
				If Len(.Fields("subcatID").value & "") > 0 Then pdicCategories.Add CStr(.Fields("subcatID").value), .Fields("subcatName").value
				pblnBottom = .Fields("bottom").value = 1 OR isNull(.Fields("bottom").value)
'				pblnBottom = .Fields("bottom").value = 1
				If pblnBottom Then
					paryCategories(k,0) = .Fields("catID").value
					paryCategories(k,1) = .Fields("subcatID").value
					paryCategories(k,2) = .Fields("catName").value
					paryCategories(k,3) = .Fields("subcatName").value
					paryCategories(k,4) = .Fields("CatHierarchy").value
					paryCategories(k,5) = .Fields("Depth").value
					paryCategories(k,6) = pblnBottom
					k = k + 1
				End If
				.MoveNext
			Next
			.MoveFirst

			'this creates the first dropdown
			Dim pstrTemp
			Dim pstrCatValue
			Dim pstrCatName
			Dim pbytCatDepth
			Dim pstrsubCatValue
			Dim pstrsubCatName
			Dim pstrSelected

			If Len(mbytCategoryFilter) = 0 And Len(mbytsubCategoryFilter) = 0 Then pstrSelected = "selected"
			pstrTemp = "<select id='CategoryFilter' name='CategoryFilter' size='10' multiple>" & vbcrlf _
					 & "  <option value='' " & pstrSelected & ">- All -</Option>" & vbcrlf

			For i=1 to .RecordCount

				If len(.Fields("Depth").value & "") > 0 Then 'Then No sub-categories so don't display
					pbytCatDepth = .Fields("Depth").value
					
					If pstrCatName <> Trim(.Fields("catName").value) Then
						pstrCatValue = .Fields("catID").value & "."
						pstrCatName = .Fields("catName").value
						pstrSelected = ""
						If pstrCatValue = CStr(mbytCategoryFilter & "." & mbytsubCategoryFilter) Then pstrSelected = "selected"
						pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrCatValue & chr(34) & pstrSelected & ">" & pstrCatName & "</option>" & vbcrlf 
						
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & .Fields("subcatID").value
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & .Fields("subcatName").value
							pstrSelected = ""
							If pstrsubCatValue = CStr(mbytCategoryFilter & "." & mbytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrsubCatValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					Else
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & .Fields("subcatID").value
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & .Fields("subcatName").value
							pstrSelected = ""
							If pstrsubCatValue = CStr(mbytCategoryFilter & "." & mbytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrsubCatValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					End If
				End If
				
				.MoveNext
			Next
			pstrTemp = pstrTemp & "</select>" & vbcrlf
			.MoveFirst

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

	If len(pstrprodID) > 0 Then
		Set prsAEProductCategories = GetRS("Select subcatDetailID,subcatCategoryId from sfSubCatDetail Where ProdID='" & SQLSafe(pstrprodID) & "'")
		With prsAEProductCategories
			For i=1 to .RecordCount
				pstrProdDictionaryList = pstrProdDictionaryList & "mdicCategory.add (" & chr(34) & "" & .Fields("subcatCategoryId").value & chr(34) & "," & chr(34) & "" & chr(34) & ");" & vbcrlf
				.MoveNext
			Next
		End With
	End If
	ReleaseObject(prsAEProductCategories) 
	
End Sub	'AEProductCategory

'***********************************************************************************************

Public Function Load()

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
				
		If Len(mbytCategoryFilter) > 0 And cblnSF5AE Then
	        pstrSQL = "SELECT sfProducts.* FROM sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID " & mstrSQLWhere
	        pstrSQL = "SELECT sfProducts.*" _
					& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " & mstrSQLWhere
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
		
	end with

	If mblnShowAttributes Then
		On Error Resume Next
		If pblnAttributeCategoryOrderable Then
			Set prsAttributes = GetRS("Select * from sfAttributes Order By attrDisplayOrder")
			If Err.number = -2147217904 Then '-2147217904 = No value given for one or more required parameters
				pblnAttributeCategoryOrderable = False
				Set prsAttributes = GetRS("Select * from sfAttributes Order By attrName")
				Err.Clear
			ElseIf Err.number > 0 Then
				debugprint "Error: " & Err.number,Err.Description
			End If
		Else
			Set prsAttributes = GetRS("Select * from sfAttributes Order By attrName")
		End If
		On Error Goto 0
		
		Set prsAttributeDetails = GetRS("Select * from sfAttributeDetail Order By attrdtOrder")
	Else
		If Len(mstrprodID) = 0 Then
			If Not prsProducts.EOF Then	
				If pblnAttributeCategoryOrderable Then
					Set prsAttributes = GetRS("Select * from sfAttributes where attrProdId = '" & prsProducts.Fields("prodID").Value & "' Order By attrDisplayOrder")
				Else
					Set prsAttributes = GetRS("Select * from sfAttributes where attrProdId = '" & prsProducts.Fields("prodID").Value & "' Order By attrName")
				End If
			End If
		Else
			If pblnAttributeCategoryOrderable Then
				Set prsAttributes = GetRS("Select * from sfAttributes where attrProdId = '" & mstrprodID & "' Order By attrDisplayOrder")
			Else
				Set prsAttributes = GetRS("Select * from sfAttributes where attrProdId = '" & mstrprodID & "' Order By attrName")
			End If
		End If
		
		If isObject(prsAttributes) Then 
			For i = 1 to prsAttributes.RecordCount
				If len(p_strWhere) > 0 Then
					p_strWhere = p_strWhere & " OR attrdtAttributeId=" & prsAttributes("attrID").value
				Else
					p_strWhere = "Where attrdtAttributeId=" & prsAttributes("attrID").value
				End If
				prsAttributes.MoveNext
			Next
			If len(p_strWhere) > 0 Then	Set prsAttributeDetails = GetRS("Select * from sfAttributeDetail " & p_strWhere & " Order By attrdtOrder")

		End If
		
	End If

	If Not prsProducts.EOF Then pstrProdID = prsProducts.Fields("prodID").Value
    Load = (Not prsProducts.EOF)

End Function    'Load

'***********************************************************************************************

Public Function Activate(strProdID,blnActivate)

Dim sql

'On Error Resume Next

	if blnActivate then
		sql = "Update sfProducts Set prodEnabledIsActive=1 where prodID='" & strProdID & "'"
        pstrMessage = "Product successfully activated."
    else
		sql = "Update sfProducts Set prodEnabledIsActive=0 where prodID='" & strProdID & "'"
        pstrMessage = "Product successfully deactivated."
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        Activate = True
    Else
        pstrMessage = Err.Description
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
    
	If mblnIsAEPM Then

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

Public Function DeleteAttribute(lngID)

Dim sql
Dim rs
Dim p_strID, p_strName
Dim p_intCount

'On Error Resume Next

	sql = "Select attrProdId,attrName from sfAttributes where attrID=" & lngID
	Set rs = GetRS(sql)
	if not rs.eof Then 
		p_strID = rs("attrProdId").value
		p_strName = rs("attrName").value
	end if
	rs.Close
    Set rs = Nothing
    
	If len(lngID) = 0 Then Exit Function
    sql = "Delete from sfAttributeDetail where attrdtAttributeId=" & lngID
    cnn.Execute sql, , 128

    sql = "Delete from sfAttributes where attrID=" & lngID
    cnn.Execute sql, , 128

	sql = "Select * from sfAttributes where attrProdId='" & sqlSafe(p_strID) & "'"
	Set rs = GetRS(sql)
	If rs.EOF Then
		sql = "Update sfProducts set prodAttrNum=0 where prodID='" & sqlSafe(p_strID) & "'"
		cnn.Execute sql,,128
	Else
		sql = "Update sfProducts set prodAttrNum=" & rs.RecordCount & " where prodID='" & sqlSafe(p_strID) & "'"
		cnn.Execute sql,,128
	End If
	rs.Close
    Set rs = Nothing
    
   If (Err.Number = 0) Then
        pstrMessage = "The product attribute " & p_strName & " was successfully deleted."
		If mblnIsAEPM Then Call UpdateInventoryFields
        DeleteAttribute = True
    Else
        pstrMessage = Err.Description
        DeleteAttribute = False
    End If

End Function    'DeleteAttribute

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
    
	If mblnIsAEPM Then

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
	
	'Update MTP
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
	
	'Update InventoryInfo
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

'On Error Resume Next

	If len(strProductID) = 0 Then Exit Function
	
	sql = "Select prodID from sfProducts where prodID = '" & SQLSafe(strNewProductID) & "'"
	Set prsSource = GetRS(sql)

	If prsSource.EOF Then
	    pstrMessage = "Product " & strNewProductID & " does not exist."
	    CopyProductAttributesToExistingProduct = False
	Else
		
		If Len(lngAttrID) = 0 Then
			sql = "Select * from sfAttributes where attrProdId='" & strProductID & "'"
			Set rsAttribute = GetRS(sql)
			For i = 1 to rsAttribute.RecordCount
				Call CopyAttribute(rsAttribute.Fields("attrID").Value,strNewProductID,"")
				rsAttribute.MoveNext
			Next
		Else
			sql = "Select * from sfAttributes where attrID=" & lngAttrID
			Set rsAttribute = GetRS(sql)
			If rsAttribute.EOF Then
				pstrMessage = "Error Locating Attribute."
				CopyProductAttributesToExistingProduct = False
			Else
				p_strAttName = rsAttribute.Fields("attrName").Value
				Call CopyAttribute(rsAttribute.Fields("attrID").Value,strNewProductID,"")
			End If
		End If
		
		rsAttribute.Close
		Set rsAttribute = Nothing
		
		If (Err.Number = 0) Then
			If Len(lngAttrID) > 0 Then
				pstrMessage = "Attribute " & p_strAttName & " was successfully copied to Product " & strNewProductID & "."
			Else
				pstrMessage = "Product " & strProductID & "'s attributes were successfully copied to Product " & strNewProductID & "."
			End If
		    CopyProductAttributesToExistingProduct = True
		ElseIf Err.number <> 0 Then
		    pstrMessage = Err.Description
		    CopyProductAttributesToExistingProduct = False
		End If
		
	End If
	
	prsSource.Close
	set prsSource = Nothing
	
End Function    'CopyProductAttributesToExistingProduct

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
			If prsSource(pstrField).Type <> 128 Then prsTarget(pstrField).value = prsSource(pstrField).value
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
	
	If mblnIsAEPM Then Call CopyAE(strProductID,strNewProductID)
	
	If (Err.Number = 0) Then
	    pstrMessage = strNewProductID & " was successfully created."
	    DuplicateProduct = True
	Else
	    pstrMessage = Err.Description
	    DuplicateProduct = False
	End If

End Function    'DuplicateProduct

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
					prsTarget.Fields(i).Value = prsSource.Fields(i).Value
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
					prsTarget.Fields(i).Value = prsSource.Fields(i).Value
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

On Error Resume Next

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
			
			If mblnIsAEPM Then 

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

'        rs("prodAttrNum") = plngprodAttrNum
        rs("prodID") = pstrprodID
        rs("prodCountryTaxIsActive") = pblnprodCountryTaxIsActive * -1
        rs("prodDescription") = pstrprodDescription
        rs("prodEnabledIsActive") = pblnprodEnabledIsActive * -1
        rs("prodHeight") = pdblprodHeight
        rs("prodWeight") = pdblprodWeight
        rs("prodWidth") = pdblprodWidth
        rs("prodLength") = pdblprodLength
        rs("prodImageLargePath") = pstrprodImageLargePath
        rs("prodImageSmallPath") = pstrprodImageSmallPath
        rs("prodFileName") = pstrprodFileName
        rs("prodLink") = pstrprodLink
        rs("prodCategoryId") = plngprodCategoryId
        rs("prodManufacturerId") = plngprodManufacturerId
        rs("prodVendorId") = plngprodVendorId
        rs("prodMessage") = pstrprodMessage
        rs("prodName") = pstrprodName
        rs("prodNamePlural") = pstrprodNamePlural
        rs("prodPrice") = pstrprodPrice
        rs("prodSaleIsActive") = pblnprodSaleIsActive * -1
        rs("prodSalePrice") = pstrprodSalePrice
        rs("prodShip") = pstrprodShip
        rs("prodShipIsActive") = pblnprodShipIsActive * -1
        rs("prodShortDescription") = pstrprodShortDescription
        rs("prodStateTaxIsActive") = pblnprodStateTaxIsActive * -1
        
        rs("prodDateModified") =  Now()
        
		If Len(pdtprodDateAdded) = 0 Then
			rs("prodDateAdded") =  Now()
		Else
			rs("prodDateAdded") =  pdtprodDateAdded
		End If
		
		If pblnPricingLevel Then
			rs("prodPLPrice") = pstrprodPLPrice
			rs("prodPLSalePrice") = pstrprodPLSalePrice
		End If

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
			
			rs("attrName") = pstrattrName
			rs("attrProdId") = pstrProdId
			
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
				
				rs("attrdtAttributeId") = plngattrID
				rs("attrdtName") = pstrattrdtName
				rs("attrdtPrice") = pstrattrdtPrice
				rs("attrdtType") = pintattrdtType
				If pblnPricingLevel Then rs("attrdtPLPrice") = pstrattrdtPLPrice

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
 
		Set rs = Server.CreateObject("ADODB.RECORDSET")
		sql = "Select * from sfAttributes where attrProdId='" & sqlSafe(pstrprodID) & "'"
		rs.open sql, cnn, 1, 3
		If Not rs.EOF Then
 			sql = "Update sfProducts set prodAttrNum=" & rs.RecordCount & " where prodID='" & sqlSafe(pstrOrigprodID) & "'"
			cnn.Execute sql,,128
		End If
		rs.Close
        Set rs = Nothing
        
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

		If mblnIsAEPM Then Call UpdateAE

    Else
        pblnError = True
    End If

    Update = (not pblnError)

End Function    'Update

'***********************************************************************************************

Public Function jsOutputValue(strValue)
	jsOutputValue = chr(34) & Replace(strValue & "",Chr(34), Chr(34) & " + cstrQuote + " & Chr(34)) & chr(34)
End Function

Public Sub OutputAttrValues()

Dim i,j
Dim pstrReplaceQuote
Dim pbytDisplayStyle

On Error Resume Next

	If Not isObject(prsAttributes) Then Exit Sub
	With prsAttributes
		If not .EOF Then .filter = "attrProdId='" & pstrProdID & "'"
		if .RecordCount > 0 then 
			.MoveFirst
			Response.Write "var aryAttributeNames = new Array("
			Response.Write prsAttributes("attrID") & "," & jsOutputValue(prsAttributes("attrName"))		
			.MoveNext
			For i = 2 to .RecordCount
				Response.Write "," & prsAttributes("attrID") & "," & jsOutputValue(prsAttributes("attrName"))		
				.MoveNext
			Next
			Response.Write ");" & vbcrlf
			.MoveFirst

			'added for text based attributes
			Dim pblnError
			Dim pstrTextBasedAttributes
			pblnError = False
			
			If pblnTextBasedAttribute Then
				pstrTextBasedAttributes = "var aryAttributeDisplay = new Array();" & vbcrlf
				For i = 1 to .RecordCount
					If isNull(prsAttributes("attrDisplayStyle")) Then
						pbytDisplayStyle = 0
					Else
						pbytDisplayStyle = prsAttributes("attrDisplayStyle")
					End If
					pstrTextBasedAttributes = pstrTextBasedAttributes & "aryAttributeDisplay[" & prsAttributes("attrID") & "]=" & pbytDisplayStyle & ";" & vbcrlf
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

Dim i,j

	If Not isObject(prsAttributes) Then Exit Sub

	prsAttributes.filter = "attrProdId='" & pstrProdID & "'"
	Response.Write "var aryAttributeDetails = new Array();" & vbcrlf
	If pblnPricingLevel Then Response.Write "var aryAttributePLDetails = new Array();" & vbcrlf
	If prsAttributes.RecordCount > 0 then prsAttributes.MoveFirst
	Response.Write vbcrlf
	For i = 1 to prsAttributes.RecordCount
		prsAttributeDetails.filter = "attrdtAttributeId=" & prsAttributes("attrID")
		If prsAttributeDetails.RecordCount > 0 then 
			prsAttributeDetails.MoveFirst
			For j = 1 to prsAttributeDetails.RecordCount
				Response.Write "aryAttributeDetails[" & prsAttributeDetails("attrdtID") & "] = ["
				Response.Write jsOutputValue(prsAttributeDetails("attrdtName"))
				Response.Write "," & prsAttributeDetails("attrdtPrice")
				Response.Write "," & prsAttributeDetails("attrdtType")
				Response.Write "]" & vbcrlf
				If pblnPricingLevel Then
					Response.Write "aryAttributePLDetails[" & prsAttributeDetails("attrdtID") & "] = ["
					Response.Write chr(34) & prsAttributeDetails("attrdtPLPrice") & chr(34)
					Response.Write "]" & vbcrlf
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
Dim aSortHeader(4,1)
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr
Dim pstrSelect, pstrHighlight
Dim pstrID
Dim pblnSelected, pblnHasDetails, pblnExpanded

	With Response

		.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke' id='tblSummary'>"
		.Write "<colgroup align='left' width='5%'>"
		.Write "<colgroup align='left' width='10%'"
		.Write "<colgroup align='left' width='85%'>"
		.Write "<colgroup align='center'>"
		.Write "	<tr class='tblhdr'>"

	If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort Product IDs in descending order"
		aSortHeader(2,0) = "Sort Product Names in descending order"
		aSortHeader(3,0) = "Sort Prices in descending order"
		aSortHeader(4,0) = "Sort Active Categories first"
		pstrSortOrder = "ASC"
	Else
		aSortHeader(1,0) = "Sort Product IDs in ascending order"
		aSortHeader(2,0) = "Sort Product Names in ascending order"
		aSortHeader(3,0) = "Sort Prices in ascending order"
		aSortHeader(4,0) = "Sort Inactive Categories first"
		pstrSortOrder = "DESC"
	End If
	aSortHeader(1,1) = "Product ID"
	aSortHeader(2,1) = "Product Name"
	aSortHeader(3,1) = "Price"
	aSortHeader(4,1) = "Active&nbsp;&nbsp;"

	.Write "<TH>&nbsp;</TH>"
	if len(mstrOrderBy) > 0 Then
		pstrOrderBy = mstrOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
	for i = 1 to 4
		If cInt(pstrOrderBy) = i Then
			If (pstrSortOrder = "ASC") Then
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'DESC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
								"&nbsp;&nbsp;&nbsp;<img src='images/up.gif' border=0 align=bottom></TH>" & vbCrLf
			Else
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'ASC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
								"&nbsp;&nbsp;&nbsp;<img src='images/down.gif' border=0 align=bottom></TH>" & vbCrLf
			End If
		Else
		    .Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'" & pstrSortOrder & "');" & chr(34) & _
							" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
							" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & "</TH>" & vbCrLf
		End If
	next 'i

		.Write "	</tr>"

		.Write "<tr><td colspan=5>"
		.Write "<div name='divSummary' style='height:" & mbytSummaryTableHeight & "; overflow:scroll;'>"
		.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary'" _
				 & ">"
		.Write "<colgroup align='left' width='1%'>"
		.Write "<colgroup align='left' width='4%'>"
		.Write "<colgroup align='left' width='25%'>"
		.Write "<colgroup align='left' width='32%'>"
		.Write "<colgroup align='left' width='19%'>"
		.Write "<colgroup align='left' width='12%'>"
    If prsProducts.RecordCount > 0 Then
        prsProducts.MoveFirst

'Need to calculate current recordset page and upper bound to loop through
dim plnguBound, plnglbound, pstrDisplay

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
	plnglbound = (mlngAbsolutePage - 1) * prsProducts.PageSize + 1
	plnguBound = mlngAbsolutePage * prsProducts.PageSize

	If plnguBound > prsProducts.RecordCount Then plnguBound = prsProducts.RecordCount
		prsProducts.AbsolutePosition = plnglbound
        For i = plnglbound To plnguBound
        
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
            pblnExpanded = pblnSelected

			prsAttributes.Filter = "attrProdId='" & SQLSafe(pstrID) & "'"
			pblnHasDetails = (Not prsAttributes.EOF)

            If pblnSelected Then
                '.Write "<TR class='Selected' onmouseover='doMouseOverRow(this);' onmouseout='doMouseOutRow(this);'>"
				'.Write " <TD>&nbsp;</TD>"
                .Write "<TR class='Selected' " & pstrHighlight & ">"
				.Write " <TD " & pstrSelect & ">&nbsp;</TD>"
            Else
				if ConvertBoolean(prsProducts("prodEnabledIsActive")) then
					.Write " <TR class='Active' " & pstrHighlight & ">"
				else
					.Write " <TR class='Inactive' " & pstrHighlight & ">"
        		end if
				.Write " <TD " & pstrSelect & ">&nbsp;</TD>"
            End If
            
			if pblnHasDetails then
				If pblnExpanded Then
					.Write " <TD onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Hide attributes' onclick='return ExpandProduct(this," & chr(34) & pstrID & chr(34) & ");'>-</TD>"
				Else
					.Write " <TD onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Show attributes' onclick='return ExpandProduct(this," & chr(34) & pstrID & chr(34) & ");'>+</TD>"
				End If
			else
				If pblnSelected Then
					'.Write " <TD>&nbsp;</TD>"
					.Write " <TD " & pstrSelect & ">&nbsp;</TD>"
				Else
					.Write " <TD " & pstrSelect & ">&nbsp;</TD>"
				End If
       		end if
       		
            If pblnSelected Then
        		'.Write "<TD " & pstrSelect & "><a onclick='return false;' href='" & pstrURL & "' " & pstrSelect & ">" & prsProducts("prodID") & "</a></TD>"
        		'.Write "<TD>" & prsProducts("prodName") & "&nbsp;</TD>"
       			'.Write "<TD>" & WriteCurrency(prsProducts("prodPrice")) & "</TD>"

        		.Write "<TD " & pstrSelect & "><a onclick='return false;' href='" & pstrURL & "' " & pstrSelect & ">" & prsProducts("prodID") & "</a></TD>"
        		.Write "<TD " & pstrSelect & ">" & prsProducts("prodName") & "&nbsp;</TD>"
        		.Write "<TD " & pstrSelect & ">" & WriteCurrency(prsProducts("prodPrice")) & "</TD>"
            Else
        		.Write "<TD " & pstrSelect & "><a onclick='return false;' href='" & pstrURL & "' " & pstrSelect & ">" & prsProducts("prodID") & "</a></TD>"
        		.Write "<TD " & pstrSelect & ">" & prsProducts("prodName") & "&nbsp;</TD>"
        		.Write "<TD " & pstrSelect & ">" & WriteCurrency(prsProducts("prodPrice")) & "</TD>"
            End If
            
			if ConvertBoolean(prsProducts("prodEnabledIsActive")) then
	       		.Write "<TD><a href='sfProductAdmin.asp?Action=Deactivate&prodID=" & prsProducts("prodID") & _
										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to deactivate " & prsProducts("prodName") & "'>Active</a></TD></TR>" & vbCrLf
			else
	       		.Write "<TD><a href='sfProductAdmin.asp?Action=Activate&prodID=" & prsProducts("prodID") & _
										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to activate " & prsProducts("prodName") & "'>Inactive</a></TD></TR>" & vbCrLf
	       	end if
'			if ConvertBoolean(prsProducts("prodEnabledIsActive")) then
'        		.Write "<TD><a onclick='theDataForm.ViewID.value = """ & pstrID & """; theDataForm.Action.value = 'ViewProduct'; theDataForm.submit();' href='sfProductAdmin.asp?Action=Deactivate&prodID=" & prsProducts("prodID") & _
'										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to deactivate " & prsProducts("prodName") & "'>Active</a></TD></TR>" & vbCrLf
'			else
'        		.Write "<TD><a href='sfProductAdmin.asp?Action=Activate&prodID=" & prsProducts("prodID") & _
'										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to activate " & prsProducts("prodName") & "'>Inactive</a></TD></TR>" & vbCrLf
'        	end if
        	.Write "</TR>"
        	
			If pblnHasDetails Then 
	        	.Write "<TR><TD COLSPAN=6>"
				Call OutputAttr(pstrID,pblnExpanded,pblnSelected)
        		.Write "</TD></TR>"
			End If

            prsProducts.MoveNext
        Next
    Else
			.Write "<TR><TD align=center COLSPAN=6><h3>There are no Products</h3></TD></TR>"
    End If
		.Write "</TABLE></div>"
		.Write "<tr class='tblhdr'><TH COLSPAN=5 align=center>"
		If prsProducts.RecordCount = 0 Then
			.Write "No Products match your search criteria"
		Elseif prsProducts.RecordCount = 1 Then
			.Write "1 product matches your search criteria"
		Else 
			.Write prsProducts.RecordCount & " products match your search criteria<br>"

		dim pstrCheck
		pstrCheck = "return isInteger(this, true, ""Please enter a positive integer for the recordset page size."");"
		.Write "Show&nbsp;<input type='Text' id='PageSize' name='PageSize' value='" & prsProducts.PageSize & "' maxlength='4' size='4' style='text-align:center;' onblur='" & pstrCheck & "'>&nbsp;records at a time.&nbsp;&nbsp;"
		For i=1 to mlngPageCount
			plnglbound = (i-1) * mlngMaxRecords + 1
			plnguBound = i * mlngMaxRecords
			if plnguBound > prsProducts.RecordCount Then plnguBound = prsProducts.RecordCount
			pstrDisplay = plnglbound & " - " & plnguBound & "&nbsp;"
			If i = cInt(mlngAbsolutePage) Then
				Response.Write pstrDisplay
			Else
				Response.Write "<a href='#' onclick='return ViewPage(" & i & ");'>" & pstrDisplay & "</a>&nbsp;"
			End If
		Next
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
    If Not IsNumeric(pdblprodHeight) Then strError = strError & "Please enter a number for the product height." & cstrDelimeter
    If Not IsNumeric(pdblprodWeight) Then strError = strError & "Please enter a number for the product weight." & cstrDelimeter
    If Not IsNumeric(pdblprodWidth) Then strError = strError & "Please enter a number for the product width." & cstrDelimeter
    If Not IsNumeric(pdblprodLength) Then strError = strError & "Please enter a number for the product length." & cstrDelimeter
    If Not IsNumeric(pstrprodPrice) Then strError = strError & "Please enter a number for the product price." & cstrDelimeter
    If Not IsNumeric(pstrprodSalePrice) Then strError = strError & "Please enter a number for the sale price." & cstrDelimeter
    If Not IsNumeric(pstrprodShip) Then strError = strError & "Please enter a number for the ship amount." & cstrDelimeter

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
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="ssStorePath.asp"-->
<!--#include file="ssProduct_CommonFilter.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

Function SetImagePath(strImage)

	If len(trim(strImage)) > 0 Then
		SetImagePath = mstrBaseHRef & strImage
	Else
		SetImagePath = "images/NoImage.gif"
	End If

End Function	'SetImagePath

'******************************************************************************************************************************************************************

Function ShortsImageManager()

dim pobjFSO
dim pstrTestPath

	pstrTestPath = Replace(Server.MapPath("sfProductAdmin.asp"),"sfProductAdmin.asp","ImageManager\imageupload.asp")
	set pobjFSO = server.CreateObject("Scripting.FileSystemObject")
	ShortsImageManager = pobjFSO.FileExists(pstrTestPath)
	set pobjFSO = nothing
	
End Function	'ShortsImageManager


mstrPageTitle = "Product Administration"

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsProduct
Dim mstrprodID, mlngattrID, mlngattrdtID
Dim mblnShortsImageManager

Dim mblnShowFilter, mblnShowSummary, mstrShow

Dim mblnShowHeader
Dim mblnShowDetail

	mblnShowHeader = True
	mblnShowDetail = True

	mstrShow = Request.Form("Show")
	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	
	If mblnDetailInNewWindow Then
		mblnShowHeader = False
		mblnShowDetail = True
	End If
	
	mstrprodID = LoadRequestValue("prodID")
	mlngattrID = LoadRequestValue("attrID")
	mlngattrdtID = LoadRequestValue("attrdtID")
	mAction = LoadRequestValue("Action")
	If Len(mAction) = 0 Then mAction = "Filter"
	mlngAbsolutePage = LoadRequestValue("AbsolutePage")
	mblnShortsImageManager = ShortsImageManager
    Call LoadFilter
    Call LoadSort
    
'debugprint "mstrprodID",mstrprodID
'debugprint "pstrSQL",pstrSQL
    Set mclsProduct = New clsProduct
    With mclsProduct
	'debugprint "mAction",mAction
    Select Case mAction
        Case "New", "Update"
            .Update
        Case "DeleteProduct"
			.DeleteProduct mstrprodID
			mstrprodID = ""
        Case "DeleteAttribute"
			.DeleteAttribute mlngattrID
			.prodID = mstrProdID
			.UpdateInventoryFields
        Case "DeleteAttrDetail"
			.DeleteAttributeDetail mlngattrdtID
			.prodID = mstrProdID
			.UpdateInventoryFields
		Case "Filter"
			mstrprodID = ""
			If mblnDetailInNewWindow Then 
				mblnShowHeader = True
				mblnShowDetail = False
			End If
        Case "Activate", "Deactivate"
			If len(mstrprodID) > 0 Then	.Activate mstrprodID, mAction = "Activate"
		Case "CopyProd"
			.DuplicateProduct mstrprodID, Request.Form("CopyProduct")
			mstrprodID = Request.Form("CopyProduct")
			mlngattrID = ""
			mlngattrdtID = ""
		Case "CopyAttributesToProd" 'copies all product attributes to existing product
			.CopyProductAttributesToExistingProduct mstrprodID, Request.Form("CopyProduct"), ""
			mstrprodID = Request.Form("CopyProduct")
			mlngattrID = ""
			mlngattrdtID = ""
			.prodID = mstrProdID
			.UpdateInventoryFields
		Case "CopyAttr" 'copies single selected attribute to existing product
			.CopyProductAttributesToExistingProduct mstrprodID, Request.Form("CopyProduct"), mlngattrID
			mstrprodID = Request.Form("CopyProduct")
			mlngattrID = ""
			mlngattrdtID = ""
			.prodID = mstrProdID
			.UpdateInventoryFields
		Case "DuplicateAttr" 'creates a duplicate of the selected product attribute within the current product
			.CopyAttribute mlngattrID,"",Request.Form("CopyProduct")
			.prodID = mstrProdID
			.UpdateInventoryFields
		Case "ViewProduct"
			mstrprodID = LoadRequestValue("ViewID")
			mlngattrID = ""
			mlngattrdtID = ""
		Case "ViewAttribute"
			mstrprodID = ""
			mlngattrID = Request.Form("ViewID")
			mlngattrdtID = ""
		Case "ViewAttrDetail"
			mstrprodID = ""
			mlngattrID = ""
			mlngattrdtID = Request.Form("ViewID")
    End Select
    
    If mblnDetailInNewWindow And (mAction = "ViewProduct") Then
		mstrsqlWhere = " where sfProducts.prodID='" & sqlSafe(mstrprodID) & "'"
		mblnShowAttributes = False
	End If
	.Load

	If len(mstrprodID) > 0 Then
		.FindProduct mstrprodID
	Elseif len(mlngattrID) > 0 Then
		.FindAttribute mlngattrID
	Elseif len(mlngattrdtID) > 0 Then
		.FindAttrDetail mlngattrdtID
	Else
		Call .LoadProductValues
	End If
	If mblnIsAEPM Then .LoadAE
	
	dim mrsCategory, mrsVendor, mrsManufacturer

	Set mrsCategory = GetRS("Select catID,catName from sfCategories Order By catName")
	Set mrsVendor = GetRS("Select vendID,vendName from sfVendors Order By vendName")
	Set mrsManufacturer = GetRS("Select mfgID,mfgName from sfManufacturers Order By mfgName")

	If .PricingLevel Then
		Dim mstrHeaderRow

		Dim maryPLPrices
		Dim mstrPLPrice
		Dim mobjrsPricingLevels
		Dim mlngNumPricingLevels
		Set mobjrsPricingLevels = GetRS("Select PricingLevelID,PricingLevelName from PricingLevels Order By PricingLevelID")
		If mobjrsPricingLevels.EOF Then
			mlngNumPricingLevels = 0
		Else
			mlngNumPricingLevels = mobjrsPricingLevels.RecordCount
		End If

		mstrHeaderRow = "<tr>"
		For i = 1 To mlngNumPricingLevels 
			mstrHeaderRow = mstrHeaderRow & "<td><font size=-1><i><b>" & Trim(mobjrsPricingLevels.Fields("PricingLevelName").Value) & "</b></i></font></td>"
			mobjrsPricingLevels.MoveNext
		Next 'i
		mstrHeaderRow = mstrHeaderRow & "</tr>"
		mobjrsPricingLevels.MoveFirst
	End If
%>
<script LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></script>
<% If mblnShortsImageManager Then %><script language="javascript" src="ImageManager/SOSLibrary/incPureUpload.js"></script><% End If %>
<script LANGUAGE=javascript>
<!--
<% If .PricingLevel Then Response.Write "var mlngPricingLevels = " & mlngNumPricingLevels & ";" %>

var cstrQuote = '"';
var cstrDash = 'DASH';

var theDataForm;
var theKeyField;
var strDetailTitle = "<% If len(.prodID) > 0 Then Response.Write .prodID & ": " & EncodeString(.prodName,False) %>";
var blnIsDirty;

function MakeDirty(theItem)
{
var theForm = theItem.form;

	theForm.btnReset.disabled = false;
	blnIsDirty = true;
}

function body_onload()
{
	theDataForm = document.frmData;
	theKeyField = theDataForm.prodID;
	blnIsDirty = false;
<% 
If mblnShowSummary then Response.Write "DisplaySummary();" & vbcrlf
If mblnShowFilter then Response.Write "DisplayFilter();" & vbcrlf
If Not mblnShowDetail then Response.Write "return false;" & vbcrlf
If len(.attrdtID) > 0 Then 
	Response.Write "SelectAttrDetail(" & .attrID & "," & .attrdtID & ");"
ElseIf len(.attrID) > 0 Then 
	Response.Write "SelectAttr(" & .attrID & ");"
End If
If len(mstrShow)>0 then 
	Response.Write "DisplaySection(" & chr(34) & mstrShow & chr(34) & ");"
else
	Response.Write "DisplaySection(" & chr(34) & "General" & chr(34) & ");"
end if
%>
	document.all("spanprodName").innerHTML = strDetailTitle;

<% If mblnIsAEPM Then %>
var arySections = new Array("General","Detail","Attributes","Shipping","MTP","Inventory","Category");
InitializeCategory();
FillCategory();
<% End If %>
}

var gobjImage;
var gblnSwitch;

function SelectImage(theImage)
{
	gblnSwitch = true;
	gobjImage = theImage;
	document.frmData.tempFile.click();
	return false;
}

function ProcessPath(theFile)
{
var pstrFilePath = theFile.value;
var pstrBaseHRef = document.frmData.strBaseHRef.value;
var pstrBasePath = document.frmData.strBasePath.value;
var pstrHREF;
var pstrItem;
var xyz = "\\";

	if (gblnSwitch)
	{
	gobjImage.src = pstrFilePath;
	pstrItem = gobjImage.name.replace("img","");
	pstrHREF = pstrFilePath.replace(pstrBasePath,"");
	eval("document.frmData." + pstrItem).value = pstrHREF.replace(xyz,"/");
	MakeDirty(eval("document.frmData." + pstrItem));
	document.frmData.btnReset.disabled = false;
	gblnSwitch = false;
	theFile.value = "";
	}
}

function btnNewProduct_onclick(theButton)
{
var theForm = theButton.form;

	theForm.OrigprodID.value = "";
	DisplaySection("Attributes");
	theForm.attrID[0].selected = true;
	theForm.attrID.length = 1;
	ChangeAttr(theForm.attrID);
//	theForm.attrdtID[0].selected = true;
//	theForm.attrName.value = "";

    theForm.btnUpdate.value = "Add Product";
    theForm.btnDeleteProduct.disabled = true;
    theForm.btnCopyProduct.disabled = true;
	theForm.btnReset.disabled = false;
	
	SetDefaults(theForm);
	
	// AE Specific
	if (<%= LCase(CStr(mblnIsAEPM)) %>)
	{
		var MTPTable = document.all("tblMTPInput");
		for (var i=MTPTable.rows.length-1; i>0; i--)
		{
		MTPTable.deleteRow(i);
		}

		var InventoryTable = document.all("tblInventoryLevels");
		for (i=InventoryTable.rows.length-1; i>0; i--)
		{
		InventoryTable.deleteRow(i);
		}

		var theSelect = theDataForm.Categories;
		var theKey;
		for (var i=theSelect.length-1; i >=0 ;i--)
		{
			theKey = theSelect.options[i].value;
			mdicCategory.Remove(theKey);
			theSelect.options.remove(i);
		}
	//	mdicCategory.Add (3,"")
		FillCategory();
		
		theForm.ChangeInventory.value = true;
		theForm.ChangeMTP.value = true;
		theForm.ChangeCategory.value = true;
		theForm.invenInStockDEF.value = 0;
		theForm.invenLowFlagDEF.value = 0;

		theForm.invenbTracked.checked = false;
		theForm.invenbNotify.checked = false;
		theForm.invenbBackOrder.checked = false;
		theForm.invenbStatus.checked = false;
		
		theForm.gwPrice.value = 0;
		theForm.gwActivate.checked = false;
	}
	
	DisplaySection("General");
    theForm.prodID.focus();
    document.all("spanprodName").innerHTML = theDataForm.btnUpdate.value;

}

function btnDeleteProduct_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete the product " + theForm.prodID.value + ": " + theForm.prodName.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "DeleteProduct";
    theForm.submit();
    return(true);
    }
    else
    {
    return(false);
    }
}

function btnReset_onclick(theButton)
{
var theForm = theButton.form;

	blnIsDirty = false;
    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    theForm.btnDeleteProduct.disabled = false;
    
		if (aryAttributeNames.length != 0)
		{
			theForm.attrID.size = aryAttributeNames.length/2 + 1;
			for (var i=0; i < aryAttributeNames.length/2;i++)
			{
				theForm.attrID.options[i+1] = new Option(aryAttributeNames[i*2+1], aryAttributeNames[i*2]);
			}
		}
<%
If len(.attrdtID) > 0 Then 
	Response.Write "SelectAttrDetail(" & .attrID & "," & .attrdtID & ");"
ElseIf len(.attrID) > 0 Then 
	Response.Write "SelectAttr(" & .attrID & ");"
End If
%>
}

function SetDefaults(theForm)
{
    theForm.prodID.value = "";
    theForm.prodName.value = "";
    theForm.prodNamePlural.value = "";
    theForm.prodShortDescription.value = "";
    theForm.prodDescription.value = "";
    theForm.prodPrice.value = "0";
    theForm.prodEnabledIsActive.checked = false;
    theForm.prodSaleIsActive.checked = false;
    theForm.prodSalePrice.value = "0";
    theForm.prodDateAdded.value = "";

    theForm.prodImageSmallPath.value = "";
    theForm.prodImageLargePath.value = "";
    theForm.prodFileName.value = "";
    theForm.prodLink.value = "";
    theForm.prodMessage.value = "";
	SetSelect(theForm.prodCategoryId,0);
	SetSelect(theForm.prodManufacturerId,0);
	SetSelect(theForm.prodVendorId,0);

    theForm.prodWeight.value = "0";
    theForm.prodHeight.value = "0";
    theForm.prodWidth.value = "0";
    theForm.prodLength.value = "0";
    theForm.prodShip.value = "0";
    theForm.prodCountryTaxIsActive.checked = false;
    theForm.prodShipIsActive.checked = false;
    theForm.prodStateTaxIsActive.checked = false;
    
<%  
Dim i
Dim paryCustomValues

paryCustomValues = mclsProduct.CustomValues
If isArray(paryCustomValues) Then 
	For i = 0 To UBound(paryCustomValues)
		Response.Write "theForm." & paryCustomValues(i)(1) & ".value = " & Chr(34) & Chr(34) & ";" & vbcrlf
	Next 'i
End If
%>
    
    
return(true);
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.Action.value = "Filter";
	theDataForm.OrderBy.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.submit();
	return false;
}

function ViewPage(theValue)
{
	theDataForm.AbsolutePage.value = theValue;
	theDataForm.Action.value = "Filter";
	theDataForm.submit();
	return false;
}

function ViewProduct(theValue)
{
	var pblnDetailInNewWindow = (getRadio(document.frmData.chkDetailInNewWindow) == 1);

	if (pblnDetailInNewWindow)
	{
		var strURL = "sfProductAdmin.asp?Action=ViewProduct&chkDetailInNewWindow=1&ViewID=" + theValue;
		var DetailWindow = window.open(strURL,"ProductDetail","height=600,width=800,copyhistory=0,scrollbars=1");
		DetailWindow.focus();
		return false;
	}else{
		theDataForm.ViewID.value = theValue;
		theDataForm.Action.value = "ViewProduct";
		theDataForm.submit();
		return false;
	}
}

function ViewAttribute(theValue)
{
	theDataForm.ViewID.value = theValue;
	theDataForm.Action.value = "ViewAttribute";
	theDataForm.submit();
	return false;
}
function ViewAttrDetail(theValue)
{
	theDataForm.ViewID.value = theValue;
	theDataForm.Action.value = "ViewAttrDetail";
	theDataForm.submit();
	return false;
}

function SelectAttr(lngID)
{
	var pblnDetailInNewWindow = (getRadio(document.frmData.chkDetailInNewWindow) == 1);

	if (!pblnDetailInNewWindow)
	{
		theDataForm.attrID.value = lngID;
		DisplaySection("Attributes");
		ChangeAttr(theDataForm.attrID);
	}

}

function SelectAttrDetail(lngAttrID,lngID)
{
	var pblnDetailInNewWindow = (getRadio(document.frmData.chkDetailInNewWindow) == 1);

	if (!pblnDetailInNewWindow)
	{
		DisplaySection("Attributes");
		SelectAttr(lngAttrID);
		theDataForm.attrdtID.value = lngID;
		ChangeAttrDetail(theDataForm.attrdtID);
	}
}

<% 
.OutputAttrValues 
.OutputAttrDetailValues
%>

function ChangeAttr(theSelect)
{
var theForm = theSelect.form;
var intIndex = theSelect.selectedIndex;
var plngSelectedValue = theSelect.value;

	if (intIndex == 0)
	{
		theForm.btnDeleteAttr.disabled = true;
		theForm.btnDuplicateAttr.disabled = true;
		theForm.btnCopyAttr.disabled = true;
		theForm.attrdtID.length = 1;
		theForm.attrdtID.size = 3;
		theForm.attrdtID[0].selected = true;
		ChangeAttrDetail(theForm.attrdtID);
		theForm.attrName.value = "";
		theForm.btnReset.disabled = false;
		theForm.attrName.focus();
		document.all("divAttrOptions").innerHTML = "&nbsp;";
		document.all("spanAttrOptions").innerHTML = "&nbsp;";
		document.all("imgUp").disabled = true;
		document.all("imgDown").disabled = true;		
	}
	else
	{
		theForm.btnDeleteAttr.disabled = false;
		theForm.btnDuplicateAttr.disabled = false;
		theForm.btnCopyAttr.disabled = false;
		theForm.attrdtID.length = 1;
		if (aryAttributes instanceof Array)
		{
		if (aryAttributes[intIndex] instanceof Array)
		{
			theForm.attrdtID.size = aryAttributes[intIndex].length/2 + 1;
			for (var i=0; i < aryAttributes[intIndex].length/2;i++)
			{
				theForm.attrdtID.options[i+1] = new Option(aryAttributes[intIndex][i*2], aryAttributes[intIndex][i*2+1]);
			}
		}
		document.all("imgUp").disabled = false;
		document.all("imgDown").disabled = false;		

		//added for text based attributes
<% If .TextBasedAttribute Then Response.Write "		SetRadio(theForm.attrDisplayStyle,aryAttributeDisplay[plngSelectedValue]);" %>
		}
	theForm.attrdtID[0].selected = true;
	ChangeAttrDetail(theForm.attrdtID);
	theForm.attrName.value = theSelect.item(intIndex).text;
	document.all("divAttrOptions").innerHTML = theSelect.item(intIndex).text;
	document.all("spanAttrOptions").innerHTML = theSelect.item(intIndex).text + " Attributes";
	}
}

function ChangeAttrDetail(theSelect)
{
var theForm = theSelect.form;
var intValue = theSelect.value;

	if (theSelect.selectedIndex == 0)
	{
		theForm.btnDeleteAttrDetail.disabled = true;
		theForm.attrdtName.value = "";
		theForm.attrdtPrice.value = "";
		SetRadio(theForm.attrdtType,0);
		theForm.attrdtID.value = "";
		theForm.attrdtName.focus();

		<% If .PricingLevel Then %>
		if (mlngPricingLevels == 1)
		{
			theForm.attrdtPLPrice.value = "";
		}else{
			for (var i=0; i < mlngPricingLevels;i++){theForm.attrdtPLPrice[i].value = "";}
		}
		<% End If %>
	}
	else
	{
		theForm.btnDeleteAttrDetail.disabled = false;
		if (aryAttributeDetails[intValue] instanceof Array)
		{
			theForm.attrdtName.value = aryAttributeDetails[intValue][0];
			theForm.attrdtPrice.value = aryAttributeDetails[intValue][1];
			SetRadio(theForm.attrdtType,aryAttributeDetails[intValue][2]);
		}
		<% If .PricingLevel Then %>
		if (mlngPricingLevels == 1)
		{
			theForm.attrdtPLPrice.value = "";
		}else{
			for (var i=0; i < mlngPricingLevels;i++){theForm.attrdtPLPrice[i].value = "";}
		}
		
		// added for Pricing Levels
		var paryPrices;
		var pstrPrices = new String(aryAttributePLDetails[intValue]);
		paryPrices = pstrPrices.split(";");
			
		for (var i=0; i < mlngPricingLevels;i++)
		{
			if (i < paryPrices.length)
			{
				if (mlngPricingLevels == 1)
				{
					theForm.attrdtPLPrice.value = paryPrices[i];
				}else{
					theForm.attrdtPLPrice[i].value = paryPrices[i];
				}
			}else{
				if (mlngPricingLevels == 1)
				{
					theForm.attrdtPLPrice.value = "";
				}else{
					theForm.attrdtPLPrice[i].value = "";
				}
			}
		}
		<% End If %>
	}
}

function ChangeAttr_MTP(theSelect)
{
var theForm = theSelect.form;
var intIndex = theSelect.selectedIndex + 1;

	theForm.attrdtID_MTP.length = 1;
	if (aryAttributes instanceof Array)
	{
	if (aryAttributes[intIndex] instanceof Array)
	{
		theForm.attrdtID_MTP.size = aryAttributes[intIndex].length/2;
		for (var i=0; i < aryAttributes[intIndex].length/2;i++)
		{
			theForm.attrdtID_MTP.options[i] = new Option(aryAttributes[intIndex][i*2], aryAttributes[intIndex][i*2+1]);
		}
	}

	theForm.attrdtID_MTP[0].selected = true;
	ChangeAttrDetail_MTP(theForm.attrdtID_MTP);
	}
}

function ChangeAttrDetail_MTP(theSelect)
{
var theForm = theSelect.form;
var intValue = theSelect.value;

	if (theSelect.selectedIndex == 0)
	{
//		theForm.attrdtName.value = "";
//		theForm.attrdtPrice.value = "";
//		SetRadio(theForm.attrdtType,0);
//		theForm.attrdtID.value = "";
//		theForm.attrdtName.focus();
	}
	else
	{
		if (aryAttributeDetails[intValue] instanceof Array)
		{
//			theForm.attrdtName.value = aryAttributeDetails[intValue][0];
//			theForm.attrdtPrice.value = aryAttributeDetails[intValue][1];
//			SetRadio(theForm.attrdtType,aryAttributeDetails[intValue][2]);
		}
	}
}

function UpItem(strTarget)
{

var theSelect;
if (strTarget == "attribute")
{
	theSelect = theDataForm.attrID;
}else{
	theSelect = theDataForm.attrdtID;
}
var intSelected = theSelect.selectedIndex;

	if (intSelected > 1)
	{
		var optText = theSelect.options[intSelected].text;
		var optValue = theSelect.options[intSelected].value;
		
		theSelect.options[intSelected].value = theSelect.options[intSelected-1].value;
		theSelect.options[intSelected].text = theSelect.options[intSelected-1].text;
		theSelect.options[intSelected-1].value = optValue;
		theSelect.options[intSelected-1].text = optText;
		theSelect.selectedIndex = intSelected - 1;
		MakeDirty(theSelect);
	}
}

function DownItem(strTarget)
{

var theSelect;
if (strTarget == "attribute")
{
	theSelect = theDataForm.attrID;
}else{
	theSelect = theDataForm.attrdtID;
}

var intSelected = theSelect.selectedIndex;

	if (intSelected > 0)
	{
	if (intSelected < (theSelect.length - 1))
	{
		var optText = theSelect.options[intSelected].text;
		var optValue = theSelect.options[intSelected].value;
		
		theSelect.options[intSelected].value = theSelect.options[intSelected+1].value;
		theSelect.options[intSelected].text = theSelect.options[intSelected+1].text;
		theSelect.options[intSelected+1].value = optValue;
		theSelect.options[intSelected+1].text = optText;
		theSelect.selectedIndex = intSelected + 1;
		MakeDirty(theSelect);
	}
	}
}

function GetSortOrder(theForm)
{
var strOrder = "";

	if (theForm.attrID.length > 1)
	{
	for (var i=1; i < theForm.attrID.length;i++)
	{
	strOrder += theForm.attrID.options[i].value + ","
	theForm.attrDisplayOrder.value = strOrder;
	}
	}

	if (theForm.attrdtID.length > 1)
	{
	for (var i=1; i < theForm.attrdtID.length;i++)
	{
	strOrder += theForm.attrdtID.options[i].value + ","
	theForm.attrdtOrder.value = strOrder;
	}
	}
}

function ValidInput(theForm)
{
var  strSection = frmData.Show.value;

//  if (!blnIsDirty)
//  {
//  alert("No changes");
//  return(false);
//  }

  DisplaySection("General");
  if (theDataForm.prodID.value == "")
  {
    alert("Please enter a Product ID.")
    theDataForm.prodID.focus();
    return(false);
  }
    
  if (theDataForm.prodName.value == "")
  {
    alert("Please enter a Product Name.")
    theDataForm.prodName.focus();
    return(false);
  }
    
  if (!isNumeric(theForm.prodPrice,false,"Please enter a number for the product price.")) {return(false);}
  if (!isNumeric(theForm.prodSalePrice,false,"Please enter a number for the product sale price.")) {return(false);}

	DisplaySection("Shipping");
  if (!isNumeric(theForm.prodWeight,false,"Please enter a number for the product weight.")) {return(false);}
  if (!isNumeric(theForm.prodHeight,false,"Please enter a number for the product height.")) {return(false);}
  if (!isNumeric(theForm.prodWidth,false,"Please enter a number for the product width.")) {return(false);}
  if (!isNumeric(theForm.prodLength,false,"Please enter a number for the product length.")) {return(false);}
  if (!isNumeric(theForm.prodShip,false,"Please enter a number for the product ship price.")) {return(false);}
	
	DisplaySection("Attributes");
  if (!isNumeric(theForm.attrdtPrice,true,"Please enter a number for the price variance.")) {return(false);}

  if (theForm.attrdtPrice.value != 0)
  {
	  if (theForm.attrdtType[2].checked)
	  {
	    alert("You have entered a price variance AND selected No Change.\n\n Please choose either an increase or decrease in the price.");
	    theForm.attrdtPrice.focus();
	    theForm.attrdtPrice.select();
	    return(false);
	  }
  }

	GetSortOrder(theForm);
	frmData.Show.value = strSection;
	
<% If mblnIsAEPM Then %>
	//set categories values
	var theSelect = theDataForm.Categories;
	for (var i=0; i < theSelect.length;i++){theSelect.options[i].selected = true;}
<% End If %>
	
	theDataForm.submit();
    return(true);
}

function ExpandProduct(theLink,lngID)
{

	if (theLink.innerHTML == "+")
	{
		theLink.innerHTML = "-";
		theLink.title = "Hide attributes";
		eval("tbl" + lngID).style.display = "";
	}
	else
	{
		theLink.innerHTML = "+";
		theLink.title = "Show attributes";
		eval("tbl" + lngID).style.display = "none";
	}
	return false;

}

function ExpandAttr(theLink,lngID)
{
	if (theLink.innerHTML == "+")
	{
		theLink.innerHTML = "-";
		theLink.title = "Hide attributes";
		eval("tblAttDetail" + lngID).style.display = "";
	}
	else
	{
		theLink.innerHTML = "+";
		theLink.title = "Show attributes";
		eval("tblAttDetail" + lngID).style.display = "none";
	}
	return false;
}

function DisplaySection(strSection)
{
<% If Not mblnShowTabs Then Response.Write "return false;" %>

<% 
Dim pstrTempHeaderRow

pstrTempHeaderRow = "'General','Detail','Attributes','Shipping'"
If mblnIsAEPM Then pstrTempHeaderRow = pstrTempHeaderRow & ",'MTP','Inventory','Category'"
If mclsProduct.CustomMTP Then pstrTempHeaderRow = pstrTempHeaderRow & ",'MTP'"
If isArray(mclsProduct.CustomValues) Then pstrTempHeaderRow = pstrTempHeaderRow & ",'Custom'"

%>
var arySections = new Array(<%= pstrTempHeaderRow %>);

  frmData.Show.value = strSection;

 for (var i=0; i < arySections.length;i++)
 {
	if (arySections[i] == strSection)
	{
		document.all("tbl" + arySections[i]).style.display = "";
		document.all("td" + arySections[i]).className = "hdrSelected";
	}else{
		document.all("tbl" + arySections[i]).style.display = "none";
		document.all("td" + arySections[i]).className = "hdrNonSelected";
	}
 }	
 
return(false);
}

var mdicCategory = new ActiveXObject("Scripting.Dictionary");
var mblnFoundNoCatKey = false;
var mstrNoCatKey = "3";

<%= mclsProduct.ProdDictionaryList %>

function InitializeCategory()
{
	var theSelect = theDataForm.CatSource;
	
	for (var i=0; i < theSelect.length;i++)
	{
		if (mdicCategory.Exists(theSelect.options[i].value))
		{
		mdicCategory(theSelect.options[i].value) = theSelect.options[i].text;
		}
	}

}

function CleanCategory()
{

	if (!mblnFoundNoCatKey)
	{
		var theSelect = theDataForm.CatSource;	
		for (var i=0; i < theSelect.length;i++)
		{
			if (theSelect.options[i].text == "No Category")
			{
				mstrNoCatKey = theSelect.options[i].value;
				mblnFoundNoCatKey = true;
				break;
			}
		}
//		alert(mstrNoCatKey);
	}
	
	if (mdicCategory.Count > 1)
	{
		if (mdicCategory.Exists(mstrNoCatKey)){mdicCategory.Remove(mstrNoCatKey)}
	}else{
		if (mdicCategory.Count == 0){mdicCategory.Add (mstrNoCatKey,"No Category")}
	}
}

function FillCategory()
{
	CleanCategory();
	var theSelect = theDataForm.Categories;
	var pary = (new VBArray(mdicCategory.Keys())).toArray();
	var plngKey;
	var theOption;
	
	theSelect.length = 0;
	
	try
	{
	for (var i=0; i < pary.length;i++)
	{
		plngKey = pary[i];
		theOption = new Option(mdicCategory(plngKey), plngKey);
		theSelect.options.add(theOption);
	}
	}
	catch(e)
	{
	return false;
	}
}

function AddCategory()
{
	var theSelect = theDataForm.CatSource;
	var mblnAdded = false;
	if (theSelect.length > 0)
	{
		for (var i=0; i < theSelect.length;i++)
		{
			if (theSelect.options[i].selected)
			{
				if (!mdicCategory.Exists(theSelect.options[i].value))
				{
				mblnAdded = true;
//alert(theSelect.options[i].value + ": " + theSelect.options[i].text);
				mdicCategory.Add (theSelect.options[i].value,theSelect.options[i].text)
				}
			}
		}
	}
	if (mblnAdded){FillCategory();}
}

function DeleteCategory()
{
	var theSelect = theDataForm.Categories;
	var mblnAdded = false;
	
	for (var i=theSelect.length-1; i >=0 ;i--)
	{
		if (theSelect.options[i].selected)
		{
			mdicCategory.Remove(theSelect.options[i].value);
			mblnAdded = true;
		}
	}
	if (mblnAdded){FillCategory();}

}

function ChangeAE(strSection)
{
document.all("Change" + strSection).value = "True";
}

//-->
</script>
<center>
<%
End With

	If mblnShowHeader Then
		Call WriteHeader("body_onload();",True)
		Call WritePageHeader
		Call WriteFormOpener
		Call WriteProductFilter(True)
		Response.Write mclsProduct.OutputMessage
		If (len(mAction) > 0 or mblnAutoShowTable) Then Response.Write "<BR>" & mclsProduct.OutputSummary & "<BR>"
		If mblnShowDetail Then Call WriteProductDetail
	Else
		Call WriteHeader("body_onload();",False)
		Call WriteFormOpener
		Response.Write mclsProduct.OutputMessage
		Call WriteProductDetail
	End If
	
%>
</FORM>

</center>
</BODY>
</HTML>
<%

	Call ReleaseObject(mrsCategory)
	Call ReleaseObject(mrsVendor)
	Call ReleaseObject(mrsManufacturer)
	If mclsProduct.PricingLevel Then Call ReleaseObject(mobjrsPricingLevels)
	
	Call ReleaseObject(cnn)

    Response.Flush

'************************************************************************************************************************************
'
'	SUPPORTING ROUTINES
'
'************************************************************************************************************************************

Sub WriteFormOpener
%>
<form action="sfProductAdmin.asp" id=frmData name=frmData onsubmit="return ValidInput(this);" method=post>
<input type=hidden id=OrigprodID name=OrigprodID value="<%= mclsProduct.prodID %>">
<input type=hidden id=CopyProduct name=CopyProduct>
<input type=hidden id="attrDisplayOrder" name="attrDisplayOrder" value="">
<input type=hidden id="attrdtOrder" name="attrdtOrder" value="">
<input type=hidden id=ViewID name=ViewID>
<input type=hidden id=Action name=Action value="Update">
<input type=hidden id=blnShowSummary name=blnShowSummary value="">
<input type=hidden id=blnShowFilter name=blnShowFilter value="">
<input type=hidden id=Show name=Show value="<%= mstrShow %>">
<input type=hidden id=AbsolutePage name=AbsolutePage value="<%= mlngAbsolutePage %>">
<input type=hidden id=OrderBy name=OrderBy value="<%= mstrOrderBy %>">
<input type=hidden id=SortOrder name=SortOrder value="<%= mstrSortOrder %>">
<input type=hidden id="chkDetailInNewWindow2" name="chkDetailInNewWindow" value="<%= mblnDetailInNewWindow %>">

<input type=hidden id=strBaseHRef name=strBaseHRef Value="<%= mstrBaseHRef %>">
<input type=hidden id=strBasePath name=strBasePath Value="<%= mstrBasePath %>">
<% End Sub	'WriteFormOpener %>
<% Sub WriteProductDetail %>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblProduct">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
  <tr class="tblhdr">
	<th align=center><span id="spanprodName"></span>&nbsp;</th>
  </tr>
  <tr>
    <td>
	<% If mblnShowTabs Then %>
	<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" ID="Table2">
		<tr class="tblhdr" align=center>
			<td ID="tdGeneral" class="hdrNonSelected" onclick="return DisplaySection('General');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View General Product Information">General</td>
			<td ID="tdDetail" class="hdrNonSelected" onclick="return DisplaySection('Detail');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Product Details" >Detail</td>
			<td ID="tdAttributes" class="hdrNonSelected" onclick='return DisplaySection("Attributes");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Product Attributes" >Attributes</td>
			<td ID='tdShipping' class="hdrNonSelected" onclick='return DisplaySection("Shipping");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Shipping/Tax Settings" >Shipping/Tax</td>
		<% If mblnIsAEPM Then %>
			<td ID="tdMTP" class="hdrNonSelected" onclick='return DisplaySection("MTP");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Multi-Tier Pricing" >Multi-Tier Pricing</td>
			<td ID="tdInventory" class="hdrNonSelected" onclick='return DisplaySection("Inventory");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Product Inventory" >Inventory</td>
			<td ID='tdCategory' class="hdrNonSelected" onclick='return DisplaySection("Category");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Category Settings" >Category</td>
		<% End If %>
		<% If mclsProduct.CustomMTP Then %>
			<td ID='tdMTP' class="hdrNonSelected" onclick='return DisplaySection("MTP");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Multi-Tier Settings" >Multi-Tier Pricing</td>
		<% End If 'mblnIsAEPM %>
		<% If isArray(mclsProduct.CustomValues) Then %>
			<td ID='tdCustom' class="hdrNonSelected" onclick='return DisplaySection("Custom");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Custom Settings" >Custom</td>
		<% End If 'mblnIsAEPM %>
		</tr>
	</table>
	<% Else %>
	<table width="100%" border="0" rules="none" ID="Table3"><tr><td></td></tr></table>
	<% End If	'mblnShowTabs %>

	<% 
	Call WriteGeneralTable
	Call WriteDetailTable
	Call WriteShippingTable
	Call WriteAttributeTable
	If mblnIsAEPM Then Call WriteMTPTable
	If mclsProduct.CustomMTP Then Call WriteCustomMTPTable
	If mblnIsAEPM Then Call WriteCategoryTable
	If mblnIsAEPM Then Call WriteInventoryTable
	Call WriteCustomTable
	Call WriteFooterTable
	%>

</td>
</tr>
</table>
<%
End Sub	'WriteProductDetail

'************************************************************************************************************************************

Sub WriteGeneralTable
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblGeneral">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
      <tr>
        <td class="Label">Product ID:</td>
        <td><input id=prodID onchange='MakeDirty(this);' name=prodID Value="<%= mclsProduct.prodID %>" maxlength=50 size=50></td>
      </tr>
      <tr>
        <td class="Label">Product Name:</td>
        <td><input id=prodName onchange="MakeDirty(this);" onblur="if (frmData.prodNamePlural.value == ''){frmData.prodNamePlural.value = frmData.prodName.value + 's'}" name=prodName Value="<%= EncodeString(mclsProduct.prodName,True) %>" maxlength=50 size=50></td>
      </tr>
      <tr>
        <td class="Label">Product Name (plural):</td>
        <td><input id=prodNamePlural onchange="MakeDirty(this);" name=prodNamePlural Value="<%= EncodeString(mclsProduct.prodNamePlural,True) %>" maxlength=50 size=50></td>
      </tr>
      <tr>
        <td class="Label" title="appears on search results" onMouseOver="DisplayTitle(this);" onMouseOut="ClearTitle();">Short Description:</td>
        <td><textarea id=prodShortDescription onchange="MakeDirty(this);" name=prodShortDescription rows="5" cols="80"><%= mclsProduct.prodShortDescription %></textarea></td>
      </tr>
      <tr>
        <td class="Label" title="appears on detail page" onMouseOver="DisplayTitle(this);" onMouseOut="ClearTitle();">Long Description:</td>
        <td><textarea id=prodDescription onchange="MakeDirty(this);" name=prodDescription rows="5" cols="50"><%= mclsProduct.prodDescription %></textarea></td>
      </tr>
      <tr>
        <td class="Label">Price:</td>
        <td><input id=prodPrice onchange="MakeDirty(this);" name=prodPrice Value="<%= mclsProduct.prodPrice %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'>&nbsp;
            <input type=checkbox id=prodEnabledIsActive onchange="MakeDirty(this);" name=prodEnabledIsActive <% If mclsProduct.prodEnabledIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label FOR="prodEnabledIsActive">Is Active</label>
        </td>
      </tr>
	  <% If mblnBasePrice Then 'added for pricing levels %>
      <tr>
        <td class="Label">&nbsp;</td>
        <td>
           <table border=1 cellspacing=0 cellpadding=0 ID="Table1">
			 <%= mstrHeaderRow %>
             <tr>
             <%
				maryPLPrices = Split(mclsProduct.prodPLPrice & "",";")
				For i = 0 To (mlngNumPricingLevels - 1)
					If i > UBound(maryPLPrices) Then
						mstrPLPrice = ""
					Else
						mstrPLPrice = maryPLPrices(i)
					End If
					Response.Write "<td><INPUT id=prodPLPrice name=prodPLPrice onchange=" & Chr(34) & "MakeDirty(this);" & Chr(34) & " Value=" & Chr(34) & mstrPLPrice & Chr(34) & " size=6 onblur='return isNumeric(this, true, " & Chr(34) & "Please enter a number" & Chr(34) & ");'></td>"
				Next 'i
             %>
             </tr>
           </table>
        </td>
      </tr>
	  <% End If %>
      <tr>
        <td class="Label">Sale Price:</td>
        <td><input id=prodSalePrice onchange="MakeDirty(this);" name=prodSalePrice Value="<%= mclsProduct.prodSalePrice %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'>&nbsp;
            <input type=checkbox id=prodSaleIsActive onchange="MakeDirty(this);" name=prodSaleIsActive <% If mclsProduct.prodSaleIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label FOR="prodSaleIsActive">Sale Is Active</label>
        </td>
      </tr>
	  <% If mblnSalePrice Then	'added for pricing levels %>
      <tr>
        <td class="Label">&nbsp;</td>
        <td>
           <table border=1 cellspacing=0 cellpadding=0 ID="Table13">
			 <%= mstrHeaderRow %>
             <tr>
             <%
				maryPLPrices = Split(mclsProduct.prodPLSalePrice & "",";")
				For i = 0 To (mlngNumPricingLevels - 1)
					If i > UBound(maryPLPrices) Then
						mstrPLPrice = ""
					Else
						mstrPLPrice = maryPLPrices(i)
					End If
					Response.Write "<td><INPUT id=prodPLSalePrice name=prodPLSalePrice onchange=" & Chr(34) & "MakeDirty(this);" & Chr(34) & " Value=" & Chr(34) & mstrPLPrice & Chr(34) & " size=6 onblur='return isNumeric(this, true, " & Chr(34) & "Please enter a number" & Chr(34) & ");'></td>"
				Next 'i
             %>
             </tr>
           </table>
        </td>
      </tr>
	  <% End If %>
	  <% If mblnIsAEPM Then %>
      <tr>
        <td class="Label">Gift Wrap Charge:</td>
        <td><input id=gwPrice onchange="MakeDirty(this);" name=gwPrice Value="<%= mclsProduct.gwPrice %>" size=6 onblur='return isNumeric(this, true, "Please enter a number");'>&nbsp;
            <input type=checkbox id=gwActivate onchange="MakeDirty(this);" name=gwActivate <% WriteCheckboxValue mclsProduct.gwActivate %> value="ON">&nbsp;Gift Wrap Is Active
        </td>
      </tr>
	  <% End If %>
      <tr>
        <td class="Label">Date Added:</td>
        <td><input name="prodDateAdded" id="prodDateAdded" onchange="MakeDirty(this);" value="<%= mclsProduct.prodDateAdded %>" maxlength=50 size=50></td>
      </tr>
      <tr>
        <td class="Label">Date Modified:</td>
        <td><%= mclsProduct.prodDateModified %></td>
      </tr>
</table>
<%
End Sub	'WriteGeneralTable

'************************************************************************************************************************************

Sub WriteDetailTable
%>
<span id=spantempFile style="display:none">
<input type=file id=tempFile name=tempFile onchange="ProcessPath(this);" size="20">
</span>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblDetail">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
       <tr>
        <% If mblnShortsImageManager Then %>
        <td class="Label"><span style="cursor:hand" onclick="PickImage('prodImageSmallPath');" title="Select small image using Image Manager">Small Image</span>:</td>
        <% Else %>
        <td class="Label">Small Image:</td>
		<% End If %>
        <td><input id=prodImageSmallPath onchange="MakeDirty(this);" name=prodImageSmallPath Value="<%= mclsProduct.prodImageSmallPath %>" maxlength=255 size=60>
			<img style="cursor:hand" name=imgprodImageSmallPath id=imgprodImageSmallPath border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"return ClearTitle();" src="<%= SetImagePath(mclsProduct.prodImageSmallPath) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit the small Product image">
		</td>
       <tr>
        <% If mblnShortsImageManager Then %>
        <td class="Label"><span style="cursor:hand" onclick="PickImage('prodImageLargePath');" title="Select large image using Image Manager">Large Image</SPAN>:</td>
        <% Else %>
        <td class="Label">Large Image:</td>
		<% End If %>
        <td><input id=prodImageLargePath onchange="MakeDirty(this);" name=prodImageLargePath Value="<%= mclsProduct.prodImageLargePath %>" maxlength=255 size=60>
			<img style="cursor:hand" name=imgprodImageLargePath id=imgprodImageLargePath border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"return ClearTitle();" src="<%= SetImagePath(mclsProduct.prodImageLargePath) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit the Product image">
		</td>
      </tr>
      <tr>
        <td class="Label">File Path:</td>
        <td><input id=prodFileName onchange="MakeDirty(this);" name=prodFileName Value="<%= mclsProduct.prodFileName %>" maxlength=255 size=60>
		</td>
      </tr>

      <tr>
        <td class="Label"><span style="cursor:hand" ondblclick="document.frmData.prodLink.value = 'detail.asp?Product_ID=' + document.frmData.prodID.value;" title="Automatically set link to detail.asp">Link:</span></td>
        <td><input id=prodLink onchange="MakeDirty(this);" name=prodLink Value="<%= EncodeString(mclsProduct.prodLink,True) %>" maxlength=255 size=60>&nbsp;<img src="images/preview.gif" title="view this page" style="cursor:hand" onclick="OpenHelp('../../../' + document.frmData.prodLink.value)"></td>
      </tr>
      <tr>
        <td class="Label">Confirmation Message:</td>
        <td><textarea id=prodMessage onchange="MakeDirty(this);" name=prodMessage rows="5" cols="50"><%= mclsProduct.prodMessage %></textarea></td>
      </tr>
      <tr><td>&nbsp;</td><td>
      <table ID="Table4">
        <colgroup align="left" width="33%">
        <colgroup align="left" width="33%">
        <colgroup align="left" width="34%">
      <tr>
        <td>Category&nbsp;</td>
        <td>Manufacturer&nbsp;</td>
        <td>Vendor&nbsp;</td>
      </tr>
      <tr>
        <td>
			<select size="1"  id=prodCategoryId name=prodCategoryId onchange="MakeDirty(this);">
			<% Call MakeCombo(mrsCategory,"catName","catID",mclsProduct.prodCategoryId) %>
			</select>
		</td>        
        <td>
			<select size="1"  id=prodManufacturerId name=prodManufacturerId onchange="MakeDirty(this);">
			<% Call MakeCombo(mrsManufacturer,"mfgName","mfgID",mclsProduct.prodManufacturerId) %>
			</select>
		</td>        
        <td>
			<select size="1"  id=prodVendorId name=prodVendorId onchange="MakeDirty(this);">
			<% Call MakeCombo(mrsVendor,"vendName","vendID",mclsProduct.prodVendorId) %>
			</select>
		</td>        
      </tr>
      </table>
      </td>
      </tr>
</table>
<%
End Sub	'WriteDetailTable 

'************************************************************************************************************************************

Sub WriteShippingTable
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblShipping">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
      <tr>
        <td class="Label">Weight:</td>
        <td><input id=prodWeight onchange="MakeDirty(this);" name=prodWeight Value="<%= mclsProduct.prodWeight %>" size=6></td>
      </tr>
      <tr>
        <td class="Label">Height:</td>
        <td><input id=prodHeight onchange="MakeDirty(this);" name=prodHeight Value="<%= mclsProduct.prodHeight %>" size=6></td>
      </tr>
      <tr>
        <td class="Label">Width:</td>
        <td><input id=prodWidth onchange="MakeDirty(this);" name=prodWidth Value="<%= mclsProduct.prodWidth %>" size=6></td>
      </tr>
      <tr>
        <td class="Label">Length:</td>
        <td><input id=prodLength onchange="MakeDirty(this);" name=prodLength Value="<%= mclsProduct.prodLength %>" size=6></td>
      </tr>
      <tr>
        <td class="Label">Shipping Cost:</td>
        <td><input id=prodShip onchange="MakeDirty(this);" name=prodShip Value="<%= mclsProduct.prodShip %>" size=6>&nbsp;
            <input type=checkbox id=prodShipIsActive onchange="MakeDirty(this);" name=prodShipIsActive <% If mclsProduct.prodShipIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label FOR="prodShipIsActive">This item is shipped</label>
        </td>
      </tr>
      <tr>
        <td class="Label">&nbsp;</td>
        <td>
        <input type=checkbox id=prodStateTaxIsActive onchange="MakeDirty(this);" name=prodStateTaxIsActive <% If mclsProduct.prodStateTaxIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label FOR="prodStateTaxIsActive">Apply State Tax to this item</label></td>
      </tr>
      <tr>
        <td class="Label">&nbsp;</td>
        <td>
        <input type=checkbox id=prodCountryTaxIsActive onchange="MakeDirty(this);" name=prodCountryTaxIsActive <% If mclsProduct.prodCountryTaxIsActive Then Response.Write "Checked" %> value="ON">&nbsp;<label FOR="prodCountryTaxIsActive">Apply Country Tax to this item</label></td>
      </tr>
</table>
<%
End Sub	'WriteShippingTable 

'************************************************************************************************************************************

Sub WriteAttributeTable
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblAttributes">
      <tr>
        <td>&nbsp;</td>
        <td>
        <table width=100% border=0 ID="Table5">
        <tr>
        <td>
			<table border=0 cellpadding=0 cellspacing=0 ID="Table6">
			  <tr>
				<td valign=middle>
					<select size=3 id=attrID name=attrID onchange="ChangeAttr(this);">
					<option value="">Create New Attribute Category</option>
					<%
						If isObject(mclsProduct.rsAttributes) Then 
							mclsProduct.rsAttributes.Filter = "attrProdId='" & mclsProduct.prodID & "'"
							Call MakeCombo(mclsProduct.rsAttributes,"attrName","attrID",mclsProduct.attrID)
						End If
					%>
					</select>
				 </td>
				 <td valign=middle>
<% If mclsProduct.AttributeCategoryOrderable Then %>
					<br>
					<input type=image id=imgUp1 src="images/up.gif" onclick="UpItem('attribute'); return false;" title="Move Attribute Up" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle()" NAME="imgUp1">
					<br><input type=image id=imgDown1 src="images/down.gif" onclick="DownItem('attribute'); return false;" title="Move Attribute Down" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle()" NAME="imgDown1">
<% Else %>
&nbsp;
<% End If %>
				 </td>
			</tr>
			</table>

		</td>
		<td>&nbsp;</td>
		<td>
		Attribute Category<br>
        <input id=attrName onchange="MakeDirty(this);" name=attrName value="<%= mclsProduct.attrName  %>" maxlength=50 size=50>

		<% If mclsProduct.TextBasedAttribute Then %>
        <table border=1 cellspacing=0 cellpadding=2 ID="Table7">
			<tr><td colspan=7>Display this attribute as:&nbsp&nbsp&nbsp&nbsp<font size=-1>(<sup>*</sup>Field is optional)</font></td></tr>
			<tr>
			  <td align=center>Select</td>
			  <td align=center>Radio</td>
			  <td align=center>Text</td>
			  <td align=center>Text<sup>*</sup></td>
			  <td align=center>Textarea</td>
			  <td align=center>Textarea<sup>*</sup></td>
			  <td align=center>Checkbox</td>
			</tr>
			<tr>
			  <td align=center><input type="radio" value="0" <% if mclsProduct.attrDisplayStyle=0 then Response.Write "Checked" %> id="attrDisplayStyle" name="attrDisplayStyle" onchange='MakeDirty(this);'></td>
			  <td align=center><input type="radio" value="1" <% if mclsProduct.attrDisplayStyle=1 then Response.Write "Checked" %> id="Radio9" name="attrDisplayStyle" onchange='MakeDirty(this);'></td>
			  <td align=center><input type="radio" value="2" <% if mclsProduct.attrDisplayStyle=2 then Response.Write "Checked" %> id="Radio10" name="attrDisplayStyle" onchange='MakeDirty(this);'></td>
			  <td align=center><input type="radio" value="3" <% if mclsProduct.attrDisplayStyle=3 then Response.Write "Checked" %> id="Radio11" name="attrDisplayStyle" onchange='MakeDirty(this);'></td>
			  <td align=center><input type="radio" value="4" <% if mclsProduct.attrDisplayStyle=4 then Response.Write "Checked" %> id="Radio12" name="attrDisplayStyle" onchange='MakeDirty(this);'></td>
			  <td align=center><input type="radio" value="5" <% if mclsProduct.attrDisplayStyle=5 then Response.Write "Checked" %> id="Radio13" name="attrDisplayStyle" onchange='MakeDirty(this);'></td>
			  <td align=center><input type="radio" value="6" <% if mclsProduct.attrDisplayStyle=6 then Response.Write "Checked" %> id="Radio14" name="attrDisplayStyle" onchange='MakeDirty(this);'></td>
			</tr>
		</table>
		<% End If 'pblnTextBasedAttribute %>
		
		</td>
		<td>&nbsp;</td>
		<td align=center>
			<input class='butn' title="Delete this attribute category" id=btnDeleteAttr name=btnDeleteAttr type=button value='Delete Category' onclick='var theForm = this.form; var blnConfirm=confirm("Are you sure you wish to delete attribute category " + document.frmData.attrName.value + "?"); if (blnConfirm){theForm.Action.value = "DeleteAttribute"; theForm.submit();}' disabled><br>
			<input class='butn' title="Copy this attribute category and attributes to a new category for this product" id=btnDuplicateAttr name=btnDuplicateAttr type=button value='Duplicate Category' onclick="var theForm = this.form; var pstrNewprodName = prompt('Enter New Attribute Category Name','New Attribute Category');if (pstrNewprodName != null){theForm.CopyProduct.value = pstrNewprodName;theForm.Action.value = 'DuplicateAttr';if (<%= LCase(CStr(mblnIsAEPM)) %>){theForm.ChangeInventory.value = true;}theForm.submit();}" disabled><br>
			<input class='butn' title="Copy just this attribute category and attributes to an existing product" id=btnCopyAttr name=btnCopyAttr type=button value='Copy Category' onclick='var theForm = this.form; var pstrNewprodName = prompt("Enter Product ID to copy attribute to","Enter Product ID");if (pstrNewprodName != null){theForm.CopyProduct.value = pstrNewprodName;theForm.Action.value = "CopyAttr";theForm.submit();}' disabled>
			<input class='butn' title="Copy all of this product's attribute category and attributes to an existing product" id=btnCopyProduct name=btnCopyProduct type=button value='Copy Attributes' onclick='var theForm = this.form; var pstrNewprodName = prompt("Enter Product ID to copy attributes to","Enter Product ID"); if (pstrNewprodName != null){ theForm.CopyProduct.value = pstrNewprodName; theForm.Action.value = "CopyAttributesToProd"; theForm.submit();}'>&nbsp;
		</td>
		</tr>
		</table>
		</td>        
      </tr>
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanAttrOptions"><%= mclsProduct.attrName %> &nbsp;</span></th>
  </tr>
      <tr>
        <td>&nbsp;</td>
        <td>
        <table width=100% border=0 ID="Table8">
        <tr>
        <td valign=top>
			<table border=0 cellpadding=0 cellspacing=0 ID="Table9">
			  <tr>
				<td valign=middle>
				  <div id="divAttrOptions">&nbsp;</div>
				  <select size=3  id=attrdtID name=attrdtID onchange="ChangeAttrDetail(this);">
				  <option value="">Create New Attribute</option>
				  </select>
				 </td>
				 <td valign=middle>
					<br>
					<input type=image id=imgUp src="images/up.gif" onclick="UpItem('attributeDetail'); return false;" title="Move Attribute Up" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle()" disabled NAME="imgUp">
					<br><input type=image id=imgDown src="images/down.gif" onclick="DownItem('attributeDetail'); return false;" title="Move Attribute Down" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle()" disabled NAME="imgDown">
				 </td>
			</tr>
			</table>
		</td>
		<td valign=top>&nbsp;</td>
		<td>
		<table ID="Table10">
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td class="Label">Attribute:</td>
			<td><input id=attrdtName onchange='this.form.ChangeInventory.value = true;MakeDirty(this);' name=attrdtName value="<%= mclsProduct.attrdtName  %>" maxlength=50 size=50></td>
		</tr>
		<tr>
			<td class="Label">Price Variance:</td>
			<td><input id=attrdtPrice onchange='MakeDirty(this);' name=attrdtPrice value="<%= mclsProduct.attrdtPrice  %>" size=6></td>
		</tr>
		<% If mblnAttrPrice Then %>
		<tr>
		<td class="Label">&nbsp;</td>
		<td>
			<table border=1 cellspacing=0 cellpadding=0 ID="Table14">
				<%= mstrHeaderRow %>
				<tr>
				<%
				maryPLPrices = Split(mclsProduct.attrdtPLPrice & "",";")
				For i = 0 To (mlngNumPricingLevels - 1)
					If i > UBound(maryPLPrices) Then
						mstrPLPrice = ""
					Else
						mstrPLPrice = maryPLPrices(i)
					End If
					Response.Write "<td><INPUT id=attrdtPLPrice name=attrdtPLPrice onchange=" & Chr(34) & "MakeDirty(this);" & Chr(34) & " Value=" & Chr(34) & mstrPLPrice & Chr(34) & " size=6 onblur='return isNumeric(this, true, " & Chr(34) & "Please enter a number" & Chr(34) & ");'></td>"
				Next 'i
				%>
				</tr>
			</table>
		</td>
		</tr>
		<% End If %>
		<tr>
			<td>&nbsp;</td>
			<td>
			<input type="radio" value="1" <% if mclsProduct.attrdtType=1 then Response.Write "Checked" %> id="attrdtType1" name="attrdtType" onchange='MakeDirty(this);'><label for="attrdtType1">Increase</label><br>
			<input type="radio" value="2" <% if mclsProduct.attrdtType=2 then Response.Write "Checked" %> id="attrdtType2" name="attrdtType" onchange='MakeDirty(this);'><label for="attrdtType2">Decrease</label><br>
			<input type="radio" value="0" <% if mclsProduct.attrdtType=0 then Response.Write "Checked" %> id="attrdtType0" name="attrdtType" onchange='MakeDirty(this);'><label for="attrdtType0">No Change</label>
			</td>
		</tr>
		</table>
		</td>
		<td>&nbsp;</td>
		<td align=center>
			<input class='butn' id=btnDeleteAttrDetail name=btnDeleteAttrDetail type=button value='Delete Attribute' onclick='var theForm = this.form; var blnConfirm = confirm("Are you sure you wish to delete attribute " + document.frmData.attrdtName.value + "?"); if (blnConfirm){theForm.Action.value = "DeleteAttrDetail"; theForm.submit();}' disabled><br>
		</td>
		</tr>
		</table>
		</td>        
      </tr>
</table>
<%
End Sub	'WriteAttributeTable

'************************************************************************************************************************************

 Sub WriteInventoryTable 

	On Error Resume Next
%>
<input type=hidden id=ChangeInventory name=ChangeInventory value=False>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblInventory">
      <tr>
        <td>&nbsp;</td>
        <td>
        <table class="tbl" width=100% border=0 ID="Table11">
          <tr>
            <td class="Label"><label id='lblinvenInStockDEF' name='lblinvenInStockDEF' for='invenInStockDEF' class="Label">Default Inventory Qty</label>:</td>
            <td>
              <input type='text' id='invenInStockDEF' name='invenInStockDEF' value='<%= mclsProduct.rsInventoryInfo.Fields("invenInStockDEF").value %>' onblur='return isInteger(this, false, "Please enter an integer for the quantity");' size="20">
              <input type='checkbox' id='invenbTracked' name='invenbTracked' <% WriteCheckboxValue mclsProduct.rsInventoryInfo.Fields("invenbTracked").value %> value="ON">&nbsp;<label id='lblinvenbTracked' name='lblinvenbTracked' for='invenbTracked'>Track Inventory</label>
            </td>
          </tr>
          <tr>
            <td class="Label"><label id='lblinvenLowFlagDEF' name='lblinvenLowFlagDEF' for='lblinvenLowFlagDEF'>Default Notify Qty</label>:</td>
            <td>
              <input type='text' id='invenLowFlagDEF' name='invenLowFlagDEF' value='<%= mclsProduct.rsInventoryInfo.Fields("invenLowFlagDEF").value %>' onblur='return isInteger(this, false, "Please enter an integer for the quantity");' size="20">
              <input type='checkbox' id='invenbNotify' name='invenbNotify' <% WriteCheckboxValue mclsProduct.rsInventoryInfo.Fields("invenbNotify").value %> value="ON">&nbsp;<label id='lblinvenbNotify' name='lblinvenbNotify' for='invenbNotify'>Notify when stock reaches this level</label>
            </td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>
              <input type='checkbox' id='invenbBackOrder' name='invenbBackOrder' <% WriteCheckboxValue mclsProduct.rsInventoryInfo.Fields("invenbBackOrder").value %> value="ON">&nbsp;<label id='lblinvenbBackOrder' name='lblinvenbBackOrder' for='invenbBackOrder'>Allow Back Order</label><br>
			  <input type='checkbox' id='invenbStatus' name='invenbStatus' <% WriteCheckboxValue mclsProduct.rsInventoryInfo.Fields("invenbStatus").value %> value="ON">&nbsp;<label id='lblinvenbStatus' name='lblinvenbStatus' for='invenbStatus'>Show Stock Status on Search Page</label>
            </td>
          </tr>
		</table>
		</td></tr>
		<tr><td colspan=5><hr></td></tr>
		<tr>
		  <td colspan=5 align="center">
		  <table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblInventoryLevels">
			<tr>
			  <th>Attribute</th>
			  <th>Qty In Stock</th>
			  <th>Notify When Qty Reaches</th>
			</tr>
<%
If isObject(mclsProduct.rsInventory) Then
  With mclsProduct.rsInventory
	Do While Not .EOF
%>
   <tr>
    <input type='hidden' id='invenId' name='invenId' value='<%= .Fields("invenId").value %>'>
    <td class="Label"><%= .Fields("invenAttName").value %>:&nbsp;&nbsp;</td>
    <td align="center">
    <input type='text' id='invenInStock' name='invenInStock' value='<%= .Fields("invenInStock").value %>' onblur='return isInteger(this, true, "Please enter an integer for the quantity");' onchange="ChangeAE('Inventory');" size="20"></td>
    <td align="center">
    <input type='text' id='invenLowFlag' name='invenLowFlag' value='<%= .Fields("invenLowFlag").value %>' onblur='return isInteger(this, true, "Please enter an integer for the quantity");' onchange="ChangeAE('Inventory');" size="20"></td>
  </tr>
<%	
	.MoveNext 
	Loop
  End With
End If
%>
		  </table>
		  </td></tr>
		</TD>        
      </TR>
</table>
<%
End Sub	'WriteInventoryTable

'************************************************************************************************************************************

Sub WriteMTPTable
%>
<script>
function AddMTP()
{
var pNewRow;
var pNewCell;
var cstrQuote = '"';
var pstrCell1 = "<input type='text' id='mtQuantity' name='mtQuantity' value='0' onblur='return isInteger(this, true, " + cstrQuote + "Please enter an integer for the quantity" + cstrQuote + ");'>"
var pstrCell2 = "<input type='text' id='mtValue' name='mtValue' value='0' onblur='return isNumeric(this, true, " + cstrQuote + "Please enter an integer for the discount" + cstrQuote + ");'>"
var pstrCell3 = "<select id='mtType' name='mtType'><option>Amount</option><option>Percent</option></select>"      
var pstrCell4 = "<INPUT class='butn' id=btnDeleteMTP name=btnDeleteMTP type='button' value='Delete Discount Level' onclick='DeleteMTP();'>" 

	pNewRow = document.all("tblMTPInput").insertRow();
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell1;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell2;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell3;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell4;
	
}

function DeleteMTP(theCell)
{
var ptheRow = window.event.srcElement.parentElement.parentElement;
ptheRow.parentElement.deleteRow(ptheRow.rowIndex);
}

</script>
<input type=hidden id=ChangeMTP name=ChangeMTP value=False>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblMTP">
<tr><td>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblMTPInput">
      <tr>
        <td>Quantity</td>
        <td>Discount Amount</td>
        <td>Discount Type<br><i>(checked: value, unchecked: percentage)</i></td>
        <td>&nbsp;</td>
      </tr>
<% 
'mtProdID
'mtIndex
'mtQuantity
'mtValue
'mtType

Dim pblnAmount
Dim j

If isObject(mclsProduct.rsMTP) Then
  With mclsProduct.rsMTP
	For j = 1 To .RecordCount
%>
      <tr>
        <td valign=top><input type='text' id='mtQuantity' name='mtQuantity' value='<%= .Fields("mtQuantity").value %>' onblur='return isInteger(this, true, "Please enter an integer for the quantity");' onchange="ChangeAE('MTP');"></td>
		<td valign=top>
		  <input type='text' id='mtValue' name='mtValue' value='<%= .Fields("mtValue").value %>' onblur='return isNumeric(this, true, "Please enter an integer for the discount");' onchange="ChangeAE('MTP');">
			<% If mblnMTPrice Then %>
			<table border=1 cellspacing=0 cellpadding=0 ID="Table15">
			 <%= mstrHeaderRow %>
             <tr>
             <%
				maryPLPrices = Split(.Fields("mtPLValue").value & "",";")
				For i = 0 To (mlngNumPricingLevels - 1)
					If i > UBound(maryPLPrices) Then
						mstrPLPrice = ""
					Else
						mstrPLPrice = maryPLPrices(i)
					End If
					Response.Write "<td><INPUT id=mtPLValue" & j & " name=mtPLValue" & j & " onchange=" & Chr(34) & "ChangeAE('MTP');" & Chr(34) & " Value=" & Chr(34) & mstrPLPrice & Chr(34) & " size=6 onblur='return isNumeric(this, true, " & Chr(34) & "Please enter a number" & Chr(34) & ");'></td>"
				Next 'i
             %>
             </tr>
			</table>
			<% End If %>
		</td>
		<% If (.Fields("mtType").value="Amount") Then %>
		<td valign=top><select id='mtType' name='mtType' onchange="ChangeAE('MTP');"><option selected>Amount</option><option>Percent</option></select></td>        
		<% Else %>
		<td valign=top><select id="Select1" name='mtType' onchange="ChangeAE('MTP');"><option>Amount</option><option selected>Percent</option></select></td>        
		<% End If %>
		<td valign=top><input class='butn' id=btnDeleteMTP name=btnDeleteMTP type='button' value='Delete Discount Level' onclick="DeleteMTP(this); ChangeAE('MTP');"></td>        
      </tr>
<%	
	mclsProduct.rsMTP.MoveNext 
	Next 'j
  End With
End If
%>
</table>
</td><tr>
<tr>
	<td><input class='butn' id=btnNewMTP name=btnNewMTP type='button' value='New Discount Level' onclick="AddMTP(); ChangeAE('MTP');"></td></tr>
</table>
<% 
End Sub	'WriteMTPTable 

'************************************************************************************************************************************

Sub WriteCategoryTable
%>

<input type=hidden id=ChangeCategory name=ChangeCategory value=False>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblCategory">
  <tr>
    <td>
      <table width="100%" border="0" ID="Table12">
		<tr>
		  <th>Categories</th>
		  <th>&nbsp;</th>
		  <th>This product is in the following categories:</th>
		</tr>
		<tr>
		  <td align=center><%= mclsProduct.CategoryList %></td>
		  <td valign=middle align=center>
			<input class="butn" type=button id="btnAddCategory" name="btnAddCategory" onclick="AddCategory(); ChangeAE('Category');" value="-->"><br>
			<input class="butn" type=button id="btnDeleteCategory" name="btnDeleteCategory" onclick="DeleteCategory(); ChangeAE('Category');" value="<--"><br>
		  </td>
		  <td align=center>
			<select id=Categories name=Categories size=10 multiple>
			</select>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
</table>
<% 
End Sub	'WriteCategoryTable

'************************************************************************************************************************************

Sub WriteCustomTable

Dim i
Dim paryCustomValues

paryCustomValues = mclsProduct.CustomValues
If Not isArray(paryCustomValues) Then Exit Sub
%>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblCustom">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
	<% For i = 0 To UBound(paryCustomValues) %>
      <tr>
        <td class="Label"><%= paryCustomValues(i)(0) %>:</td>
        <td>
        <%	Select Case paryCustomValues(i)(3) %>
        <%		Case "select" %>
        <%		Case "radio" %>
        <%		Case "textbox" %>
        <%		Case "checkbox" %>
        <input  type="checkbox" name="<%= paryCustomValues(i)(1) %>" id="<%= paryCustomValues(i)(1) %>" onchange='MakeDirty(this);' Value="1" <% WriteCheckboxValue(paryCustomValues(i)(2)) %>>
        <%		Case "listbox" %>
        <%		Case Else %>
        <input  type="text" name="<%= paryCustomValues(i)(1) %>" id="<%= paryCustomValues(i)(1) %>" onchange='MakeDirty(this);' Value="<%= Server.HTMLEncode(paryCustomValues(i)(2) & "") %>" size=50>
        <%	End Select %>
        
        </td>
      </tr>
    <% Next 'i %>
</table>
<% 
End Sub	'WriteCustomTable 

'**************************************************************************************************************************************************

Sub WriteFooterTable
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFooter">
  <tr>
    <td>&nbsp;</td>
    <td>
		<input class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/ProductMgr5_help.htm'); return false;" id=btnHelp name=btnHelp title="help">&nbsp;
        <input class='butn' title='Create a new product' id=btnNewProduct name=btnNewProduct type=button value='New' onclick='return btnNewProduct_onclick(this)'>&nbsp;
        <input class='butn' title='Create a new product based on this product' id=btnDuplicateProduct name=btnDuplicateProduct type=button value='Duplicate' onclick='DuplicateProduct(this);'>&nbsp;
        <input class='butn' title="Reset" id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)' disabled>&nbsp;&nbsp;
        <input class='butn' title="Delete this product" id=btnDeleteProduct name=btnDeleteProduct type=button value='Delete' onclick='return btnDeleteProduct_onclick(this)'>
        <input class='butn' title="Save changes" id=btnUpdate name=btnUpdate type=button value='Save Changes' onclick='return ValidInput(this.form);'>
    </td>
  </tr>
</table>
<script language="javascript">
function DuplicateProduct(theField)
{

if (false)
{
var pbytH = window.screen.availHeight;
var pbytW = window.screen.availWidth;
var theTable = document.all("tblDuplicateProduct");

theTable.style.top = (pbytH - 100)/2;
theTable.style.left = (pbytW - 200)/2;
theTable.style.display = "";
document.frmData.CopyProduct1.focus();

return false;
}

	var theForm = theField.form;
	var pstrNewprodName = prompt("Enter New Product ID","New Product ID");
	if (pstrNewprodName != null){
		theForm.CopyProduct.value = pstrNewprodName;
		theForm.Action.value = "CopyProd";
		theForm.submit();
	}
}
</SCRIPT>
<table onblur="alert('this');" style="display: none; position: absolute; left: 50; top: 50; width: 200; height: 100; border-style: outset; border-width: 3" cellpadding="3" cellspacing="0" border="1"  bgcolor="#FFFFFF" id="tblDuplicateProduct">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
      <tr>
        <td class="Label"><span title="Enter the product ID to copy this product to">New Product ID:</SPAN></td>
        <td><input name="CopyProduct1" id="CopyProduct1" onchange='' value="" size=50></td>
      </tr>
      <tr>
        <td class="Label">ProductID:</td>
        <td>
          <input class='butn' title='Create a new product based on this product' id="btnDuplicateProduct2" name=btnDuplicateProduct2 type=button value='Duplicate Product' onclick='DuplicateProduct(this);'>
          <input class='butn' title='Cancel this operation' id="btnCancelDuplicateProduct" name="btnCancelDuplicateProduct" type="button" value="Cancel" onclick="document.all('tblDuplicateProduct').style.display = 'none';">
        </td>
      </tr>
</table>
<%
End Sub	'WriteFooterTable

'**************************************************************************************************************************************************

Function WriteCheckboxValue(vntValue)

	If len(Trim(vntValue) & "") > 0 Then
		If cBool(vntValue) Then Response.Write "CHECKED"
	End If


End Function	'WriteCheckboxValue

'************************************************************************************************************************************

Sub WritePageHeader
%>
<table border=0 cellPadding=5 cellSpacing=1 width="95%" ID="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
    <th>&nbsp;</th>
    <th align='right'>
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Filter">Hide Filter</div></a><br>
<% If (len(mAction) > 0 or mblnAutoShowTable) Then %>
		<a href="#"><div id="divSummary" onclick="return DisplaySummary();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Summary">Hide Summary</div></a>
<% End If %>
	</th>
  </tr>
</table>
<% 
End Sub	'WritePageHeader 

'************************************************************************************************************************************

Sub WriteCustomMTPTable
%>
<script>
function AddPricingLevel()
{
var pNewRow;
var pNewCell;
var cstrQuote = '"';
var pstrCell1 = "<input type='text' id='PricingLevel' name='PricingLevel' value='0' onblur='return isInteger(this, true, " + cstrQuote + "Please enter an integer for the quantity" + cstrQuote + ");'>"
var pstrCell2 = "<input type='text' id='PricingAmount' name='PricingAmount' value='0' onblur='return isNumeric(this, true, " + cstrQuote + "Please enter an integer for the discount" + cstrQuote + ");'>"
var pstrCell3 = "<INPUT class='butn' id=btnDeletePricingLevel name=btnDeletePricingLevel type='button' value='Delete Pricing Level' onclick='DeletePricingLevel();'>" 

	pNewRow = document.all("tblPricingLevelInput").insertRow();
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell1;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell2;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell3;
	
}

function DeletePricingLevel(theCell)
{
var ptheRow = window.event.srcElement.parentElement.parentElement;
ptheRow.parentElement.deleteRow(ptheRow.rowIndex);
}

</script>
<input type=hidden id=ChangePricingLevel name=ChangePricingLevel value=False>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblMTP">
<tr><td>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblPricingLevelInput">
      <tr>
        <td>Quantity</td>
        <td>Price</td>
        <td>&nbsp;</td>
      </tr>
<% 
Dim pblnAmount

If isObject(mclsProduct.rsPricingLevels) Then
  With mclsProduct.rsPricingLevels
	Do While Not .EOF
%>
      <tr>
        <td><input type='text' id='PricingLevel' name='PricingLevel' value='<%= .Fields("PricingLevel").value %>' onblur='return isInteger(this, true, "Please enter an integer for the quantity");'"></td>
		<td><input type='text' id='PricingAmount' name='PricingAmount' value='<%= .Fields("PricingAmount").value %>' onblur='return isNumeric(this, true, "Please enter an integer for the discount");'"></td>
		<td><input class='butn' id=btnDeletePricingLevel name=btnDeletePricingLevel type='button' value='Delete Pricing Level' onclick="DeletePricingLevel(this);"></td>        
      </tr>
<%	
	mclsProduct.rsPricingLevels.MoveNext 
	Loop
  End With
End If
%>
</table>

</td><tr>
<tr>
	<td><input class='butn' id=btnNewPricingLevel name=btnNewPricingLevel type='button' value='New Pricing Level' onclick="AddPricingLevel();"></td></tr>
</table>
<%
End Sub	'WriteCustomMTPTable
%>
