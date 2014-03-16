<% Option Explicit 
'********************************************************************************
'*   Promotion Manager for StoreFront 5.0										*
'*   Release Version:	2.00.003 												*
'*   Release Date:		August 10, 2003											*
'*   Revision Date:		October 30, 2003										*
'*																				*
'*	 Relese Notes:																*
'*																				*
'*	 2.00.003 (October 30, 2003)												*
'*	 - Note - updated to use new common file set								*
'*																				*
'*	 2.00.002 (August 29, 2003)													*
'*	 - Bug Fix - updated ssPromotionsAdmin.asp to fix filter error returning	*
'*			     no records														*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.				*
'********************************************************************************

Response.Buffer = True

Class clsPromotion
'Assumptions:
'	Connection: defines a previously opened connection to the database
'

'database variables
dim pPromotionID
dim pStartDate
dim pEndDate
dim pDuration
dim pPromoCode
dim pNumUses
Dim pNumUsesByCustomer
dim pMaxUses
dim pDiscount
dim pMaxAllowableValue
dim pMaxAllowableValuePerItem
dim pblnPercentage
Dim pblnApplyToBasePrice
dim pblnofferFreeGiftAutomatically
dim pMinSubTotal
dim pCombineable
dim pPromoTitle
dim pModifiedOn
dim pPromoRules
dim pInactive
dim pApplyAutomatically
dim pblnExcludeSaleItems

dim pstrFreeShippingCode
dim pstrProductID
dim pstrFreeProductID
dim pstrProductIDExclusion
dim pstrCategory
dim pstrCategoryExclusion
dim pstrManufacturer
dim pstrManufacturerExclusion
dim pstrVendor
dim pstrVendorExclusion

dim pstrProductCountLimit
dim plngBuyX
dim plngGetY
dim pdblFreeShippingLimit
Dim pblnLikeItem

'working variables
dim pConnection
dim prsPromo
dim pstrMessage
dim pblnError

'variable for  handling Order parameters
dim aProductsInCart
dim pstrPromotions
dim psubTotal
dim pBestDiscountAmount
dim pBestDiscountCode

dim cstrPromoDelimiter
Private Sub class_Initialize()
	cstrPromoDelimiter = ";"
End Sub

Public Property Let Connection(objConnection)
	Set pConnection = objConnection
End Property

Public Property Get PromoDelimiter
	PromoDelimiter = cstrPromoDelimiter
End Property

Public Property Get Message
	Message = pstrMessage
End Property

Public Sub OutputMessage

dim i
dim aError

	aError = split(pstrMessage,";")
	for i = 0 to ubound(aError)
		if pblnError then
			Response.Write "<P align='center'><H4><FONT color=Red>" & aError(i) & "</FONT></H4></P>"	
		else
			Response.Write "<P align='center'><H4>" & aError(i) & "</H4></P>"	
		end if
	next 'i

End Sub	'OutputMessage

Public Property Get PromotionID
	PromotionID = pPromotionID
End Property

Public Property Let StartDate(dtStartDate)
	pStartDate = StartDate
End Property

Public Property Get StartDate
	StartDate = pStartDate
End Property

Public Property Let EndDate(dtEndDate)
	pEndDate = EndDate
End Property

Public Property Get EndDate
	EndDate = pEndDate
End Property

Public Property Let Duration(lngDuration)
	ppDuration = pDuration
End Property

Public Property Get Duration
	Duration = pDuration
End Property

Private Function WrapString(strToWrap,blnWrap)

dim strTemp

	if len(strToWrap) = 0 then Exit Function
	strTemp = strToWrap
	if blnWrap then
		strTemp = Replace(strToWrap, ",", cstrPromoDelimiter)
		strTemp = Replace(strTemp, " ", "")
		
		'if left(strTemp,1) <> cstrPromoDelimiter then strTemp = cstrPromoDelimiter & strTemp
		'if right(strTemp,1) <> cstrPromoDelimiter then strTemp = strTemp & cstrPromoDelimiter
	else
		'if left(strTemp,1) = cstrPromoDelimiter then strTemp = right(strTemp,len(strTemp)-1)
		'if right(strTemp,1) = cstrPromoDelimiter then strTemp = left(strTemp,len(strTemp)-1)
	end if

	WrapString = strTemp
	
End Function

Public Property Let PromoCode(strPromoCode)
	pPromoCode = strPromoCode
End Property

Public Property Get PromoCode
	PromoCode = pPromoCode
End Property

Public Property Let NumUses(lngNumUses)
	pNumUses = NumUses
End Property

Public Property Get NumUses
	NumUses = pNumUses
End Property

Public Property Let NumUsesByCustomer(lngNumUsesByCustomer)
	NumUsesByCustomer = lngNumUsesByCustomer
End Property

Public Property Get NumUsesByCustomer
	NumUsesByCustomer = pNumUsesByCustomer
End Property

Public Property Let MaxUses(lngMaxUses)
	pMaxUses = MaxUses
End Property

Public Property Get MaxUses
	MaxUses = pMaxUses
End Property

Public Property Let Discount(lngDiscount)
	pDiscount = Discount
End Property

Public Property Get Discount
	Discount = pDiscount
End Property

Public Property Let MaxAllowableValue(byVal vntValue)
	pMaxAllowableValue = vntValue
End Property

Public Property Get MaxAllowableValue
	MaxAllowableValue = pMaxAllowableValue
End Property

Public Property Let MaxAllowableValuePerItem(byVal vntValue)
	pMaxAllowableValuePerItem = vntValue
End Property

Public Property Get MaxAllowableValuePerItem
	MaxAllowableValuePerItem = pMaxAllowableValuePerItem
End Property

Public Property Let Percentage(blnPercentage)
	pblnPercentage = Percentage
End Property

Public Property Get Percentage
	Percentage = pblnPercentage
End Property

Public Property Let ApplyToBasePrice(blnApplyToBasePrice)
	pblnApplyToBasePrice = ApplyToBasePrice
End Property

Public Property Get ApplyToBasePrice
	ApplyToBasePrice = pblnApplyToBasePrice
End Property

Public Property Let offerFreeGiftAutomatically(blnofferFreeGiftAutomatically)
	pblnofferFreeGiftAutomatically = offerFreeGiftAutomatically
End Property

Public Property Get offerFreeGiftAutomatically
	offerFreeGiftAutomatically = pblnofferFreeGiftAutomatically
End Property

Public Property Let MinSubTotal(curSubTotal)
	pMinSubTotal = MinSubTotal
End Property

Public Property Get MinSubTotal
	MinSubTotal = pMinSubTotal
End Property

Public Property Let Combineable(blnCombineable)
	pCombineable = Combineable
End Property

Public Property Get Combineable
	Combineable = pCombineable
End Property

Public Property Let PromoTitle(strPromoTitle)
	pPromoTitle = PromoTitle
End Property

Public Property Get PromoTitle
	PromoTitle = pPromoTitle
End Property

Public Property Let PromoRules(strPromoRules)
	pPromoRules = PromoTitle
End Property

Public Property Get PromoRules
	PromoRules = pPromoRules
End Property

Public Property Let FreeShippingCode(strID)
	pstrFreeShippingCode = strID
End Property
Public Property Get FreeShippingCode
	FreeShippingCode = pstrFreeShippingCode
End Property

Public Property Let ProductID(strID)
	pstrProductID = strID
End Property
Public Property Get ProductID
	ProductID = pstrProductID
End Property
Public Property Let FreeProductID(strID)
	pstrFreeProductID = strID
End Property
Public Property Get FreeProductID
	FreeProductID = pstrFreeProductID
End Property

Public Property Let ProductIDExclusion(strID)
	pstrProductIDExclusion = strID
End Property
Public Property Get ProductIDExclusion
	ProductIDExclusion = pstrProductIDExclusion
End Property

Public Property Let Category(strID)
	pstrCategory = strID
End Property
Public Property Get Category
	Category = pstrCategory
End Property
Public Property Let CategoryExclusion(strID)
	pstrCategoryExclusion = strID
End Property
Public Property Get CategoryExclusion
	CategoryExclusion = pstrCategoryExclusion
End Property

Public Property Let Manufacturer(strID)
	pstrManufacturer = strID
End Property
Public Property Get Manufacturer
	Manufacturer = pstrManufacturer
End Property
Public Property Let ManufacturerExclusion(strID)
	pstrManufacturerExclusion = strID
End Property
Public Property Get ManufacturerExclusion
	ManufacturerExclusion = pstrManufacturerExclusion
End Property

Public Property Let Vendor(strID)
	pstrVendor = strID
End Property
Public Property Get Vendor
	Vendor = pstrVendor
End Property
Public Property Let VendorExclusion(strID)
	pstrVendorExclusion = strID
End Property
Public Property Get VendorExclusion
	VendorExclusion = pstrVendorExclusion
End Property

Public Property Let ProductCountLimit(strID)
	pstrProductCountLimit = strID
End Property
Public Property Get ProductCountLimit
	ProductCountLimit = pstrProductCountLimit
End Property

Public Property Let BuyX(lngID)
	plngBuyX = lngID
End Property
Public Property Get BuyX
	BuyX = plngBuyX
End Property

Public Property Let GetY(lngID)
	plngGetY = lngID
End Property
Public Property Get GetY
	GetY = plngGetY
End Property

Public Property Let FreeShippingLimit(dblID)
	pdblFreeShippingLimit = dblID
End Property
Public Property Get FreeShippingLimit
	FreeShippingLimit = pdblFreeShippingLimit
End Property

Public Property Let LikeItem(blnID)
	pblnLikeItem = blnID
End Property
Public Property Get LikeItem
	LikeItem = pblnLikeItem
End Property


Public Property Let Inactive(blnInactive)
	pInactive = blnInactive
End Property

Public Property Get Inactive
	Inactive = pInactive
End Property

Public Property Let ApplyAutomatically(blnApplyAutomatically)
	pApplyAutomatically = blnApplyAutomatically
End Property

Public Property Get ApplyAutomatically
	ApplyAutomatically = pApplyAutomatically
End Property

Public Property Let ExcludeSaleItems(blnExcludeSaleItems)
	pblnExcludeSaleItems = blnExcludeSaleItems
End Property

Public Property Get ExcludeSaleItems
	ExcludeSaleItems = pblnExcludeSaleItems
End Property

Public Property Get ModifiedOn
	ModifiedOn = pModifiedOn
End Property

Public Property Get Promotion

dim pstrPromotion

	pstrPromotion = pDiscount
	if pblnPercentage then
		pstrPromotion = pstrPromotion & "%"
	else
		pstrPromotion = "$" & pstrPromotion
	end if
	pstrPromotion = pstrPromotion & " off purchase on orders of $" & pMinSubTotal & " or more."
	Promotion = pstrPromotion
	
End Property

'***********************************************************************************************

Public Function Clone(byVal strCloneCodes)

Dim i, j
Dim prsSource
Dim prsTarget
Dim paryCodes
Dim pstrDelimiter
Dim pstrField
Dim pstrNewCode
Dim pstrResult

	If InStr(1, strCloneCodes, ",") Then
		pstrDelimiter = ","
	ElseIf InStr(1, strCloneCodes, ";") Then
		pstrDelimiter = ";"
	Else
		pstrDelimiter = ","
	End If
	paryCodes = Split(Trim(strCloneCodes), pstrDelimiter)
	
	Set prsSource = server.CreateObject("adodb.Recordset")
	prsSource.open "Select * from Promotions where PromotionID=" & pPromotionID, cnn, 1, 3

	For i = 0 To UBound(paryCodes)
		pstrNewCode = Trim(paryCodes(i))
		If Len(pstrNewCode) > 0 Then
			Set prsTarget = server.CreateObject("adodb.Recordset")
			prsTarget.open "Select * from Promotions where PromoCode = '" & pstrNewCode & "'", cnn, 1, 3
			If prsTarget.EOF Then
				prsTarget.AddNew
			
				For j=0 to prsSource.fields.count-1
					pstrField = prsTarget.fields(j).name
					'skip autonumber and timestamps
					If LCase(pstrField) <> "promotionid" And prsSource.Fields(pstrField).Type <> 128 Then
						'On Error Resume Next
						prsTarget.Fields(pstrField).value = prsSource.Fields(pstrField).value
						If Err.number <> 0 Then
							Response.Write "<fieldset><legend>Error</legend>" _
										& "Error " & Err.number & ": " & Err.Description & "<br />" _
										& "Field: " & prsSource(pstrField).Name & "<br />" _
										& "Type : " & prsSource(pstrField).Type & "<br />" _
										& "prsSource(pstrField): " & prsSource(pstrField) & "<br />" _
										& "</fieldset>"
							Err.Clear
						End If
					Else
						'Response.Write "<font color=red>" & pstrField & " was skipped</font><br />"
					End If
				Next
				prsTarget.Fields("PromoCode").value = pstrNewCode
				prsTarget.Update
			
				pstrResult = pstrResult & "<li><strong>" & pstrNewCode & "</strong> was successfully cloned.</font></li>"
			Else
				pstrResult = pstrResult & "<li><font color=red><strong>" & pstrNewCode & "</strong> could not be created as it already exists!</font></li>"
			End If
			prsTarget.Close
			Set prsTarget = Nothing
		End If
		'Response.Write i & ": " &  & "<br />"
	
	Next 'i
	
	pstrResult = "<ul>" & pstrResult & "</ul>"
	Response.Write pstrResult
	
End Function	'Clone

'***********************************************************************************************

Public Function FindByPromoCode(strPromoCode)

	with prsPromo
		.Find "PromoCode = '" & strPromoCode & "'"
		if not .eof then 
			LoadValues(prsPromo)
			FindByPromoCode = True
		end if
	end with
	
End Function

'***********************************************************************************************

Public Function FindByPromoID(lngPromoID)

	if len(lngPromoID)=0 or lngPromoID=0 then Exit Function
	with prsPromo
		.Find "PromotionID = " & lngPromoID
		if not .eof then 
			LoadValues(prsPromo)
			FindByPromoID = True
		end if
	end with
	
End Function

'***********************************************************************************************

Public Function LoadByPromoCode(strPromoCode)

dim sql
dim rs

	On Error Resume Next

	sql = "Select * from Promotions where PromoCode = '" & strPromoCode & "'"
	set rs = server.CreateObject("adodb.recordset")
	set rs = pConnection.Execute(sql)
	
	
	if not (rs.EOF or rs.BOF) then 
		LoadValues(rs)
		LoadByPromoCode = True
	end if
	rs.Close
	set rs = Nothing
	
End Function	'LoadByPromoCode

'***********************************************************************************************

Public Function LoadByPromoID(lngPromoID)

dim sql
dim rs

	On Error Resume Next

	sql = "Select * from Promotions where PromotionID = " & lngPromoID
	set rs = server.CreateObject("adodb.recordset")
	set rs = pConnection.Execute(sql)
	
	
	if not (rs.EOF or rs.BOF) then 
		LoadValues(rs)
		LoadByPromoID = True
	end if
	rs.Close
	set rs = Nothing
	
End Function	'LoadByPromoID

Public Function LoadAll

	On Error Resume Next

	set prsPromo = server.CreateObject("adodb.recordset")
	with prsPromo
		.ActiveConnection = pConnection
		.CursorLocation = 2 'adUseClient
		.CursorType = 3 'adOpenStatic
		.LockType = 1 'adLockReadOnly
		.Source = "Select * from Promotions"  & mstrsqlWhere
		.Open
	end with

	If Err.number <> 0 Then
		Response.Write "<h3><font color=red>The Promotion Manager 2 add-on upgrade does not appear to have been performed.</font>&nbsp;&nbsp;"
		Response.Write "<a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=PromoMgrII'>Click here to upgrade</a></h3>"
		Response.Write "<h4>Error " & Err.number & ": " & Err.Description & "</h4>"
		Response.Write "<h4>SQL: " & pstrSQL & "</h4>"
		Response.Flush
		Err.Clear
		Load = False
		Exit Function
	ElseIf prsPromo.EOF Then
		If Not isCorrectVersion Then
			Response.Write "<h3><font color=red>The Promotion Manager 2 add-on upgrade does not appear to have been performed.</font>&nbsp;&nbsp;"
			Response.Write "<a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=PromoMgrII'>Click here to upgrade</a></h3>"
			Response.Write "<h4>Error " & Err.number & ": " & Err.Description & "</h4>"
			Response.Write "<h4>SQL: " & pstrSQL & "</h4>"
			Response.Flush
		End If
		LoadAll = False
	Else
		Call LoadValues(prsPromo)
		LoadAll = (err.number = 0)
	End If

End Function	'LoadAll

'***********************************************************************************************

Private Function isCorrectVersion

Dim pobjrsTest

	On Error Resume Next

	set pobjrsTest = server.CreateObject("adodb.recordset")
	with pobjrsTest
		.ActiveConnection = pConnection
		.CursorLocation = 2 'adUseClient
		.CursorType = 3 'adOpenStatic
		.LockType = 1 'adLockReadOnly
		.Source = "Select FreeShippingCode from Promotions"
		.Open
		If Err.number <> 0 Then
			isCorrectVersion = False
		Else
			isCorrectVersion = True
		End If
		.Close
	end with
	Set pobjrsTest = Nothing

End Function	'isCorrectVersion

'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

    With prsPromo
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "PromotionID=" & lngID
            Else
                .MoveLast
            End If
            If Not .EOF Then LoadValues(prsPromo)
        End If
    End With

End Function    'Find

'***********************************************************************************************

Public Function Delete(lngPromotionID)

dim sql

On Error Resume Next

	if len(lngPromotionID)=0 or lngPromotionID="0" then Exit Function

	if LoadByPromoID(lngPromotionID) then	
		sql = "Delete from Promotions where PromotionID = " & lngPromotionID
		pConnection.Execute sql,,128
		pstrMessage = pPromoTitle & " was successfully deleted."
	End If
	
	Delete = (Err.number=0)
	
End Function

'***********************************************************************************************

Public Function Update

dim strErrorMessage
dim sql
dim rs
dim blnAdd

	Call LoadFromRequest
	strErrorMessage = ValidateValues
	If len(strErrorMessage) = 0 then
		if len(pPromotionID) = 0 then pPromotionID = 0
		set rs = server.CreateObject("adodb.recordset")
		sql = "Select * from Promotions where PromotionID=" & pPromotionID
		rs.open sql, pConnection, 1,	3
		
		blnAdd = False
		If rs.eof then 
			rs.addnew
			blnAdd = True
		End IF

		rs("StartDate") = pStartDate
    	if len(pEndDate)<> 0 then 
    		rs("EndDate") = pEndDate
		else
    		rs("EndDate") = Null
    	end if
    	if len(pDuration)<> 0 then 
    		rs("Duration") = pDuration
		else
    		rs("Duration") = Null
		end if
    	rs("PromoCode") = pPromoCode
    	if len(pNumUses)<> 0 then 
    		rs("NumUses") = pNumUses
		else
    		rs("NumUses") = Null
		end if
    	if len(pNumUsesByCustomer)<> 0 then 
    		rs("NumUsesByCustomer") = pNumUsesByCustomer
		else
    		rs("NumUsesByCustomer") = Null
		end if
    	if len(pMaxUses)<> 0 then 
    		rs("MaxUses") = pMaxUses
		else
    		rs("MaxUses") = Null
		end if
    	rs("Discount") = pDiscount
    	rs.Fields("MaxAllowableValue").Value = pMaxAllowableValue
    	rs.Fields("MaxAllowableValuePerItem").Value = pMaxAllowableValuePerItem
    	rs("Percentage") = pblnPercentage
    	rs("ApplyToBasePrice") = pblnApplyToBasePrice
    	rs("offerFreeGiftAutomatically") = pblnofferFreeGiftAutomatically
    	rs("MinSubTotal") = pMinSubTotal
    	rs("Combineable") = pCombineable
    	rs("PromoTitle") = pPromoTitle
    	rs("ModifiedOn") = Now()
    	if len(pPromoRules)<> 0 then 
    		rs("PromoRules") = pPromoRules
		else
    		rs("PromoRules") = Null
		end if
		
		rs.Fields("FreeShippingCode").Value = wrapSQLValue(pstrFreeShippingCode, False, enDatatype_NA)
		rs.Fields("ProductID").Value = wrapSQLValue(pstrProductID, False, enDatatype_NA)
		rs.Fields("FreeProductID").Value = wrapSQLValue(pstrFreeProductID, False, enDatatype_NA)
		rs.Fields("ProductIDExclusion").Value = wrapSQLValue(pstrProductIDExclusion, False, enDatatype_NA)
		rs.Fields("Category").Value = wrapSQLValue(pstrCategory, False, enDatatype_NA)
		rs.Fields("CategoryExclusion").Value = wrapSQLValue(pstrCategoryExclusion, False, enDatatype_NA)
		rs.Fields("Manufacturer").Value = wrapSQLValue(pstrManufacturer, False, enDatatype_NA)
		rs.Fields("ManufacturerExclusion").Value = wrapSQLValue(pstrManufacturerExclusion, False, enDatatype_NA)
		rs.Fields("Vendor").Value = wrapSQLValue(pstrVendor, False, enDatatype_NA)
		rs.Fields("VendorExclusion").Value = wrapSQLValue(pstrVendorExclusion, False, enDatatype_NA)

		rs.Fields("ProductCountLimit").Value = wrapSQLValue(pstrProductCountLimit, False, enDatatype_NA)
		rs.Fields("BuyX").Value = wrapSQLValue(plngBuyX, False, enDatatype_NA)
		rs.Fields("GetY").Value = wrapSQLValue(plngGetY, False, enDatatype_NA)
		rs.Fields("FreeShippingLimit").Value = wrapSQLValue(pdblFreeShippingLimit, False, enDatatype_NA)
		rs.Fields("LikeItem").Value = wrapSQLValue(pblnLikeItem, False, enDatatype_boolean)

    	rs("Inactive") = pInactive
    	rs("ApplyAutomatically") = pApplyAutomatically
    	rs("ExcludeSaleItems") = pblnExcludeSaleItems

		On Error Resume Next

    	rs.Update
   	
    	if Err.number = -2147217887 then
    		if Err.description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." then
				pstrMessage = "<H4>The Promo Code you entered is already in use.<br />Please enter a different code.</H4><br />"	
				pblnError = True
    		end if
    	elseif Err.number = -2147217873 then 
			pstrMessage = "<H4>The Promo Code you entered is already in use.<br />Please enter a different code.</H4><br />"	
			pblnError = True
    	elseif Err.number <> 0 then 
			Response.Write "Error: " & Err.number & " - " & Err.description & "<br />"
'    		Err.Raise Err.number,"modPromo.Update",Err.description
    	end if
    	
    	pPromotionID = rs("PromotionID")
    	rs.close
    	set rs = nothing
    	
   		if Err.number=0 then
    		if  blnAdd then
    			pstrMessage = pPromoTitle & " was successfully added."
    		else
    			pstrMessage = "The changes to " & pPromoTitle & " were successfully saved."
    		end if
    	end if
	End If
	
	Update = strErrorMessage
	
End Function 'Update

'***********************************************************************************************

Private Sub LoadValues(rs)

	pPromotionID = rs("PromotionID")
    pStartDate = rs("StartDate")
    pEndDate = rs("EndDate")
    pDuration = rs("Duration")
    pPromoCode = trim(rs("PromoCode"))
    pNumUses = rs("NumUses")
    pNumUsesByCustomer = rs("NumUsesByCustomer")
    pMaxUses = rs("MaxUses")
    pDiscount = rs("Discount")
    pMaxAllowableValue = rs.Fields("MaxAllowableValue").Value
    pMaxAllowableValuePerItem = rs.Fields("MaxAllowableValuePerItem").Value
    pblnPercentage = rs("Percentage")
    pblnApplyToBasePrice = rs("ApplyToBasePrice")
    pMinSubTotal = rs("MinSubTotal")
    pCombineable = rs("Combineable")
    pPromoTitle = trim(rs("PromoTitle"))
    pModifiedOn = rs("ModifiedOn")
    pPromoRules = trim(rs("PromoRules"))
	pInactive = rs("Inactive")
	pApplyAutomatically = rs("ApplyAutomatically")
	pblnExcludeSaleItems = rs("ExcludeSaleItems")
	pstrProductID = trim(rs.Fields("ProductID").Value)

	On Error Resume Next
	pstrFreeShippingCode = trim(rs.Fields("FreeShippingCode").Value)
	If Err.number <> 0 Then
		Response.Write "<h3><font color=red>The Promotion Manager 2 add-on upgrade does not appear to have been performed.</font>&nbsp;&nbsp;"
		Response.Write "<a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=PromoMgrII'>Click here to upgrade</a></h3>"
		Response.Write "<h4>Error " & Err.number & ": " & Err.Description & "</h4>"
		Response.Write "<h4>SQL: " & pstrSQL & "</h4>"
		Response.Flush
		Err.Clear
	End If

	pstrFreeProductID = trim(rs.Fields("FreeProductID").Value)
	pstrProductIDExclusion = trim(rs.Fields("ProductIDExclusion").Value)
	pstrCategory = trim(rs.Fields("Category").Value)
	pstrCategoryExclusion = trim(rs.Fields("CategoryExclusion").Value)
	pstrManufacturer = trim(rs.Fields("Manufacturer").Value)
	pstrManufacturerExclusion = trim(rs.Fields("ManufacturerExclusion").Value)
	pstrVendor = trim(rs.Fields("Vendor").Value)
	pstrVendorExclusion = trim(rs.Fields("VendorExclusion").Value)

    pblnofferFreeGiftAutomatically = rs("offerFreeGiftAutomatically")
	pstrProductCountLimit = trim(rs.Fields("ProductCountLimit").Value)
	plngBuyX = trim(rs.Fields("BuyX").Value)
	plngGetY = trim(rs.Fields("GetY").Value)
	pdblFreeShippingLimit = trim(rs.Fields("FreeShippingLimit").Value)
	pblnLikeItem = trim(rs.Fields("LikeItem").Value)

	If Err.number <> 0 Then Err.Clear
    
End Sub

Private Sub LoadFromRequest

	with Request.Form
		pPromotionID = trim(.Item("PromotionID"))
		pStartDate = trim(.Item("StartDate"))
		pEndDate = trim(.Item("EndDate"))
		pDuration = trim(.Item("Duration"))
		pPromoCode = trim(.Item("PromoCode"))
		pNumUses = trim(.Item("NumUses"))
		pNumUsesByCustomer = trim(.Item("NumUsesByCustomer"))
		pMaxUses = trim(.Item("MaxUses"))
		pDiscount = trim(.Item("Discount"))
		pMaxAllowableValue = trim(.Item("MaxAllowableValue"))
		pMaxAllowableValuePerItem = trim(.Item("MaxAllowableValuePerItem"))
		pblnPercentage = (LCase(.Item("Percentage")) = "on")
		pblnApplyToBasePrice = (LCase(.Item("ApplyToBasePrice")) = "on")
		pblnofferFreeGiftAutomatically = (LCase(.Item("offerFreeGiftAutomatically")) = "on")
		pMinSubTotal = trim(.Item("MinSubTotal"))
		pCombineable = (LCase(.Item("Combineable")) = "on")
		pPromoTitle = trim(.Item("PromoTitle"))
		pModifiedOn = trim(.Item("ModifiedOn"))
		pPromoRules = trim(.Item("PromoRules"))
		pInactive =  (.Item("Inactive") = "on")
		pApplyAutomatically =  (LCase(.Item("ApplyAutomatically")) = "on")
		pblnExcludeSaleItems =  (LCase(.Item("ExcludeSaleItems")) = "on")
		pstrFreeShippingCode =  WrapString(trim(.Item("FreeShippingCode")),True)
		pstrProductID =  WrapString(trim(.Item("ProductID")),True)
		pstrFreeProductID =  WrapString(trim(.Item("FreeProductID")),True)
		pstrProductIDExclusion =  WrapString(trim(.Item("ProductIDExclusion")),True)
		pstrCategory =  WrapString(trim(.Item("Category")),True)
		pstrCategoryExclusion =  WrapString(trim(.Item("CategoryExclusion")),True)
		pstrManufacturer =  WrapString(trim(.Item("Manufacturer")),True)
		pstrManufacturerExclusion =  WrapString(trim(.Item("ManufacturerExclusion")),True)
		pstrVendor =  WrapString(trim(.Item("Vendor")),True)
		pstrVendorExclusion =  WrapString(trim(.Item("VendorExclusion")),True)

		pstrProductCountLimit = trim(.Item("ProductCountLimit"))
		plngBuyX = trim(.Item("BuyX"))
		plngGetY = trim(.Item("GetY"))
		pdblFreeShippingLimit = trim(.Item("FreeShippingLimit"))
		pblnLikeItem =  (.Item("LikeItem") = "on")
		
	end with
	
	'Now force lower case since script compares against lower case
	pstrProductID = LCase(pstrProductID)
	pstrProductIDExclusion = LCase(pstrProductIDExclusion)
	pstrFreeProductID = LCase(pstrFreeProductID)
	
End Sub	'LoadFromRequest

Private Function ValidateValues

dim strError

	if len(pPromoTitle)=0 then
		strError = strError & "Please enter a title for the promotion;"
	end if
	if len(pPromoCode)=0 then
		strError = strError & "Please enter a promotion code;"
	end if
	if not isDate(pStartDate) then
		strError = strError & "Please enter a valid date for the Start Date;"
	end if
	if not isDate(pEndDate) and len(pEndDate)<>0 then
		strError = strError & "Please enter a valid date for the End Date;"
	end if
	if not isNumeric(pDiscount) and len(pDiscount)<>0 then
		strError = strError & "Please enter a number for the Discount amount;"
	elseif len(pDiscount)=0 then
		strError = strError & "Please enter a Discount amount;"
	end if
	if not isNumeric(pDuration) and len(pDuration)<>0 then
		strError = strError & "Please enter a number for the Duration;"
	end if
	if not isNumeric(pNumUses) and len(pNumUses)<>0 then
		strError = strError & "Please enter a number for the number of uses;"
	end if
	if not isNumeric(pNumUsesByCustomer) and len(pNumUsesByCustomer)<>0 then
		strError = strError & "Please enter a number for the number of uses;"
	end if
	if not isNumeric(pMaxUses) and len(pMaxUses)<>0 then
		strError = strError & "Please enter a number for the maximum number of uses;"
	end if
	if not isNumeric(pMinSubTotal) and len(pMinSubTotal)<>0 then
		strError = strError & "Please enter a number for the minimum subTotal;"
	elseif len(pMinSubTotal)=0 then
		strError = strError & "Please enter a minimum subTotal;"
	end if

	pstrMessage = strError
	pblnError = (len(strError)<>0)
	ValidateValues = strError

End Function	'ValidateValues

'***********************************************************************************************

Public Sub OutputSummary()

Dim i
Dim aSortHeader(3,1)
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr

	With Response

    .Write "<table class='tbl' id='tblSummary' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<COLGROUP align='center' width='10%'>"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='40%'>"
    .Write "<COLGROUP align='center' width='30%'>"
    .Write "  <tr class='tblhdr'>"

	If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort by Code in descending order"
		aSortHeader(2,0) = "Sort by Title in descending order"
		aSortHeader(3,0) = "Sort by Promotion in descending order"
		pstrSortOrder = "ASC"
	Else
		aSortHeader(1,0) = "Sort by Code in ascending order"
		aSortHeader(2,0) = "Sort by Title in ascending order"
		aSortHeader(3,0) = "Sort by Promotion  Sites in ascending order"
		pstrSortOrder = "DESC"
	End If
	aSortHeader(1,1) = "&nbsp;&nbsp;Code"
	aSortHeader(2,1) = "Title"
	aSortHeader(3,1) = "Promotion"

	if len(mstrOrderBy) > 0 Then
		pstrOrderBy = mstrOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
	.Write "<TH>&nbsp;</TH>"
	for i = 1 to 3
		If cInt(pstrOrderBy) = i Then
			If (pstrSortOrder = "ASC") Then
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'DESC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
								"&nbsp;&nbsp;&nbsp;<img src='Images/up.gif' border=0 align=bottom></TH>" & vbCrLf
			Else
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'ASC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
								"&nbsp;&nbsp;&nbsp;<img src='Images/down.gif' border=0 align=bottom></TH>" & vbCrLf
			End If
		Else
		    .Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'" & pstrSortOrder & "');" & chr(34) & _
							" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
							" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & "</TH>" & vbCrLf
		End If
	next 'i

    .Write "  </tr>"
	.Write "<tr><td colspan=4>"
    .Write "<div name='divSummary' style='height:400; overflow:scroll;'>"
	.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' " _
				 & ">"
    .Write "<COLGROUP align='center' width='10%'>"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='40%'>"
    .Write "<COLGROUP align='center' width='30%'>"
				  				  
	If prsPromo.RecordCount > 0 Then
        prsPromo.MoveFirst
        For i = 1 To prsPromo.RecordCount
			pstrAbbr = Trim(prsPromo("PromotionID"))
 			pstrTitle = "Click to edit " & prsPromo("PromoCode") & "."
			pstrURL = "ssPromotionsAdmin.asp?Action=View&PromotionID=" & pstrAbbr

			if cLng(pstrAbbr) = cLng(PromotionID) then
        		.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
				.Write "    <TD><A href='ssPromotionsReport.asp?PromoCode=" & prsPromo("PromoCode") & "'><img src='Images/Note12.ico' height='8' width='6'></A><br />" & vbcrlf
				.Write "    <TD>" & prsPromo("PromoCode") & "</TD>" & vbcrlf
			else
				.Write "<TR title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
				.Write "    <TD><A href='ssPromotionsReport.asp?PromoCode=" & prsPromo("PromoCode") & "'><img src='Images/Note12.ico' height='8' width='6'></A><br />" & vbcrlf
				.Write "    <TD><A href='ssPromotionsAdmin.asp?Action=View&PromotionID=" & prsPromo("PromotionID") & "'>" & prsPromo("PromoCode") & "</A></TD>" & vbcrlf
        	end if
			.Write "    <TD>&nbsp;" & prsPromo("PromoTitle") & "</TD>" & vbcrlf
			.Write "    <TD>&nbsp;" & prsPromo("PromoRules") & "</TD>" & vbcrlf

            Response.Write "</TR>" & vbCrLf
            prsPromo.MoveNext
        Next
    Else
        Response.Write "<TR><TD align=center><h3>There are no Promotions</h3></TD></TR>"
    End If
    .Write "</td></tr></TABLE></div>"
    .Write "</TABLE>"
	End With
	
End Sub      'OutputSummary

Private Sub writeFields

Dim i

	debugprint "count", prsPromo.Fields.Count
	For i = 0 To prsPromo.Fields.Count-1
		debugprint i, prsPromo.Fields(i).Name
	Next 'i

End Sub	'writeFields

End Class	'clsPromotion

'***********************************************************************************************

%>
<!--#include file="SSLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

Function SummaryFilter

Dim pstrOrderBy
Dim pstrsqlWhere

	'load the text filter
	mbytText_Filter = Request.Form("optText_Filter")
	mstrText_Filter = Request.Form("Text_Filter")
	If len(mstrText_Filter) > 0 Then
		Select Case mbytText_Filter
			Case "0"	'Do Not Include
			Case "1"	'Code
				pstrsqlWhere = pstrsqlWhere & " AND (PromoCode Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "2"	'Name
				pstrsqlWhere = pstrsqlWhere & " AND (PromoTitle Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "3"	'Message
				pstrsqlWhere = pstrsqlWhere & " AND (PromoRules Like '%" & mstrText_Filter & "%')"
		End Select	
	End If

	mbytradStartDate = Request.Form("radStartDate")
	If len(mbytradStartDate) = 0 Then mbytradStartDate = 0	
	mstrfilterStartDate = Request.Form("filterStartDate")
	If len(mstrfilterStartDate) > 0 then 
		If mbytradStartDate = 1 Then	'after
			pstrsqlWhere = pstrsqlWhere & " AND (StartDate>=" & wrapSQLValue(mstrfilterStartDate & " 12:00:00 AM", False, enDatatype_date)
		Else
			pstrsqlWhere = pstrsqlWhere & " AND (StartDate<=" & wrapSQLValue(mstrfilterStartDate & " 11:59:59 PM", False, enDatatype_date)
		End If
	End If
	
	mbytradEndDate = Request.Form("radEndDate")
	If len(mbytradEndDate) = 0 Then mbytradEndDate = 0	
	mstrfilterEndDate = Request.Form("filterEndDate")
	If len(mstrfilterEndDate) > 0 then 
		If mbytradEndDate = 1 Then	'after
			pstrsqlWhere = pstrsqlWhere & " AND (EndDate>=" & wrapSQLValue(mstrfilterEndDate & " 12:00:00 AM", False, enDatatype_date)
		Else
			pstrsqlWhere = pstrsqlWhere & " AND (EndDate<=" & wrapSQLValue(mstrfilterEndDate & " 11:59:59 PM", False, enDatatype_date)
		End If
	End If

	'load the radio filters and set the defaults
	mbytActive_Filter = Request.Form("optActive_Filter")
	If len(mbytActive_Filter) = 0 Then mbytActive_Filter = 1
	mbytApplyAuto_Filter = Request.Form("optApplyAuto_Filter")
	If len(mbytApplyAuto_Filter) = 0 Then mbytApplyAuto_Filter = 0
	mbytExclusive_Filter	= Request.Form("optExclusive_Filter")
	If len(mbytExclusive_Filter) = 0 Then mbytExclusive_Filter = 0

	Select Case mbytApplyAuto_Filter
		Case "0"	'Do Not Include
		Case "1"	'Active
			pstrsqlWhere = pstrsqlWhere & " AND ApplyAutomatically is Not Null"
		Case "2"	'Inactive
			pstrsqlWhere = pstrsqlWhere & " AND ApplyAutomatically is Null"
	End Select	

	Select Case mbytExclusive_Filter
		Case "0"	'Do Not Include
		Case "1"	'Flagged
			pstrsqlWhere = pstrsqlWhere & " AND (Combineable=1 Or Combineable=-1)"
		Case "2"	'unflagged
			pstrsqlWhere = pstrsqlWhere & " AND (Combineable is Null OR Combineable=0)"
	End Select	

	Select Case mbytActive_Filter
		Case "0"	'Do Not Include
		Case "1"	'Active
			pstrsqlWhere = pstrsqlWhere & " AND (Inactive is Null OR Inactive=0)"
		Case "2"	'Inactive
			pstrsqlWhere = pstrsqlWhere & " AND (Inactive=1 Or Inactive=-1)"
	End Select	
	
	'Build  the Order By
	mstrOrderBy = Request.Form("OrderBy")
	If len(mstrOrderBy) = 0 Then mstrOrderBy = 0
	
	mstrSortOrder = Request.Form("SortOrder")
	If len(mstrSortOrder) = 0 Then mstrSortOrder = "Desc"

	dim paryOrderBy(3)
	paryOrderBy(0) = "PromoCode"	'Default
	paryOrderBy(1) = "PromoCode"
	paryOrderBy(2) = "PromoTitle"
	paryOrderBy(3) = "PromoRules"

	pstrOrderBy = " Order By " & paryOrderBy(mstrOrderBy) & " " & mstrSortOrder 

	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	
	pstrsqlWhere = Trim(pstrsqlWhere)
	If Left(pstrsqlWhere,3) = "AND" Then pstrsqlWhere = Replace(pstrsqlWhere, "AND", "", 1, 1)
	
	If Len(pstrsqlWhere) > 0 Then pstrsqlWhere = " Where " & pstrsqlWhere
	
	If len(pstrOrderBy) > 0 then pstrsqlWhere = pstrsqlWhere & pstrOrderBy
	
	SummaryFilter = pstrsqlWhere
	
End Function    'SummaryFilter

'***********************************************************************************************
'***********************************************************************************************

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsPromo
Dim mblnShowSummary
Dim mstrOrderBy
Dim mstrSortOrder
Dim mstrsqlWhere
Dim mbytText_Filter
Dim mstrText_Filter
Dim mbytradStartDate
Dim mstrfilterStartDate
Dim mbytradEndDate
Dim mstrfilterEndDate
Dim mbytActive_Filter
Dim mbytApplyAuto_Filter
Dim mbytExclusive_Filter
Dim mstrShow
Dim mblnShowDetail
Dim mblnShowFilter
Dim mbytSummaryTableHeight

Const cblnUsePicker = True

Dim mvntID

mstrPageTitle = "Promotion Administration"

	mstrsqlWhere = SummaryFilter

	mvntID = LoadRequestValue("PromotionID")
	mAction = LoadRequestValue("Action")
    
	mstrShow = Request.Form("Show")
	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	
	set mclsPromo = new clsPromotion
    mclsPromo.Connection = cnn
    
    Select Case mAction
        Case "New", "Update"
            mclsPromo.Update
            If mclsPromo.LoadAll Then mclsPromo.Find mvntID
        Case "Clone"
            If mclsPromo.LoadAll Then
				mclsPromo.Find mvntID
				mclsPromo.Clone Request.Form("CloneCodes")
			End If
        Case "Delete"
            mclsPromo.Delete mvntID
            mclsPromo.LoadAll
            mblnShowSummary = True
            mblnShowFilter = False
        Case "View"
            If mclsPromo.LoadAll Then mclsPromo.Find mvntID
        Case "Activate", "Deactivate"
            mclsPromo.Activate mvntID, mAction= "Activate"
            If mclsPromo.LoadAll Then mclsPromo.Find mvntID
            mblnShowSummary = True
            mblnShowFilter = False
        Case Else
            mclsPromo.LoadAll
            mblnShowSummary = True
            mblnShowFilter = False
    End Select
    
    'If Len(mclsPromo.PromoCode) = 0 Then mclsPromo.LoadAll
    
	Call WriteHeader("body_onload();",True)
    With mclsPromo

%>
<script LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></script>
<script LANGUAGE="JavaScript" SRC="SSLibrary/calendar.js"></script>
<script LANGUAGE="JavaScript">
<!--
var theDataForm;
var blnIsDirty;
var strModifiedOn;
blnIsDirty = false;

function body_onload()
{
	theDataForm = document.frmData;

	theKeyField = theDataForm.PromotionID;

<% If cblnUsePicker Then %>
//Fill product dictionary
<%= setCustomDictionary(.productID, .PromoDelimiter, "ProductID", "product") %>
<%= setCustomDictionary(.FreeProductID, .PromoDelimiter, "FreeProductID", "product") %>
<%= setCustomDictionary(.ProductIDExclusion, .PromoDelimiter, "ProductIDExclusion", "product") %>
<%= setCustomDictionary(.Category, .PromoDelimiter, "Category", "category") %>
<%= setCustomDictionary(.CategoryExclusion, .PromoDelimiter, "CategoryExclusion", "category") %>
<%= setCustomDictionary(.Manufacturer, .PromoDelimiter, "Manufacturer", "manufacturer") %>
<%= setCustomDictionary(.ManufacturerExclusion, .PromoDelimiter, "ManufacturerExclusion", "manufacturer") %>
<%= setCustomDictionary(.Vendor, .PromoDelimiter, "Vendor", "vendor") %>
<%= setCustomDictionary(.VendorExclusion, .PromoDelimiter, "VendorExclusion", "vendor") %>

FillItem("ProductID");
FillItem("FreeProductID");
FillItem("ProductIDExclusion");
FillItem("Category");
FillItem("CategoryExclusion");
FillItem("Manufacturer");
FillItem("ManufacturerExclusion");
FillItem("Vendor");
FillItem("VendorExclusion");
<% End If %>

<%
If mblnShowFilter Then
	Response.Write "DisplayMainSection('Summary');" & vbcrlf
ElseIf mblnShowFilter Then
	Response.Write "DisplayMainSection('Filter');" & vbcrlf
Else
	Response.Write "DisplayMainSection('itemDetail');" & vbcrlf
	Response.Write "ScrollToElem('selectedSummaryItem');" & vbcrlf
	'Response.Write "frmData.PromoCode.focus();" & vbcrlf
	
	'Response.Write "DisplaySection(" & chr(34) & mstrShow & chr(34) & ");"
End If
%>

}

function btnNew_onclick() 
{
var pDate = new Date();
var pstrDate = (pDate.getMonth()+1)+"/"+pDate.getDate()+"/"+pDate.getYear();

	theDataForm.Action.value = "New";
	theDataForm.PromotionID.value = 0;
	theDataForm.StartDate.value = pstrDate;
	theDataForm.EndDate.value = "";
	theDataForm.Duration.value = "";
	theDataForm.PromoCode.value = "";
	theDataForm.MaxUses.value = "";
	theDataForm.NumUses.value = "";
	theDataForm.NumUsesByCustomer.value = "";
	theDataForm.Discount.value = "0";
	theDataForm.MinSubTotal.value = "0";
	theDataForm.PromoTitle.value = "";
	theDataForm.PromoRules.value = "";
	strModifiedOn = document.getElementById("ModifiedOn").innerHTML;
	document.getElementById("ModifiedOn").innerHTML = "";
	theDataForm.Percentage.checked = false;
	theDataForm.ApplyToBasePrice.checked = false;
	theDataForm.offerFreeGiftAutomatically.checked = false;
	theDataForm.Combineable.checked = false;
	theDataForm.ApplyAutomatically.checked = false;
	theDataForm.ExcludeSaleItems.checked = false;

	theDataForm.btnUpdate.value = "Add Promotion";
	theDataForm.btnDelete.disabled = true;
	
	theDataForm.PromoCode.focus();

}

function btnDelete_onclick() 
{
var blnConfirm;

	blnConfirm = confirm("Are you sure you wish to delete " + theDataForm.PromoTitle.value + "?");
	if (blnConfirm)
	{
	theDataForm.Action.value = "Delete";
	theDataForm.submit();
	return(true);
	}
	else
	{
	return(false);
	}
}

function btnClone_onclick() 
{
var blnConfirm;

	blnConfirm = prompt("Please enter the codes you wish to create. Mulitple codes may be separated by a comma or semi-colon.");
	
	if (blnConfirm.length > 0)
	{
	theDataForm.Action.value = "Clone";
	theDataForm.CloneCodes.value = blnConfirm;
	theDataForm.submit();
	return(true);
	}
	else
	{
	return(false);
	}
}

function btnReset_onclick() 
{
	document.getElementById("ModifiedOn").innerHTML = strModifiedOn;
	theDataForm.Action.value = "Update";
	theDataForm.btnUpdate.value = "Save Changes";
	theDataForm.btnDelete.disabled = false;
}

function DisplaySection(strSection)
{
<% 'Response.Write "return false;" %>

<% 
Dim pstrTempHeaderRow

pstrTempHeaderRow = "'General'"
pstrTempHeaderRow = pstrTempHeaderRow & ",'Custom'"
pstrTempHeaderRow = pstrTempHeaderRow & ",'ProductSales'"
%>
var arySections = new Array(<%= pstrTempHeaderRow %>);

	frmData.Show.value = strSection;

	for (var i=0; i < arySections.length;i++)
	{
		if (arySections[i] == strSection)
		{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "";
				document.all("td" + arySections[i]).className = "hdrSelected";
			}
		}else{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "none";
				document.all("td" + arySections[i]).className = "hdrNonSelected";
			}
		}
	}
}

function DisplayMainSection(strSection)
{

	var arySections = new Array('Filter', 'Summary', 'itemDetail');

	for (var i=0; i < arySections.length;i++)
	{
		if (arySections[i] == strSection)
		{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "";
				document.all("td" + arySections[i]).className = "hdrSelected";
			}
		}else{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "none";
				document.all("td" + arySections[i]).className = "hdrNonSelected";
			}
		}
	}
	
	if (document.all("tblSummaryFunctions") != null)
	{
 		if (strSection == "Summary")
		{
			document.all("tblSummaryFunctions").style.display = "";
		}else{
			document.all("tblSummaryFunctions").style.display = "none";
		}
	}

	return(false);
}

function ValidInput() 
{
  if (isEmpty(theDataForm.PromoTitle,"Please enter a title for the promotion.")) {return(false);}
  if (isEmpty(theDataForm.PromoCode,"Please enter a promotion code.")) {return(false);}
  if (isEmpty(theDataForm.StartDate,"Please enter a starting date.")) {return(false);}
  if (!isInteger(theDataForm.Duration,true,"Please enter a integer for the duration.")) {return(false);}
  if (!isNumeric(theDataForm.Discount,false,"Please enter a value for the discount.")) {return(false);}
  if (!isNumeric(theDataForm.MinSubTotal,false,"Please enter a value for the minimum subTotal.")) {return(false);}
  if (!isInteger(theDataForm.NumUses,true,"Please enter a integer for the number of uses.")) {return(false);}
  if (!isInteger(theDataForm.NumUsesByCustomer,true,"Please enter a integer for the number of uses.")) {return(false);}
  
<% If cblnUsePicker Then %>
	setMultiSelect(theDataForm.ProductID);
	setMultiSelect(theDataForm.FreeProductID);
	setMultiSelect(theDataForm.ProductIDExclusion);
	setMultiSelect(theDataForm.Category);
	setMultiSelect(theDataForm.CategoryExclusion);
	setMultiSelect(theDataForm.Manufacturer);
	setMultiSelect(theDataForm.ManufacturerExclusion);
	setMultiSelect(theDataForm.Vendor);
	setMultiSelect(theDataForm.VendorExclusion);
<% End If %>

return(true);
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.Action.value = "";
	theDataForm.OrderBy.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.submit();
	return false;
}

function ViewDetail(theValue)
{
	theKeyField.value = theValue;
	theDataForm.Action.value = "View";
	theDataForm.submit();
	return false;
}

//-->
</script>

<center>

<form action='ssPromotionsAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type=hidden id=PromotionID name=PromotionID value="<%= mclsPromo.PromotionID %>">
<input type=hidden id=Action name=Action value='Update'>
<input type=hidden id=blnShowSummary name=blnShowSummary value="">
<input type=hidden id=blnShowFilter name=blnShowFilter value="">
<input type=hidden id=Show name=Show value="<%= mstrShow %>">
<input type=hidden id=OrderBy name=OrderBy value="<%= mstrOrderBy %>">
<input type=hidden id=SortOrder name=SortOrder value="<%= mstrSortOrder %>">
<input type=hidden id="CloneCodes" name=CloneCodes value="">

<%= .OutputMessage %>
<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th id="tdFilter" class="hdrSelected" nowrap onclick="return DisplayMainSection('Filter');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Set your filter criteria here.">&nbsp;Filter&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdSummary" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('Summary');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View items which meet the specified filter criteria">&nbsp;Summaries&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tditemDetail" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('itemDetail');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View the selected item's detail">&nbsp;Detail&nbsp;</th>
	<th width="90%" align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/PromotionManagerII/help_PromotionManagerII.htm')" id="Button2" name="btnHelp" title="Release Version <%= mstrssAddonVersion %>"></th>
  </tr>
  <tr>
	<td colspan="6" class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td colspan="6" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
	<%
		Call WriteFilter
		Response.Write .OutputSummary
	%>
	<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" id="tblitemDetail">
<!--
	<tr>
		<td>
			<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" ID="Table4">
				<tr class="tblhdr" align=center>
					<td nowrap ID="tdGeneral" class="hdrNonSelected" onclick="return DisplaySection('General');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View General Product Information">General</td>
					<td ID='tdSpacer1' bgcolor="white">&nbsp;</td>
					<td nowrap ID="tdDetail" class="hdrNonSelected" onclick="return DisplaySection('Detail');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Product Details" >Detail</td>
					<td ID='tdSpacer2' bgcolor="white">&nbsp;</td>
					<td nowrap ID="tdAttributes" class="hdrNonSelected" onclick='return DisplaySection("Attributes");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Product Attributes" >Attributes</td>
					<td ID='tdSpacer3' bgcolor="white">&nbsp;</td>
					<td nowrap ID='tdShipping' class="hdrNonSelected" onclick='return DisplaySection("Shipping");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Shipping Settings" >Shipping</td>
				</tr>
			</table>
		</td>
	</tr>
-->
	<tr>
		<td class="Label">&nbsp;<label id=lblPromoCode for=PromoCode title="This is the discount code you send to your customers for registration and what you use to track the promotion's status.">Promotion Code:</label><sup><font color=Red>*</font></sup></td>
		<td>&nbsp;<input id=PromoCode name=PromoCode Value="<%= mclsPromo.PromoCode %>"></td>
	</tr>
	<tr>
		<td class="Label">&nbsp;<label id=lblPromoTitle for=PromoTitle title="This is the name of the promotion the customer will see during registration and on the order summary">Name:</label><sup><font color=Red>*</font></sup></td>
		<td>&nbsp;<input id=PromoTitle name=PromoTitle style="HEIGHT: 22px; WIDTH: 496px" Value="<%= mclsPromo.PromoTitle %>" maxlength=50></td>
	</tr>
	<tr>
		<td class="Label">&nbsp;<label id=lblPromoRules for=PromoRules title="This is the message the customer will see during registration. It should include the general terms and restrictions of the promotion.">Message:</label>&nbsp;</td>
		<td>
		&nbsp;<textarea id=PromoRules name=PromoRules style="HEIGHT: 71px; WIDTH: 496px"><%= mclsPromo.PromoRules %></textarea>
			<a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.PromoRules);" title="Edit this field with the HTML Editor">
			<img SRC="images/prop.bmp" BORDER=0></a>
		</td>
	</tr>
	<tr>
		<td class="Label">&nbsp;<label id=lblStartDate for=StartDate>Start Date:</label><sup><font color=Red>*</font></sup></td>
		<td>&nbsp;<input id=StartDate name=StartDate Value="<%= mclsPromo.StartDate %>">&nbsp;
			<a HREF="javascript:doNothing()" title="Select start date"
			onClick="setDateField(theDataForm.StartDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
			<img SRC="images/calendar.gif" BORDER=0></a>&nbsp;(mm/dd/yyyy)
		</td>
	</tr>
	<tr>
		<td class="Label"><label id=lblEndDate for=EndDate title="Use the field to end the promotion on a specific date.">End Date:</label></td>
		<td>&nbsp;<input id=EndDate name=EndDate Value="<%= mclsPromo.EndDate %>">&nbsp;
			<a HREF="javascript:doNothing()" title="Select end date"
			onClick="setDateField(theDataForm.EndDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
			<img SRC="images/calendar.gif" BORDER=0></a>&nbsp;(mm/dd/yyyy)
		</td>
	</tr>
	<tr>
		<td class="Label"><label id=lblDuration for=Duration title="Use the field to limit the duration of the promotion to a certain number of days.">Duration:</label>&nbsp;</td>
		<td>&nbsp;<input id=Duration name=Duration Value="<%= mclsPromo.Duration %>" onblur="return isInteger(theDataForm.Duration,true,'Please enter a integer for the duration.');">&nbsp;days</td>
	</tr>
	<tr>
		<td class="Label"><label id=lblMaxUses for=MaxUses title="Use this field to limit the promotion to a maximum number of uses as a whole. This is NOT a customer specific limit.">Max Uses:</label>&nbsp;</td>
		<td>&nbsp;<input id=MaxUses name=MaxUses Value="<%= mclsPromo.MaxUses %>" onblur="return isInteger(theDataForm.NumUses,true,'Please enter a integer for the max number of uses.');">&nbsp;&nbsp;&nbsp;
			<label id=lblNumUses for=NumUses>Current Usage</label>&nbsp;&nbsp;&nbsp;
			<input id=NumUses name=NumUses Value="<%= mclsPromo.NumUses %>" onblur="return isInteger(theDataForm.NumUses,true,'Please enter a integer for the number of uses.');"></td>
	</tr>
	<tr>
		<td class="Label"><label id="NumUsesByCustomer" for=MaxUses title="Use this field to limit the promotion to a maximum number by a customer.">Max Uses by Customer:</label>&nbsp;</td>
		<td>&nbsp;<input id="NumUsesByCustomer" name=NumUsesByCustomer Value="<%= mclsPromo.NumUsesByCustomer %>" onblur="return isInteger(theDataForm.NumUsesByCustomer,true,'Please enter a integer for the max number of uses.');">&nbsp;&nbsp;&nbsp;
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;<input id=Combineable name=Combineable type=checkbox <% if (mclsPromo.Combineable="True") then Response.Write "Checked" %>>
		<label id=lblCombineable for=Combineable title="Check this box if you want to enable this promotion to be used in conjunction with other available promotions that are similarly marked.">Can be combined with other promotions</label></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;<input id=ApplyAutomatically name=ApplyAutomatically type=checkbox <% if (mclsPromo.ApplyAutomatically="True") then Response.Write "Checked" %>>
		<label id=lblApplyAutomatically for=ApplyAutomatically title="Check this box if you want the promotion to appear without requiring the customer to register.">Apply Automatically</label></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;<input id=ExcludeSaleItems name=ExcludeSaleItems type=checkbox <% if (mclsPromo.ExcludeSaleItems="True") then Response.Write "Checked" %>>
		<label id=lblExcludeSaleItems for=ExcludeSaleItems title="Check this box if you do not want the promotion to apply to sale items. This means the discount will not be applied to items which are on sale AND items in the cart which are on sale will NOT be counted in the order subTotal used to determine the promotion qualification, if used.">Exclude Sale Items</label></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;<input id=Inactive name=Inactive type=checkbox <% if (mclsPromo.Inactive="True") then Response.Write "Checked" %>>
		<label id=lblInactive for=Inactive title="Check this box if you wish to deactivate this promotion.">Inactive</label></td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<th>&nbsp;</th>
		<th align="left"><div title="Use this section for all types of discounts except free shipping.">Discount Information</div></th>
	</tr>
	<tr>
		<td class="Label"><label id=lblDiscount for=Discount>Discount</label><sup><font color=Red>*</font></sup></td>
		<td>&nbsp;<input id=Discount name=Discount Value="<%= mclsPromo.Discount %>" onblur="return isNumeric(theDataForm.Discount,true,'Please enter a value for the discount.');">&nbsp;&nbsp;&nbsp;
			<div>
			<input type=checkbox name=Percentage id=Percentage <%= isChecked(mclsPromo.Percentage="True") %>>
			<label id=lblPercentage for=Percentage>Check if a percentage off</label><br />
			<input type=checkbox name=ApplyToBasePrice id=ApplyToBasePrice <%= isChecked(mclsPromo.ApplyToBasePrice="True") %>>
			<label id="lblApplyToBasePrice" for=ApplyToBasePrice>Check to apply to base price only (* Applies only to percentage discounts)</label>
			</div>
		</td>
	</tr>
	<tr>
		<td class="Label"><label id="lblMaxAllowableValue" for=MaxAllowableValue>Max. Allowable Value (Total Discount)</label></td>
		<td>&nbsp;<input id="MaxAllowableValue" name=MaxAllowableValue Value="<%= mclsPromo.MaxAllowableValue %>" onblur="return isNumeric(theDataForm.MaxAllowableValue,true,'Please enter a value for the discount.');"></td>
	</tr>
	<tr>
		<td class="Label"><label id="MaxAllowableValuePerItem" for=MaxAllowableValuePerItem>Max. Allowable Value (Per Order Item)</label></td>
		<td>&nbsp;<input id="MaxAllowableValuePerItem" name=MaxAllowableValuePerItem Value="<%= mclsPromo.MaxAllowableValuePerItem %>" onblur="return isNumeric(theDataForm.MaxAllowableValuePerItem,true,'Please enter a value for the discount.');"></td>
	</tr>
	<tr>
		<td class="Label"><label id="lblproductCountLimit" for="productCountLimit">Limit Product Count</label></td>
		<td>&nbsp;<input id="productCountLimit" name="productCountLimit" value="<%= mclsPromo.ProductCountLimit %>" maxlength=50 onblur="return isNumeric(theDataForm.productCountLimit,true,'Please enter a value for the product count limit.');"></td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<th>&nbsp;</th>
		<th align="left"><div title="Use this section for a free shipping promotion.">Free Shipping</div></th>
	</tr>
	<tr>
		<td class="Label"><label id="lblFreeShippingCode" for="FreeShippingCode">Free Shipping</label></td>
		<td>&nbsp;
		<select name="FreeShippingCode" id="FreeShippingCode">
		<%
		Dim mbytshipType
		Dim mbytshipPremiumIsActive

		Call GetSupportedShippingType(mbytshipType, mbytshipPremiumIsActive)

		If Len(mclsPromo.FreeShippingCode) = 0 Then
			Response.Write "<option value=''>Free Shipping Not Enabled</option>"
		Else
			Response.Write "<option value='' selected>Free Shipping Not Enabled</option>"
		End If
	    
		If mbytshipType <> 2 Then
			If mclsPromo.FreeShippingCode = 0 Then
				Response.Write "<option value=0 selected>Regular Shipping</option>"
			Else
				Response.Write "<option value=0>Regular Shipping</option>"
			End If
		    
			If mbytshipPremiumIsActive <> 0 Then
				If mclsPromo.FreeShippingCode = 1 Then
					Response.Write "<option value=1 selected>Premium Shipping</option>"
				Else
					Response.Write "<option value=1>Premium Shipping</option>"
				End If
			End If
		Else
			If cblnAddon_PostageRate Then
				Call MakeCombo("Select ssShippingMethodCode, ssShippingMethodName From ssShippingMethods Where ssShippingMethodEnabled<>0 Order By ssShippingMethodName","ssShippingMethodName","ssShippingMethodCode",mclsPromo.FreeShippingCode)
			Else
				Call MakeCombo("Select shipID, shipMethod From sfShipping Where shipIsActive=1 Order By shipMethod","shipMethod","shipID",mclsPromo.FreeShippingCode)
			End If
		End If
	    
		%>
		</select>
	</tr>
	<tr>
		<td class="Label"><label id="lblFreeShippingLimit" for="FreeShippingLimit" title="Use this field to limit the amount of free shipping. Leave it blank to limit it to the amount allowed in the discount information section. You should set the discount to 100% in the Discount Information section if you use this.">Free Shipping Limit</label></td>
		<td>&nbsp;<input id="FreeShippingLimit" name="FreeShippingLimit" value="<%= mclsPromo.FreeShippingLimit %>" size=5 style="text-align:center;" onblur="return isNumeric(theDataForm.FreeShippingLimit,true,'Please enter a value for the amount to get off.');">
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<th>&nbsp;</th>
		<th align="left"><div title="Use this section for a buy a product, get another product at a discount.">Buy a product, get a product at a discount</div></th>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
		&nbsp;<label id="lblbuyX" for="buyX">Buy</label>&nbsp;<input id="buyX" name="buyX" value="<%= mclsPromo.buyX %>" size=5 style="text-align:center;" onblur="return isNumeric(theDataForm.buyX,true,'Please enter a value.');">
		&nbsp;<label id="lblgetY" for="getY">Get</label>&nbsp;<input id="getY" name="getY" value="<%= mclsPromo.getY %>" size=5 style="text-align:center;" onblur="return isNumeric(theDataForm.getY,true,'Please enter a value.');">
		&nbsp;at the discount set in <em>Discount Information</em>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<input type=checkbox name="likeItem" id="likeItem" <% if (mclsPromo.likeItem="True") then Response.Write "checked" %>>
			<label id="lbllikeItem" for="likeItem">Apply only to like items</label>
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
	</tr>
	<tr>
		<th>&nbsp;</th>
		<th align="left"><div title="Use this section for to select which products are free with purchase. The purchase must meet the criteria specified.">Free Product with purchase</div></th>
	</tr>
		<tr>
			<td align="right" valign="top">Free Products:&nbsp;</td>
			<td align=left>
			<select id="FreeProductID" name="FreeProductID" size=5 ondblclick="openMovementWindow('FreeProductID','product');" multiple></select>
			<a href="" onclick="openMovementWindow('FreeProductID','product'); return false;"><img src="images/properites.gif" border="0"></a>
			</td>
		</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<input type="checkbox" name="offerFreeGiftAutomatically" id="offerFreeGiftAutomatically" <% if (mclsPromo.offerFreeGiftAutomatically="True") then Response.Write "Checked" %>>
			<label id="lblofferFreeGiftAutomatically" for="offerFreeGiftAutomatically">Offer free gift automatically at order.asp</label>
		</td>
	</tr>
	<script LANGUAGE=javascript>
	<!--
	var mdicProductID = new ActiveXObject("Scripting.Dictionary");;
	var mdicCategory = new ActiveXObject("Scripting.Dictionary");;
	var mdicManufacturer = new ActiveXObject("Scripting.Dictionary");;
	var mdicVendor = new ActiveXObject("Scripting.Dictionary");;

	var mdicFreeProductID = new ActiveXObject("Scripting.Dictionary");;
	var mdicProductIDExclusion = new ActiveXObject("Scripting.Dictionary");;
	var mdicCategoryExclusion = new ActiveXObject("Scripting.Dictionary");;
	var mdicManufacturerExclusion = new ActiveXObject("Scripting.Dictionary");;
	var mdicVendorExclusion = new ActiveXObject("Scripting.Dictionary");;

	//-->
	</script>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
		<table border="0" cellpadding="0" cellspacing="0" ID="Table2">
			<tr>
			<th align="left" colspan="3"><div title="This section sets the criteria an order must meet to qualify for the promotion. The customer must have items in the cart from the following product listing, category, manufacturer, or vendor as selected below AND meet the minimum order amount (subTotal) if specified. Items marked as excluded by either the product listing, category, manufacturer, or vendor are excluded. In the event of a general discount the criteria below determines which products are discounted.">Order Requirements</div></th>
			</tr>
			<tr>
				<td class="Label"><label id=lblMinSubTotal for=MinSubTotal>MinSubTotal</label><sup><font color=Red>*</font></sup></td>
				<td>&nbsp;<input id=MinSubTotal name=MinSubTotal Value="<%= mclsPromo.MinSubTotal %>" onblur="return isNumeric(theDataForm.MinSubTotal,true,'Please enter a value for the minimum subTotal.');"></td>
			<th>&nbsp;</th>
			</tr>
			<tr>
			<th>&nbsp;</th>
			<th align="left">Include items:</th>
			<th align="left">Exclude items:</th>
			</tr>
			<tr>
			<td align="right" valign="top" class="Label">Products:&nbsp;</td>
			<td align=left>
				<% If cblnUsePicker Then %>
				<select id="ProductID" name="ProductID" size=5 ondblclick="openMovementWindow('ProductID','product');" multiple></select>
				<a href="" onclick="openMovementWindow('ProductID','product'); return false;"><img src="images/properites.gif" border="0"></a>
				<% Else %>
				<textarea id="Textarea1" name="ProductID" cols="40" rows="4"><%= .ProductID %></textarea>
				<% End If %>
			</td>
			<td align=left>
				<% If cblnUsePicker Then %>
				<select id="ProductIDExclusion" name="ProductIDExclusion" size=5 ondblclick="openMovementWindow('ProductIDExclusion','product');" multiple></select>
				<a href="" onclick="openMovementWindow('ProductIDExclusion','product'); return false;"><img src="images/properites.gif" border="0"></a>
				<% Else %>
				<textarea cols="40" rows="4" ID="ProductIDExclusion" NAME="ProductIDExclusion"><%= .ProductIDExclusion %></textarea>
				<% End If %>
			</td>
			</tr>
			<tr>
			<td align="right" valign="top">Categories:&nbsp;</td>
			<td align=left>
				<% If cblnUsePicker Then %>
				<select id="Category" name="Category" size="5" ondblclick="openMovementWindow('Category','category');" multiple></select>
				<a href="" onclick="openMovementWindow('Category','category'); return false;"><img src="images/properites.gif" border="0"></a>
				<% Else %>
				<textarea cols="40" rows="4" ID="Category" NAME="Category"><%= .Category %></textarea>
				<% End If %>
			</td>
			<td align=left>
				<% If cblnUsePicker Then %>
				<select id="CategoryExclusion" name="CategoryExclusion" size="5" ondblclick="openMovementWindow('CategoryExclusion','category');" multiple></select>
				<a href="" onclick="openMovementWindow('CategoryExclusion','category'); return false;"><img src="images/properites.gif" border="0"></a>
				<% Else %>
				<textarea cols="40" rows="4" ID="CategoryExclusion" NAME="CategoryExclusion"><%= .CategoryExclusion %></textarea>
				<% End If %>
			</td>
			</tr>
			<tr>
			<td align="right" valign="top">Manufacturers:&nbsp;</select>
			</td>
			<td align=left>
				<% If cblnUsePicker Then %>
				<select id="Manufacturer" name="Manufacturer" size="5" ondblclick="openMovementWindow('Manufacturer','manufacturer');" multiple></select>
				<a href="" onclick="openMovementWindow('Manufacturer','manufacturer'); return false;"><img src="images/properites.gif" border="0"></a>
				<% Else %>
				<textarea cols="40" rows="4" ID="Manufacturer" NAME="Manufacturer"><%= .Manufacturer %></textarea>
				<% End If %>
			</td>
			<td align=left>
				<% If cblnUsePicker Then %>
				<select id="ManufacturerExclusion" name="ManufacturerExclusion" size="5" ondblclick="openMovementWindow('ManufacturerExclusion','manufacturer');" multiple></select>
				<a href="" onclick="openMovementWindow('ManufacturerExclusion','manufacturer'); return false;"><img src="images/properites.gif" border="0"></a>
				<% Else %>
				<textarea cols="40" rows="4" ID="ManufacturerExclusion" NAME="ManufacturerExclusion"><%= .ManufacturerExclusion %></textarea>
				<% End If %>
			</td>
			</tr>
			<tr>
			<td align="right" valign="top">Vendors:&nbsp;</td>
			<td align=left>
				<% If cblnUsePicker Then %>
				<select id="Vendor" name="Vendor" size="5" ondblclick="openMovementWindow('Vendor','vendor');" multiple></select>
				<a href="" onclick="openMovementWindow('Vendor','vendor'); return false;"><img src="images/properites.gif" border="0"></a>
				<% Else %>
				<textarea cols="40" rows="4" ID="Vendor" NAME="Vendor"><%= .Vendor %></textarea>
				<% End If %>
			</td>
			<td align=left>
				<% If cblnUsePicker Then %>
				<select id="VendorExclusion" name="VendorExclusion" size="5" ondblclick="openMovementWindow('VendorExclusion','vendor');" multiple></select>
				<a href="" onclick="openMovementWindow('VendorExclusion','vendor'); return false;"><img src="images/properites.gif" border="0"></a>
				<% Else %>
				<textarea cols="40" rows="4" ID="VendorExclusion" NAME="VendorExclusion"><%= .VendorExclusion %></textarea>
				<% End If %>
			</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td class="Label"><label id=lblModifiedOn>Modified On</LABEL></td>
		<td>&nbsp;<div id=ModifiedOn><%= mclsPromo.ModifiedOn %>&nbsp;</div></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><font color=Red>*Required</font></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<input class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/PromotionManagerII/help_PromotionManagerII.htm'); return false;" id=btnHelp name=btnHelp title="help">&nbsp;
			<input class='butn' id=btnNew name=btnNew type=button value=New onclick="return btnNew_onclick()">&nbsp;
			<input class='butn' name=btnReset type=reset value=Reset onclick="return btnReset_onclick()" ID="Reset1">&nbsp;&nbsp; 
			<input class='butn' name=btnDelete type=button value=Delete onclick="return btnDelete_onclick()" ID="Button1"> 
			<input class='butn' name=btnClone type=button value=Clone onclick="return btnClone_onclick()" ID="btnClone"> 
			<input class='butn' name=btnUpdate type=submit value="Save Changes" ID="Submit1"> 
		</td>
	</tr>
	</TABLE>
	</td>
  </tr>
</table>



</form>
<p><a href="ssPromotionsReport.asp">Promotion Report</a></p>
</center>

</body>
</HTML>
<% 

	End With
	
	set mclsPromo = Nothing
	set cnn = Nothing

Response.Flush

'***********************************************************************************************

Sub GetSupportedShippingType(shipType, shipPremiumIsActive)

Dim pobjRS

	Set pobjRS = Server.CreateObject("ADODB.RecordSet")
	With pobjRS
		.CursorLocation = 2 'adUseClient
		.Open "SELECT adminShipType, adminPrmShipIsActive FROM sfAdmin", cnn, 3, 1, &H0001		'adOpenStatic, adLockReadOnly, adCmdText
	
		If Not .EOF Then
			shipType = .Fields("adminShipType").Value
			shipPremiumIsActive = .Fields("adminPrmShipIsActive").Value
		Else
			shipType = 1
			shipPremiumIsActive = 0
		End If
		.Close
	End With
	Set pobjRS = Nothing

End Sub	'GetSupportedShippingType

'***********************************************************************************************

Sub WriteFilter()
%>
<script LANGUAGE=javascript>
<!--

function btnFilter_onclick(theButton)
{
  theDataForm.Action.value = "Filter";
  theDataForm.submit();
  return(true);
}

//-->
</script>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
<colgroup align="left">
<colgroup align="left">
  <tr>
    <td valign="top">
        <input type="radio" value="1" <% if mbytText_Filter="1" then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter1"><label for="optText_Filter1">Code</label><br />
        <input type="radio" value="2" <% if mbytText_Filter="2" then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter2"><label for="optText_Filter2">Name</label><br />
        <input type="radio" value="3" <% if mbytText_Filter="3" then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter3"><label for="optText_Filter3">Message</label><br />
        <input type="radio" value="0" <% if (mbytText_Filter="0" or mbytText_Filter="") then Response.Write "Checked" %> name="optText_Filter" ID="optText_Filter0"><label for="optText_Filter0">Do Not Include</label>
        <p>containing the text<br />
        <input type="text" id="Text_Filter" name="Text_Filter" size="20" value="<%= Server.HTMLEncode(mstrText_Filter) %>">
	</td>
    <td valign="top">
        Show Only Promotions with<br />
		<label for="filterStartDate">a start date</label> <input type="radio" name="radStartDate" id="radStartDate0" value="0" <% if mbytradStartDate="0" then Response.Write "Checked" %>><label for="radStartDate0">before</label>&nbsp;<input type="radio" name="radStartDate" id="radStartDate1" value="1" <% if mbytradStartDate="1" then Response.Write "Checked" %>><label for="radStartDate1">after</label>&nbsp;<input name="filterStartDate" id="filterStartDate" value="<%= mstrfilterStartDate %>">
		<a HREF="javascript:doNothing()" title="Select start date"
		onClick="setDateField(document.frmData.filterStartDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<img SRC="images/calendar.gif" BORDER=0></a><br />

		<label for="filterEndDate">an end date</label> <input type="radio" name="radEndDate" id="radEndDate0" value="0" <% if mbytradEndDate="0" then Response.Write "Checked" %>><label for="radEndDate0">before</label>&nbsp;<input type="radio" name="radEndDate" id="radEndDate1" value="1" <% if mbytradEndDate="1" then Response.Write "Checked" %>><label for="radEndDate1">after</label>&nbsp;<input name="filterEndDate" id="filterEndDate" value="<%= mstrfilterEndDate %>">
		<a HREF="javascript:doNothing()" title="Select end date"
		onClick="setDateField(document.frmData.filterEndDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<img SRC="images/calendar.gif" BORDER=0></a>

	</td>
    <td valign="top">
		<table cellpadding="0" cellspacing="0" border="0" ID="Table3">
		  <caption align=left><font size="-1">Show Promotions that are:</font></caption>
		  <tr>
		    <td><input type="radio" value="1" <% if mbytActive_Filter="1" then Response.Write "Checked" %> name="optActive_Filter" ID="optActive_Filter1"><label for="optActive_Filter1"><font size="-1">Active</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <% if mbytActive_Filter="2" then Response.Write "Checked" %> name="optActive_Filter" ID="optActive_Filter2"><label for="optActive_Filter2"><font size="-1">Inactive</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <% if (mbytActive_Filter="0" or mbytActive_Filter="") then Response.Write "Checked" %> name="optActive_Filter" ID="optActive_Filter0"><label for="optActive_Filter0"><font size="-1">Either</font></label></td>
		  </tr>
		  <tr>
		    <td><input type="radio" value="2" <% if mbytApplyAuto_Filter="2" then Response.Write "Checked" %> name="optApplyAuto_Filter" ID="optApplyAuto_Filter2"><label for="optApplyAuto_Filter2"><font size="-1">Applied Automatically</font></label></td>
		    <td>&nbsp;<input type="radio" value="1" <% if mbytApplyAuto_Filter="1" then Response.Write "Checked" %> name="optApplyAuto_Filter" ID="optApplyAuto_Filter1"><label for="optApplyAuto_Filter1"><font size="-1">Require Registration</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <% if (mbytApplyAuto_Filter="0" or mbytApplyAuto_Filter="") then Response.Write "Checked" %> name="optApplyAuto_Filter" ID="optApplyAuto_Filter0"><label for="optApplyAuto_Filter0"><font size="-1">Either</font></label></td>
		  </tr>
		  <tr>
		    <td><input type="radio" value="1" <% if mbytExclusive_Filter="1" then Response.Write "Checked" %> name="optExclusive_Filter" ID="optExclusive_Filter1"><label for="optExclusive_Filter1"><font size="-1">Can be combined</font></label></td>
		    <td>&nbsp;<input type="radio" value="2" <% if mbytExclusive_Filter="2" then Response.Write "Checked" %> name="optExclusive_Filter" ID="optExclusive_Filter2"><label for="optExclusive_Filter2"><font size="-1">Exclusive Use</font></label></td>
		    <td>&nbsp;<input type="radio" value="0" <% if (mbytExclusive_Filter="0" or mbytExclusive_Filter="") then Response.Write "Checked" %> name="optExclusive_Filter" ID="optExclusive_Filter0"><label for="optExclusive_Filter0"><font size="-1">Either</font></label></td>
		  </tr>
		</table>
	</td>
	<td valign="middle">
		<p><input class="butn" id=btnFilter name=btnFilter type=button value="Apply Filter" onclick="btnFilter_onclick(this);"></p>
	</td>
  </tr>
</table>
<% End Sub	'WriteFilter %>