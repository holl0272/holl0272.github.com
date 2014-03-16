<% Option Explicit 
'********************************************************************************
'*   Product Import Tool For StoreFront 6.0
'*   Release Version:	1.02.004
'*   Release Date:		August 22, 2003
'*   Revision Date:		November 11, 2004
'*
'*   Release Notes:                                                             *
'*   -- See Product Documentation                                               *
'*
'*	 Revision 2.00.001 Beta (December 30, 2004)
'*   - Delete existing products moved inside initialization check
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.
'********************************************************************************

Response.Buffer = True

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Server.ScriptTimeout = 600			'this is the maximum time the page can run. Note: IIS may be configured such there is an additional limit in place. Consult your host for specific values
Const clngBufferTimeout = 30		'this is the maximum estimated time it will take for a single product to import. Note: if your products have a lot of attributes this should be higher or if your running SQL Server
Const clngPageRefreshDelay = 5		'this is the time in seconds the page will wait before reposting on large imports; longer delays slow the import but have been know to be necessary on some servers to let them "catch up"

Const cblnAssignProductsToAllCategoryLevels = True	'False will set to only last category/subcategory in tree

Const cblnNumericCategoriesAreUID = True	'False
Const cblnNumericManufacturersAreUID = True	'False
Const cblnNumericVendorsAreUID = True	'False

maryCustomBooleanValues = Array("1", "true", "yes", "y", "on", "-1", "active")

'The import tool can map many import values to a True/False condition. Mappings are set to False by default. Common settings which map to True
'are found above. Simply add your custom mappings to the end of the list
'Notes:
' - all text entries must be in lower case and the import test is case insensitive
' - If an entry is in your data which does not map to one of the values below it will map to False
' - the above discussion refers to True/False values and/or fields which map to 0/1 only

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Global variables
'**********************************************************

Const cenCreateCat_Default = 0
Const cenCreateCat_Create = 1
Const cenCreateCat_CreateAndDelete = 2

Const cenCreateMfg_Default = 0
Const cenCreateMfg_Create = 1
Const cenCreateMfg_CreateAndDelete = 2

Const cenCreateVend_Default = 0
Const cenCreateVend_Create = 1
Const cenCreateVend_CreateAndDelete = 2

Const enTargetFieldName = 0
Const enSourceFieldName = 1
Const enDisplayFieldName = 2
Const enDefaultValue = 3
Const enFieldDataType = 4
Const enDisplayType = 5

Const cstrTarget_Code = "prodID"
Const cstrTarget_Manufacturer = "prodManufacturerId"
Const cstrTarget_Vendor = "prodVendorId"

Dim maryAttributeSupportingStyles(4)

maryAttributeSupportingStyles(0) = Array("No Attribute", "")
maryAttributeSupportingStyles(1) = Array("ShopSite Style", "")
maryAttributeSupportingStyles(2) = Array("Attribute Template", "")
maryAttributeSupportingStyles(3) = Array("Auto Select", "")
maryAttributeSupportingStyles(4) = Array("Price Options", "")

Const enImportAll = 0
Const enImportAllDeleteCategories = 1
Const enImportAll_DeleteExisting = 2
Const enImportNewOnly = 3
Const enImportNewOnlyAllowDuplicateCodes = 4
Const enImportInvPriceOnly = 5
Const enImportUpdateSelectedFieldsOnly = 6
Const enImportInformationOnly = 7
Const enImportDelete = 8

Dim maryImportTypes(8)
maryImportTypes(enImportAll) = Array("Import all products - Overwrite existing", "Check this option to import all products from the data source. Existing products will be overwritten with the new information", "")
maryImportTypes(enImportAll_DeleteExisting) = Array("Import all products - Delete existing", "Check this option to import all products from the data source. All existing products will be deleted first.", "")
maryImportTypes(enImportNewOnly) = Array("Import new products only", "Check this option to only add new products. Product codes already in database will be ignored.", "")
maryImportTypes(enImportInvPriceOnly) = Array("Update cost and inventory only", "Check this option to only update pricing and inventory levels for products which are already in the database. New products will still be added.", "")
maryImportTypes(enImportInformationOnly) = Array("Import category/manufacturer/vendor information only", "Check this option to import category/manufacturer/vendor information only.", "")
maryImportTypes(enImportAllDeleteCategories) = Array("Import all products - Overwrite existing, Remove Existing category assignments", "Check this option to import the entire source. Existing products will be overwritten with the new information. Any previously assigned categories will be removed and the new ones used.", "")
maryImportTypes(enImportUpdateSelectedFieldsOnly) = Array("Update selected fields only", "Check this option if you only want to update existing products using the specified fields through the drop down. New products will still be added normally.", "")
maryImportTypes(enImportNewOnlyAllowDuplicateCodes) = Array("Import new products only - Duplicate codes allowed", "Check this option to add new products. Product codes already in database will not be verified and products with duplicate product codes may result.", "")
maryImportTypes(enImportDelete) = Array("Delete", "Delete products.", "")

Dim mstrAvailableTables

'**********************************************************
'*	Functions
'**********************************************************

'Function checkForCorruptData(byVal strData)
'Function checkForSubCategories(byVal strValue, byRef dicCategories, byVal pblnLocalDebug, byRef pstrLocalDebugOut)
'Function checkReplacements(byRef strSource, byRef strProdID)
'Function checkRequiredValue(ByVal strFieldName, ByVal vntValue, ByRef blnRequiredValueSet)
'Function ClearCategories()
'Function ClearManufacturers()
'Function ClearVendors()
'Sub cleanCategoryData(byRef strCategoryName)
'Function Counter(ByRef lngCounter)
'Function createCategory(byRef lngUID, byVal strCategoryName, byVal lngParentLevel, byVal ParentID, byVal bytIsActive)
'Function customBoolean(byVal strValue)
'Function deleteXMLProfile(byVal strProfileID)
'Function dbFieldName(byVal strSpreadSheetFieldName)
'Sub DoImport(objcnnSource, strTableSource)
'Sub DoImport_Custom()
'Function getCategoryUIDbyName(byVal vntCategoryName, byRef objrsCategories, byRef dicCategories)
'Function getdbFieldName(byVal strSpreadSheetFieldName, byRef arySearch)
'Sub GetNodeValues(objNode)
'Function getProductID(byRef objRSSource)
'Sub getProductDetail(byRef objXMLElement, byRef aryProductDetail)
'Function getProductName(byRef objRSSource)
'Function GetSelectOptions(aryData, vntValue)
'Function getValueFrom(byRef arySource, byRef objrsProducts, byVal lngIndex, byVal blnSpecial, byVal vntDefault)
'Function hasCorruptData(byVal strData)
'Sub IdentifyKeyProductFields()
'Function importTypeFromText(byVal strImportType)
'Function InitializeCustomFields
'Sub InitializeDefaults(byRef strProfileID)
'Sub LoadFormValues(byVal blnComplete)
'Sub loadNamedProfiles()
'Function loadSettingsFromProfile(byRef objXMLProfile)
'Function LoadXMLProfiles(byVal strProfileID, byRef XMLDoc)
'Function newXMLProfile(byVal strProfileID, byVal strProfileName)
'Function OpenTableSQL(ByVal strTableSource)
'Function profilePath()
'Sub RecordTime(byVal strMessage, byVal dtStartTime, byRef dtCurrentTime, byRef dtLastTime)
'Sub saveAttributes(byRef objrsProducts, byRef aryProduct)
'Sub saveCategories(byRef objrsProducts, byVal lngProductID, byRef objrsCategories, byRef dicCategories)
'Sub saveCustom(byRef objrsProducts, byVal lngProductUID)
'Sub SaveFailedImport(strProductID, strProductName)
'Sub saveGiftWrap(byRef objrsProducts, byVal lngProductUID)
'Function saveXMLProfile(byVal strProfileID, byVal strProfileName)
'Sub SetColumns(byRef objRS)
'Sub setProductDetail(byRef objXMLElement, byRef aryProductDetail)
'Function setInventoryLevels(byRef objrsProducts, byVal aryProduct, byRef blnInventoryUpdateOnly)
'Function setSubCategories(byVal lngProductID, byVal strValue, byRef objrsCategories, byRef dicCategories)
'Function setVolumePricing(byRef objrsProducts, byVal lngProductUID, byVal strBasePrice)
'Function spreadsheetFieldName(byVal strdbFieldName)
'Function spreadsheetIndexByFieldName(byVal strdbFieldName)
'Function updateDetailLink(byRef objrsProducts, byVal lngProductUID, byVal strProductID)
'Sub verifyCategories(byRef objrsNewProducts, byRef strFieldName, byRef objrsCategories, byRef dicCategories)
'Function verifyCategory(byVal strCategoryName, byRef objrsCategories, byRef dicCategories, byRef blnAdded)
'Sub VerifyConnections(ByRef objCnn, ByRef strTableSource)
'Function verifyMfgVends(byRef objrsNewProducts, byRef strFieldName, byRef dicMfgVends, byVal lngType)
'Sub WriteFailedImports()
'Sub WriteOutput(byVal strOut)

'****************************************************************************************************************************************************************

Function checkForCorruptData(byVal strData)

Dim pstrTemp

	If isNull(strData) Then
		pstrTemp = strData
	Else
		pstrTemp = CStr(strData)
	
	End If

	pstrTemp = Trim(strSource & "")
	pstrTemp = Replace(pstrTemp, "<prodID>", strProdID)
	pstrTemp = Replace(pstrTemp, "<code>", strProdID)
	pstrTemp = Replace(pstrTemp, "<Date()>", Date())
	
	checkForCorruptData = pstrTemp

End Function	'checkForCorruptData

'****************************************************************************************************************************************************************

Function checkForSubCategories(byVal strValue, byRef dicCategories, byVal pblnLocalDebug, byRef pstrLocalDebugOut)

Dim i
Dim parySubcategory
Dim parySubcategoryIDs
Dim plngParentID
Dim plngUID
Dim plngLevel
Dim pstrCategoryTrailToCheck

	pblnLocalDebug = False	'True	False

	parySubcategory = Split(strValue, cstrSubcategoryDelimiter)
	parySubcategoryIDs = Split(strValue, cstrSubcategoryDelimiter)
	plngParentID = 0
	
	If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "<table class=tbl border=1 cellpadding=2 cellspacing=0>" _
																 & "<tr class=tblhdr><th colspan=3 align=left>checkForSubCategories (<i>" & strValue & "</i>)</th></tr>" _
																 & "<tr class=tblhdr><th>Category</th><th>ID</th><th>ParentID</th></tr>"

	For i = 0 To UBound(parySubcategory)
		If i = 0 Then
			pstrCategoryTrailToCheck = Trim(parySubcategory(i))
		Else
			pstrCategoryTrailToCheck = pstrCategoryTrailToCheck & cstrSubcategoryDelimiter & Trim(parySubcategory(i))
		End If
		If i > 0 Then plngParentID = parySubcategoryIDs(i-1)

		If createCategory(plngUID, Trim(parySubcategory(i)), i, plngParentID, 1) Then
			If Not dicCategories.Exists(pstrCategoryTrailToCheck) Then dicCategories.add pstrCategoryTrailToCheck, plngUID
			
			If i = 0 Then
				WriteOutput "Category <em>" & strValue & "</em> being checked . . .<br />" & vbcrlf
				WriteOutput "<ul><li><strong><em>" & parySubcategory(i) & "</em> added.</strong></li>" & vbcrlf
			Else
				WriteOutput "<ul><li><strong><em>" & parySubcategory(i) & "</em> added.</strong></li>" & vbcrlf
			End If
		Else
			If Not dicCategories.Exists(pstrCategoryTrailToCheck) Then dicCategories.add pstrCategoryTrailToCheck, plngUID
			If i = 0 Then
				WriteOutput "Category <em>" & strValue & "</em> being checked . . .<br />" & vbcrlf
				WriteOutput "<ul><li><em>" & parySubcategory(i) & "</em> exists.</li>" & vbcrlf
			Else
				WriteOutput "<ul><li><em>" & parySubcategory(i) & "</em> exists.</li>" & vbcrlf
			End If
		End If
		If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "<tr><td>" & pstrCategoryTrailToCheck & "</td><td>" & plngUID & "</td><td>" & plngParentID & "</td></tr>"
		parySubcategoryIDs(i) = plngUID
	Next 'i
	
	For i = 0 To UBound(parySubcategory)
		WriteOutput "</ul>" & vbcrlf
	Next 'i
	
	If pblnLocalDebug Then
		pstrLocalDebugOut = pstrLocalDebugOut & "<tr class=tblhdr><th colspan=3>checkForSubCategories( " & strValue & ") = " & plngUID & "</tr>"
		pstrLocalDebugOut = pstrLocalDebugOut & "</table>"
		Response.Write "<fieldset><legend>checkForSubCategories: {" & strValue & "}</legend>" & pstrLocalDebugOut & "</fieldset>"
	End If
	
	checkForSubCategories = plngUID

End Function	'checkForSubCategories

'****************************************************************************************************************************************************************

Function checkReplacements(byRef strSource, byRef strProdID)

Dim pstrTemp

	pstrTemp = Trim(strSource & "")
	pstrTemp = Replace(pstrTemp, "<prodID>", strProdID)
	pstrTemp = Replace(pstrTemp, "<code>", strProdID)
	pstrTemp = Replace(pstrTemp, "<Date()>", Date())
	
	checkReplacements = pstrTemp

End Function	'checkReplacements

'****************************************************************************************************************************************************************

Function checkRequiredValue(ByVal strFieldName, ByVal vntValue, ByRef blnRequiredValueSet)

Dim pvntOut

	pvntOut = Trim(vntValue & "")
	pvntOut = Replace(pvntOut, vbcrlf, "")

	Select Case strFieldName
		Case "ManufacturerId", _
			 "VendorId":
				If CBool(Len(Trim(pvntOut & "")) = 0) Then
					pvntOut = 1
					blnRequiredValueSet = True
				End If
		Case "IsActive", _
			 "IsOnSale", _
			 "IsShipable", _
			 "HasCountryTax", _
			 "HasStateTax", _
			 "HasLocalTax", _
			 "DropShip", _
			 "DownloadOneTime", _
			 "DownloadExpire", _
			 "DealTimeIsActive", _
			 "MMIsActive":
				If ConvertToBoolean(pvntOut, False) Then
					pvntOut = 1
				Else
					pvntOut = 0
				End If
		Case "Cost", _
			 "Price", _
			 "SalePrice", _
			 "ShipPrice", _
			 "Weight", _
			 "Length", _
			 "Width", _
			 "Height":
				If CBool(Len(pvntOut) = 0) Or Not isNumeric(pvntOut) Then
					WriteOutput "&nbsp;&nbsp;<font color=red>Invalid data for field <em>" & strFieldName & "</em>. {" & pvntOut & "} replaced with {" & "0" & "}</font><br />" & vbcrlf
					pvntOut = 0
					blnRequiredValueSet = True
				End If
		Case "ProductType", _
			 "ProductionTime":
				If CBool(Len(pvntOut) = 0) Or Not isNumeric(pvntOut) Then
					WriteOutput "&nbsp;&nbsp;<font color=red>Invalid data for field <em>" & strFieldName & "</em>. {" & pvntOut & "} replaced with {" & "0" & "}</font><br />" & vbcrlf
					pvntOut = 0
				End If
		Case "MinQty":
				If CBool(Len(pvntOut) = 0) Or Not isNumeric(pvntOut) Then
					WriteOutput "&nbsp;&nbsp;<font color=red>Invalid data for field <em>" & strFieldName & "</em>. {" & pvntOut & "} replaced with {" & "1" & "}</font><br />" & vbcrlf
					pvntOut = 1
				End If
	End Select	'strFieldName

	checkRequiredValue = pvntOut
	
End Function	'checkRequiredValue

'****************************************************************************************************************************************************************

Sub cleanCategoryData(byRef strCategoryName)
'If the input category name is greater than 255 characters the category may be corrupted
'This checks for this situation and if present, removes the final category to be created

Dim i
Dim paryCategory
Dim pstrTemp

	If hasCorruptData(strCategoryName) Then
		'Note this is split across three lines due to writing the corrupt category name
		WriteOutput "&nbsp;&nbsp;<font color=red><strong>Category data error.</strong> {"
		WriteOutput strCategoryName
		WriteOutput "}.</font><br />This may occur if your category field to import is greater than 255 characters<br />" & vbcrlf
		paryCategory = Split(strCategoryName, cstrMultipleCategoryDelimiter)
		If UBound(paryCategory) > 0 Then
			ReDim Preserve paryCategory(UBound(paryCategory) - 1)
			strCategoryName = paryCategory(0)
			For i = 1 To UBound(paryCategory)
				strCategoryName = strCategoryName & cstrMultipleCategoryDelimiter & paryCategory(i)
			Next
			WriteOutput "&nbsp;&nbsp;<font color=red>Attempting recovery of partial category assignments - <b>Please verify results</b></font><br />" & vbcrlf
			For i = 0 To UBound(paryCategory)
				WriteOutput "&nbsp;&nbsp;" & i + 1 & ": " & paryCategory(i) & "<br />" & vbcrlf
			Next
		Else
			strCategoryName = ""
		End If		
	End If	'hasCorruptData(strCategoryName)

End Sub	'cleanCategoryData

'****************************************************************************************************************************************************************

Function ClearCategories()
'This function deletes all existing categories 
'Existing products are reset to default category value

Dim pblnSuccess
Dim pstrLocalError
Dim pstrSQL

	pblnSuccess = True

	pstrSQL = "Delete From sfSubCatDetail"
	'debugprint "pstrSQL", pstrSQL
	pblnSuccess = pblnSuccess And Execute_NoReturn(pstrSQL, pstrLocalError)
	
	pstrSQL = "Delete From sfSub_Categories WHERE subcatCategoryId<>1"
	'debugprint "pstrSQL", pstrSQL
	pblnSuccess = pblnSuccess And Execute_NoReturn(pstrSQL, pstrLocalError)
	
	pstrSQL = "Delete From sfCategories WHERE catID<>1"
	'debugprint "pstrSQL", pstrSQL
	pblnSuccess = pblnSuccess And Execute_NoReturn(pstrSQL, pstrLocalError)
	
	If pblnSuccess Then
		WriteOutput "<b>Existing Categories removed</b><br />" & vbcrlf
	Else
		WriteOutput "<b><font color=red>Error removing existing Categories</font></b><br />" & pstrLocalError & vbcrlf
	End If
	
	ClearCategories = pblnSuccess

End Function	'ClearCategories

'****************************************************************************************************************************************************************

Function ClearManufacturers()

Dim pblnSuccess
Dim pstrLocalError
Dim pstrSQL

	pblnSuccess = True

	pstrSQL = "DELETE FROM sfManufacturers WHERE mfgID <> 1"
	pblnSuccess = pblnSuccess And Execute_NoReturn(pstrSQL, pstrLocalError)
	
	pstrSQL = "Update sfProducts Set prodManufacturerId=1"
	pblnSuccess = pblnSuccess And Execute_NoReturn(pstrSQL, pstrLocalError)
	
	If pblnSuccess Then
		WriteOutput "<b>Existing manufacturers removed. All products have been updated to <em>No Manufacturer</em></b><br />" & vbcrlf
	Else
		WriteOutput "<b><font color=red>Error removing existing manufacturers</font></b><br />" & pstrLocalError & vbcrlf
	End If
	
	ClearManufacturers = pblnSuccess

End Function	'ClearManufacturers

'****************************************************************************************************************************************************************

Function ClearVendors()

Dim pblnSuccess
Dim pstrLocalError
Dim pstrSQL

	pblnSuccess = True

	pstrSQL = "DELETE FROM sfVendors WHERE vendID <> 1"
	pblnSuccess = pblnSuccess And Execute_NoReturn(pstrSQL, pstrLocalError)
	
	pstrSQL = "Update sfProducts Set prodVendorId=1"
	pblnSuccess = pblnSuccess And Execute_NoReturn(pstrSQL, pstrLocalError)
	
	If pblnSuccess Then
		WriteOutput "<b>Existing manufacturers removed. All products have been updated to <em>No Manufacturer</em></b><br />" & vbcrlf
	Else
		WriteOutput "<b><font color=red>Error removing existing manufacturers</font></b><br />" & pstrLocalError & vbcrlf
	End If
	
	ClearVendors = pblnSuccess

End Function	'ClearVendors

'****************************************************************************************************************************************************************

Function Counter(ByRef lngCounter)
	lngCounter = lngCounter + 1
	Counter = lngCounter
End Function	'Counter

'***********************************************************************************************

Function createEmptyProduct(byVal strProductCode)

Dim plngUID
Dim pobjRS
Dim pstrLocalError
Dim pstrSQL
Dim pstrTempCode

	If Len(strProductCode) = 0 Then
		createEmptyProduct = -1
		Exit Function
	End If

	pstrTempCode = "ssTempCode" & Session.SessionID & Now()
	pstrSQL = "Insert Into sfProducts (prodID, prodEnabledIsActive) Values ('" & pstrTempCode & "', 0)"
	If Execute_NoReturn(pstrSQL, pstrLocalError) Then
		pstrSQL = "Select sfProductID From sfProducts Where prodID='" & pstrTempCode & "'"
		Set pobjRS = GetRS(pstrSQL)
		If pobjRS.EOF Then
			plngUID = -1
		Else
			plngUID = pobjRS.Fields("sfProductID").Value
			pstrSQL = "Update sfProducts Set prodID=" & wrapSQLValue(strProductCode, False, enDatatype_string) & " Where sfProductID=" & plngUID
			Call Execute_NoReturn(pstrSQL, pstrLocalError)
		End If
		Call ReleaseObject(pobjRS)
	Else
		plngUID = -1
	End If

	createEmptyProduct = plngUID

End Function	'createEmptyProduct

'****************************************************************************************************************************************************************

Function customBoolean(byVal strValue)

Dim pstrTemp
Dim i
Dim pblnValue

	'Set initial condition
	pblnValue = False
	pstrTemp = LCase(Trim(strValue) & "")
	
	For i = 0 To UBound(maryCustomBooleanValues)
		If CBool(CStr(pstrTemp) = CStr(maryCustomBooleanValues(i))) Then
			pblnValue = True
			Exit For
		End If
	Next 'i
	
	customBoolean = pblnValue

End Function	'customBoolean

'****************************************************************************************************************************************************************

Function deleteXMLProfile(byVal strProfileID)

Dim pblnResult
Dim pobjNodeList

		pblnResult = False
		Set pobjNodeList = mobjXMLProfiles.selectNodes("profiles/profile")
		If pobjNodeList.length > 1 Then
			mobjXMLProfiles.documentElement.removeChild mobjXMLProfiles.selectSingleNode("profiles/profile[@profileID='" & strProfileID & "']")
			pblnResult = True
		Else
			Response.Write "<h4><font color=red>Error deleting profile. There is only one saved profile remaining.</font></h4>"
		End If
		Set pobjNodeList = Nothing

		pblnResult = WriteXMLProfile
		
		deleteXMLProfile = pblnResult

End Function	'deleteXMLProfile

'****************************************************************************************************************************************************************

Function dbFieldName(byVal strSpreadSheetFieldName)

Dim i
Dim pstrTemp

	For i = 0 To UBound(maryFields)
		'debugprint maryFields(i)(enSourceFieldName) & " - " & strSpreadSheetFieldName, maryFields(i)(enSourceFieldName) = strSpreadSheetFieldName
		If maryFields(i)(enSourceFieldName) = strSpreadSheetFieldName Then
			pstrTemp = maryFields(i)(enTargetFieldName)
		End If
	Next 'i
	
	If Len(pstrTemp) = 0 Then
		WriteOutput "<h4><font color=red>Your spreadsheet mappings are incorrect</font></h4>" & vbcrlf
		Response.Flush
	End If
	
	dbFieldName = pstrTemp

End Function	'dbFieldName

'****************************************************************************************************************************************************************

Sub DoImport(objcnnSource, strTableSource)

Dim paryAttDetailNames
Dim paryAttDetailUIDs
Dim paryProduct(2)	'product UID, product attribute detail uid(s), product attribute detail name(s)
Dim pblnAlreadyExists
Dim pblnAborted
Dim pblnSuccess
Dim pblnRequiredValueSet
Dim pdicCategories
Dim pdicManufacturers
Dim pdicVendors
Dim pdtCurrentTime
Dim pdtLastTime
Dim pdtProductStartTime
Dim pdtMustFinishBy
Dim plngCurrentRecord
Dim plngFieldCount
Dim plngPos1, plngPos2
Dim plngProductUID
Dim plngRecordsToBeImported
Dim plngWorkingUID
Dim pobjRSCategory
Dim pobjRSSource
Dim pstrBasePrice
Dim pstrDSN_Target
Dim pstrFieldValue
Dim pstrFieldname_temp
Dim pstrFieldname_temp_wrapped
Dim pstrFilter
Dim pstrProductID
Dim pstrProductCode
Dim pstrProductName
Dim pstrSQL
Dim pstrSQLInsert_Part1
Dim pstrSQLInsert_Values
Dim pstrSQLUpdate_Alternate
Dim pstrSQLUpdate

	'Added for tracking time
	pdtCurrentTime = Time()
	pdtLastTime = pdtCurrentTime
	pdtProductStartTime = pdtCurrentTime

	Session("ssImportProduct_StartImport") = Now()
	Call RecordTime("<b>Starting Import</b>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)

	pdtMustFinishBy = DateAdd("s", Server.ScriptTimeout - clngBufferTimeout, Now())	'give it a 30 second buffer
	'For Testing
	'pdtMustFinishBy = DateAdd("s", Server.ScriptTimeout - Server.ScriptTimeout, Now())	'give it a 30 second buffer
	
	pblnSuccess = True
	pblnAborted = False
	plngCurrentRecord = 0

	'Open the source data
	pstrSQL = strTableSource
    Set pobjRSSource = Server.CreateObject("ADODB.Recordset")
	pobjRSSource.CursorLocation = 2 'adUseClient - need to use server to do sorting later
	Call RecordTime("<i>Opening data source . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
	pobjRSSource.Open pstrSQL, objcnnSource, 3, 2	'adOpenStatic, Pessimistic Lock.
	plngRecordsToBeImported = pobjRSSource.RecordCount
	
	Call RecordTime("<i>Data source opened</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
	plngFieldCount = pobjRSSource.Fields.Count - 1
	
	If False Then
		Do While Not pobjRSSource.EOF
			Response.Write "<strong>Field " & maryFields(0)(enSourceFieldName) & ": " & pobjRSSource.Fields(maryFields(0)(enSourceFieldName)).Value & "</strong><br />" & vbcrlf
			pobjRSSource.MoveNext
		Loop
		pobjRSSource.MoveFirst
	End If
			
	Call RecordTime("<i>Opening category table . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)

	If cblnSF5AE Then
		Set pobjRSCategory = GetRS("Select subcatID, subcatCategoryId, subcatName, Depth from sfSub_Categories Order By subcatName")
	Else
		Set pobjRSCategory = GetRS("Select catID, catName from sfCategories Order By catName")
	End If
	If pobjRSCategory.State = 1 Then
		Call RecordTime("<i>Category table opened. " & pobjRSCategory.RecordCount & " record(s).</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
	Else
		Call RecordTime("<i>Error opening category table.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
	End If
	
	'Check to see if this is resuming a previous import
	If Len(Session("ssNextProductIDToImport")) = 0 Then

		'Delete existing products
		If mlngImportType = enImportAll_DeleteExisting Then
			If DeleteAllProducts Then
				WriteOutput "<strong>The existing product catalog has been deleted</strong><br />" & vbcrlf
			Else
				WriteOutput "<strong><font color=red>There was an error deleting the existing product catalog</font></strong><br />" & vbcrlf
			End If
		End If

		'Now make sure the Categories are in the database
		Set pdicCategories = Server.CreateObject("Scripting.Dictionary")
		If CLng(mbytCreateCat) = cenCreateCat_CreateAndDelete Then Call ClearCategories
		If (CLng(mbytCreateCat) = cenCreateCat_Create) Or (CLng(mbytCreateCat) = cenCreateCat_CreateAndDelete) Then
			If Len(mstrCategoryColumn) > 0 Then
				'disabled since it is now handled individually for each product
				If mlngImportType = enImportInformationOnly Then Call verifyCategories(pobjRSSource, mstrCategoryColumn, pobjRSCategory, pdicCategories)
			End If
		End If
		
		'Now make sure the Manufacturers are in the database
		Set pdicManufacturers = Server.CreateObject("SCRIPTING.DICTIONARY")
		If CLng(mbytCreateMfg) = cenCreateMfg_CreateAndDelete Then Call ClearManufacturers
		If (CLng(mbytCreateMfg) = cenCreateMfg_Create) Or (CLng(mbytCreateMfg) = cenCreateMfg_CreateAndDelete) Then
			If Len(spreadsheetFieldName(cstrTarget_Manufacturer)) > 0 Then
				'disabled since it is now handled individually for each product
				If mlngImportType = enImportInformationOnly Then Call verifyManufacturers(pobjRSSource, spreadsheetFieldName(cstrTarget_Manufacturer), pdicManufacturers)
			End If
		End If
		
		'Now make sure the vendors are in the database
		Set pdicVendors = Server.CreateObject("SCRIPTING.DICTIONARY")
		If CLng(mbytCreateVend) = cenCreateVend_CreateAndDelete Then Call ClearVendors
		If (CLng(mbytCreateVend) = cenCreateVend_Create) Or (CLng(mbytCreateVend) = cenCreateVend_CreateAndDelete) Then
			If Len(spreadsheetFieldName(cstrTarget_Vendor)) > 0 Then	
				'disabled since it is now handled individually for each product
				If mlngImportType = enImportInformationOnly Then Call verifyVendors(pobjRSSource, spreadsheetFieldName(cstrTarget_Vendor), pdicVendors)
			End If
		End If
	
		If mlngImportType = enImportInformationOnly Then
			WriteOutput "<h4>Category, manufacturer and vendor information imported</h4>" & vbcrlf
			Exit Sub		
		End If
	
	Else
		'move to next product to import
		pstrProductID = Trim(Session("ssNextProductIDToImport"))
		
		If isObject(Session("ssImportedCategories")) Then
			Set pdicCategories = Session("ssImportedCategories")
		Else
			Set pdicCategories = Server.CreateObject("SCRIPTING.DICTIONARY")
		End If
		
		If isObject(Session("ssImportedManufacturers")) Then
			Set pdicManufacturers = Session("ssImportedManufacturers")
		Else
			Set pdicManufacturers = Server.CreateObject("SCRIPTING.DICTIONARY")
		End If
		
		If isObject(Session("ssImportedVendors")) Then
			Set pdicVendors = Session("ssImportedManufacturers")
		Else
			Set pdicVendors = Server.CreateObject("SCRIPTING.DICTIONARY")
		End If
		
		WriteOutput "<strong>Starting import with " & pstrProductID & " . . .</strong><br />" & vbcrlf
		Do While Not pobjRSSource.EOF
			If Trim(pobjRSSource.Fields(maryFields(0)(enSourceFieldName)).Value & "") <> pstrProductID Then
				plngCurrentRecord = plngCurrentRecord + 1
				pobjRSSource.MoveNext
			Else
				Session("ssNextProductIDToImport") = ""
				Exit Do
			End If
		Loop
	End If	'Len(Session("ssNextProductIDToImport")) = 0
	
	'Set up the inital SQL strings
	For i = 0 To UBound(maryFields)
		If Len(pstrSQLInsert_Part1) = 0 Then
			pstrSQLInsert_Part1 = "Insert Into sfProducts (" & WrapFieldName(maryFields(i)(enTargetFieldName))
		Else
			pstrSQLInsert_Part1 = pstrSQLInsert_Part1 & ", " & WrapFieldName(maryFields(i)(enTargetFieldName))
		End If
	Next 'i
	pstrSQLInsert_Part1 = pstrSQLInsert_Part1 & ") Values "
	'debugprint "pstrSQLInsert_Part1",pstrSQLInsert_Part1
	
	Do While Not pobjRSSource.EOF
	
		If Len(mstrImportTypeColumn) > 0 Then
			pstrFieldValue = Trim(pobjRSSource.Fields(mstrImportTypeColumn).Value)
			Select Case LCase(pstrFieldValue)
				Case ""
					mlngImportType = enImportAll
				Case ""
					mlngImportType = enImportAllDeleteCategories
				Case ""
					mlngImportType = enImportAll_DeleteExisting
				Case ""
					mlngImportType = enImportNewOnly
				Case ""
					mlngImportType = enImportNewOnlyAllowDuplicateCodes
				Case ""
					mlngImportType = enImportInvPriceOnly
				Case ""
					mlngImportType = enImportUpdateSelectedFieldsOnly
				Case ""
					mlngImportType = enImportInformationOnly = 7
				Case "delete"
					mlngImportType = enImportDelete
				Case Else
					mlngImportType = mlngDefaultImportType
			End Select
			WriteOutput "Import style to use: " & maryImportTypes(mlngImportType)(0) & "<br />" & vbcrlf
		Else
			mlngImportType = mlngDefaultImportType
		End If
	
		'Start product import time tracking
		pdtProductStartTime = Time()
		Call RecordTime("<br /><b>Starting Product</b>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
		
		pstrProductID = getProductID(pobjRSSource)
		pstrProductName = getProductName(pobjRSSource)
		'Response.Write "<strong>Field " & maryFields(spreadsheetIndexByFieldName("Code"))(enSourceFieldName) & ": " & pobjRSSource.Fields(maryFields(0)(enSourceFieldName)).Value & "</strong><br />" & vbcrlf
		
		'Custom Rule
		'If Len(pstrProductID) = 0 Then pstrProductID = pobjRSSource.Fields("IV_KEYS").Value
		
		If mlngImportType = enImportDelete And Len(pstrProductID) > 0 Then
			WriteOutput "<font size='-1'>(" & plngCurrentRecord + 1 & " of " & plngRecordsToBeImported & ")</font><strong> Deleting " & pstrProductID & " . . .</strong><br />" & vbcrlf
		
		ElseIf Len(pstrProductID) > 0 Then
		
			If False Then
				Response.Write "<fieldset><legend>Source data: " & pstrProductID & " (" & pstrProductID & ")</legend>" & vbcrlf
				For i = 0 To UBound(maryFields)
					If Len(maryFields(i)(enSourceFieldName)) > 0 Then
						pstrFieldValue = Trim(pobjRSSource.Fields(maryFields(i)(enSourceFieldName)).Value)
						Response.Write "<fieldset><legend>Source Field: " & maryFields(i)(enSourceFieldName) & "</legend>" & vbcrlf
						Response.Write "{" & pstrFieldValue & "}" & vbcrlf
						Response.Write "</fieldset>" & vbcrlf
					End If
				Next 'i
				Response.Write "</fieldset>" & vbcrlf
			End If

			'On Error Resume Next
			If Len(pstrProductName) = 0 Then
				WriteOutput "<font size='-1'>(" & plngCurrentRecord + 1 & " of " & plngRecordsToBeImported & ")</font><strong> Importing " & pstrProductID & " . . .</strong><br />" & vbcrlf
			Else
				WriteOutput "<font size='-1'>(" & plngCurrentRecord + 1 & " of " & plngRecordsToBeImported & ")</font><strong> Importing " & pstrProductID & ": " & pstrProductName & " . . .</strong><br />" & vbcrlf
			End If
			
			'Clear the variables
			pstrSQLInsert_Values = ""
			pstrSQLUpdate = ""
			pblnAlreadyExists = False
			
			For i = 0 To UBound(maryFields)
				pblnRequiredValueSet = False
				'Custom implementation - unique to prodID
				If Len(maryFields(i)(enSourceFieldName)) = 0 Then
					pstrFieldValue = Trim(maryFields(i)(enDefaultValue))
				Else
					'this section added for 255+ character fields
					On Error Resume Next
					pstrFieldValue = Trim(pobjRSSource.Fields(maryFields(i)(enSourceFieldName)).Value)
					If Err.number <> 0 Then
						WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
						WriteOutput "Error reading field <b>" & maryFields(i)(enSourceFieldName) & "</b>. Please note, if this field contains more than 255 characters and you are importing from an Excel spreadsheet you will need to manually update this field. There is a limition with this data source. Please refer to the help guide for additional details.<br />" & vbcrlf
						Err.Clear
					End If
					On Error Goto 0

					Select Case maryFields(i)(enTargetFieldName)
						Case cstrTarget_Manufacturer
							Call RecordTime("<i>Verifying manufacturer <em>" & pstrFieldValue & "</em> . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
							If Not isNumeric(pstrFieldValue) Then
								If Len(pstrFieldValue) = 0 Then
									pstrFieldValue = mlngDefaultManufacturerID
								ElseIf pdicManufacturers.Exists(pstrFieldValue) Then
									pstrFieldValue = pdicManufacturers(pstrFieldValue)
									Call RecordTime("<i>Manufacturer <em>" & pstrFieldValue & "</em> previously retrieved (" & pstrFieldValue & ").</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
								Else
									plngWorkingUID = getManufacturerByName(pstrFieldValue, CBool((CLng(mbytCreateMfg) = cenCreateMfg_Create) Or (CLng(mbytCreateMfg) = cenCreateMfg_CreateAndDelete)))
									Call RecordTime("<i>Manufacturer <em>" & pstrFieldValue & "</em> created and/or retrieved from database (" & plngWorkingUID & ").</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
									If plngWorkingUID = -1 Then
										If isNumeric(mlngDefaultManufacturerID) Then
											pstrFieldValue = mlngDefaultManufacturerID
										Else
											If pdicManufacturers.Exists(mlngDefaultManufacturerID) Then
												pstrFieldValue = pdicManufacturers(mlngDefaultManufacturerID)
											Else
												pstrFieldValue = 1
											End If
										End If
										WriteOutput "&nbsp;&nbsp;&nbsp;<font color=red>Manufacturer " & pstrFieldValue & " does not exist. Default Manufacturer ID of " & mlngDefaultManufacturerID & " used.</font><br />" & vbcrlf
									Else
										pdicManufacturers.Add pstrFieldValue, plngWorkingUID
										pstrFieldValue = plngWorkingUID
									End If
								End If
							End If
							'debugprint maryFields(i)(enTargetFieldName), pstrFieldValue
							Call RecordTime("<i>Manufacturer verified (" & pstrFieldValue & ").</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						Case cstrTarget_Vendor
							Call RecordTime("<i>Verifying vendor <em>" & pstrFieldValue & "</em> . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
							If Not isNumeric(pstrFieldValue) Then
								If Len(pstrFieldValue) = 0 Then
									pstrFieldValue = mlngDefaultVendorID
								ElseIf pdicVendors.Exists(pstrFieldValue) Then
									pstrFieldValue = pdicVendors(pstrFieldValue)
								Else
									plngWorkingUID = getVendorByName(pstrFieldValue, CBool((CLng(mbytCreateVend) = cenCreateVend_Create) Or (CLng(mbytCreateVend) = cenCreateVend_CreateAndDelete)))
									If plngWorkingUID = -1 Then
										If isNumeric(mlngDefaultVendorID) Then
											pstrFieldValue = mlngDefaultVendorID
										Else
											If pdicVendors.Exists(mlngDefaultVendorID) Then
												pstrFieldValue = pdicVendors(mlngDefaultVendorID)
											Else
												pstrFieldValue = 1
											End If
										End If
										WriteOutput "&nbsp;&nbsp;&nbsp;<font color=red>Vendor " & pstrFieldValue & " does not exist. Default Vendor ID of " & mlngDefaultVendorID & " used.</font><br />" & vbcrlf
									Else
										pdicVendors.Add pstrFieldValue, plngWorkingUID
										pstrFieldValue = plngWorkingUID
									End If
								End If
							End If
							Call RecordTime("<i>Vendor verified (" & pstrFieldValue & ").</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						Case cstrTarget_Code
							pstrFieldValue = pstrProductID			
						Case Else
							'debugprint "Else: " & maryFields(i)(enTargetFieldName), pstrFieldValue					
					End Select
				End If	'Len(maryFields(i)(enSourceFieldName)) = 0
				
				On error goto 0
				'Check for textboxs and defaults
				If CBool(maryFields(i)(enDisplayType)=enDisplayType_checkbox) And CBool(maryFields(i)(enFieldDataType)=enDatatype_boolean) Then
					pstrFieldValue = customBoolean(pstrFieldValue)
					'Response.Write maryFields(i)(enSourceFieldName) & " (" & maryFields(i)(enTargetFieldName) & "): " & pstrFieldValue & "<br />"
				End If

				pstrFieldValue = checkReplacements(pstrFieldValue, pstrProductID)
				'Response.Write maryFields(i)(enSourceFieldName) & " (" & maryFields(i)(enTargetFieldName) & "): " & pstrFieldValue & "<br />"
				'Check for some common import errors
				pstrFieldValue = checkRequiredValue(maryFields(i)(enTargetFieldName), pstrFieldValue, pblnRequiredValueSet)
				Select Case i
					Case clngTarget_Cost, clngTarget_Price, clngTarget_SalePrice 
						If CBool(Len(pstrFieldValue) = 0) Then
							pstrFieldValue = 0
							pblnRequiredValueSet = True
						End If
				End Select
				If Len(pstrSQLInsert_Values) = 0 Then
					pstrSQLInsert_Values = "(" & wrapSQLValue(pstrFieldValue, True, maryFields(i)(enFieldDataType))
				Else
					pstrSQLInsert_Values = pstrSQLInsert_Values & ", " & wrapSQLValue(pstrFieldValue, True, maryFields(i)(enFieldDataType))
				End If
				
				'No reason to update codes
				If maryFields(i)(enTargetFieldName) <> cstrTarget_Code Then
					If (mlngImportType = enImportUpdateSelectedFieldsOnly AND Len(maryFields(i)(enSourceFieldName)) > 0) Or (mlngImportType <> enImportUpdateSelectedFieldsOnly) Then
						If Not pblnRequiredValueSet Then
							If Len(pstrSQLUpdate) = 0 Then
								pstrSQLUpdate = makeSQLUpdate(maryFields(i)(enTargetFieldName),pstrFieldValue, True, maryFields(i)(enFieldDataType))
							Else
								pstrSQLUpdate = pstrSQLUpdate & ", " & makeSQLUpdate(maryFields(i)(enTargetFieldName),pstrFieldValue, True, maryFields(i)(enFieldDataType))
							End If
						End If
					End If
				End If
				
			Next 'i
			pstrSQLInsert_Values = pstrSQLInsert_Values & ")"
			pstrSQL = pstrSQLInsert_Part1 & pstrSQLInsert_Values
			
			plngProductUID = getProductUIDByCode(pstrProductID)			
			pblnAlreadyExists = CBool(plngProductUID <> -1)
			If mlngImportType = enImportNewOnlyAllowDuplicateCodes And pblnAlreadyExists Then
				plngProductUID = createEmptyProduct(pstrProductID)			
				WriteOutput "&nbsp;&nbsp;&nbsp;Product created with duplicate product code <em>" & pstrProductID & "</em><br />" & vbcrlf
			End If

			i = clngTarget_Price
			If Len(maryFields(i)(enSourceFieldName)) = 0 Then
				pstrBasePrice = maryFields(i)(enDefaultValue)
			Else
				pstrBasePrice = pobjRSSource.Fields(maryFields(i)(enSourceFieldName)).Value
			End If

			'debugprint "mlngImportType",mlngImportType
			If mlngImportType = enImportInvPriceOnly Then
				pstrSQLUpdate = ""
				
				Dim maryFieldsToCheck
				maryFieldsToCheck = Array(clngTarget_Cost, clngTarget_Price, clngTarget_SaleIsActive, clngTarget_SalePrice)

				For j = 0 To UBound(maryFieldsToCheck)
					i = maryFieldsToCheck(j)
					If Len(i) > 0 Then
						If Len(maryFields(i)(enSourceFieldName)) = 0 Then
							pstrFieldValue = maryFields(i)(enDefaultValue)
						Else
							pstrFieldValue = pobjRSSource.Fields(maryFields(i)(enSourceFieldName)).Value
						End If
					Else
						pstrFieldValue = ""
					End If

					If Len(pstrFieldValue) > 0 Then
						If Len(pstrSQLUpdate) = 0 Then
							pstrSQLUpdate = makeSQLUpdate(maryFields(i)(enTargetFieldName),pstrFieldValue, True, maryFields(i)(enFieldDataType))
						Else
							pstrSQLUpdate = pstrSQLUpdate & ", " & makeSQLUpdate(maryFields(i)(enTargetFieldName), pstrFieldValue, True, maryFields(i)(enFieldDataType))
						End If
					End If
				Next 'j

			End If	'mlngImportType = enImportInvPriceOnly
			If Len(pstrSQLUpdate) > 0 Then pstrSQLUpdate = "Update sfProducts Set " & pstrSQLUpdate & " Where sfProductID=" & wrapSQLValue(plngProductUID, True, enDatatype_number)

			If False Then
				Response.Write "<fieldset><legend></legend>"
				Response.Write "pstrSQL: " & pstrSQL & "<br />"
				Response.Write "pstrSQLUpdate: " & pstrSQLUpdate & "<br />"
				Response.Write "pblnAlreadyExists: " & pblnAlreadyExists & "<br />"
				Response.Write "</fieldset>"
			End If
			
			'On Error Goto 0
			
			'Implement Error handling for update section
			On Error Resume Next
			If pblnAlreadyExists Then
				If mlngImportType = enImportNewOnly Then
					WriteOutput "&nbsp;&nbsp;&nbsp;Ignored " & pstrProductID & ", a product with this code already exists.<br />" & vbcrlf
				Else
					If mlngImportType = enImportInvPriceOnly Then
						Call RecordTime("<i>Issuing product update to database . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						If Len(pstrSQLUpdate) > 0 Then cnn.Execute pstrSQLUpdate,, 128
						Call RecordTime("<i>Product updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						If Err.number = 0 Then
							If Len(pstrSQLUpdate) > 0 Then
								WriteOutput "&nbsp;&nbsp;&nbsp;Updated " & pstrProductID & " pricing<br />" & vbcrlf
							Else
								WriteOutput "&nbsp;&nbsp;&nbsp;No pricing updates issued for " & pstrProductID & ".<br />" _
										  & "&nbsp;&nbsp;&nbsp;- This will occur if no columns are specified for prices/inventory levels or if the values for the product are empty.<br />" & vbcrlf
							End If
						Else
							If Not commonError(Err) Then
								WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
								WriteOutput "pstrSQLUpdate: " & pstrSQLUpdate & "<br />" & vbcrlf
								Err.Clear
							Else
								WriteOutput "<font color=red>&nbsp;&nbsp;&nbsp;Failed updating " & pstrProductID & " pricing<br /></font><br />" & vbcrlf
							End If
						End If
					Else
						If Len(pstrSQLUpdate) > 0 Then
							Call RecordTime("<i>Issuing product update to database . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
							cnn.Execute pstrSQLUpdate,, 128
							Call RecordTime("<i>Product updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
							If Err.number = 0 Then
								WriteOutput "&nbsp;&nbsp;&nbsp;Updated " & pstrProductID & "<br />" & vbcrlf
							Else
								'This section added because the "Description" field causes an error on some imports - updates only
								If InStr(1, Err.Description, "The search key was not found in any record") Then

									plngPos1 = Instr(1, pstrSQLUpdate, " [Description]")
									plngPos2 = Instr(1, pstrSQLUpdate, ", [UpSellMessage]")
									pstrSQLUpdate_Alternate = Left(pstrSQLUpdate, plngPos1 - 1) & Right(pstrSQLUpdate, Len(pstrSQLUpdate) - plngPos2)
									
									Err.Clear
									Call RecordTime("<i>Issuing product update (alternate insertion because of long fields) to database . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
									cnn.Execute pstrSQLUpdate_Alternate,, 128
									Call RecordTime("<i>Product updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
									
									If Err.number = 0 Then
										WriteOutput "<font color=red>The long description (<i>Description</i> field in the Products table) for this product could not be updated. You will need to manually enter this information. All other fields were updated normally.</font><br />" & vbcrlf
									Else
										'Continue with routine error handling below
										WriteOutput "<font color=red>&nbsp;&nbsp;&nbsp;Error updating product description. This error may be caused if you're importing a long description or if the current description is long. The update for this product will need to be manually set.</font><br />" & vbcrlf
									End If
								Else
									'Continue with routine error handling below
								End If
							End If	'Err.number = 0
						Else
							Call RecordTime("<i>Product updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
							WriteOutput "<font color=red>&nbsp;&nbsp;&nbsp;No update issued for this product; no fields specified.</font><br />" & vbcrlf
						End If	'Len(pstrSQLUpdate) > 0
					End If	'mlngImportType = enImportInvPriceOnly
				End If	'mlngImportType = enImportNewOnly
			Else
				Call RecordTime("<i>Issuing product insertion to database . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
				cnn.Execute pstrSQL,, 128
				Call RecordTime("<i>Product inserted.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
				If Err.number = 0 Then WriteOutput "&nbsp;&nbsp;&nbsp;Inserted " & pstrProductID & "<br />" & vbcrlf
			End If	'pblnAlreadyExists

			'Retrieve the uid
			If plngProductUID = -1 And Err.number = 0 Then
				Call RecordTime("<i>Retrieving product UID . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
				plngProductUID = getProductUIDByCode(pstrProductID)
				Call RecordTime("<i>Product UID retrieved.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
			End If
			If Err.number <> 0 Then
				'WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
				If InStr(1, Err.Description, "The search key was not found in any record") Or InStr(1, Err.Description, "Unclosed quotation mark before the character string") Then
					WriteOutput "<font color=red>This error may be caused if you're importing a large field (> 255 characters) from an Excel spreadsheet or the field you're importing contains more information that the database is set for (ex. Short description is limited to 255 characters).</font><br />" & vbcrlf
					If mblnssDebugImport Then
						If pblnAlreadyExists Then
							WriteOutput "pstrSQLUpdate: " & pstrSQLUpdate & "<br />" & vbcrlf
						Else
							WriteOutput "pstrSQL: " & pstrSQL & "<br />" & vbcrlf
						End If	'pblnAlreadyExists
					End If	'cblnExtendedDebugging
				ElseIf Not commonError(Err) Then
					WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
					If pblnAlreadyExists Then
						WriteOutput "pstrSQLUpdate: " & pstrSQLUpdate & "<br />" & vbcrlf
					Else
						WriteOutput "pstrSQL: " & pstrSQL & "<br />" & vbcrlf
					End If	'pblnAlreadyExists
				End If
				Err.Clear
			Else
				If Not (CBool(pblnAlreadyExists) And CBool(mlngImportType = enImportNewOnly)) Then
					On Error Goto 0

					Call RecordTime("<i>Updating attributes . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
					'paryProduct(0) = plngProductUID
					paryProduct(0) = pstrProductID
					paryProduct(1) = ""	'paryAttDetailUIDs
					paryProduct(2) = ""	'paryAttDetailNames
					Call saveAttributes(pobjRSSource, paryProduct)
					
					paryAttDetailUIDs = paryProduct(1)
					paryAttDetailNames = paryProduct(2)
					Dim j
					Dim pvntTemp

					If False Then	'True	,False
						If isArray(paryAttDetailUIDs) Then
							Response.Write "<fieldset><legend>Pre-sort</legend>"
							For i = 0 To UBound(paryAttDetailUIDs)
								debugprint i, paryAttDetailUIDs(i)
							Next 'i
							Response.Write "</fieldset>"
						End If
					End If
					
					'sort the attributes by uid, inefficient sort but effective
					If isArray(paryAttDetailNames) Then
						For i = 0 To UBound(paryAttDetailUIDs)-1
							For j = 1 To UBound(paryAttDetailUIDs)
								If paryAttDetailUIDs(j) < paryAttDetailUIDs(i) Then
									pvntTemp = paryAttDetailUIDs(j)
									paryAttDetailUIDs(j) = paryAttDetailUIDs(i)
									paryAttDetailUIDs(i) = pvntTemp
									
									pvntTemp = paryAttDetailNames(j)
									paryAttDetailNames(j) = paryAttDetailNames(i)
									paryAttDetailNames(i) = pvntTemp
								End If
							Next 'j
						Next 'i
					Else
						paryAttDetailUIDs = plngProductUID
						paryAttDetailNames = pstrProductName
					End If
					
					paryProduct(1) = paryAttDetailUIDs
					paryProduct(2) = paryAttDetailNames

					If False Then	'True	,False
						If isArray(paryAttDetailUIDs) Then
							Response.Write "<fieldset><legend>Post-sort</legend>"
							For i = 0 To UBound(paryAttDetailUIDs)
								debugprint i, paryAttDetailUIDs(i)
							Next 'i
							Response.Write "</fieldset>"
						End If
					End If
					Call RecordTime("<i>Attributes updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
				
					If Not mlngImportType = enImportInvPriceOnly Then
						If CBool(Len(pstrSQLUpdate) > 0 And CBool(mlngImportType = enImportUpdateSelectedFieldsOnly)) _
						   OR CBool(CBool(mlngImportType = enImportAll) Or CBool(mlngImportType = enImportAll_DeleteExisting) Or CBool(mlngImportType = enImportNewOnly) Or CBool(mlngImportType = enImportAllDeleteCategories)) Then
						   
							Call RecordTime("<i>Updating product detail link . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
							pblnSuccess = pblnSuccess And updateDetailLink(pobjRSSource, plngProductUID, pstrProductID)
							Call RecordTime("<i>Product detail link updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						End If

						If Len(mstrCategoryColumn) > 0 Then
							Call RecordTime("<i>Updating categories . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
							Call saveCategories(pobjRSSource, pstrProductID, pobjRSCategory, pdicCategories)
							Call RecordTime("<i>Categories updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						End If
						
						If cblnSF5AE Then
							Call RecordTime("<i>Updating gift wrap . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
							Call saveGiftWrap(pobjRSSource, pstrProductID)
							Call RecordTime("<i>Gift wrap updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						End If

						Call RecordTime("<i>Updating custom . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						Call saveCustom(pobjRSSource, pstrProductID)
						Call RecordTime("<i>Custom updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)

						pstrFieldValue = ""	'reset value

						'This line added because sometimes the empty column imports as False
						If pstrFieldValue = "False" Then pstrFieldValue = ""
					End If	'Not mlngImportType = enImportInvPriceOnly

					If cblnSF5AE Then
						Call RecordTime("<i>Updating inventory . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						pblnSuccess = pblnSuccess And setInventoryLevels(pobjRSSource, paryProduct, CBool(mlngImportType = enImportInvPriceOnly))
						Call RecordTime("<i>Inventory updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						
						Call RecordTime("<i>Updating volume pricing . . .</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
						pblnSuccess = pblnSuccess And setVolumePricing(pobjRSSource, pstrProductID, pstrBasePrice)
						Call RecordTime("<i>Volume pricing updated.</i>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
					End If
				End If	'Not(pblnAlreadyExists And mlngImportType = enImportNewOnly)
			End If	'Err.number <> 0
			
			'Turn error handling off
			On Error Goto 0

		Else
			If Len(pstrProductName) = 0 Then
				WriteOutput "<font color=red>Could not import this product because the required <i>code</i> field was not present. No product code or name is available to identify this item.</font><br />" & vbcrlf
			Else
				WriteOutput "<font color=red>Could not import <i>" & pstrProductName & "</i> because the required <i>code</i> field was not present.</font><br />" & vbcrlf
			End If
			WriteOutput "This condition can occur if <ol><li>No column was selected for the code field</li>" _
						 & "<li>You're importing from a .xls data source and your product codes contain a mixture of numeric and alphanumeric entries. See the help file for this situation</li>" & vbcrlf _
						 & "<li>You're importing from a .xls data source and you have empty rows or you delete rows by clearing the cells rather than deleting them. See the help file for this situation</li>" & vbcrlf _
						 & "</ol>" & vbcrlf
		End If	'Len(pstrProductID) > 0
	
		'Note: The EOF check was added because during testing of forced restart (timeout check set to 0) EOF error occured
		If Not pobjRSSource.EOF Then
			pobjRSSource.MoveNext
			plngCurrentRecord = plngCurrentRecord + 1
		End If
		
		Response.Write  "<script language=javascript>window.scrollTo(0, 999999999);</script>" & vbcrlf
		Response.Write  "<script language=javascript>if (OutputWindow != null) OutputWindow.scrollTo(0, 999999999);</script>" & vbcrlf
		Response.Flush
		
		If pblnSuccess Then
			If Not pobjRSSource.EOF Then
				Session("ssNextProductIDToImport") = Trim(pobjRSSource.Fields(maryFields(clngTarget_ProductID)(enSourceFieldName)).Value & "")
			Else
				Session("ssNextProductIDToImport") = ""
			End If
			If CBool(Now() > pdtMustFinishBy) Then
				pblnAborted = True
				Exit Do
			End If
		Else
			'this check added to handle database connection timeouts
			If Not pobjRSSource.EOF Then 
				WriteOutput "<font color=red>Unexpected error. Attempting to restart with product <i>" & Trim(pobjRSSource.Fields(maryFields(0)(enSourceFieldName)).Value & "") & "</i></font><br />" & vbcrlf
			Else
			End If
			pblnAborted = True
			Exit Do
		End If

		Call RecordTime("<b>Product Finished</b>", pdtProductStartTime, pdtCurrentTime, pdtLastTime)
	Loop
    'pobjRSSource.MoveFirst
    
	Call ReleaseObject(pobjRSCategory)
	Call ReleaseObject(pobjRSSource)
	'Call ReleaseObject(pobjcnnSource)
	
	'Finished import so clean up session
	If pblnAborted Then
		Set Session("ssImportedCategories") = pdicCategories
		Set Session("ssImportedManufacturers") = pdicManufacturers
		Set Session("ssImportedManufacturers") = pdicVendors
	Else
		Dim pElaspedHours
		Dim pElapsedSeconds
		
		pElaspedHours = DateDiff("h", Now(), Session("ssImportProduct_StartImport"))
		pElapsedSeconds = pElaspedHours * 60 - DateDiff("s", Now(), Session("ssImportProduct_StartImport"))
		
		WriteOutput "<hr><h4>Import Completed</h4>"
		WriteOutput "<b>Import Started</b>: " & FormatDateTime(Session("ssImportProduct_StartImport"), 3) & "<br />"
		WriteOutput "<b>Import Completed</b>: " & FormatDateTime(Now(), 3) & "<br />"
		WriteOutput "<b>Elapsed Time</b>: " & FormatNumber(pElaspedHours, 2) & " hours, " _
						& FormatNumber(pElapsedSeconds, 2) & " seconds.<br />"
		WriteOutput "<hr>"
		
		Call WriteFailedImports
		
		Session.Contents.Remove("ssImportProduct_StartImport")
	
		Session.Contents.Remove("ssImportedCategories")
		Session.Contents.Remove("ssImportedManufacturers")
		Session.Contents.Remove("ssImportedVendors")
		
		Session.Contents.Remove("ssFailedImports")
		
		'The following items are defined in ssProduct_CommonFilter.asp
		Session.Contents.Remove("category" & "Combo")
		Session.Contents.Remove("manufacturer" & "Combo")
		Session.Contents.Remove("vendor" & "Combo")
		Session.Contents.Remove("product" & "Combo")
		Session.Contents.Remove("CategoryList")

	End If
	
	Set pdicCategories = Nothing
	Set pdicManufacturers = Nothing
	Set pdicVendors = Nothing
    
End Sub	'DoImport

'****************************************************************************************************************************************************************

Sub DoImport_Custom()

Dim pobjcnnSource
Dim pobjRS
Dim pobjRSCategory, pobjRSVendor, pobjRSManufacturer
Dim pstrSQLInsert_Part1, pstrSQLInsert_Values
Dim pstrSQLUpdate
Dim pstrSQL
Dim plngFieldCount
Dim pstrFieldValue
Dim pstrProductID
Dim pstrFilter
Dim pstrDSN_Target
Dim pstrFieldname_temp
Dim pblnAlreadyExists
Dim pobjRSInventory
Dim pobjRSInventoryInfo

    
	Call ReleaseObject(pobjRSCategory)
	Call ReleaseObject(pobjRSVendor)
	Call ReleaseObject(pobjRSManufacturer)
	Call ReleaseObject(pobjRS)
	Call ReleaseObject(pobjRSInventory)
	Call ReleaseObject(pobjRSInventoryInfo)
	Call ReleaseObject(pobjcnnSource)
    
End Sub	'DoImport_Custom

'****************************************************************************************************************************************************************

Function getCategoryUIDbyName(byVal vntCategoryName, byRef objrsCategories, byRef dicCategories)
'This function finds the category ID based on the input
'If a numeric input is found it is assumed to be a UID and simply returned
'For 1st level categories if it is not found in the dictionary than the recordset is checked and if found, added to the dictionary
'For subCategories the dictionary is checked only
'-1 is returned in the event no match is found

Dim pblnNumericCategoryID
Dim plngUID

	If Len(vntCategoryName) = 0 Then
		plngUID = -1
	ElseIf isNumeric(vntCategoryName) Then
		plngUID = vntCategoryName
	ElseIf instr(1, vntCategoryName, cstrSubcategoryDelimiter) > 0 Then
		If dicCategories.Exists(vntCategoryName) Then 
			plngUID = dicCategories(vntCategoryName)
		Else
			plngUID = -1
		End If
	Else
		If dicCategories.Exists(vntCategoryName) Then 
			plngUID = dicCategories(vntCategoryName)
		Else
			'Category is assumed to be a top level category
			plngUID = -1
			If cblnSF5AE Then
				objrsCategories.Filter = "subcatName = " & wrapSQLValue(vntCategoryName, False, enDatatype_string) & " AND Depth=0"
				If Not objrsCategories.EOF Then plngUID = objrsCategories.Fields("subcatID").Value
			Else
				objrsCategories.Filter = "catName = " & wrapSQLValue(vntCategoryName, False, enDatatype_string)
				If Not objrsCategories.EOF Then plngUID = objrsCategories.Fields("catID").Value
			End If
			If plngUID <> - 1 Then dicCategories.Add vntCategoryName, plngUID
		End If
	End If	'Not isNumeric(vntCategoryName)

	'Response.Write "getCategoryUIDbyName: " & vntCategoryName & " = " & plngUID & "<br />"
	
	getCategoryUIDbyName = plngUID
	
End Function	'getCategoryUIDbyName

'****************************************************************************************************************************************************************

Function getdbFieldName(byVal strSpreadSheetFieldName, byRef arySearch)

Dim i
Dim pstrTemp

	For i = 0 To UBound(arySearch)
		'debugprint arySearch(i)(enSourceFieldName) & " - " & strSpreadSheetFieldName, arySearch(i)(enSourceFieldName) = strSpreadSheetFieldName
		If arySearch(i)(enSourceFieldName) = strSpreadSheetFieldName Then
			pstrTemp = arySearch(i)(enTargetFieldName)
		End If
	Next 'i
	
	If Len(pstrTemp) = 0 Then
		WriteOutput "<h4><font color=red>Your spreadsheet mappings are incorrect</font></h4>" & vbcrlf
		Response.Flush
	End If
	
	getdbFieldName = pstrTemp

End Function	'getdbFieldName

'****************************************************************************************************************************************************************

Sub getProductDetail(byRef objXMLElement, byRef aryProductDetail)

Dim pvntTemp

	aryProductDetail = Array(enTargetFieldName, enSourceFieldName, enDisplayFieldName, enDefaultValue, enFieldDataType, enDisplayType)

	aryProductDetail(enTargetFieldName) = objXMLElement.attributes.item(0).value
	aryProductDetail(enSourceFieldName) = GetXMLElementValue(objXMLElement, "columnName")
	aryProductDetail(enDisplayFieldName) = GetXMLElementValue(objXMLElement, "displayName")
	aryProductDetail(enDefaultValue) = GetXMLElementValue(objXMLElement, "default")
	
	pvntTemp = GetXMLElementValue(objXMLElement, "dataType")
	If Len(pvntTemp) = 0 Or Not isNumeric(pvntTemp) Then pvntTemp = dataTypeFromText(pvntTemp)
	aryProductDetail(enFieldDataType) = CLng(pvntTemp)
	
	aryProductDetail(enDisplayType) = displayTypeFromText(GetXMLElementValue(objXMLElement, "displayType"))
	
	If False Then
		Response.Write "<fieldset><legend>" & aryProductDetail(enTargetFieldName) & "</legend>"
		Response.Write "columnName: " & aryProductDetail(enSourceFieldName) & "<br />"
		Response.Write "displayName: " & aryProductDetail(enDisplayFieldName) & "<br />"
		Response.Write "default: " & aryProductDetail(enDefaultValue) & "<br />"
		Response.Write "dataType: " & dataTypeToText(aryProductDetail(enFieldDataType)) & "(" & aryProductDetail(enFieldDataType) & ")<br />"
		Response.Write "displayType: " & displayTypeToText(aryProductDetail(enDisplayType)) & "(" & aryProductDetail(enDisplayType) & ")<br />"
		Response.Write "</fieldset>"
	End If

End Sub	'getProductDetail

'****************************************************************************************************************************************************************

Sub GetNodeValues(objNode)

Dim i
Dim e

	mstrcatID = objNode.attributes.item(0).nodeValue
	For i = 0 To objNode.childNodes.length
		Set e = objNode.childNodes.Item(i)
		Select Case e.nodeName
			Case "catParent": mstrcatParent = e.text
			Case "catDepth": mbytDepth = CInt(e.text)
			Case "catName": mstrcatName = e.text
			Case "catDescription": mstrcatDescription = e.text
			Case "catImage": mstrcatImage = e.text
			Case "catIsActive":  mblncatIsActive = e.text
			Case "catHeirarchy": mstrCatHeir = e.text
			Case "catBottom": mstrcatBottom = e.text
			Case "catStatus": mstrStatus = e.text
		End Select
	Next 'i

End Sub	'GetNodeValues

'****************************************************************************************************************************************************************

Function getProductID(byRef objRSSource)

Dim pstrProductID

	On Error Resume Next
	
	If Err.number <> 0 Then Err.Clear
	'Error handling added since the product code is a required field
	pstrProductID = Trim(objRSSource.Fields(maryFields(clngTarget_ProductID)(enSourceFieldName)).Value & "")
	If Err.number <> 0 Then
		WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
		If InStr(1, Err.Description, "Item cannot be found in the collection corresponding to the requested name or ordinal.") Then
			WriteOutput "<font color=red>The field <strong><i>Code</i></code> is required for import. This error may be caused if your datasource does not contain this column. Please recheck this column exists.</font><br />" & vbcrlf
		End If
		Err.Clear
		pstrProductID = ""
	End If
	
	On Error Goto 0
	
	getProductID = pstrProductID

End Function	'getProductID

'****************************************************************************************************************************************************************

Function getProductName(byRef objRSSource)

Dim pstrProductName

	On Error Resume Next
	
	If Err.number <> 0 Then Err.Clear
	'Error handling added since the product name is a optional field which may not be present
	pstrProductName = Trim(objRSSource.Fields(maryFields(clngTarget_ProductName)(enSourceFieldName)).Value & "")
	If Err.number <> 0 Then
		Err.Clear
		pstrProductName = ""
	End If
	
	On Error Goto 0
	
	getProductName = pstrProductName

End Function	'getProductName

'**************************************************************************************************************************************************

Function GetSelectOptions(aryData, vntValue)

Dim pstrDefaultText
Dim pstrOut
Dim pblnFound
Dim j

	pstrDefaultText = "- Use Default -"

	pstrOut = ""
	pblnFound = False
	
    For j = 0 To UBound(marySheetColumns)

		If LCase(vntValue) = LCase(aryData(j)) Then
			pstrOut = pstrOut & "  <option selected>" & aryData(j) & "</option>" & vbcrlf
			pblnFound = True
		Else
			pstrOut = pstrOut & "  <option>" & aryData(j) & "</option>" & vbcrlf
		End If
    Next 'j
	If pblnFound Then 
		pstrOut = pstrOut & "  <option value="""">" & pstrDefaultText & "</option>" & vbcrlf
	Else
		pstrOut = pstrOut & "  <option value="""" selected>" & pstrDefaultText & "</option>" & vbcrlf
	End If
	
	GetSelectOptions = pstrOut
 
End Function	'GetSelectOptions

'**************************************************************************************************************************************************

Function getAvailableTables

Dim paryTables
Dim pstrOut

	paryTables = getAvailableTables_Array
	For i = 0 To UBound(paryTables)
		pstrOut = pstrOut & "<option value=""" & paryTables(i)(0) & """>" & paryTables(i)(1) & "</option>"
	Next 'i

	getAvailableTables = pstrOut

End Function	'getAvailableTables

'**************************************************************************************************************************************************

Function getAvailableTables_Array

Dim objSchema
Dim pblnInserted
Dim plngPos
Dim pstrTableValue
Dim pstrTableName
Dim pstrType
Dim pstrprevTableName
Dim paryPrimary
Dim paryTemp
Dim plngCounter
Dim plngPointer

	Set objSchema = mobjCnn.OpenSchema(20) 
	'adSchemaTables = 20
	'adSchemaPrimaryKeys = 28

	plngCounter = -1
	Do Until objSchema.EOF
		'debugprint "TABLE_TYPE", objSchema("TABLE_TYPE")
		pstrType = objSchema("TABLE_TYPE")
		Select Case UCase(objSchema("TABLE_TYPE"))
			Case "TABLE", "VIEW"
				pstrTableName = objSchema("TABLE_NAME")
				plngPos = InStr(1, pstrTableName, "$")
				If plngPos > 1 Then pstrTableName = Left(pstrTableName, plngPos-1)
				If pstrprevTableName <> pstrTableName Then
					plngCounter = plngCounter + 1
					pstrTableValue = UCase(objSchema("TABLE_TYPE")) & " - " & pstrTableName
					pstrTableValue = pstrTableName & " (" & objSchema("TABLE_TYPE") & ")" 
					If plngCounter = 0 Then
						ReDim paryPrimary(plngCounter)
						paryPrimary(plngCounter) = Array(pstrTableName, pstrTableValue)
					Else
						ReDim paryTemp(plngCounter)
						plngPointer = 0
						pblnInserted  = False

						For i = 0 To UBound(paryPrimary)
							'If pstrTableValue > paryPrimary(i)(1) And Not pblnInserted Then
							If LCase(pstrTableName) < LCase(paryPrimary(i)(0)) And Not pblnInserted Then
								paryTemp(plngPointer) = Array(pstrTableName, pstrTableValue)
								plngPointer = plngPointer + 1
								pblnInserted = True
							End If
							paryTemp(plngPointer) = paryPrimary(i)
							plngPointer = plngPointer + 1
						Next
						
						'check if item was at bottom of list
						If Not isArray(paryTemp(plngCounter)) Then 
							paryTemp(plngCounter) = Array(pstrTableName, pstrTableValue)
						End If
						
						paryPrimary = paryTemp
					End If					

				End If
				pstrprevTableName = pstrTableName
			
			Case Else
				'do nothing
		End Select
		
		objSchema.MoveNext
	Loop
	
	Call ReleaseObject(objSchema)
	
	getAvailableTables_Array = paryPrimary

End Function	'getAvailableTables_Array

'****************************************************************************************************************************************************************

Function getValueFrom(byRef arySource, byRef objrsProducts, byVal lngIndex, byVal blnSpecial, byVal vntDefault)

Dim pvntTemp

	On Error Resume Next
	If Err.number <> 0 Then Err.Clear
	
	If Len(arySource(lngIndex)(enSourceFieldName)) > 0 Then
		pvntTemp = Trim(objrsProducts.Fields(arySource(lngIndex)(enSourceFieldName)).Value & "")
		If Err.number <> 0 Then
			WriteOutput "<font color=red><strong>The column <em>" & arySource(lngIndex)(enSourceFieldName) & "<em> does not appear to be in your data source. Please double check your saved profile<br /></strong></font><br />" & vbcrlf
			Err.Clear
		End If
	Else
		pvntTemp = Trim(arySource(lngIndex)(enDefaultValue) & "")
	End If
	
	If Len(pvntTemp) = 0 Then pvntTemp = vntDefault
	
	If blnSpecial Then pvntTemp = customBoolean(pvntTemp)
	
	getValueFrom = pvntTemp

End Function	'getValueFrom

'****************************************************************************************************************************************************************

Function hasCorruptData(byVal strData)

Dim pblnResult
Dim pbytLastChar

	If isNull(strData) Then
		pblnResult = False
	ElseIf InStr(1, strData, Chr(0)) > 0 Then
		pblnResult = True
	Else
		pblnResult = False
	End If
	
	hasCorruptData = pblnResult

End Function	'hasCorruptData
'****************************************************************************************************************************************************************

Function hasMTP()

	On Error Resume Next
	If Err.number <> 0 Then Err.Clear
	hasMTP = CBool(UBound(maryMTPs) >= 0)
	If Err.number <> 0 Then Err.Clear

End Function	'hasMTP

'****************************************************************************************************************************************************************

Sub IdentifyKeyProductFields()

Dim i

	For i = 0 To UBound(maryFields)
		Select Case maryFields(i)(enTargetFieldName)
			Case "prodCode": clngTarget_ProductID = i
			Case "prodName": clngTarget_ProductName = i
			Case "prodPrice": clngTarget_Price = i
			Case "Cost": clngTarget_Cost = i
			Case "prodSaleIsActive": clngTarget_SaleIsActive = i
			Case "prodSalePrice": clngTarget_SalePrice = i
		End Select
	Next 'i

End Sub	'IdentifyKeyProductFields

'****************************************************************************************************************************************************************

Function importTypeFromText(byVal strImportType)

	Select Case strImportType
		Case "enImportAll": importTypeFromText = enImportAll
		Case "enImportAll_DeleteExisting": importTypeFromText = enImportAll_DeleteExisting
		Case "enImportNewOnly": importTypeFromText = enImportNewOnly
		Case "enImportInvPriceOnly": importTypeFromText = enImportInvPriceOnly
		Case "enImportInformationOnly": importTypeFromText = enImportInformationOnly
		Case "enImportAllDeleteCategories": importTypeFromText = enImportAllDeleteCategories
		Case "enImportUpdateSelectedFieldsOnly": importTypeFromText = enImportUpdateSelectedFieldsOnly
	End Select

End Function	'importTypeFromText

'**************************************************************************************************************************************************

Function InitializeCustomFields

	InitializeCustomFields = isArray(maryCustomFields)
	Exit Function

	'This section is now handled in the XML section - kept here for reference only
	If Not isArray(maryCustomFields) Then
		ReDim maryCustomFields(1)
		maryCustomFields(0) = Array("ProductImprintAreas", "ImprintArea", "Imprint Area", "", enDatatype_string, enDisplayType_textbox)
		maryCustomFields(1) = Array("ProductColors", "ProductColors", "Product Colors", "", enDatatype_string, enDisplayType_textbox)
	End If
	
	InitializeCustomFields = isArray(maryCustomFields)
	
End Function	'InitializeCustomFields

'****************************************************************************************************************************************************************

Sub InitializeDefaults(byRef strProfileID)
'Purpose: Create the data arrays. Note this only creates/sizes them from store profiles. It does not load any values from the request object

Dim i
Dim pobjXMLProfile

	Call loadNamedProfiles

	If Len(strProfileID) = 0 Then strProfileID = cstrDefaultProfile

	If LoadXMLProfiles(strProfileID, pobjXMLProfile) Then
		Call loadSettingsFromProfile(pobjXMLProfile)
	ElseIf LoadXMLProfiles(maryProfiles(0)(0), pobjXMLProfile) Then
		Response.Write "<h4><font color='red'>Error loading profile <i>" & strProfileID & "</i>. First saved profile used.</font></h4>"
		Call loadSettingsFromProfile(pobjXMLProfile)
	Else
		Response.Write "<h4><font color='red'>Error loading profile <i>" & strProfileID & "</i>. Back-up settings used.</font></h4>"
		%><!--#include file="ProductImportTool_Support/ssImportProducts_FieldMappings_XML.asp"--><%
	End If	'LoadXMLProfiles("")

	Set pobjXMLProfile = Nothing
	Call IdentifyKeyProductFields

	If Len(mstrDSN_Target) = 0 Then mstrDSN_Target = connectionString

End Sub	'InitializeDefaults

'****************************************************************************************************************************************************************

Sub LoadFormValues(byVal blnComplete)

Dim i

	mstrDSN_Source = Trim(LoadRequestValue("DSN_Source"))
	mstrSourceTable = Trim(LoadRequestValue("SourceTable"))
	mstrDSN_Target = Trim(LoadRequestValue("DSN_Target"))
	
	'Remove any extra carriage returns
	mstrDSN_Source = Replace(mstrDSN_Source, vbcrlf, "")
	mstrDSN_Target = Replace(mstrDSN_Target, vbcrlf, "")
	
	mbytCreateCat = LoadRequestValue("createCat")
	If Len(mbytCreateCat) = 0 Then mbytCreateCat = cenCreateCat_Default
	
	mbytCreateMfg = LoadRequestValue("createMfg")
	If Len(mbytCreateMfg) = 0 Then mbytCreateMfg = cenCreateMfg_Default
	
	mbytCreateVend = LoadRequestValue("createVend")
	If Len(mbytCreateVend) = 0 Then mbytCreateVend = cenCreateVend_Default
	
	If Not blnComplete Then Exit Sub

	For i = 0 To UBound(maryFields)
		maryFields(i)(enSourceFieldName) = LoadRequestValue("fieldMap" & i)
		maryFields(i)(enDefaultValue) = LoadRequestValue("fieldDefault" & i)
		
		If False Then
			Response.Write "<fieldset><legend>" & maryFields(i)(enTargetFieldName) & "</legend>"
			Response.Write "columnName: " &  maryFields(i)(enSourceFieldName) & "<br />"
			Response.Write "default: " &  maryFields(i)(enDefaultValue) & "<br />"
			Response.Write "</fieldset>"
		End If
		
	Next 'i

	For i = 0 To UBound(maryInventoryFields)
		maryInventoryFields(i)(enSourceFieldName) = LoadRequestValue("fieldMapInventory" & i)
		maryInventoryFields(i)(enDefaultValue) = LoadRequestValue("fieldDefaultInventory" & i)
	Next 'i

	For i = 0 To UBound(maryGiftWrap)
		maryGiftWrap(i)(enSourceFieldName) = LoadRequestValue("fieldMapGiftWrap" & i)
		maryGiftWrap(i)(enDefaultValue) = LoadRequestValue("fieldDefaultGiftWrap" & i)
	Next 'i
	
	For i = 0 To UBound(maryAttributes)
		maryAttributes(i)(enSourceFieldName) = LoadRequestValue("fieldMapAttributes" & i)
		maryAttributes(i)(enDefaultValue) = LoadRequestValue("fieldAttributeImportStyle" & i)
	Next 'i

	mlngImportType = LoadRequestValue("ImportType")
	If isNumeric(mlngImportType) Then
		mlngImportType = CLng(mlngImportType)
	Else
		mlngImportType = mlngDefaultImportType
	End If
	mlngDefaultImportType = mlngImportType
	
	mstrCategoryColumn = LoadRequestValue("CategoryColumn")
	mlngDefaultCategoryID = LoadRequestValue("DefaultCategoryValue")
	
	mstrImportTypeColumn = LoadRequestValue("ImportTypeColumn")
	mblnDeleteExistingMTPs = CBool(LoadRequestValue("DeleteExistingMTPs") = "1")
	
	cstrSubcategoryDelimiter = LoadRequestValue("SubcategoryDelimiter")
	cstrMultipleCategoryDelimiter = LoadRequestValue("MultipleCategoryDelimiter")
	'Now for the Custom Fields
	If InitializeCustomFields Then
		For i = 0 To UBound(maryCustomFields)
			maryCustomFields(i)(enSourceFieldName) = LoadRequestValue("customField" & i)
			maryCustomFields(i)(enDefaultValue) = LoadRequestValue("customFieldDefault" & i)
		Next 'i
	End If
	
End Sub	'LoadFormValues

'****************************************************************************************************************************************************************

Sub loadNamedProfiles()

Dim i
Dim plngSavedProfileCount
Dim pobjNodeList
Dim Item

	If LoadXMLProfiles("", mobjXMLProfiles) Then
		
		Set pobjNodeList = mobjXMLProfiles.selectNodes("profiles/defaultProfileToUse")
		cstrDefaultProfile = pobjNodeList.item(0).text
		If Len(cstrDefaultProfile) = 0 Then cstrDefaultProfile = "default"

		Set pobjNodeList = mobjXMLProfiles.selectNodes("profiles/profile")
		plngSavedProfileCount = pobjNodeList.length - 1
		ReDim maryProfiles(plngSavedProfileCount)
		For i = 0 To plngSavedProfileCount
			Set Item = pobjNodeList.item(i)
			maryProfiles(i) = Array(Item.attributes.item(0).value, Item.attributes.item(1).value)
		Next
		
		Set Item = Nothing
		Set pobjNodeList = Nothing

	Else
		Response.Write "<h4><font color='red'>Error opening profiles.</h4>"
	End If	'LoadXMLProfiles

End Sub	'loadNamedProfiles

'****************************************************************************************************************************************************************

Function loadSettingsFromProfile(byRef objXMLProfile)

Dim oNodeList
Dim pobjXML_Profile
Dim plngProductFieldCount
Dim pblnSuccess

	pblnSuccess = True

	With objXMLProfile
		'Profile Specific
		mstrDSN_Source = .selectSingleNode("SourceDSN").Text
		mstrSourceTable = .selectSingleNode("SourceTable").Text

		'Product Table data
		Set oNodeList = .selectNodes("ProductFields/field")
		plngProductFieldCount = oNodeList.length - 1
		ReDim maryFields(plngProductFieldCount)
		For i = 0 To plngProductFieldCount	
			Call getProductDetail(oNodeList.item(i), maryFields(i))
		Next

		'Attributes
		Set oNodeList = .selectSingleNode("AttributeColumn")
		mstrCategoryColumn = GetXMLElementValue(oNodeList, "columnName")

		ReDim maryAttributes(0)
		maryAttributes(0) = Array("N/A", enSourceFieldName, enDisplayFieldName, enDefaultValue)
		maryAttributes(0)(enSourceFieldName) = GetXMLElementValue(oNodeList, "columnName")
		maryAttributes(0)(enDisplayFieldName) = GetXMLElementValue(oNodeList, "displayName")
		maryAttributes(0)(enDefaultValue) = GetXMLElementValue(oNodeList, "default")

		'Import Type
		Set oNodeList = .selectSingleNode("ImportTypeColumn")
		mstrImportTypeColumn = GetXMLElementValue(oNodeList, "columnName")

		'Category Assignments
		Set oNodeList = .selectSingleNode("CategoryColumn")
		mstrCategoryColumn = GetXMLElementValue(oNodeList, "columnName")
		cstrMultipleCategoryDelimiter = GetXMLElementValue(oNodeList, "MultipleCategoryDelimiter")
		cstrSubcategoryDelimiter = GetXMLElementValue(oNodeList, "SubcategoryDelimiter")

		'Inventory (AE only)
		Set oNodeList = .selectNodes("InventoryFields/field")
		plngProductFieldCount = oNodeList.length - 1
		ReDim maryInventoryFields(plngProductFieldCount)
		For i = 0 To plngProductFieldCount	
			Call getProductDetail(oNodeList.item(i), maryInventoryFields(i))
		Next

		'Gift Wrap (AE only)
		Set oNodeList = .selectNodes("GiftWrapFields/field")
		plngProductFieldCount = oNodeList.length - 1
		ReDim maryGiftWrap(plngProductFieldCount)
		For i = 0 To plngProductFieldCount	
			Call getProductDetail(oNodeList.item(i), maryGiftWrap(i))
		Next

		'Volume Pricing (AE only)
		Set oNodeList = .selectSingleNode("MTPColumn")
		cstrMTPImportPrefix = GetXMLElementValue(oNodeList, "MTPImportPrefix")
		cstrMTPImportSeparator = GetXMLElementValue(oNodeList, "MTPImportSeparator")
		mblnDeleteExistingMTPs = GetXMLElementValue(oNodeList, "DeleteExistingMTPs")

		'Import Options
		Set oNodeList = .selectSingleNode("ImportOptions")
		mbytCreateCat = GetXMLElementValue(oNodeList, "CreateCat")
		mbytCreateMfg = GetXMLElementValue(oNodeList, "CreateMfg")
		mbytCreateVend = GetXMLElementValue(oNodeList, "CreateVend")
		mlngDefaultImportType = GetXMLElementValue(oNodeList, "DefaultImportType")
		mlngDefaultCategoryID = GetXMLElementValue(oNodeList, "DefaultCategoryID")
		mlngDefaultManufacturerID = GetXMLElementValue(oNodeList, "DefaultManufacturerID")
		mlngDefaultVendorID = GetXMLElementValue(oNodeList, "DefaultVendorID")
		
		'Custom options
		Set oNodeList = .selectNodes("CustomSections/field")
		plngProductFieldCount = oNodeList.length - 1
		If plngProductFieldCount >= 0 Then
			ReDim maryCustomFields(plngProductFieldCount)
			For i = 0 To plngProductFieldCount	
				Call getProductDetail(oNodeList.item(i), maryCustomFields(i))
			Next
		End If

	End With	'objXMLProfile

	Set oNodeList = Nothing
	
	loadSettingsFromProfile = pblnSuccess
		
End Function	'loadSettingsFromProfile

'****************************************************************************************************************************************************************

Function LoadXMLProfiles(byVal strProfileID, byRef XMLDoc)

Dim pblnFound
Dim pobjNodeList

	'Load profiles from file, only needed if not already loaded
	If isObject(mobjXMLProfiles) Then
		pblnFound = True
	Else
		pblnFound = getXMLDoc(mobjXMLProfiles, profilePath)
	End If
	
	'Now look for a specific profile, if identified
	If pblnFound Then
		'Response.Write "Loading profile " & strProfileID & " . . .<br />"
		If Len(strProfileID) > 0 Then
			Set pobjNodeList = mobjXMLProfiles.selectNodes("profiles/profile")
			For i = 0 To pobjNodeList.length - 1
				Set XMLDoc = pobjNodeList.item(i)
				pblnFound = CBool(CStr(XMLDoc.attributes.item(0).value) = CStr(strProfileID))
				If pblnFound Then
					'Response.Write "<fieldset><legend>Located profile " & strProfileID & "</legend><pre>" & Server.HTMLEncode(XMLDoc.xml) & "</pre></fieldset>"
					Exit For
				Else
					Response.Write GetXMLElementValue(XMLDoc, "profileName") & "<br />"
					'Response.Write Item.xml & "<br />"
				End If
			Next	'i
			Set pobjNodeList = Nothing
		Else
			'send out the root document
			Set XMLDoc = mobjXMLProfiles
			pblnFound = True
		End If	'Len(strProfileID) > 0
	End If	'pblnFound
	
	LoadXMLProfiles = pblnFound
		
End Function	'LoadXMLProfiles

'****************************************************************************************************************************************************************

Function newXMLProfile(byVal strProfileID, byVal strProfileName)

Dim pobjXMLProfile
Dim pobjXMLProfile_New
Dim pblnSuccess

	pblnSuccess = False
	If LoadXMLProfiles(strProfileID, pobjXMLProfile) Then
		'take node and add to root
		Set pobjXMLProfile_New = pobjXMLProfile.cloneNode(True)
		
		pobjXMLProfile_New.attributes.item(0).value = strProfileName
		pobjXMLProfile_New.attributes.item(1).value = strProfileName
		mobjXMLProfiles.documentElement.appendChild pobjXMLProfile_New
		If WriteXMLProfile Then
			pblnSuccess = saveXMLProfile(strProfileName, strProfileName)
		End If
	End If

	newXMLProfile = pblnSuccess

End Function	'newXMLProfile

'****************************************************************************************************************************************************************

Function OpenTableSQL(ByVal strTableSource)

Dim pstrSQL

	If LCase(Left(strTableSource, 6)) = "select" Then
		pstrSQL = strTableSource
	Else
		pstrSQL = "Select * from [" & strTableSource & "]"
	End If
	
	OpenTableSQL = pstrSQL

End Function	'OpenTableSQL

'****************************************************************************************************************************************************************

Function profilePath()
	profilePath = ssAdminPath & "ProductImportTool_Support\ssImportProducts_Profiles.xml"
End Function	'profilePath

'**************************************************************************************************************************************************

Sub RecordTime(byVal strMessage, byVal dtStartTime, byRef dtCurrentTime, byRef dtLastTime)

Dim pCurrentTime:	pCurrentTime = Time()
Dim pElapsedTime:	pElapsedTime = DateDiff("s", dtStartTime, pCurrentTime)
Dim pIncrementTime:	pIncrementTime = DateDiff("s", dtLastTime, pCurrentTime)
Dim pstrOut

	If mblnssDebugShowTime Then
		If pElapsedTime = 0 Then
			pstrOut = strMessage & ": " & dtStartTime
		Else
			pstrOut = strMessage & ": Elaspsed time " & FormatNumber(pElapsedTime, 4) & " seconds."
			
			If pIncrementTime <> 0 Then
				pstrOut = pstrOut & " - (Increment " & FormatNumber(pIncrementTime, 4) & " seconds)"
			End If
		End If
		
		pstrOut = pstrOut & "<br />"
		WriteOutput pstrOut
	End If
	
	'Update time
	dtLastTime = dtCurrentTime
	dtCurrentTime = pCurrentTime

End Sub	'RecordTime

'****************************************************************************************************************************************************************

Sub saveAttributes(byRef objrsProducts, byRef aryProduct)

Dim i, j
Dim plngAttributeCounter
Dim plngCategoryCounter
Dim plngDetailCounter

Dim paryAttributes
Dim paryAttribute_Temp
Dim paryAttribute_Category
Dim paryAttribute_Detail
Dim pstrName
Dim pstrPrice
Dim paryTemp

Dim plngAttributeID
Dim plngAttributeDetailID

Dim pblnAttributesImported
Dim pstrAttributeTemplate
Dim pbytAttributeImportStyle
Dim pstrAttributeSourceValue
Dim plngProductUID
Dim paryAttDetailUIDs
Dim paryAttDetailNames

Dim cstrAttributeCategorySeparator
Dim cstrAttributeDetailSeparator
Dim cstrAttributeItemSeparator

	plngProductUID = aryProduct(0)

	pblnAttributesImported = False
	If Len(plngProductUID) = 0 Then Exit Sub

	'maryAttributeSupportingStyles(0) = Array("No Attribute", "")
	'maryAttributeSupportingStyles(1) = Array("ShopSite Style", "")
	'maryAttributeSupportingStyles(2) = Array("Attribute Template", "")
	'maryAttributeSupportingStyles(3) = Array("Auto Select", "")

	For i = 0 To UBound(maryAttributes)
		pbytAttributeImportStyle = CStr(maryAttributes(i)(enDefaultValue))
		If Len(pbytAttributeImportStyle) = 0 Then pbytAttributeImportStyle = "3"
		
		On Error Resume Next
		pstrAttributeSourceValue = Trim(objrsProducts.Fields(maryAttributes(i)(enSourceFieldName)).Value & "")
		If Err.number <> 0 Then
			Err.Clear
			Exit Sub
		End If
		On Error Goto 0
		
		If Len(pstrAttributeSourceValue) = 0 Then
			pbytAttributeImportStyle = "0"
		ElseIf pbytAttributeImportStyle = "3" Then
			If InStr(1, pstrAttributeSourceValue, "|n|") Then
				WriteOutput "&nbsp;&nbsp;&nbsp;Automatic attribute selection chosen - ShopSite style being used<br />" & vbcrlf
				pbytAttributeImportStyle = "1"
			ElseIf InStr(1, pstrAttributeSourceValue, ":") Then
				WriteOutput "&nbsp;&nbsp;&nbsp;Automatic attribute selection chosen - Pricing style being used<br />" & vbcrlf
				pbytAttributeImportStyle = "4"
			Else
				WriteOutput "&nbsp;&nbsp;&nbsp;Automatic attribute selection chosen - StoreFront template being used<br />" & vbcrlf
				pbytAttributeImportStyle = "2"
			End If
		End If
		
		Select Case pbytAttributeImportStyle
			Case "0": 'No Import
				aryProduct(1) = ""
				aryProduct(2) = ""
			Case "1": 'ShopSite
				'Attribute Category||Type||Required|n|Attribute Name||Price||PriceType||Weight||WeightType||AttrOrder||SmallImage||LargeImage||FileLocation|n||n|Select a Flex:|n|Light|n|Regular|n|Firm|n||n|Select a Hand:|n|Right|n|Left
				'On Error Resume Next
				
				cstrAttributeCategorySeparator = "|n||n|"
				cstrAttributeDetailSeparator = "|n|"
				cstrAttributeItemSeparator = ""

				paryAttributes = Split(pstrAttributeSourceValue, cstrAttributeCategorySeparator)
				If Err.number <> 0 Then
					If Instr(1, Err.Description, "Item cannot be found in the collection corresponding") > 0 Then
						WriteOutput "&nbsp;&nbsp;&nbsp;<font color=red>Error Importing Attributes - No Column identified for attribute import</font><br />" & vbcrlf
					Else
						WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
						WriteOutput "<font color=red>Source DSN: " & mstrDSN_Source & "</font><br />" & vbcrlf
					End If
					Err.Clear
					paryAttributes = Array("")
					pblnAttributesImported = False
				End If
				On Error Goto 0

				If isArray(paryAttributes) Then WriteOutput "&nbsp;&nbsp;&nbsp;Importing attributes . . .<br />" & vbcrlf
				ReDim paryAttDetailUIDs(UBound(paryAttributes))
				ReDim paryAttDetailNames(UBound(paryAttributes))
				
				For plngAttributeCounter = 0 To UBound(paryAttributes)
					paryAttribute_Temp = Split(paryAttributes(plngAttributeCounter), cstrAttributeDetailSeparator)
					If UBound(paryAttribute_Temp) > 0 Then
						paryAttribute_Category = Split(paryAttribute_Temp(0),"||")
						ReDim paryAttribute_Detail(UBound(paryAttribute_Temp)-1)
						For plngCategoryCounter = 1 To UBound(paryAttribute_Temp)
							paryAttribute_Detail(plngCategoryCounter - 1) = Split(paryAttribute_Temp(plngCategoryCounter),"||")
						Next 'plngCategoryCounter

						'For plngCategoryCounter = 0 To UBound(paryAttribute_Category)
							plngAttributeID = setAttribute(plngProductUID, paryAttribute_Category)
							WriteOutput "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<em>" & paryAttribute_Category(0) & "</em><br />" & vbcrlf
							'WriteOutput "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<em>" & paryAttribute_Category(0) & "</em> (" & plngAttributeID & ")<br />" & vbcrlf
							For plngDetailCounter = 0 To UBound(paryAttribute_Detail)
								If isArray(paryAttribute_Detail(plngDetailCounter)) Then
									If UBound(paryAttribute_Detail(plngDetailCounter)) >= 0 Then
										plngAttributeDetailID = setAttributeDetail(plngAttributeID, paryAttribute_Detail(plngDetailCounter))
										paryAttDetailUIDs(plngAttributeCounter) = plngAttributeDetailID
										paryAttDetailNames(plngAttributeCounter) = paryAttribute_Detail(plngDetailCounter)(0)
										WriteOutput "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- <em>" & paryAttribute_Detail(plngDetailCounter)(0) & "</em><br />" & vbcrlf
										'WriteOutput "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- <em>" & paryAttribute_Detail(plngDetailCounter)(0) & "</em> (" & plngAttributeDetailID & ")<br />" & vbcrlf
									End If
								End If
							Next 'plngDetailCounter
						'Next 'plngCategoryCounter
					End If
					pblnAttributesImported = True
				Next 'plngAttributeCounter

				aryProduct(1) = paryAttDetailUIDs
				aryProduct(2) = paryAttDetailNames

			Case "2": 'StoreFront Attribute Template
				On Error Resume Next

				cstrAttributeCategorySeparator = "|"
				cstrAttributeDetailSeparator = ""
				cstrAttributeItemSeparator = ""

				paryAttributes = Split(pstrAttributeSourceValue, cstrAttributeCategorySeparator)
				If Err.number <> 0 Then
					If Instr(1, Err.Description, "Item cannot be found in the collection corresponding") > 0 Then
						WriteOutput "&nbsp;&nbsp;&nbsp;<font color=red>Error Importing Attributes - No Column identified for attribute import</font><br />" & vbcrlf
					Else
						WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
						WriteOutput "<font color=red>Source DSN: " & mstrDSN_Source & "</font><br />" & vbcrlf
					End If
					Err.Clear
					pblnAttributesImported = False
				Else
					For j = 0 To UBound(paryAttributes)
						pstrAttributeTemplate = Trim(paryAttributes(j))
						
						Dim XMLDoc
						Dim e
						Dim pstrPath
						Dim plngAttributeCategoryCounter
						Dim plngAttributeDetailCounter
						Dim plngNodeCounter
						Dim xmlAttributeCategories
						Dim xmlAttributeDetails
						Dim attrUID
						
						pstrPath = managementPath & "SFTEMPLATES\" & pstrAttributeTemplate & ".att"
						Set XMLDoc = CreateObject("MSXML.DOMDocument")
						If getXMLDoc(XMLDoc, pstrPath) Then
							Set xmlAttributeCategories = XMLDoc.getElementsByTagName("Attributes")
							Set xmlAttributeDetails = XMLDoc.getElementsByTagName("AttributeDetail")
							If xmlAttributeDetails.length = 0 Then Set xmlAttributeDetails = XMLDoc.getElementsByTagName("AttributeDetails")

							ReDim paryAttribute_Category(xmlAttributeCategories.length)	
							ReDim paryAttribute_Detail(2)
							For plngAttributeCategoryCounter = 0 To xmlAttributeCategories.length - 1
								For plngNodeCounter = 0 To xmlAttributeCategories.item(plngAttributeCategoryCounter).childNodes.length - 1
									Set e = xmlAttributeCategories.item(plngAttributeCategoryCounter).childNodes.Item(plngNodeCounter)
									Select Case e.nodeName
										Case "Name":			paryAttribute_Detail(0) = e.text
										Case "Type":			paryAttribute_Detail(1) = e.text
										Case "Required":		paryAttribute_Detail(2) = e.text
										Case "uid":				
																attrUID = CLng(e.text)
																If attrUID > UBound(paryAttribute_Category) Then ReDim Preserve paryAttribute_Category(attrUID)
									End Select
								Next 'plngNodeCounter
								plngAttributeID = setAttribute(plngProductUID, paryAttribute_Detail)
								If plngAttributeID = -1 Then
									WriteOutput "&nbsp;&nbsp;&nbsp;Error inserting attribute category <i>" & paryAttribute_Detail(0) & "</i><br />" & vbcrlf
								Else
									WriteOutput "&nbsp;&nbsp;&nbsp;Attribute category <i>" & paryAttribute_Detail(0) & "</i> updated.<br />" & vbcrlf
								End If
								paryAttribute_Category(attrUID) = plngAttributeID
							Next 'plngAttributeCategoryCounter
							
							'Load attribute details
							ReDim paryAttribute_Detail(11)
							For plngAttributeDetailCounter = 0 To xmlAttributeDetails.length - 1
								For plngNodeCounter = 0 To xmlAttributeDetails.item(plngAttributeDetailCounter).childNodes.length - 1
									Set e = xmlAttributeDetails.item(plngAttributeDetailCounter).childNodes.Item(plngNodeCounter)
									Select Case e.nodeName
										Case "AttributeID":		plngAttributeID = paryAttribute_Category(CLng(e.text))
										Case "Name":			paryAttribute_Detail(0) = e.text
										Case "Price":			paryAttribute_Detail(1) = e.text
										Case "Weight":			paryAttribute_Detail(2) = e.text
										Case "WeightType":		paryAttribute_Detail(4) = e.text
										Case "PriceType":		paryAttribute_Detail(3) = e.text
										Case "AttributeOrder":	paryAttribute_Detail(5) = e.text
										Case "FileLocation":	paryAttribute_Detail(8) = e.text
										Case "Lines":			paryAttribute_Detail(10) = e.text
										Case "Characters":		paryAttribute_Detail(11) = e.text
									End Select
								Next 'plngNodeCounter
								
								'add check to set Price to 0 if PriceType = 0
								If CStr(paryAttribute_Detail(3)) = "0" Then paryAttribute_Detail(1) = 0
								plngAttributeDetailID = setAttributeDetail(plngAttributeID, paryAttribute_Detail)
								If plngAttributeDetailID = -1 Then
									WriteOutput "&nbsp;&nbsp;&nbsp;Error inserting attribute <i>" & paryAttribute_Detail(0) & "</i><br />" & vbcrlf
								Else
									WriteOutput "&nbsp;&nbsp;&nbsp;Attribute <i>" & paryAttribute_Detail(0) & "</i> updated.<br />" & vbcrlf
								End If
							Next 'plngAttributeCategoryCounter
							Set e = Nothing
							Set xmlAttributeDetails = Nothing
							Set xmlAttributeCategories = Nothing
						Else
							Dim myErr
							Set myErr = objXSL.parseError
							WriteOutput "<h4><font color=red>Error loading <i>" & pstrAttributeTemplate & "</i> template at " & pstrPath & ". Error " & myErr.errorCode & ": " & myErr.reason & "</font></h4>"
						End If
						Set XMLDoc = Nothing
						
					Next 'j
				End If
			Case "4": 'Size: 1(-$2), 1 1/2(-$2),  2(-$2), 2 1/2(-$2), 3(-$2), 3 1/2(-$2), 4(-$2), 4 1/2(-$2), 5, 5 1/2, 6, 6 1/2, 7(+2), 7 1/2(+2), 8(+2), 8 1/2(+2), 9(+4), 9 1/2(+4), 10(+4), 10 1/2(+4), 11(+6), 11 1/2(+6), 12(+6), 12 1/2(+6), 13(+6)
				'Attribute Category||Type||Required|n|Attribute Name||Price||PriceType||Weight||WeightType||AttrOrder||SmallImage||LargeImage||FileLocation|n||n|Select a Flex:|n|Light|n|Regular|n|Firm|n||n|Select a Hand:|n|Right|n|Left
				'On Error Resume Next

				cstrAttributeCategorySeparator = "|"
				cstrAttributeDetailSeparator = ":"
				cstrAttributeItemSeparator = ","
				
				paryAttributes = Split(pstrAttributeSourceValue, cstrAttributeCategorySeparator)
				If Err.number <> 0 Then
					If Instr(1, Err.Description, "Item cannot be found in the collection corresponding") > 0 Then
						WriteOutput "&nbsp;&nbsp;&nbsp;<font color=red>Error Importing Attributes - No Column identified for attribute import</font><br />" & vbcrlf
					Else
						WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
						WriteOutput "<font color=red>Source DSN: " & mstrDSN_Source & "</font><br />" & vbcrlf
					End If
					Err.Clear
					paryAttributes = Array("")
					pblnAttributesImported = False
				End If
				On Error Goto 0

				If isArray(paryAttributes) Then WriteOutput "&nbsp;&nbsp;&nbsp;Importing attributes . . .<br />" & vbcrlf
				ReDim paryAttDetailUIDs(UBound(paryAttributes))
				ReDim paryAttDetailNames(UBound(paryAttributes))
				
				For plngAttributeCounter = 0 To UBound(paryAttributes)
					paryAttribute_Temp = Split(paryAttributes(plngAttributeCounter), cstrAttributeDetailSeparator)
					If UBound(paryAttribute_Temp) > 0 Then
						paryAttribute_Category = paryAttribute_Temp(0)
						paryAttribute_Detail = Split(paryAttribute_Temp(1), cstrAttributeItemSeparator)
						For plngCategoryCounter = 0 To UBound(paryAttribute_Detail)
							paryTemp = Split(paryAttribute_Detail(plngCategoryCounter), "(")
							pstrName = Trim(getArrayValue(paryTemp, 0, ""))
							pstrPrice = getArrayValue(paryTemp, 1, "0")
							pstrPrice = Replace(pstrPrice, "(", "")
							pstrPrice = Replace(pstrPrice, ")", "")
							pstrPrice = Replace(pstrPrice, "$", "")

							paryAttribute_Detail(plngCategoryCounter) = Array(pstrName, pstrPrice, 0, 0, 0, plngCategoryCounter)
	
						Next 'plngCategoryCounter

						plngAttributeID = setAttribute(plngProductUID, Array(paryAttribute_Category))
						WriteOutput "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<em>" & paryAttribute_Category & "</em><br />" & vbcrlf
						'WriteOutput "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<em>" & paryAttribute_Category(0) & "</em> (" & plngAttributeID & ")<br />" & vbcrlf
						For plngDetailCounter = 0 To UBound(paryAttribute_Detail)
							If isArray(paryAttribute_Detail(plngDetailCounter)) Then
								If UBound(paryAttribute_Detail(plngDetailCounter)) >= 0 Then
									plngAttributeDetailID = setAttributeDetail(plngAttributeID, paryAttribute_Detail(plngDetailCounter))
									paryAttDetailUIDs(plngAttributeCounter) = plngAttributeDetailID
									paryAttDetailNames(plngAttributeCounter) = paryAttribute_Detail(plngDetailCounter)(0)
									WriteOutput "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- <em>" & paryAttribute_Detail(plngDetailCounter)(0) & "</em><br />" & vbcrlf
									'WriteOutput "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- <em>" & paryAttribute_Detail(plngDetailCounter)(0) & "</em> (" & plngAttributeDetailID & ")<br />" & vbcrlf
								End If
							End If
						Next 'plngDetailCounter
					End If
					pblnAttributesImported = True
				Next 'plngAttributeCounter

				aryProduct(1) = paryAttDetailUIDs
				aryProduct(2) = paryAttDetailNames

			Case Else 'Ignore
		End Select
	Next 'i
	
	If pblnAttributesImported Then WriteOutput "&nbsp;&nbsp;&nbsp;Updated Attributes<br />" & vbcrlf
		
End Sub	'saveAttributes

'****************************************************************************************************************************************************************

Sub saveCategories(byRef objrsProducts, byVal strProductID, byRef objrsCategories, byRef dicCategories)

Dim pblnAdded
Dim pstrSQL
Dim pstrTemp
Dim pvntCategory
Dim paryCategory
Dim plngCategoryID
Dim i

	If Len(strProductID) = 0 Then Exit Sub

	If mlngImportType = enImportAllDeleteCategories Then
		If deleteProductCategoryAssignments(strProductID) Then
			WriteOutput "&nbsp;&nbsp;&nbsp;Removed existing category assignment(s)<br />" & vbcrlf
		Else
			WriteOutput "&nbsp;&nbsp;&nbsp;Unable to remove category assignment(s)<br />" & vbcrlf
		End If
	End If

	If Len(mstrCategoryColumn) > 0 Then pstrTemp = Trim(objrsProducts.Fields(mstrCategoryColumn).Value & "")
	
	If False Then
		Response.Write "<fieldset><legend>saveCategories - " & strProductID & ": (" & Asc(Right(pstrTemp, 1)) & ")</legend>" & vbcrlf
		Response.Write "pstrTemp: {" & pstrTemp & "}<br />" & vbcrlf
		Response.Write "hasCorruptData: " & hasCorruptData(pstrTemp) & "<br />" & vbcrlf
		Response.Write "</fieldset>" & vbcrlf
	End If

	'No category column so need to use default value
	If Len(pstrTemp) = 0 Then
		pstrTemp = mlngDefaultCategoryID
	Else
		Call cleanCategoryData(pstrTemp)
	End If

	If Len(pstrTemp) > 0 Then
		Call verifyCategory(pstrTemp, objrsCategories, dicCategories, pblnAdded)
		paryCategory = Split(pstrTemp, cstrMultipleCategoryDelimiter)
		For i = 0 To UBound(paryCategory)
			pvntCategory = Trim(paryCategory(i))
			plngCategoryID = -1
			If Len(pvntCategory) > 0 Then
				If instr(1, pvntCategory, cstrSubcategoryDelimiter) > 0 Then
					plngCategoryID = setSubCategories(strProductID, pvntCategory, objrsCategories, dicCategories)
				Else
					plngCategoryID = getCategoryUIDbyName(pvntCategory, objrsCategories, dicCategories)
				End If
				
				If plngCategoryID = mlngDefaultCategoryID Then
					If createProductCategoryAssignment(strProductID, plngCategoryID) Then
						WriteOutput "&nbsp;&nbsp;&nbsp;Created Product Category Assignment using default category ID of <em>" & plngCategoryID & "</em><br />" & vbcrlf
					End If
				ElseIf plngCategoryID <> -1 Then
					If createProductCategoryAssignment(strProductID, plngCategoryID) Then
						WriteOutput "&nbsp;&nbsp;&nbsp;Created Product Category Assignment for <em>" & pvntCategory & "</em><br />" & vbcrlf
					End If
				Else
					WriteOutput "&nbsp;&nbsp;<font color=red>Error creating category <i>" & pvntCategory & "</i>(" & plngCategoryID & ")</font><br />" & vbcrlf
				End If
			End If
		Next 'i
	End If
	
End Sub	'saveCategories

'**************************************************************************************************************************************************

Sub saveCustom(byRef objrsProducts, byVal lngProductUID)

Dim i

	If Len(lngProductUID) = 0 Then Exit Sub

	If InitializeCustomFields Then
		For i = 0 To UBound(maryCustomFields)
			If Len(maryCustomFields(i)(enSourceFieldName)) > 0 Or mlngImportType <> enImportUpdateSelectedFieldsOnly Then
				Call Execute("Call saveCustom_" & maryCustomFields(i)(enTargetFieldName) & "(objrsProducts, lngProductUID, i)")
			End If
		Next 'i
	End If
		
End Sub	'saveCustom

'**************************************************************************************************************************************************

Sub SaveFailedImport(strProductID, strProductName)

Dim pstrToAdd
Dim pstrSession

	If Len(strProductID) > 0 Then
		pstrToAdd = strProductID & "|" & strProductName
		
		pstrSession = Session("ssFailedImports")
		If Len(pstrSession) = 0 Then
			pstrSession = pstrToAdd
		Else
			pstrSession = "{new}" & pstrToAdd
		End If

		Session("ssFailedImports") = pstrSession
	End If
		
End Sub	'SaveFailedImport

'****************************************************************************************************************************************************************

Sub saveGiftWrap(byRef objrsProducts, byVal strProductID)

Dim pbytgwActivate
Dim pstrgwPrice

	If Len(strProductID) = 0 Then Exit Sub
	
	If CBool(mlngImportType = enImportUpdateSelectedFieldsOnly) AND CBool((Len(maryGiftWrap(0)(enSourceFieldName)) = 0) OR (Len(maryGiftWrap(1)(enSourceFieldName)) = 0))  Then Exit Sub

	pstrgwPrice = getValueFrom(maryGiftWrap, objrsProducts, 1, False, "0")
	pbytgwActivate = getValueFrom(maryGiftWrap, objrsProducts, 0, True, "0")
	
	If False Then
		Response.Write "<fieldset><legend></legend>"
		Response.Write "pbytgwActivate: " & pbytgwActivate & "(" & maryGiftWrap(0)(enSourceFieldName) & ": " & Trim(objrsProducts.Fields(maryGiftWrap(0)(enSourceFieldName)).Value & "") & ")<br />"
		Response.Write "</fieldset>"
	End If
		
	If setGiftWrap(strProductID, pbytgwActivate, pstrgwPrice) Then
		WriteOutput "&nbsp;&nbsp;&nbsp;Updated Gift Wrap<br />" & vbcrlf
	Else
		WriteOutput "<font color=red>&nbsp;&nbsp;&nbsp;Error updating Gift Wrap</font><br />" & vbcrlf
	End If
		
End Sub	'saveGiftWrap

'****************************************************************************************************************************************************************

Function saveXMLProfile(byVal strProfileID, byVal strProfileName)

Dim oNodeList
Dim pblnSuccess
Dim pobjItem
Dim pobjProfileNode
Dim pobjXML_Profile
Dim plngProductFieldCount

	If LoadXMLProfiles(strProfileID, pobjProfileNode) Then
		'Response.Write "<fieldset><legend>saveXMLProfile - pobjXML_Profile</legend><pre>" & Server.HTMLEncode(pobjProfileNode.xml) & "</pre></fieldset>"
	
		With pobjProfileNode
		
			'Profile Specific
			pobjProfileNode.attributes.item(0).value = strProfileName
			pobjProfileNode.attributes.item(1).value = strProfileName
			pobjProfileNode.selectSingleNode("SourceDSN").Text = mstrDSN_Source
			pobjProfileNode.selectSingleNode("SourceTable").Text = mstrSourceTable
			
			If Not mblnProfileOnly Then

				'Product Table data
				Set oNodeList = .selectNodes("ProductFields/field")
				plngProductFieldCount = oNodeList.length - 1
				For i = 0 To plngProductFieldCount	
					Call setProductDetail(oNodeList.item(i), maryFields(i))
				Next

				'Attributes
				pobjProfileNode.selectSingleNode("AttributeColumn/columnName").Text = maryAttributes(0)(enSourceFieldName)
				pobjProfileNode.selectSingleNode("AttributeColumn/displayName").Text = maryAttributes(0)(enDisplayFieldName)
				pobjProfileNode.selectSingleNode("AttributeColumn/default").Text = maryAttributes(0)(enDefaultValue)
				Set pobjItem = pobjProfileNode.selectSingleNode("AttributeColumn/default")

				'Category Assignments
				pobjProfileNode.selectSingleNode("CategoryColumn/columnName").Text = mstrCategoryColumn
				pobjProfileNode.selectSingleNode("CategoryColumn/MultipleCategoryDelimiter").Text = cstrMultipleCategoryDelimiter
				pobjProfileNode.selectSingleNode("CategoryColumn/SubcategoryDelimiter").Text = cstrSubcategoryDelimiter
				
				'Import Type Assignments
				Call setXMLValue_selectSingleNode(pobjProfileNode, "ImportTypeColumn/columnName", mstrImportTypeColumn)
				'Inventory (AE only)
				Set oNodeList = .selectNodes("InventoryFields/field")
				plngProductFieldCount = oNodeList.length - 1
				For i = 0 To plngProductFieldCount	
					Call setProductDetail(oNodeList.item(i), maryInventoryFields(i))
				Next

				'Gift Wrap (AE only)
				Set oNodeList = .selectNodes("GiftWrapFields/field")
				plngProductFieldCount = oNodeList.length - 1
				For i = 0 To plngProductFieldCount	
					Call setProductDetail(oNodeList.item(i), maryGiftWrap(i))
				Next

				'Volume Pricing (AE only)
				pobjProfileNode.selectSingleNode("MTPColumn/MTPImportPrefix").Text = cstrMTPImportPrefix
				pobjProfileNode.selectSingleNode("MTPColumn/MTPImportSeparator").Text = cstrMTPImportSeparator
				pobjProfileNode.selectSingleNode("MTPColumn/DeleteExistingMTPs").Text = mblnDeleteExistingMTPs
				
				'Import Options
				pobjProfileNode.selectSingleNode("ImportOptions/CreateCat").Text = mbytCreateCat
				pobjProfileNode.selectSingleNode("ImportOptions/CreateMfg").Text = mbytCreateMfg
				pobjProfileNode.selectSingleNode("ImportOptions/CreateVend").Text = mbytCreateVend
				pobjProfileNode.selectSingleNode("ImportOptions/DefaultImportType").Text = mlngDefaultImportType
				pobjProfileNode.selectSingleNode("ImportOptions/DefaultCategoryID").Text = mlngDefaultCategoryID
				pobjProfileNode.selectSingleNode("ImportOptions/DefaultManufacturerID").Text = mlngDefaultManufacturerID
				pobjProfileNode.selectSingleNode("ImportOptions/DefaultVendorID").Text = mlngDefaultVendorID
				
				'Custom options
				Set oNodeList = .selectNodes("CustomSections/field")
				plngProductFieldCount = oNodeList.length - 1
				If plngProductFieldCount >= 0 Then
					For i = 0 To plngProductFieldCount	
						Call setProductDetail(oNodeList.item(i), maryCustomFields(i))
					Next
				End If

			End If	'Not mblnProfileOnly
			
		End With	'mobjXMLProfiles

		Set oNodeList = Nothing
		Set pobjProfileNode = Nothing
		
		'Now update the default
		If Len(cstrDefaultProfile) > 0 Then mobjXMLProfiles.selectSingleNode("profiles/defaultProfileToUse").Text = cstrDefaultProfile

		pblnSuccess = WriteXMLProfile
	Else
		Response.Write "<h4><font color=red>Unable to locate profile <i>" & strProfileID & "</i> for saving.</h4>"
		pblnSuccess = False
	End If

	saveXMLProfile = pblnSuccess
		
End Function	'saveXMLProfile

'****************************************************************************************************************************************************************

Sub SetColumns(byRef objRS)

Dim i
Dim plngFieldCount
Dim plngMTPColumns
Dim pstrTemp
Dim paryTemp

	plngFieldCount = objRS.Fields.Count-1
	ReDim marySheetColumns(plngFieldCount)
	
	'Do the initial load
	For i = 0 To plngFieldCount
		marySheetColumns(i) = Trim(objRS.Fields(i).Name)
	Next 'i

	'Do the sort
	For i = 1 To plngFieldCount
		For j = 0 To plngFieldCount-1
			If LCase(marySheetColumns(j)) > LCase(marySheetColumns(i)) Then
				pstrTemp = marySheetColumns(j)
				marySheetColumns(j) = marySheetColumns(i)
				marySheetColumns(i) = pstrTemp
			End If
		Next 'j
	Next 'i

	'added to identify MTP columns denoted by "MTP-" (set using cstrMTPImportPrefix constant) prefix
	'This is for figuring out MTP
	'use maryMTPs(numColums)(information)
	'Column name is in format MTP-##-??? Where
	'	MTP - identifies this as a column requiring special attention
	'	## - represents the qty level
	'	??? - represents the type of discount
	'	QtyOrDiscount - represents whether ## is the Qty or the discount - values can be QTY, DISCOUNT, or empty (empty means QTY for backward compatibility)
	plngMTPColumns = -1
	For i = 0 To UBound(marySheetColumns)
		If Left(marySheetColumns(i), Len(cstrMTPImportPrefix)) = cstrMTPImportPrefix Then
			plngMTPColumns = plngMTPColumns + 1
			'Response.Write "plngMTPColumns: " & plngMTPColumns & "<br />"					
			ReDim Preserve maryMTPs(plngMTPColumns)
			pstrTemp = Replace(marySheetColumns(i), cstrMTPImportPrefix, "")
			paryTemp = Split(pstrTemp, cstrMTPImportSeparator)
			
			'Safety check
			If Not isArray(paryTemp) Then
				WriteOutput "<font color=red>Column header <i>" & marySheetColumns(i) & "<i> is not valid for volume pricing. Default values have been loaded. Please refer to the help file for the proper format.</font>" & vbcrlf
			ElseIf UBound(paryTemp) < 1 Then
				WriteOutput "<font color=red>Column header <i>" & marySheetColumns(i) & "<i> is not valid for volume pricing. Default values have been loaded. Please refer to the help file for the proper format.</font>" & vbcrlf
			End If
			
			maryMTPs(plngMTPColumns) = Array("mtIndex", "mtQuantity", "mtValue", "mtType", "columnName", "QtyOrDiscount")
			maryMTPs(plngMTPColumns)(0) = plngMTPColumns
			maryMTPs(plngMTPColumns)(1) = getArrayValue(paryTemp, 0, 9999)	'paryTemp(0)
			maryMTPs(plngMTPColumns)(2) = ""
			maryMTPs(plngMTPColumns)(3) = getArrayValue(paryTemp, 1, "AMOUNT")	'paryTemp(1)
			maryMTPs(plngMTPColumns)(4) = marySheetColumns(i)
			
			If UBound(paryTemp) > 1 Then
				maryMTPs(plngMTPColumns)(5) = getArrayValue(paryTemp, 2, "QTY")	'paryTemp(2)
			Else
				maryMTPs(plngMTPColumns)(5) = "QTY"
			End If
			
			If False Then
				Response.Write "MTP Column: " & maryMTPs(plngMTPColumns)(0) & "<br />"					
				Response.Write "MTP Column: " & maryMTPs(plngMTPColumns)(1) & "<br />"					
				Response.Write "MTP Column: " & maryMTPs(plngMTPColumns)(3) & "<br />"					
				Response.Write "MTP Column: " & maryMTPs(plngMTPColumns)(4) & "<br />"
				Response.Write "MTP Column: " & maryMTPs(plngMTPColumns)(5) & "<hr>"
			End If
		End If
	Next 'i

End Sub	'SetColumns

'****************************************************************************************************************************************************************

Sub SetMTPColumns(byRef objRS)

Dim i
Dim pstrColName

	ReDim maryMTPs(2)
	
	For i = 0 To UBound(maryMTPs)
		pstrColName = "Quantity" & i + 2
		maryMTPs(i) = Array("mtIndex", "mtQuantity", "mtValue", "mtType", "columnName", "QtyOrDiscount")
		maryMTPs(i)(0) = 0
		maryMTPs(i)(1) = "Use Price" & i + 2 & " column"
		maryMTPs(i)(2) = ""
		maryMTPs(i)(3) = "Calculate"
		maryMTPs(i)(4) = pstrColName
		maryMTPs(i)(5) = "N/A"
			
		If False Then
			Response.Write "MTP Column: " & maryMTPs(i)(0) & "<br />"					
			Response.Write "MTP Column: " & maryMTPs(i)(1) & "<br />"					
			Response.Write "MTP Column: " & maryMTPs(i)(3) & "<br />"					
			Response.Write "MTP Column: " & maryMTPs(i)(4) & "<br />"
			Response.Write "MTP Column: " & maryMTPs(i)(5) & "<br />"
			Response.Flush
		End If
	Next 'i

End Sub	'SetMTPColumns

'****************************************************************************************************************************************************************

Sub setProductDetail(byRef objXMLElement, byRef aryProductDetail)

	objXMLElement.attributes.item(0).value = aryProductDetail(enTargetFieldName)
	
	Call setXMLElementValue(objXMLElement, "columnName", aryProductDetail(enSourceFieldName))
	Call setXMLElementValue(objXMLElement, "displayName", aryProductDetail(enDisplayFieldName))
	Call setXMLElementValue(objXMLElement, "default", aryProductDetail(enDefaultValue))
	Call setXMLElementValue(objXMLElement, "dataType", dataTypeToText(aryProductDetail(enFieldDataType)))
	Call setXMLElementValue(objXMLElement, "displayType", displayTypeToText(aryProductDetail(enDisplayType)))
	
	If False Then
		Response.Write "<fieldset><legend>" & aryProductDetail(enTargetFieldName) & "</legend>"
		Response.Write "columnName: " & aryProductDetail(enSourceFieldName) & "<br />"
		Response.Write "displayName: " & aryProductDetail(enDisplayFieldName) & "<br />"
		Response.Write "default: " & aryProductDetail(enDefaultValue) & "<br />"
		Response.Write "dataType: " & dataTypeToText(aryProductDetail(enFieldDataType)) & "(" & aryProductDetail(enFieldDataType) & ")<br />"
		Response.Write "displayType: " & displayTypeToText(aryProductDetail(enDisplayType)) & "(" & aryProductDetail(enDisplayType) & ")<br />"
		Response.Write "</fieldset>"
	End If

End Sub	'setProductDetail

'****************************************************************************************************************************************************************

Function setInventoryLevels(byRef objrsProducts, byVal aryProduct, byRef blnInventoryUpdateOnly)

Dim pstrSQL
Dim pblnTracked
Dim pblnStatus
Dim pblnNotify
Dim plngLowFlag
Dim pblnCanBackOrder
Dim plngDefaultQTY

Dim plngInventory
Dim plngInventoryLow
Dim pblnOnOrder

Dim objRSInventory
Dim objRSInventoryInfo
Dim pstrProductID
Dim paryAttDetailUIDs
Dim paryAttDetailNames
Dim pstrAttDetailUIDs
Dim pstrAttDetailNames

Dim pblnLocalDebug
Dim pstrLocalDebugOut
Dim pblnSuccess
Dim pstrLocalError

	pblnSuccess = True

	pblnLocalDebug = False	'True	False

	pstrProductID = aryProduct(0)
	paryAttDetailUIDs = aryProduct(1)
	paryAttDetailNames = aryProduct(2)

	'Protect against invalid product UIDs
	If pstrProductID = -1 Then
		setInventoryLevels = True
		Exit Function
	End If

	If isArray(paryAttDetailUIDs) Then
		pstrAttDetailUIDs = paryAttDetailUIDs(0)
		pstrAttDetailNames = paryAttDetailNames(0)
		For i = 1 To UBound(paryAttDetailUIDs)
			pstrAttDetailUIDs = pstrAttDetailUIDs & "," & paryAttDetailUIDs(i)
			pstrAttDetailNames = pstrAttDetailNames & ", " & paryAttDetailNames(i)
		Next 'i
	Else
		pstrAttDetailUIDs = paryAttDetailUIDs
		pstrAttDetailNames = paryAttDetailNames	
	End If
	If pblnLocalDebug Then WriteOutput "&nbsp;&nbsp;&nbsp;setInventoryLevels (pstrAttDetailUIDs-" & pstrAttDetailUIDs & ")<br />" & vbcrlf
	
	If Len(pstrAttDetailUIDs) = 0 Then pstrAttDetailUIDs = "0"
	'items without attributes use ProductID instead of 0 like prior version - SF6 specific
	'If Len(pstrAttDetailUIDs) = 0 Then pstrAttDetailUIDs = pstrProductID

	pblnTracked = getValueFrom(maryInventoryFields, objrsProducts, 0, True, "0")
	pblnStatus = getValueFrom(maryInventoryFields, objrsProducts, 1, True, "0")
	pblnNotify = getValueFrom(maryInventoryFields, objrsProducts, 2, True, "0")
	plngLowFlag = getValueFrom(maryInventoryFields, objrsProducts, 3, False, "0")
	pblnCanBackOrder = getValueFrom(maryInventoryFields, objrsProducts, 4, True, "0")
	plngDefaultQTY = getValueFrom(maryInventoryFields, objrsProducts, 5, False, "0")
	plngInventory = getValueFrom(maryInventoryFields, objrsProducts, 6, False, plngDefaultQTY)
	plngInventoryLow = getValueFrom(maryInventoryFields, objrsProducts, 7, False, plngLowFlag)
	pblnOnOrder = False
	
	'Convert string to boolean
	If customBoolean(pblnTracked) Then
		pblnTracked = 1
	Else
		pblnTracked = 0
	End If
	If customBoolean(pblnStatus) Then
		pblnStatus = 1
	Else
		pblnStatus = 0
	End If
	If customBoolean(pblnNotify) Then
		pblnNotify = 1
	Else
		pblnNotify = 0
	End If
	If customBoolean(pblnCanBackOrder) Then
		pblnCanBackOrder = 1
	Else
		pblnCanBackOrder = 0
	End If
	
	If pblnLocalDebug Then
		pstrLocalDebugOut = pstrLocalDebugOut & "<fieldset><legend>Inventory Info</legend>"
		pstrLocalDebugOut = pstrLocalDebugOut & "pblnTracked: " & pblnTracked & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "pblnStatus: " & pblnStatus & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "pblnNotify: " & pblnNotify & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "plngLowFlag: " & plngLowFlag & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "pblnCanBackOrder: " & pblnCanBackOrder & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "plngDefaultQTY: " & plngDefaultQTY & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "plngInventory: " & plngInventory & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "plngInventoryLow: " & plngInventoryLow & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "</fieldset>"
		WriteOutput pstrLocalDebugOut & vbcrlf
	End If	'pblnLocalDebug

	If Not pblnTracked Then
		WriteOutput "&nbsp;&nbsp;&nbsp;Inventory not tracked for this item<br />" & vbcrlf
		
		pstrSQL = "Select invenProdId from sfInventoryInfo Where invenProdId=" & wrapSQLValue(pstrProductID, False, enDatatype_string)
		Set objRSInventoryInfo = GetRS(pstrSQL)
		If objRSInventoryInfo.EOF Then
			WriteOutput "&nbsp;&nbsp;&nbsp;-&nbsp;No existing inventory defaults to remove<br />" & vbcrlf
		Else
			pstrSQL = "Delete From sfInventoryInfo Where invenProdId=" & wrapSQLValue(pstrProductID, False, enDatatype_string)
			If Execute_NoReturn(pstrSQL, pstrLocalError) Then
				WriteOutput "&nbsp;&nbsp;&nbsp;-&nbsp;Existing inventory defaults removed<br />" & vbcrlf

				pstrSQL = "Delete From sfInventory Where invenProdId=" & wrapSQLValue(pstrProductID, False, enDatatype_string)
				If Execute_NoReturn(pstrSQL, pstrLocalError) Then
					WriteOutput "&nbsp;&nbsp;&nbsp;-&nbsp;Existing inventory Options removed<br />" & vbcrlf
				Else
					pblnSuccess = False
					WriteOutput "<b><font color=red>&nbsp;&nbsp;&nbsp;-&nbsp;Error removing existing inventory options</font></b><br />" & pstrLocalError & vbcrlf
				End If
			Else
				pblnSuccess = False
				WriteOutput "<b><font color=red>&nbsp;&nbsp;&nbsp;-&nbsp;Error removing existing inventory defaults</font></b><br />" & pstrLocalError & vbcrlf
			End If
		End If	'Not objRSInventoryInfo.EOF
		
		Call ReleaseObject(objRSInventory)
		Call ReleaseObject(objRSInventoryInfo)
		
		setInventoryLevels = pblnSuccess
		Exit Function
	End If	'Not pblnTracked

	pstrSQL = "Select invenProdId from sfInventoryInfo Where invenProdId=" & wrapSQLValue(pstrProductID, False, enDatatype_string)
	Set objRSInventoryInfo = GetRS(pstrSQL)

	If pblnLocalDebug Then
		pstrLocalDebugOut = pstrLocalDebugOut & "<fieldset><legend>Check for existing inventory info</legend>"
		pstrLocalDebugOut = pstrLocalDebugOut & "pstrSQL: " & pstrSQL & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "objRSInventoryInfo.EOF: " & objRSInventoryInfo.EOF & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "</fieldset>"
		WriteOutput pstrLocalDebugOut & vbcrlf
	End If	'pblnLocalDebug

	If objRSInventoryInfo.EOF Then
		If pblnTracked Then
			pstrSQL = "Insert Into sfInventoryInfo (invenProdId, invenbTracked, invenbStatus, invenbNotify, invenLowFlagDEF, invenbBackOrder, invenInStockDEF)" _
					& " Values (" & wrapSQLValue(pstrProductID, False, enDatatype_string) & ", " & wrapSQLValue(pblnTracked, False, enDatatype_number) & ", " & wrapSQLValue(pblnStatus, False, enDatatype_number) & ", " & wrapSQLValue(pblnNotify, False, enDatatype_number) & ", " & wrapSQLValue(plngLowFlag, False, enDatatype_number) & ", " & wrapSQLValue(pblnCanBackOrder, False, enDatatype_number) & ", " & wrapSQLValue(plngDefaultQTY, False, enDatatype_number) & ")"
Response.Write "pstrSQL: " & pstrSQL & "<br>"
			If Execute_NoReturn(pstrSQL, pstrLocalError) Then
				WriteOutput "&nbsp;&nbsp;&nbsp;Inventory Options added<br />" & vbcrlf
			Else
				pblnSuccess = False
				WriteOutput "<b><font color=red>Error setting inventory options</font></b><br />" & pstrLocalError & vbcrlf
			End If
		End If
	Else
		If Not blnInventoryUpdateOnly Then
			pstrSQL = "Update sfInventoryInfo Set" _
					& " invenbTracked=" & wrapSQLValue(pblnTracked, False, enDatatype_number)  & "," _
					& " invenbStatus=" & wrapSQLValue(pblnStatus, False, enDatatype_number)  & "," _
					& " invenbNotify=" & wrapSQLValue(pblnNotify, False, enDatatype_number)  & "," _
					& " invenLowFlagDEF=" & wrapSQLValue(plngLowFlag, False, enDatatype_number)  & "," _
					& " invenbBackOrder=" & wrapSQLValue(pblnCanBackOrder, False, enDatatype_number) & "," _
					& " invenInStockDEF=" & wrapSQLValue(plngDefaultQTY, False, enDatatype_number) & "" _
					& " Where invenProdId=" & wrapSQLValue(pstrProductID, False, enDatatype_string)
			If Execute_NoReturn(pstrSQL, pstrLocalError) Then
				WriteOutput "&nbsp;&nbsp;&nbsp;Inventory Options updated<br />" & vbcrlf
			Else
				WriteOutput "<b><font color=red>Error updating inventory options</font></b><br />" & pstrLocalError & vbcrlf
				pblnSuccess = False
			End If
		End If
	End If
	
	pstrSQL = "Select invenId from sfInventory Where invenProdId=" & wrapSQLValue(pstrProductID, False, enDatatype_string) _
			& " AND (invenAttDetailID=" & wrapSQLValue(pstrAttDetailUIDs, False, enDatatype_string) & " OR invenAttDetailID='0')"
	
	Set objRSInventory = GetRS(pstrSQL)
	If pblnLocalDebug Then
		pstrLocalDebugOut = pstrLocalDebugOut & "<fieldset><legend>Check for existing inventory record(s)</legend>"
		pstrLocalDebugOut = pstrLocalDebugOut & "pstrSQL: " & pstrSQL & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "objRSInventory.EOF: " & objRSInventory.EOF & "<br />"
	End If	'pblnLocalDebug

	If objRSInventory.EOF Then
		If Not blnInventoryUpdateOnly Then
			pstrSQL = "Insert Into sfInventory (invenProdId,invenAttDetailID,invenAttName,invenInStock,invenLowFlag)" _
					& " Values (" & wrapSQLValue(pstrProductID, False, enDatatype_string) & "," & wrapSQLValue(pstrAttDetailUIDs, False, enDatatype_string) & "," & wrapSQLValue(pstrAttDetailNames, False, enDatatype_string) & "," & plngInventory & "," & plngInventoryLow & ")"
			If Execute_NoReturn(pstrSQL, pstrLocalError) Then
				WriteOutput "&nbsp;&nbsp;&nbsp;Inventory updated to quantity " & plngInventory & "<br />" & vbcrlf
			Else
				pblnSuccess = False
				WriteOutput "<b><font color=red>Error setting inventory quantity</font></b><br />" & pstrLocalError & vbcrlf
			End If
		End If
	Else
		pstrSQL = "Update sfInventory Set invenInStock=" & plngInventory _
				& " Where invenId=" & wrapSQLValue(objRSInventory.Fields("invenId").Value, False, enDatatype_number)
		If Execute_NoReturn(pstrSQL, pstrLocalError) Then
			WriteOutput "&nbsp;&nbsp;&nbsp;Inventory updated to quantity " & plngInventory & "<br />" & vbcrlf
		Else
			pblnSuccess = False
			WriteOutput "<b><font color=red>Error updating inventory quantity</font></b><br />" & pstrLocalError & vbcrlf
		End If
	End If

	If pblnLocalDebug Then 
		pstrLocalDebugOut = pstrLocalDebugOut & "pstrSQL: " & pstrSQL & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "Error: " & Err.Description & "<br />"
		pstrLocalDebugOut = pstrLocalDebugOut & "</fieldset>"
		WriteOutput pstrLocalDebugOut & vbcrlf
	End If	'pblnLocalDebug

	Call ReleaseObject(objRSInventory)
	Call ReleaseObject(objRSInventoryInfo)
	
	setInventoryLevels = pblnSuccess
	
End Function	'setInventoryLevels

'****************************************************************************************************************************************************************

Function setSubCategories(byVal lngProductID, byVal strValue, byRef objrsCategories, byRef dicCategories)

Dim i
Dim parySubcategory
Dim pstrTempName
Dim plngUID
Dim plngPrevUID
Dim plngEndPos

	parySubcategory = Split(strValue, cstrSubcategoryDelimiter)
	plngUID = 0
	
	plngEndPos = UBound(parySubcategory)
	For i = 0 To plngEndPos
	
		If Len(pstrTempName) = 0 Then
			pstrTempName = Trim(parySubcategory(i))
		Else
			pstrTempName = pstrTempName & cstrSubcategoryDelimiter & Trim(parySubcategory(i))
		End If
		
		If dicCategories.Exists(pstrTempName) Then 
			plngUID = dicCategories(pstrTempName)
		Else
			objrsCategories.Filter = "[Name] = " & wrapSQLValue(Trim(parySubcategory(i)), False, enDatatype_string) & " AND ParentID=" & plngUID
			If Not objrsCategories.EOF Then
				plngUID = objrsCategories.Fields("UID").Value
			Else
				plngUID = -1
			End If
		
			dicCategories.Add pstrTempName, plngUID
		End If

		If plngUID <> -1 Then
			If cblnAssignProductsToAllCategoryLevels Then
				If createProductCategoryAssignment(lngProductID, plngUID) Then
					WriteOutput "&nbsp;&nbsp;&nbsp;Created Product Category Assignment for <em>" & pstrTempName & "</em><br />" & vbcrlf
				End If
			ElseIf CBool(i = plngEndPos) Then
				If createProductCategoryAssignment(lngProductID, plngUID) Then
					WriteOutput "&nbsp;&nbsp;&nbsp;Created Product Category Assignment for <em>" & pstrTempName & "</em><br />" & vbcrlf
				End If
			Else
				WriteOutput "&nbsp;&nbsp;&nbsp;No Product Category Assignment for <em>" & pstrTempName & "</em> - Option selected to only create assignments at bottom level of category structure.<br />" & vbcrlf
			End If
		Else
				WriteOutput "&nbsp;&nbsp;&nbsp;No Product Category Assignment for <em>" & pstrTempName & "</em> - Unable to locate category.<br />" & vbcrlf
		End If
	Next 'i
	objrsCategories.Filter = ""
	
	setSubCategories = plngUID

End Function	'setSubCategories

'****************************************************************************************************************************************************************

Function setVolumePricing(byRef objrsProducts, byVal lngProductUID, byVal strBasePrice)
'This is for figuring out MTP
'use maryMTPs(numColums)(information)
'Column name is in format MTP-##-???-QtyOrDiscount Where
'	MTP - identifies this as a column requiring special attention
'	## - represents the qty level
'	??? - represents the type of discount
'	QtyOrDiscount - represents whether ## is the Qty or the discount - values can be QTY, DISCOUNT, or empty (empty means QTY for backward compatibility)

'Table information
'mtIndex - 
'mtQuantity - 
'mtValue - 
'mtType	- PERCENT or AMOUNT 

Dim j
Dim plngMTP_Index
Dim plngMTP_Qty
Dim plngMTP_Value
Dim plngMTP_Discount
Dim pstrMTP_Type
Dim pstrFieldValue
Dim pstrSQL
Dim pblnSuccess
Dim pstrLocalError

	pblnSuccess = True

	If Not hasMTP Then
		setVolumePricing = True
		Exit Function
	End If
	
	If mblnDeleteExistingMTPs Then
		pstrSQL = "Delete From sfMTPrices Where mtProdID=" & wrapSQLValue(lngProductUID, True, enDatatype_string)
		If Execute_NoReturn(pstrSQL, pstrLocalError) Then
			WriteOutput "&nbsp;&nbsp;&nbsp;Cleared volume pricing levels.<br />" & vbcrlf
		Else
			pblnSuccess = False
			WriteOutput "<b><font color=red>Error removing existing volume pricing levels.</font></b><br />" & pstrLocalError & vbcrlf
		End If
	End If
	
	'added switch to revert to custom import
	If Left(maryMTPs(0)(4), Len(cstrMTPImportPrefix)) <> cstrMTPImportPrefix Then
		setVolumePricing = setVolumePricing_Custom(objrsProducts, lngProductUID, strBasePrice)
		Exit Function
	End If
		
	For j = 0 To UBound(maryMTPS)
		'maryMTPs(plngMTPColumns) = Array("mtIndex", "mtQuantity", "mtValue", "mtType", "columnName", "QtyOrDiscount")
		pstrFieldValue = Trim(objrsProducts.Fields(maryMTPS(j)(4)).Value & "")
		If Len(pstrFieldValue) > 0 Then
			'debugprint "MTP " & j, maryMTPS(j)(4)
			plngMTP_Index = maryMTPS(j)(0)
			
			If maryMTPS(j)(5) = "QTY" Then
				plngMTP_Qty = maryMTPS(j)(1)
				plngMTP_Discount = pstrFieldValue
			Else
				plngMTP_Qty = pstrFieldValue
				plngMTP_Discount = maryMTPS(j)(1)
			End If
			
			pstrMTP_Type = maryMTPS(j)(3)
			If LCase(pstrMTP_Type) = "calculate" Then
				pstrMTP_Type = "0"
				If Not isNumeric(strBasePrice) Then strBasePrice = 0
				plngMTP_Value = CDbl(plngMTP_Discount) - CDbl(strBasePrice)
			Else
				If UCase(pstrMTP_Type) = "PERCENT" Then
					pstrMTP_Type = 1
				Else
					pstrMTP_Type = 0
				End If
				plngMTP_Value = plngMTP_Discount
			End If
			plngMTP_Value = Round(plngMTP_Value, 2)

			If createVolumePriceBreak(lngProductUID, j, plngMTP_Qty, plngMTP_Value, pstrMTP_Type) Then
				If pstrMTP_Type = 1 Then
					WriteOutput "&nbsp;&nbsp;&nbsp;Inserted volume pricing level of " & plngMTP_Discount & " percent for " & plngMTP_Qty & " items.<br />" & vbcrlf
				Else
					WriteOutput "&nbsp;&nbsp;&nbsp;Inserted volume pricing level of " & plngMTP_Discount & " (amount) for " & plngMTP_Qty & " items.<br />" & vbcrlf
				End If
				If CDbl(plngMTP_Value) < 0 Then
					WriteOutput "&nbsp;&nbsp;&nbsp;<font color=red>Discount amount was adjusted to be a positive value per application requirements. <em>" & CDbl(plngMTP_Value) & "</em></font><br />" & vbcrlf
				End If
			Else
				pblnSuccess = False
				WriteOutput "<b><font color=red>Error setting volume pricing</font></b><br />" & pstrLocalError & vbcrlf
			End If

		End If	'Len(pstrFieldValue) > 0 
	Next 'j
	
	setVolumePricing = pblnSuccess

End Function	'setVolumePricing

'****************************************************************************************************************************************************************

Function setVolumePricing_Custom(byRef objrsProducts, byVal lngProductUID, byVal strBasePrice)
'This is for figuring out MTP
'use maryMTPs(numColums)(information)
'Column name is in format MTP-##-???-QtyOrDiscount Where
'	MTP - identifies this as a column requiring special attention
'	## - represents the qty level
'	??? - represents the type of discount
'	QtyOrDiscount - represents whether ## is the Qty or the discount - values can be QTY, DISCOUNT, or empty (empty means QTY for backward compatibility)

'Table information
'mtIndex - 
'mtQuantity - 
'mtValue - 
'mtType	- PERCENT or AMOUNT 

Dim j
Dim plngMTP_Index
Dim plngMTP_Qty
Dim plngMTP_Value
Dim plngMTP_Discount
Dim pstrMTP_Type
Dim pstrFieldValue
Dim pstrSQL
Dim pblnSuccess
Dim pstrLocalError

	pblnSuccess = True

	'test to make sure field is present
	If Err.number <> 0 Then Err.Clear
	On Error Resume Next
	plngMTP_Qty = Trim(objrsProducts.Fields(maryMTPS(j)(4)).Value & "")
	If Err.number <> 0 Then
		Err.Clear
		setVolumePricing_Custom = True
		Exit Function
	End If
	On Error Goto 0
	
	For j = 0 To UBound(maryMTPS)
		'maryMTPs(plngMTPColumns) = Array("mtIndex", "mtQuantity", "mtValue", "mtType", "columnName", "QtyOrDiscount")
		plngMTP_Qty = Trim(objrsProducts.Fields(maryMTPS(j)(4)).Value & "")
		If Len(plngMTP_Qty) = 0 Then
			plngMTP_Qty = 0
		ElseIf Not isNumeric(plngMTP_Qty) Then
			plngMTP_Qty = 0
		Else
			plngMTP_Qty = CLng(plngMTP_Qty)
		End If
		
		If plngMTP_Qty > 0 Then
			'debugprint "MTP " & j, maryMTPS(j)(4)
			plngMTP_Index = maryMTPS(j)(0)
			
			'Quantity is from the QuantityX column
			'Price is from the PriceX column
			plngMTP_Discount = Trim(objrsProducts.Fields("Price" & j + 2).Value & "")
			pstrMTP_Type = maryMTPS(j)(3)
			
			If LCase(pstrMTP_Type) = "calculate" Then
				pstrMTP_Type = 0
				If Not isNumeric(strBasePrice) Then strBasePrice = 0
				plngMTP_Value = CDbl(strBasePrice) - CDbl(plngMTP_Discount)
			Else
				If UCase(pstrMTP_Type) = "PERCENT" Then
					pstrMTP_Type = 1
				Else
					pstrMTP_Type = 0
				End If
				plngMTP_Value = plngMTP_Discount
			End If
			
			plngMTP_Value = Round(plngMTP_Value, 2)
			
			If createVolumePriceBreak(lngProductUID, j, plngMTP_Qty, plngMTP_Value, pstrMTP_Type) Then
				If pstrMTP_Type = 1 Then
					WriteOutput "&nbsp;&nbsp;&nbsp;Inserted volume pricing level of " & plngMTP_Value & " percent for " & plngMTP_Qty & " items.<br />" & vbcrlf
				Else
					WriteOutput "&nbsp;&nbsp;&nbsp;Inserted volume pricing level of " & plngMTP_Value & " (amount) for " & plngMTP_Qty & " items.<br />" & vbcrlf
				End If
				If CDbl(plngMTP_Value) < 0 Then
					WriteOutput "&nbsp;&nbsp;&nbsp;<font color=red>Discount amount was adjusted to be a positive value per application requirements. <em>" & CDbl(plngMTP_Value) & "</em></font><br />" & vbcrlf
				End If
			Else
				pblnSuccess = False
				WriteOutput "<b><font color=red>Error setting volume pricing</font></b>" & vbcrlf
			End If

		End If	'Len(pstrFieldValue) > 0 
	Next 'j
	
	setVolumePricing_Custom = pblnSuccess

End Function	'setVolumePricing_Custom

'****************************************************************************************************************************************************************

Function spreadsheetFieldName(byVal strdbFieldName)

Dim i
Dim pstrTemp

	For i = 0 To UBound(maryFields)
		'debugprint maryFields(i)(enTargetFieldName) & " - " & strdbFieldName, maryFields(i)(enTargetFieldName) = strdbFieldName
		If maryFields(i)(enTargetFieldName) = strdbFieldName Then
			pstrTemp = maryFields(i)(enSourceFieldName)
		End If
	Next 'i
	spreadsheetFieldName = pstrTemp

End Function	'spreadsheetFieldName

'****************************************************************************************************************************************************************

Function spreadsheetIndexByFieldName(byVal strdbFieldName)

Dim i
Dim plngIndex

	plngIndex = -1
	For i = 0 To UBound(maryFields)
		'debugprint maryFields(i)(enTargetFieldName) & " - " & strdbFieldName, maryFields(i)(enTargetFieldName) = strdbFieldName
		If maryFields(i)(enTargetFieldName) = strdbFieldName Then
			plngIndex = i
			Exit For
		End If
	Next 'i
	
	spreadsheetIndexByFieldName = plngIndex

End Function	'spreadsheetIndexByFieldName

'****************************************************************************************************************************************************************

Function updateDetailLink(byRef objrsProducts, byVal lngProductUID, byVal strProductID)

Dim pstrSQLUpdate
Dim plngIndex
Dim pstrFieldValue
Dim pblnSuccess
Dim pstrLocalError

	pblnSuccess = True

	plngIndex = spreadsheetIndexByFieldName("prodLink")
	If plngIndex >= 0 Then
		If Len(maryFields(plngIndex)(enSourceFieldName)) = 0 Then
			pstrFieldValue = maryFields(plngIndex)(enDefaultValue)
		Else
			pstrFieldValue = objrsProducts.Fields(maryFields(plngIndex)(enSourceFieldName)).Value
		End If
	End If

	pstrFieldValue = checkReplacements(pstrFieldValue, strProductID)
	pstrFieldValue = Replace(pstrFieldValue, "<uid>", lngProductUID)

	pstrSQLUpdate = "Update sfProducts Set prodLink='" & pstrFieldValue & "' Where sfProductID=" & lngProductUID
	'response.Write "pstrSQLUpdate: " & pstrSQLUpdate & "<br />"
	If Execute_NoReturn(pstrSQLUpdate, pstrLocalError) Then
		WriteOutput "&nbsp;&nbsp;&nbsp;Updated detail link for uid<br />" & vbcrlf
	Else
		pblnSuccess = False
		WriteOutput "<b><font color=red>Error updating detail link</font></b><br />" & pstrLocalError & vbcrlf
	End If

	updateDetailLink = pblnSuccess

End Function	'updateDetailLink

'****************************************************************************************************************************************************************

Sub verifyCategories(byRef objrsNewProducts, byRef strFieldName, byRef objrsCategories, byRef dicCategories)
'objrsNewProducts comes in as just the Category field
'Purpose: this function accepts the products to be entered, sorted by Category and the Category recordset
'Categories are stored to a dictionary for fast retrieval as they are entered/found.
'If a category or subcategory is not located in the recordset it is created and the uid is set to the dictionary

Dim i
Dim pblnAdded
Dim pblnLocalDebug
Dim pstrLocalDebugOut
Dim pstrPrev

	pblnLocalDebug = False	'True	False
	pblnAdded = False
	
	'The purpose of this section is to speed up category access
	'Note: individual categories CANNOT be added to the dictionary since there could be duplicate names
	
	If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "<fieldset><legend>verifyCategories</legend>"

	'Check in case no Category column is specified
	If Len(strFieldName) = 0 Then Exit Sub
	
	Do While Not objrsNewProducts.EOF
		pstrPrev = Trim(objrsNewProducts(strFieldName).Value & "")
		'Check added to remove erroneous leading delimiter
		If Left(pstrPrev, 1) = cstrMultipleCategoryDelimiter Then pstrPrev = Replace(pstrPrev, cstrMultipleCategoryDelimiter, "", 1, 1)
		
		If dicCategories.Exists(pstrPrev) Then
			If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "Category <i>" & pstrPrev & "</i> exists. (uid=" & dicCategories(pstrPrev) & ")<br />"
		Else
			If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "Category <i>" & pstrPrev & "</i> not found.<br />"
			Call verifyCategory(pstrPrev, objrsCategories, dicCategories, pblnAdded)
		End If
		
		objrsNewProducts.MoveNext
	Loop
	If objrsNewProducts.RecordCount > 0 Then objrsNewProducts.MoveFirst
	objrsCategories.Filter = ""
	
	If pblnLocalDebug Then
		pstrLocalDebugOut = pstrLocalDebugOut & "</fieldset>"
		Response.Write pstrLocalDebugOut
	End If
	
	If pblnLocalDebug Then
		Dim vItem
		Response.Write "<fieldset><legend>verifyCategories: " & dicCategories.count & " categories</legend>"
		For Each vItem in dicCategories
			Response.Write vItem & " (" & dicCategories(vItem) & ")<br />"
		Next
		Response.Write "</fieldset>"
	End If	'pblnLocalDebug
	
	If pblnAdded Then 
		Set objrsCategories = GetRS("Select uid, [Name], ParentLevel, ParentID from Categories Order By [Name]")
	End If

End Sub	'verifyCategories

'****************************************************************************************************************************************************************

Function verifyCategory(byVal strCategoryName, byRef objrsCategories, byRef dicCategories, byRef blnAdded)
'Purpose: this function accepts the products to be entered, sorted by Category and the Category recordset
'Categories are stored to a dictionary for fast retrieval as they are entered/found.
'If a category or subcategory is not located in the recordset it is created and the uid is set to the dictionary

Dim i
Dim paryCategories
Dim pblnLocalDebug
Dim pblnNumericCategoryID
Dim plngCatID
Dim pstrCategoryToCheck
Dim pstrLocalDebugOut

	pblnLocalDebug = False	'True	False
	blnAdded = False
	
	If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "<fieldset><legend>verifyCategory</legend>"

	'Check to make sure default category exists
	If Not isNumeric(mlngDefaultCategoryID) Then
		If getCategoryUIDbyName(mlngDefaultCategoryID, objrsCategories, dicCategories) = -1 Then
			If createCategory(plngCatID, mlngDefaultCategoryID, 0, 0, 1) Then
				dicCategories.Add mlngDefaultCategoryID, plngCatID
				WriteOutput "Default Category <em>" & strCategoryName & "</em> added.<br />" & vbcrlf
				blnAdded = True
			Else
				plngCatID = 1
				WriteOutput "&nbsp;&nbsp;<font color=red>Error creating category <em>" & mlngDefaultCategoryID & "</em>.</font><br />" & vbcrlf
			End If
		End If
	End If
	
	'Check added to remove erroneous leading delimiter
	If Left(strCategoryName, 1) = cstrMultipleCategoryDelimiter Then strCategoryName = Replace(strCategoryName, cstrMultipleCategoryDelimiter, "", 1, 1)
	
	Call cleanCategoryData(strCategoryName)

	If dicCategories.Exists(strCategoryName) Then
		plngCatID = dicCategories(strCategoryName)
		If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "Category <i>" & strCategoryName & "</i> exists. (uid=" & plngCatID & ")<br />"
	Else
		If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "Category <i>" & strCategoryName & "</i> not found.<br />"

		paryCategories = Split(strCategoryName, cstrMultipleCategoryDelimiter)
		If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "<ol>"
		For i = 0 To UBound(paryCategories)
			pstrCategoryToCheck = Trim(paryCategories(i))
			pblnNumericCategoryID = isNumeric(pstrCategoryToCheck)
			If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "<li>" & pstrCategoryToCheck
			If pblnNumericCategoryID Then
				WriteOutput "&nbsp;&nbsp;<font color=red>Category <em>" & pstrCategoryToCheck & "</em> not added because it is a number. This message is normal if you are importing cateogries by their UID.</font><br />" & vbcrlf
			ElseIf instr(1, pstrCategoryToCheck, cstrSubcategoryDelimiter) > 0 Then
				If Not dicCategories.Exists(pstrCategoryToCheck) Then 
					plngCatID = checkForSubCategories(pstrCategoryToCheck, dicCategories, pblnLocalDebug, pstrLocalDebugOut)
					blnAdded = True
				End If
			Else
				If dicCategories.Exists(pstrCategoryToCheck) Then 
					plngCatID = dicCategories(pstrCategoryToCheck)
				Else
					'Category is assumed to be a top level category
					On Error Resume Next	'added since category section could be > 255
					If Err.number <> 0 Then Err.Clear
					plngCatID = -1
					If cblnSF5AE Then
						objrsCategories.Filter = "subcatName = " & wrapSQLValue(pstrCategoryToCheck, False, enDatatype_string) & " AND Depth=0"
						If Not objrsCategories.EOF Then plngCatID = objrsCategories.Fields("subcatID").Value
					Else
						objrsCategories.Filter = "catName = " & wrapSQLValue(pstrCategoryToCheck, False, enDatatype_string)
						If Not objrsCategories.EOF Then plngCatID = objrsCategories.Fields("catID").Value
					End If
					If Err.number = 0 Then
						If plngCatID = -1 Then
							If createCategory(plngCatID, pstrCategoryToCheck, 0, 0, 1) Then
								If Not blnAdded Then WriteOutput "Adding Categories . . .<br />" & vbcrlf
								dicCategories.Add pstrCategoryToCheck, plngCatID
								If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "<ul><li>Category <em>" & pstrCategoryToCheck & "(" & plngCatID & ")</em> added.</li></ul>"
								WriteOutput "&nbsp;&nbsp;Category <em>" & pstrCategoryToCheck & "</em> added.<br />" & vbcrlf
								blnAdded = True
							Else
								plngCatID = 1
								WriteOutput "&nbsp;&nbsp;<font color=red>Error creating category <em>" & pstrCategoryToCheck & "</em>. Default value of " & plngCatID & " used.</font><br />" & vbcrlf
							End If
						Else
							dicCategories.Add pstrCategoryToCheck, plngCatID
						End If	'objrsCategories.EOF
					Else
						plngCatID = 1
						WriteOutput "&nbsp;&nbsp;<font color=red>Error creating category <em>" & pstrCategoryToCheck & "</em>. Default value of " & plngCatID & " used.</font><br />This may occur if your category field to import is greater than 255 characters<br />" & vbcrlf
					End If
				End If
			End If	'Not isNumeric(pstrCategoryToCheck)
			If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "</li>"
		Next	'i
		If pblnLocalDebug Then pstrLocalDebugOut = pstrLocalDebugOut & "</ol>"
	End If	'Not dicCategories.Exists(strCategoryName)

	objrsCategories.Filter = ""
	
	If pblnLocalDebug Then
		pstrLocalDebugOut = pstrLocalDebugOut & "</fieldset>"
		Response.Write pstrLocalDebugOut
	End If
	
	If pblnLocalDebug Then
		Dim vItem
		Response.Write "<fieldset><legend>" & dicCategories.count & " categories</legend>"
		For Each vItem in dicCategories
			Response.Write vItem & " (" & dicCategories(vItem) & ")<br />"
		Next
		Response.Write "</fieldset>"
	End If	'pblnLocalDebug
	
	verifyCategory = plngCatID
	
End Function	'verifyCategory

'****************************************************************************************************************************************************************

Sub VerifyConnections(ByRef objCnn, ByRef strTableSource)

Dim pobjCnnTarget
Dim pobjRS
Dim pstrSQL
Dim i
Dim pstrFileExtension
Dim pstrPathToUse

	pstrPathToUse = decodeFilePath(mstrDSN_Source)

	On Error Resume Next
	
	If Err.number <> 0 Then Err.Clear
	
    pstrFileExtension = LCase(Right(mstrDSN_Source, 4))
    Select Case pstrFileExtension
	    Case ".mdb"
		    mblnDSN_SourceExists = FileExists(pstrPathToUse)
		    If Not mblnDSN_SourceExists Then 
			    mstrMessage = "<h4><font color=red>The file <em>" & Replace(pstrPathToUse, " ", "&nbsp;") & "</em> does not exist at the specified location.<br />Please check the path you entered or make sure the file is located on the server itself.</font></h4>" & vbcrlf
		    Else
			    Set objCnn = Server.CreateObject("ADODB.Connection")
			    With objCnn
					.CursorLocation = 2 'adUseServer
				    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
									  & "Data Source=" & pstrPathToUse & ";"
				    .Open
			    End With
			    If Err.number <> 0 Then
				    Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Attempted Access connection</font><br />" & vbcrlf
					Response.Write "<font color=red>Connection String: " & objCnn.ConnectionString & "</font><br />" & vbcrlf
				    Response.Write "<font color=red>Source DSN: " & pstrPathToUse & "</font><br />" & vbcrlf
				    Err.Clear
			    End If
		    End If
		    mblnDSN_SourceExists = CBool(objCnn.State = 1)
		    pstrSQL = OpenTableSQL(mstrSourceTable)
	    Case ".xls"
		    mblnDSN_SourceExists = FileExists(pstrPathToUse)
		    If Not mblnDSN_SourceExists Then 
			    mstrMessage = "<h4><font color=red>The file <em>" & Replace(pstrPathToUse, " ", "&nbsp;") & "</em> does not exist at the specified location.<br />Please check the path you entered or make sure the file is located on the server itself.</font></h4>" & vbcrlf
		    Else
			    Set objCnn = Server.CreateObject("ADODB.Connection")
			    With objCnn
				    .Provider = "Microsoft.Jet.OLEDB.4.0"
				    '.CursorLocation = 3	'adUseClient
					.CursorLocation = 2 'adUseServer
				    .ConnectionString = "Data Source=" & pstrPathToUse & ";" _
								    & "Extended Properties=Excel 8.0;"
				    .Open
			    End With
			    If Err.number <> 0 Then
				    Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Attempted Excel connection</font><br />" & vbcrlf
					Response.Write "<font color=red>Connection String: " & objCnn.ConnectionString & "</font><br />" & vbcrlf
				    Response.Write "<font color=red>Source DSN: " & pstrPathToUse & "</font><br />" & vbcrlf
				    Err.Clear
			    End If
			    mblnDSN_SourceExists = CBool(objCnn.State = 1)
		    End If
		    pstrSQL = OpenTableSQL(mstrSourceTable & "$")
	    Case ".csv"
		    mblnDSN_SourceExists = FileExists(pstrPathToUse)
		    If mblnDSN_SourceExists Then 
			    Set objCnn = Server.CreateObject("ADODB.Connection")
			    With objCnn
				    '.CursorLocation = 3	'adUseClient
					.CursorLocation = 2 'adUseServer
				    .ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & pstrPathToUse & ";" _
									  & "Extensions=csv"
				    .Open
			    End With
			    If Err.number <> 0 Then
				    Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Attempted CSV connection</font><br />" & vbcrlf
					Response.Write "<font color=red>Connection String: " & objCnn.ConnectionString & "</font><br />" & vbcrlf
				    Response.Write "<font color=red>Source DSN: " & pstrPathToUse & "</font><br />" & vbcrlf
				    Err.Clear
			    End If
			    mblnDSN_SourceExists = CBool(objCnn.State = 1)
		    Else
			    mstrMessage = "<h4><font color=red>The file <em>" & Replace(pstrPathToUse, " ", "&nbsp;") & "</em> does not exist at the specified location.<br />Please check the path you entered or make sure the file is located on the server itself.</font></h4>" & vbcrlf
		    End If
		    pstrSQL = OpenTableSQL(mstrSourceTable & "$")
	    Case Else
		    Set objCnn = Server.CreateObject("ADODB.Connection")
		    objCnn.CommandTimeout = Server.ScriptTimeout
		    objCnn.Open mstrDSN_Source
		    If Err.number <> 0 Then
			    Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
				Response.Write "<font color=red>Attempted DSN connection</font><br />" & vbcrlf
				Response.Write "<font color=red>Connection String: " & objCnn.ConnectionString & "</font><br />" & vbcrlf
			    Response.Write "<font color=red>Source DSN: " & mstrDSN_Source & "</font><br />" & vbcrlf
			    Err.Clear
		    End If
		    mblnDSN_SourceExists = CBool(objCnn.State = 1)
		    pstrSQL = OpenTableSQL(mstrSourceTable)
    End Select
	strTableSource = pstrSQL
	
	If mblnDSN_SourceExists Then
		Set pobjRS = Server.CreateObject("ADODB.RECORDSET")
		With pobjRS
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly

			If Err.number <> 0 Then
				If InStr(1, Err.Description, "$' is not a valid name") > 0 Then
					Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Source Table: " & mstrSourceTable & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Query: " & pstrSQL & "</font><br />" & vbcrlf
					Response.Write "<font color=red>This error will occur if <ul><li>Database source - you have specified an invalid table name</li><li>Excel spreadsheet - you have specified a worksheet which does not exist</li></font><br />" & vbcrlf
				Else
					Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Connection String: " & objCnn.ConnectionString & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Source Table: " & mstrSourceTable & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Query: " & pstrSQL & "</font><br />" & vbcrlf
				End If
				Err.Clear
			Else
				On Error Goto 0
				mblnSourceTableExists = CBool(pobjRS.State = 1)
				If mblnSourceTableExists Then Call SetColumns(pobjRS)
				If mblnSourceTableExists And Not hasMTP Then Call SetMTPColumns(pobjRS)
			End If
			.Close
		End With
		Set pobjRS = Nothing
		
	End If	'mblnDSN_SourceExists
	
	'pobjCnn.Close
	'Set pobjCnn = Nothing
	
	'Now verify the target connection
	On Error Resume Next
	If Err.number <> 0 Then Err.Clear
	Set pobjCnnTarget = Server.CreateObject("ADODB.Connection")
	pobjCnnTarget.CommandTimeout = Server.ScriptTimeout
	pobjCnnTarget.Open mstrDSN_Target
	If Err.number <> 0 Then
		Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
		Response.Write "<font color=red>Target DSN: " & mstrDSN_Target & "</font><br />" & vbcrlf
		Err.Clear
	End If
	mblnDSN_TargetExists = CBool(pobjCnnTarget.State = 1)
	pobjCnnTarget.Close
	Set pobjCnnTarget = Nothing	

	If Trim(mstrDSN_Target) <> Trim(connectionString) Then
		Response.Write "<h4>Using user specified connection which is not the connection to the database specified in the web.config file.</h4>"
		cnn.Close
		cnn.Open mstrDSN_Target
	End If

End Sub	'VerifyConnections

'****************************************************************************************************************************************************************

Function verifyManufacturers(byRef objrsNewProducts, byRef strFieldName, byRef dicMfgVends)
'Purpose: iterates through objrsNewProducts to ensure Mfg/Vend exists in database
'Action: populates dictionary object with name/id pair as items are found and/or created

Dim pobjRSMfgVend
Dim pstrSQL
Dim pstrName
Dim plngNewID
Dim pblnAdded
Dim pstrMessageOut

	pblnAdded = False
	
	Do While Not objrsNewProducts.EOF
		pstrName = Trim(objrsNewProducts(strFieldName).Value & "")
		If Len(pstrName) = 0 Then
			'do nothing
		ElseIf Not isNumeric(pstrName) Then
			If Not dicMfgVends.Exists(pstrName) Then
				pstrSQL = "SELECT mfgID from sfManufacturers Where mfgName=" & wrapSQLValue(pstrName, False, enDatatype_string)
				Set pobjRSMfgVend = GetRS(pstrSQL)
				If pobjRSMfgVend.EOF Then
					plngNewID = getManufacturerByName(pstrName, lngType, True)
					If plngNewID <> -1 Then
						If Not pblnAdded Then pstrMessageOut = pstrMessageOut & "Adding manufacturer . . .<br />" & vbcrlf
						pstrMessageOut = pstrMessageOut & "&nbsp;&nbsp;<em>" & pstrName & "</em> added.<br />" & vbcrlf
						pblnAdded = True
						dicMfgVends.Add pstrName, plngNewID
					End If
				Else
					dicMfgVends.Add pstrName, pobjRSMfgVend.Fields("mfgID").Value
				End If	'pobjRSMfgVend.EOF
				pobjRSMfgVend.Close
				Set pobjRSMfgVend = Nothing
			End If
		End If
		objrsNewProducts.MoveNext
	Loop
	If objrsNewProducts.RecordCount > 0 Then objrsNewProducts.MoveFirst
	
	WriteOutput pstrMessageOut 
	verifyManufacturers = pstrMessageOut

End Function	'getManufacturerByName

'****************************************************************************************************************************************************************

Function verifyVendors(byRef objrsNewProducts, byRef strFieldName, byRef dicMfgVends)
'Purpose: iterates through objrsNewProducts to ensure Mfg/Vend exists in database
'Action: populates dictionary object with name/id pair as items are found and/or created

Dim pobjRSMfgVend
Dim pstrSQL
Dim pstrName
Dim plngNewID
Dim pblnAdded
Dim pstrMessageOut

	pblnAdded = False
	
	Do While Not objrsNewProducts.EOF
		pstrName = Trim(objrsNewProducts(strFieldName).Value & "")
		If Len(pstrName) = 0 Then
			'do nothing
		ElseIf Not isNumeric(pstrName) Then
			If Not dicMfgVends.Exists(pstrName) Then
				pstrSQL = "SELECT vendID from sfVendors Where vendName=" & wrapSQLValue(pstrName, False, enDatatype_string)
				Set pobjRSMfgVend = GetRS(pstrSQL)
				If pobjRSMfgVend.EOF Then
					plngNewID = getVendorByName(pstrName, lngType, True)
					If plngNewID <> -1 Then
						If Not pblnAdded Then pstrMessageOut = pstrMessageOut & "Adding vendor . . .<br />" & vbcrlf
						pstrMessageOut = pstrMessageOut & "&nbsp;&nbsp;<em>" & pstrName & "</em> added.<br />" & vbcrlf
						pblnAdded = True
						dicMfgVends.Add pstrName, plngNewID
					End If
				Else
					dicMfgVends.Add pstrName, pobjRSMfgVend.Fields("vendID").Value
				End If	'pobjRSMfgVend.EOF
				pobjRSMfgVend.Close
				Set pobjRSMfgVend = Nothing
			End If
		End If
		objrsNewProducts.MoveNext
	Loop
	If objrsNewProducts.RecordCount > 0 Then objrsNewProducts.MoveFirst
	
	WriteOutput pstrMessageOut 
	verifyVendors = pstrMessageOut

End Function	'verifyVendors

'**************************************************************************************************************************************************

Sub WriteFailedImports()

Dim paryProducts
Dim paryProduct
Dim pstrSession
Dim plngNumFailedProducts
Dim i
Dim pstrProductID
Dim pstrProductName
Dim pstrOut

	pstrSession = Session("ssFailedImports")
	If Len(pstrSession) > 0 Then
		paryProducts = Split("{new}", pstrSession)
		
		pstrOut = "<h4><font color=red>" & plngNumFailedProducts & " product(s) failed to import.</h4>" & vbcrlf
		pstrOut = pstrOut & "<p>You will need to manually verify each of the following products:</p>" & vbcrlf
		pstrOut = pstrOut & "<ol>" & vbcrlf
		For i = 0 To UBound(paryProducts)
			If Len(paryProducts(i)) > 0 Then
				paryProduct = Split(paryProducts(i), "|")
				pstrProductID = getArrayValue(paryProduct, 0, "")
				pstrProductName = getArrayValue(paryProduct, 1, "")
				pstrOut = pstrOut & "<li>" & pstrProductID & ": " & pstrProductName & "</li>" & vbcrlf
			End If
		Next 'i
		pstrOut = pstrOut & "</ol>" & vbcrlf
	End If
	
	WriteOutput pstrOut
		
End Sub	'WriteFailedImports

'**************************************************************************************************************************************************

Sub WriteOutput(byVal strOut)
	Response.Write strOut
	Call WriteToOutputWindow(strOut)
End Sub	'WriteOutput

'****************************************************************************************************************************************************************

Function WriteXMLProfile()

	On Error Resume Next
	
	mobjXMLProfiles.Save profilePath
	If Err.number = 0 Then
		WriteXMLProfile = True
	Else
		WriteOutput "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
		WriteOutput "Error saving profiles to disk. <br />" & vbcrlf
		WriteOutput "File: <i>" & profilePath & "</i><br />" & vbcrlf
		WriteOutput "This is likely due to MSXML not having write permissions. Please contact your host.<br />" & vbcrlf
		Err.Clear
		WriteXMLProfile = False
	End If

End Function	'WriteXMLProfile

'****************************************************************************************************************************************************************
'****************************************************************************************************************************************************************

Dim mblnDSN_SourceExists
Dim mstrDSN_Source
Dim mblnSourceTableExists
Dim mstrSourceTable

Dim mblnDSN_TargetExists
Dim mstrDSN_Target

Dim mbytCreateCat
Dim mbytCreateMfg
Dim mbytCreateVend
Dim mlngImportType

Dim cstrMultipleCategoryDelimiter
Dim cstrSubcategoryDelimiter
Dim mlngDefaultCategoryID
Dim mlngDefaultManufacturerID
Dim mlngDefaultVendorID
Dim maryCustomBooleanValues

Dim mstrProfileID
Dim cstrDefaultProfile
Dim marySheetColumns

Dim mblnValidImport
Dim mblnProfileOnly

Dim mlngDefaultImportType
Dim clngTarget_ProductID
Dim clngTarget_ProductName
Dim clngTarget_Cost
Dim clngTarget_Price
Dim clngTarget_SalePrice
Dim clngTarget_SaleIsActive

Dim mstrCategoryColumn
'added with 2.0
Dim mblnAttributesDeleteExisting
Dim mblnCategoriesDeleteExisting
Dim mblnAssignProductsToAllCategoryLevels
Dim mblnRelatedProductsDeleteExisting

Dim mstrImportTypeColumn 

Dim cstrMTPImportPrefix
Dim cstrMTPImportSeparator

Dim maryProfiles
Dim maryFields()
Dim maryInventoryFields()
Dim maryGiftWrap()
Dim maryMTPs()
Dim mblnDeleteExistingMTPs

'Added for attribute import
Dim maryAttributes()

'Declared here because they are commonly used
Dim mlngMaxRecords
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="Common/ssProducts_Common.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'pstrssAddonCode = "ProductImportSF6"
mstrssAddonVersion = "2.00.001"

'**************************************************
'
'	Start Code Execution
'

'debug.Enabled = True
'debug.printform

Dim mblnssDebugImport
Dim mblnssDebugShowTime
Dim mlngDebugType
Dim mstrAction
Dim mobjCnn
Dim mstrTableSource
Dim mobjXMLProfiles

Dim maryDebugTypes(1)
maryDebugTypes(0) = Array("Basic Display (recommended)", "Check this option to display the minimum output information from the debug")
maryDebugTypes(1) = Array("Include times (debugging only)", "Check this option to display the individual times for each major step of the import")
'maryDebugTypes(2) = Array("Detailed output (full debug mode)", "Check this option to display the individual times for each major step of the import and complete debugging output")

mlngDebugType = Request.Form("DebugType")
If Len(mlngDebugType) = 0 Then
	mlngDebugType = 0
Else
	mlngDebugType = CLng(mlngDebugType)
End If
Select Case mlngDebugType
	Case 0:
				mblnssDebugShowTime = False
				mblnssDebugImport = False
	Case 1:
				mblnssDebugShowTime = True
				mblnssDebugImport = False
	Case 2:
				mblnssDebugShowTime = True
				mblnssDebugImport = True
	Case Else:
				mblnssDebugShowTime = False
				mblnssDebugImport = False
End Select

	mstrPageTitle  = "Product Import"
	'Call CheckForUpdatedVersion("ProductImportSF5", mstrssAddonVersion)
	Call WriteHeader("setPageDefaults();",True)
	
%>
<center>
<script language="javascript">

//tipMessage[...]=[title,text]
tipMessage['ProfileID']=["Data Entry Help", "Select from a preconfigured import profile"]
tipMessage['DSN_Source']=["Data Entry Help", "This either the path to the data file (usually an Excel worksheet or Access database) which must be located on the server or a complete connection string to your datasource."]
tipMessage['SourceTable']=["Data Entry Help", "This is the table, query/view, or worksheet name (Excel only) you're using for your datasource. If a valid data source is selected a drop down listing potential table names will appear. However, if your table name contains a $ it will not function correctly. Additionally some Excel workbooks with deleted worksheets will still contain the names of the deleted worksheets."]
tipMessage['DSN_Target']=["Data Entry Help", "This is the connection string as identified in the web.config file. It should not be changed and is displayed for reference only."]

tipMessage['CategoryColumn']=["Data Entry Help", "This items is the column used to identify the product's category/subcategory assignment(s)"]
tipMessage['SubcategoryDelimiter']=["Data Entry Help", "This items is used to separate levels of categories/sub-categories"]
tipMessage['MultipleCategoryDelimiter']=["Data Entry Help", "This items is used to separate multiple category assignments"]

tipMessage['imgOpenProfile']=["Data Entry Help", "Open the saved profile."]
tipMessage['imgNewProfile']=["Data Entry Help", "Create a new named profile using the current settings."]
tipMessage['imgSaveProfile']=["Data Entry Help", "Save the current settings"]
tipMessage['imgDeleteProfile']=["Data Entry Help", "Delete the currently selected named profile"]

tipMessage['aUpload']=["Data Entry Help", "Use this to upload a file from your local computer to your server. The upload will initially be set to your database directory"]

//loadProfiles();

tipMessage['spanTargetField']=["Data Entry Help", "This is the product data item to be imported. Refer to the StoreFront documentation for a description of each field."]
tipMessage['spanSourceField']=["Data Entry Help", "This is the column name found in your data source. Only <b>Product Code</b> is required."]
tipMessage['spanDefaultValue']=["Data Entry Help", "This is the default value if <b>NO</b> corresponding source field is specified. If a source field is specified the value in the data source will be used."]

function autoExpand(theSelect, maxLength, Expand)
{
var plngLength;
var plngDefaultLength = 5;

	return false;
	
	if (Expand)
	{
		if (maxLength > 0)
		{
			plngLength = maxLength;
		}else{
			plngLength = plngDefaultLength;
		}

		if (theSelect.size > plngLength)
		{
			theSelect.size = plngLength;
		}else{
			theSelect.size = theSelect.length;
		}

	}else{
		theSelect.size = theSelect.length;
	}
}

function changeSelectedProfile(theSelect)
{
var theForm = theSelect.form;

	theForm.profileName.value = getSelectText(theSelect);
	if (theSelect.selectedIndex == 0)
	{
		tipMessage['imgDeleteProfile']=["Data Entry Help", "Deletion disabled. You cannot delete the first profile."]
	}else{
		tipMessage['imgDeleteProfile']=["Data Entry Help", "Delete the currently selected named profile."]
	}
	//document.all("imgDeleteProfile").disabled = (theSelect.selectedIndex == 0);

	return false;
}

function disableImportButtons(theItem)
{
	if (theItem.form.btnImport1 != null){theItem.form.btnImport1.disabled = true;}
	if (theItem.form.btnImport2 != null){theItem.form.btnImport2.disabled = true;}
}

function loadProfiles()
{

	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM")
	var nodelist;
	var i;
	
	xmlDoc.async="false"
	if (xmlDoc.load("ProductImportTool_Support/ssImportProducts_Profiles.xml"))
	{
		document.write	("The first XML element in the file contains: ")
		document.write	(xmlDoc.documentElement.childNodes.item(0).text + "<br />")
	
		nodelist = xmlDoc.getElementsByTagName("ProductFields/field")
		document.write	("nodelist.length:" + nodelist.length + "<br />")

		for (i = 0;  i < nodelist.length;  i++)
		{
			document.write	(i + ":" + nodelist(i).attributes.item(0).nodeValue + "<br />")
		}
		
	}else{
		document.write	("Profile Loading failed: ")
	}

	
}

function saveProfileSettings(theImage)
{
var strSelection = theImage.id;
var theForm = document.frmData;
var theSelect = theForm.ProfileID;

	switch (strSelection)
	{
		case "imgOpenProfile":
			theForm.Action.value='OpenProfile';
			theForm.submit();
			break;
		case "imgNewProfile":
			var strProfileName = theForm.profileName.value;
			for (var i = 0;  i < theSelect.options.length;  i++)
			{
				if (theSelect.options[i].text == strProfileName)
				{
					alert("A profile with this name already exists.");
					document.frmData.profileName.select();
					document.frmData.profileName.focus();
					return false;
				}
			}

			theForm.Action.value='NewProfile';
			theForm.submit();
			break;
		case "imgSaveProfile":
			theForm.Action.value='SaveProfile';
			theForm.submit();
			break;
		case "imgDeleteProfile":
			if (theSelect.selectedIndex == 0)
			{
				alert("The first profile cannot be deleted");
				return false;
			}else
			{
				blnConfirm = confirm("Are you sure you wish to delete this profile?.\n\nSelect CANCEL if you do not!");
				if (!blnConfirm){return(false);}
				theForm.Action.value='DeleteProfile';
				theForm.submit();
			}
			break;
	}
	return true;
}

function setPageDefaults()
{
	changeSelectedProfile(document.frmData.ProfileID)
}

function validateImportPage(theForm)
{
// this is here to warn about category/mfg/vend/prod deletion
	var blnConfirm;
	
	var catImportOption = getRadio(theForm.createCat);
	if (catImportOption == <%= cenCreateCat_CreateAndDelete %>)
	{
		blnConfirm = confirm("Are you sure you wish to delete all existing categories?\n\n              Select CANCEL if you do not!");
		if (!blnConfirm)
		{
		return(false);
		}
	}
	
	var mfgImportOption = getRadio(theForm.createMfg);
	if (mfgImportOption == <%= cenCreateMfg_CreateAndDelete %>)
	{
		blnConfirm = confirm("Are you sure you wish to delete all existing manufacturers?\n\n              Select CANCEL if you do not!");
		if (!blnConfirm)
		{
		return(false);
		}
	}
	
	var vendImportOption = getRadio(theForm.createVend);
	if (vendImportOption == <%= cenCreateVend_CreateAndDelete %>)
	{
		blnConfirm = confirm("Are you sure you wish to delete all existing vendors?\n\n              Select CANCEL if you do not!");
		if (!blnConfirm)
		{
		return(false);
		}
	}
	
	var prodImportOption = getRadio(theForm.ImportType);
	if (prodImportOption == <%= enImportAll_DeleteExisting %>)
	{
		blnConfirm = confirm("Are you sure you wish to delete all existing products? This will affect any existing orders.\n\n              Select CANCEL if you do not!");
		if (!blnConfirm)
		{
		return(false);
		}
	}
	
	if (prodImportOption == <%= enImportDelete %>)
	{
		blnConfirm = confirm("Are you sure you wish to delete these products? This will affect any existing orders.\n\n              Select CANCEL if you do not!");
		if (!blnConfirm)
		{
		return(false);
		}
	}
	
	if (theForm.Action.value == 'Import'){openOutputWindow('ssOutputWindow.asp');}
	return true;
}

//writeToOutputWindow("<h2>test</h2>");
</script>

<xml id="" />
<form name="frmProfiles" id="frmProfiles" action="ssImportProducts.asp" method="POST">
  <input type="hidden" name="Action" id="Action2" value="Save">
  <input type="hidden" name="ProfileXML" id="ProfileXML" value="">
</form>

<form name="frmData" id="frmData" action="ssImportProducts.asp" method="POST" onsubmit="return validateImportPage(this);">
	  <input type="hidden" name="Action" id="Action" value="Import">
<table class="tbl" id="Table1" cellspacing="0" cellpadding="1" border="0">
  <tr>
    <th colspan="2" align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ProductImport/help_ProductImport.htm')" id="btnHelp" name=btnHelp title="Release Version <%= mstrssAddonVersion %>"></div></th>
    </tr>
  <tr>
    <td colspan="2">
<%
Dim mstrProfileName
	mstrProfileName = Request.Form("ProfileName")

    mstrAction = LoadRequestValue("Action")
    mstrProfileID = LoadRequestValue("ProfileID")
	mblnProfileOnly = CBool(LoadRequestValue("ProfileOnly") = "1")

	Call InitializeDefaults(mstrProfileID)
	If Len(LoadRequestValue("defaultProfileToUse")) > 0 Then cstrDefaultProfile = mstrProfileName

	mlngImportType = mlngDefaultImportType
    Select Case mstrAction
        Case "Import", "Import, Import"
			Call LoadFormValues(True)
			Call VerifyConnections(mobjCnn, mstrTableSource)
			mblnValidImport = mblnDSN_SourceExists And mblnSourceTableExists And mblnDSN_TargetExists
            If mblnValidImport Then
				Call DoImport(mobjCnn, mstrTableSource)
			End If
		Case "DeleteProfile":
			Call LoadFormValues(False)
			If deleteXMLProfile(mstrProfileID) Then
				Response.Write "<h4>Profile <i>" & mstrProfileName & "</i> deleted.</h4>"
				Call loadNamedProfiles
			Else
				Response.Write "<h4><font color='red'>Error deleting profile <i>" & mstrProfileName & "</i>.</h4>"
			End If
			Call VerifyConnections(mobjCnn, mstrTableSource)
			mblnValidImport = mblnDSN_SourceExists And mblnSourceTableExists And mblnDSN_TargetExists
		Case "NewProfile":
			Call LoadFormValues(True)
			If newXMLProfile(mstrProfileID, mstrProfileName) Then
				Response.Write "<h4>Profile <i>" & mstrProfileName & "</i> created.</h4>"
				Call loadNamedProfiles
			Else
				Response.Write "<h4><font color='red'>Error creating profile <i>" & mstrProfileName & "</i>.</h4>"
			End If
			Call VerifyConnections(mobjCnn, mstrTableSource)
			mblnValidImport = mblnDSN_SourceExists And mblnSourceTableExists And mblnDSN_TargetExists
		Case "SaveProfile":
			Call LoadFormValues(True)
			If saveXMLProfile(mstrProfileID, mstrProfileName) Then
				Response.Write "<h4>Profile <i>" & mstrProfileName & "</i> saved.</h4>"
				Call loadNamedProfiles
			Else
				Response.Write "<h4><font color='red'>Error saving profile <i>" & mstrProfileName & "</i>.</h4>"
			End If
			Call VerifyConnections(mobjCnn, mstrTableSource)
			mblnValidImport = mblnDSN_SourceExists And mblnSourceTableExists And mblnDSN_TargetExists
        Case "Verify Connections"
			If mblnProfileOnly Then
				Call LoadFormValues(False)
			Else
				Call LoadFormValues(True)
			End If
			Call VerifyConnections(mobjCnn, mstrTableSource)
			mblnValidImport = mblnDSN_SourceExists And mblnSourceTableExists And mblnDSN_TargetExists
        Case Else
			Call VerifyConnections(mobjCnn, mstrTableSource)
			mblnValidImport = mblnDSN_SourceExists And mblnSourceTableExists And mblnDSN_TargetExists
    End Select
	Set mobjXMLProfiles = Nothing
	
	If mblnDSN_SourceExists Then mstrAvailableTables = getAvailableTables
    Call ReleaseObject(mobjCnn)
    
    Response.Write mstrMessage
%>&nbsp;</td>
  </tr>
  <% If Len(Session("ssNextProductIDToImport")) > 0 Then %>
  <tr>
	<td colspan="2">
	  <input type="hidden" name="ssNextProductIDToImport" id="ssNextProductIDToImport" value="<%= Session("ssNextProductIDToImport") %>">
	  <h4 id="needToRefreshMessage">Not all products were imported because of time restrictions.</h4> 
	  This page will automatically resubmit in <strong><span id="spTimeToReload">X</span></strong> seconds. Press the <em>Import</em> button or <a href="" onclick="resubmit();return false;">here</a> to continue immediately.
	  <script language="javascript">
	  var pageRefreshDelay = <%= clngPageRefreshDelay %>;	//time in seconds
	  var scrolled = false;
	  
	  function resubmit()
	  {
		document.frmData.Action.Value='Import';
		document.frmData.submit();
	  }
	  
	  function checkTimer()
	  {
		if (pageRefreshDelay > 0)
		{
			if (!scrolled)
			{
				ScrollToElem("needToRefreshMessage");
			}
			
			pageRefreshDelay = pageRefreshDelay - 1;
			var elem = document.getElementById("spTimeToReload");
			if (elem != null)
			{
				elem.innerText = pageRefreshDelay;
			}
			window.setTimeout("checkTimer();",1000);
			
		}else{
			resubmit();
		}
	  }
	  
	  function window_onScroll()
	  {
		scrolled = true;
	  }
	  
	  checkTimer();
	  </script>
	</td>
  </tr>
  <% End If %>
  
  <tr>
    <td colspan="2">
    <table class="tbl" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" border="1" width="100%" ID="Table4">
  <tr class="tblhdr"><th colspan="3"><strong>Import Profiles</strong></th></tr>
  <tr>
    <td>&nbsp;&nbsp;<label for="ProfileID" onmouseover="showDataEntryTip(this);" onMouseOut="htm();">Saved Profiles</label></td>
    <td>
    <select name="ProfileID" id="ProfileID" onchange="return changeSelectedProfile(this);" title="Select a profile to use." onfocus="autoExpand(this, 0, true)" onblur="autoExpand(this, 0, false)">
    <% For i = 0 To UBound(maryProfiles) %>
    <option value="<%= maryProfiles(i)(0) %>"<% If CStr(maryProfiles(i)(0)) = CStr(mstrProfileID) Then Response.Write " selected" %>><%= maryProfiles(i)(1) %></option>
    <% Next 'i %>
    </select>&nbsp;<input type=checkbox name="defaultProfileToUse" id="defaultProfileToUse" value="1" <%= isChecked(Trim(cstrDefaultProfile) = Trim(mstrProfileID)) %>>&nbsp;<label for="defaultProfileToUse">Default</label><br />
    <input type=text name=profileName id=profileName value="" />
    &nbsp;<img src="images/Open.gif" id="imgOpenProfile" border="0" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onmouseout="htm();" onclick="return saveProfileSettings(this);" />
    &nbsp;<img src="images/New.gif" id="imgNewProfile" border="0" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onmouseout="htm();" onclick="return saveProfileSettings(this);" />
    &nbsp;<img src="images/Save.gif" id="imgSaveProfile" border="0" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onmouseout="htm();" onclick="return saveProfileSettings(this);" />
    &nbsp;<img src="images/Delete.gif" id="imgDeleteProfile" border="0" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onmouseout="htm();" onclick="return saveProfileSettings(this);" />
    </td>
    <td rowspan="4" align="center">
    <input class="butn" type="submit" name="btnVerifyConnections" id="btnVerifyConnections1" value="Verify Connections" onclick="this.form.Action.value='Verify Connections';"><br/>
	<% If mblnValidImport Then %>
		<input class="butn" type="submit" name="btnImport" id="btnImport1" value="Import" onclick="if(validateImportPage(this.form)){this.form.Action.value='Import';return true;} return false;"><br/>
	<% Else %>
	<input type="hidden" name="ProfileOnly" id="ProfileOnly" value="1">
	<% End If 'mblnValidImport %>
    </td>
  </tr>
  <tr>
    <td <% If Not mblnDSN_SourceExists Then Response.Write "style='color:red;'" %>>&nbsp;&nbsp;<label for="DSN_Source" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Source&nbsp;DSN</label></td>
    <td>
      <textarea name="DSN_Source" id="DSN_Source" rows="4" cols="80" onchange="disableImportButtons(this);"><%= mstrDSN_Source %></textarea>
      <a href="" id="aUpload" onclick="openAsset('db'); return false;" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onmouseout="htm();">Upload</a>  
    </td>
  </tr>
  <tr>
    <td <% If Not mblnSourceTableExists Then Response.Write "style='color:red;'" %>>&nbsp;&nbsp;<label for="SourceTable" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Source&nbsp;Table</label></td>
    <td>
      <textarea name="SourceTable" id="SourceTable" rows="1" cols="80" title="<%= mstrSourceTable  %>" onchange="disableImportButtons(this);"><%= mstrSourceTable  %></textarea>
      <% If Len(mstrAvailableTables) > 0 Then %>
      <br /><select name="sourceTableSelect" id="" onchange="this.form.SourceTable.value = getSelectValue(this)"><option>Select a table</option><%= mstrAvailableTables %></select>
      <% End If %>
    </td>
  </tr>
  <tr>
    <td <% If Not mblnDSN_TargetExists Then Response.Write "style='color:red;'" %>>&nbsp;&nbsp;<label for="DSN_Target" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Target&nbsp;DSN</label></td>
    <td><textarea name="DSN_Target" id="DSN_Target" rows="4" cols="80"><%= mstrDSN_Target %></textarea></td>
  </tr>
  </table>
  </td>
  </tr>
<% If mblnValidImport Then %>
  <tr>
    <td colspan="2">
    <table class="tbl" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" border="1" width="100%" ID="Table2">
      <tr class="tblhdr"><th colspan="3">Field Mappings</th></tr>
      <tr class="tblhdr">
        <th onmouseover="stm(tipMessage['spanTargetField'],tipStyle['dataEntry']);" onmouseout="htm();">Target Field</th>
        <th onmouseover="stm(tipMessage['spanSourceField'],tipStyle['dataEntry']);" onmouseout="htm();">Source</th>
        <th onmouseover="stm(tipMessage['spanDefaultValue'],tipStyle['dataEntry']);" onmouseout="htm();">Default</th>
      </tr>
  <tr class="tblhdr"><td colspan="3"><strong>Product table data</strong></td></tr>
  <%
  Dim i,j
  
  For i = 0 To UBound(maryFields)
  %>
  <tr>
    <td nowrap>&nbsp;
    <%
    If Len(maryFields(i)(enDisplayFieldName)) > 0 Then
	%><span title="Target Field <%= maryFields(i)(enTargetFieldName) %>"><%= maryFields(i)(enDisplayFieldName) %></span><%
    Else
	%><span title="Target Field <%= maryFields(i)(enTargetFieldName) %>"><%= maryFields(i)(enTargetFieldName) %></span><%
    End If
    %>
    </td>
    <td>
      <select name="fieldMap<%= i %>" ID="fieldMap<%= i %>">
      <%= GetSelectOptions(marySheetColumns, maryFields(i)(enSourceFieldName)) %>
      </select>
    </td>
    <td><%= writeHTMLFormElement(maryFields(i)(enDisplayType), "40", "fieldDefault" & i, "fieldDefault" & i, maryFields(i)(enDefaultValue), "", "") %></td>
    </tr>
  <%
  Next 'i
  %>
  <tr class="tblhdr"><td colspan="3"><strong>Attributes</strong></td></tr>
  <% For i = 0 To UBound(maryAttributes) %>
  <tr>
    <td nowrap>&nbsp;
    <%
    If Len(maryAttributes(i)(enDisplayFieldName)) > 0 Then
	%><span title="Target Field <%= maryAttributes(i)(enTargetFieldName) %>"><%= maryAttributes(i)(enDisplayFieldName) %></span><%
    Else
	%><span title="Target Field <%= maryAttributes(i)(enTargetFieldName) %>"><%= maryAttributes(i)(enTargetFieldName) %></span><%
    End If
    %>
    </td>
    <td>
      <select name="fieldMapAttributes<%= i %>" ID="fieldMapAttributes<%= i %>">
      <%= GetSelectOptions(marySheetColumns, maryAttributes(i)(enSourceFieldName)) %>
      </select>
    </td>
    <td>
      <select name="fieldAttributeImportStyle<%= i %>" id="fieldAttributeImportStyle<%= i %>">
		<%
		For j = 0 To UBound(maryAttributeSupportingStyles)
			If j = "0" And Len(maryAttributes(i)(enDefaultValue)) = 0 Then
				Response.Write "<option value=" & j & " selected>" & maryAttributeSupportingStyles(j)(0) & "</option>"
			ElseIf CBool(Trim(maryAttributes(i)(enDefaultValue)) = CStr(j)) Then
				Response.Write "<option value=" & j & " selected>" & maryAttributeSupportingStyles(j)(0) & "</option>"
			Else
				Response.Write "<option value=" & j & ">" & maryAttributeSupportingStyles(j)(0) & "</option>"
			End If
		Next 'j
		%>
      </select>
    </td>
    </tr>
  <% Next 'i %>
  <tr>
  <td>&nbsp;</td>
  <td colspan="3"><input type="checkbox" name="AttributesDeleteExisting" id="AttributesDeleteExisting" value="1" <%= isChecked(mblnAttributesDeleteExisting) %>>&nbsp;<label for="AttributesDeleteExisting">Delete existing attributes</label></td>
  </tr>

  <tr class="tblhdr"><td colspan="3"><strong>Category Assignments</strong></td></tr>
  <tr>
    <td>&nbsp;&nbsp;<label for="CategoryColumn" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Category</label></td>
    <td>
      <select name="CategoryColumn" ID="CategoryColumn" title="Name of column for importing categories"><%= GetSelectOptions(marySheetColumns, mstrCategoryColumn) %></select>
    </td>
    <td>
      <input type=text name="DefaultCategoryValue" ID="DefaultCategoryValue" value="<%= mlngDefaultCategoryID %>">
    </td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;<label for="SubcategoryDelimiter" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Sub-category Delimeter</label></td>
    <td colspan=2>
      <input type=text name="SubcategoryDelimiter" id="SubcategoryDelimiter" value="<%= cstrSubcategoryDelimiter %>">
    </td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;<label for="MultipleCategoryDelimiter" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Multiple Category Delimeter</label></td>
    <td colspan=2>
      <input type=text name="MultipleCategoryDelimiter" id="MultipleCategoryDelimiter" value="<%= cstrMultipleCategoryDelimiter %>">
    </td>
  </tr>
  <tr>
  <td>&nbsp;</td>
  <td colspan="3"><input type="checkbox" name="CategoriesDeleteExisting" id="CategoriesDeleteExisting" value="1" <%= isChecked(mblnCategoriesDeleteExisting) %>>&nbsp;<label for="CategoriesDeleteExisting">Delete existing category assignments</label></td>
  </tr>
  <tr>
  <td>&nbsp;</td>
  <td colspan="3"><input type="checkbox" name="AssignProductsToAllCategoryLevels" id="AssignProductsToAllCategoryLevels" value="1" <%= isChecked(mblnAssignProductsToAllCategoryLevels) %>>&nbsp;<label for="AssignProductsToAllCategoryLevels">Assign product at all category levels</label></td>
  </tr>
  <% If cblnSF5AE Then %>
  <tr class="tblhdr"><td colspan="3"><strong>Gift Wrap</strong></td></tr>
  <% For i = 0 To UBound(maryGiftWrap) %>
  <tr>
    <td nowrap>&nbsp;
    <%
    If Len(maryGiftWrap(i)(enDisplayFieldName)) > 0 Then
	%><span title="Target Field <%= maryGiftWrap(i)(enTargetFieldName) %>"><%= maryGiftWrap(i)(enDisplayFieldName) %></span><%
    Else
	%><span title="Target Field <%= maryGiftWrap(i)(enTargetFieldName) %>"><%= maryGiftWrap(i)(enTargetFieldName) %></span><%
    End If
    %>
    </td>
    <td>
      <select name="fieldMapGiftWrap<%= i %>" ID="fieldMapGiftWrap">
      <%= GetSelectOptions(marySheetColumns, maryGiftWrap(i)(enSourceFieldName)) %>
      </select>
    </td>
    <td><%= writeHTMLFormElement(maryGiftWrap(i)(enDisplayType), "40", "fieldDefaultGiftWrap" & i, "fieldDefaultGiftWrap" & i, maryGiftWrap(i)(enDefaultValue), "", "") %></td>
    </tr>
  <% Next 'i %>

  <tr class="tblhdr"><td colspan="3"><strong>Inventory</strong></td></tr>
  <% For i = 0 To UBound(maryInventoryFields) %>
  <tr>
    <td nowrap>&nbsp;
    <%
    If Len(maryInventoryFields(i)(enDisplayFieldName)) > 0 Then
	%><span title="Target Field <%= maryInventoryFields(i)(enTargetFieldName) %>"><%= maryInventoryFields(i)(enDisplayFieldName) %></span><%
    Else
	%><span title="Target Field <%= maryInventoryFields(i)(enTargetFieldName) %>"><%= maryInventoryFields(i)(enTargetFieldName) %></span><%
    End If
    %>
    </td>
    <td>
      <select name="fieldMapInventory<%= i %>" ID="fieldMapInventory<%= i %>">
      <%= GetSelectOptions(marySheetColumns, maryInventoryFields(i)(enSourceFieldName)) %>
      </select>
    </td>
    <td><%= writeHTMLFormElement(maryInventoryFields(i)(enDisplayType), "40", "fieldDefaultInventory" & i, "fieldDefaultInventory" & i, maryInventoryFields(i)(enDefaultValue), "", "") %></td>
    </tr>
  <% Next 'i %>

  <% Call DisplayCustomImportRoutine %>
  
  <%
	If hasMTP Then
  %>
  <tr class="tblhdr"><td colspan="3"><strong>Volume Pricing</strong></td></tr>
  <tr>
    <td colspan="3" align="left">
    <table class="tbl" style="border-collapse: collapse" bordercolor="#111111" cellpadding="2" cellspacing="0" border="1" ID="Table3">
      <tr>
        <th>Column</th>
        <th>Qty (or Discount)</th>
        <th>Discount Type</th>
        <th>Qty or Discount</th>
      </tr>
  <% For i = 0 To UBound(maryMTPS) %>
  <tr>
    <td><select name="MTPColumn" ID="MTPColumn" title="Name of column for identifying MTP."><%= GetSelectOptions(marySheetColumns, maryMTPS(i)(4)) %></select></td>
    <td align="center"><%= maryMTPS(i)(1) %></td>
    <td align="center"><%= maryMTPS(i)(3) %></td>
    <td align="center"><%= maryMTPS(i)(5) %></td>
  </tr>
  <% Next 'i %>
  <tr>
    <td>&nbsp;<span title="This items is used to identify the prefix which identifies a column header as being a tiered pricing column">Column Prefix</span></td>
    <td colspan=3>
      <input type=text name="MTPImportPrefix" id="MTPImportPrefix" value="<%= cstrMTPImportPrefix %>">
    </td>
  </tr>
  <tr>
    <td>&nbsp;<span title="This items is used to separate entries in the MTP column header">Column Delimeter</span></td>
    <td colspan=3>
      <input type=text name="MTPImportSeparator" id="MTPImportSeparator" value="<%= cstrMTPImportSeparator %>">
    </td>
  </tr>
  <tr>
  <td>&nbsp;</td>
  <td colspan="3"><input type="checkbox" name="DeleteExistingMTPs" id="DeleteExistingMTPs" value="1" <%= isChecked(mblnDeleteExistingMTPs) %>>&nbsp;<label for="DeleteExistingMTPs">Delete the existing volume pricing schedule</label></td>
  </tr>
    </table>
    </td>
  </tr>
  <% End If	'mblnHasMTP %>

  <% End If 'cblnSF5AE %>
  <%  
	If InitializeCustomFields Then
  %>
  <tr class="tblhdr"><td colspan="3"><strong>Custom Import</strong></td></tr>
  <%
  For i = 0 To UBound(maryCustomFields)
  %>
  <tr>
    <td nowrap>&nbsp;
    <%
    If Len(maryCustomFields(i)(enDisplayFieldName)) > 0 Then
	%><span title="Target Field <%= maryCustomFields(i)(enTargetFieldName) %>"><%= maryCustomFields(i)(enDisplayFieldName) %></span><%
    Else
	%><span title="Target Field <%= maryCustomFields(i)(enTargetFieldName) %>"><%= maryCustomFields(i)(enTargetFieldName) %></span><%
    End If
    %>
    </td>
    <td>
      <select name="customField<%= i %>" id="customField<%= i %>">
      <%= GetSelectOptions(marySheetColumns, maryCustomFields(i)(enSourceFieldName)) %>
      </select>
    </td>
    <td><%= writeHTMLFormElement(maryCustomFields(i)(enDisplayType), "40", "customFieldDefault" & i, "customFieldDefault" & i, maryCustomFields(i)(enDefaultValue), "", "") %></td>
  </tr>
  <%
  Next 'i
	End If	'InitializeCustomFields
  %>  
  <tr class="tblhdr"><td colspan="3"><strong>Import Type</strong></td></tr>
  <tr>
    <td>&nbsp;&nbsp;Import Type</td>
    <td colspan=2>
      <select name="ImportTypeColumn" ID="ImportTypeColumn" title="Name of column to use for import type. Note this is only used if the import type Use Data Source is selected"><%= GetSelectOptions(marySheetColumns, mstrImportTypeColumn) %></select>
    </td>
  </tr>

    </table>
    </td></tr>
<% End If 'mblnValidImport %>
  <tr>
    <td colspan="3" align="left">
  <table class="tbl" style="border-collapse: collapse" bordercolor="#111111" cellpadding="2" cellspacing="0" border="1" width="100%" id="Table5">
  <tr class="tblhdr"><td><strong>Import Options</strong></td></tr>
  <tr>
    <td>
      <strong>Category</strong><br />
      &nbsp;&nbsp;<input type="radio" name="createCat" id="createCat0" value="<%= cenCreateCat_Default %>" <% If CLng(mbytCreateCat) = cenCreateCat_Default Then Response.Write "checked" %>>&nbsp;<label for="createCat0">Use default category</label><br />
      &nbsp;&nbsp;<input type="radio" name="createCat" id="createCat1" value="<%= cenCreateCat_Create %>" <% If CLng(mbytCreateCat) = cenCreateCat_Create Then Response.Write "checked" %>>&nbsp;<label for="createCat1">Automatically create categories</label><br />
      &nbsp;&nbsp;<input type="radio" name="createCat" id="createCat2" value="<%= cenCreateCat_CreateAndDelete %>" <% If CLng(mbytCreateCat) = cenCreateCat_CreateAndDelete Then Response.Write "checked" %>>&nbsp;<label for="createCat2">Automatically create categories - delete existing categories</label><br />
      
      <strong>Manufacturer</strong><br />
      &nbsp;&nbsp;<input type="radio" name="createMfg" id="createMfg0" value="<%= cenCreateMfg_Default %>" <% If CLng(mbytCreateMfg) = cenCreateMfg_Default Then Response.Write "checked" %>>&nbsp;<label for="createMfg0">Use default manufacturer</label><br />
      &nbsp;&nbsp;<input type="radio" name="createMfg" id="createMfg1" value="<%= cenCreateMfg_Create %>" <% If CLng(mbytCreateMfg) = cenCreateMfg_Create Then Response.Write "checked" %>>&nbsp;<label for="createMfg1">Automatically create manufacturers</label><br />
      &nbsp;&nbsp;<input type="radio" name="createMfg" id="createMfg2" value="<%= cenCreateMfg_CreateAndDelete %>" <% If CLng(mbytCreateMfg) = cenCreateMfg_CreateAndDelete Then Response.Write "checked" %>>&nbsp;<label for="createMfg2">Automatically create manufacturers - delete existing manufacturers</label><br />
      
      <strong>Vendor</strong><br />
      &nbsp;&nbsp;<input type="radio" name="createVend" id="createVend0" value="<%= cenCreateVend_Default %>" <% If CLng(mbytCreateVend) = cenCreateVend_Default Then Response.Write "checked" %>>&nbsp;<label for="createVend0">Use default vendor</label><br />
      &nbsp;&nbsp;<input type="radio" name="createVend" id="createVend1" value="<%= cenCreateVend_Create %>" <% If CLng(mbytCreateVend) = cenCreateVend_Create Then Response.Write "checked" %>>&nbsp;<label for="createVend1">Automatically create vendors</label><br />
      &nbsp;&nbsp;<input type="radio" name="createVend" id="createVend2" value="<%= cenCreateVend_CreateAndDelete %>" <% If CLng(mbytCreateVend) = cenCreateVend_CreateAndDelete Then Response.Write "checked" %>>&nbsp;<label for="createVend2">Automatically create vendors - delete existing vendors</label><br />
    </td>
  </tr>
  <tr>
    <td>
      <strong>Import Type</strong><br />
	  <% For i = 0 To UBound(maryImportTypes) %>
      &nbsp;&nbsp;<input type="radio" name="ImportType" id="ImportType<%= i %>" value="<%= i %>" <%= isChecked(CLng(mlngImportType) = i) %>>&nbsp;<label for="ImportType<%= i %>" title="<%= maryImportTypes(i)(1) %>"><%= maryImportTypes(i)(0) %></label><br />
	  <% Next 'i %>
    </td>
  </tr>
  <tr>
    <td>
      <strong>Status Display Type</strong><br />
	  <% For i = 0 To UBound(maryDebugTypes) %>
      &nbsp;&nbsp;<input type="radio" name="DebugType" id="DebugType<%= i %>" value="<%= i %>" <% If CLng(mlngDebugType) = i Then Response.Write "checked" %>>&nbsp;<label for="DebugType<%= i %>" title="<%= maryDebugTypes(i)(1) %>"><%= maryDebugTypes(i)(0) %></label><br />
	  <% Next 'i %>
    </td>
  </tr>
  <tr>
    <td>
    <input class="butn" type="submit" name="btnVerifyConnections" id="btnVerifyConnections2" value="Verify Connections" onclick="this.form.Action.value='Verify Connections';">
<% If mblnValidImport Then %>
    <input class="butn" type="submit" name="btnImport" id="btnImport2" value="Import" onclick="if(validateImportPage(this.form)){this.form.Action.value='Import';return true;} return false;">
<% Else %>
<% End If 'mblnValidImport %>
   </td>
  </tr>
  </table>
	</td>
  </tr>
  </table>
</form>
<!--#include file="adminFooter.asp"-->
</center>
</body>
</html>

<%

If Response.Buffer Then Response.Flush

'**********************************************************
'	Developer notes
'
'	The code beneath this section was created under contract - It is provided here as an example of what can be done
'   It is provided as is with no guarantee. It is not supported.
'
'	saveCustom_??? references functions in Common/ssProducts_Custom.asp
'
'**********************************************************

'**********************************************************
'*	Global variables
'**********************************************************

Dim maryCustomFields
Const cstrCustomFieldDelimiter = ","
Const cstrCustomFieldDelimiter_Multiple = "|"

'**********************************************************
'*	Functions
'**********************************************************

'Sub DisplayCustomImportRoutine
'Function setProductColor(byVal lngProductUID, byVal strProductColor, byVal strImprintColor)
'Sub saveCustom_orig(byRef objrsProducts, byVal lngProductUID)
'Sub saveCustom(byRef objrsProducts, byVal lngProductUID)
'Function InitializeCustomFields

'**************************************************************************************************************************************************

Sub DisplayCustomImportRoutine

	If Not InitializeCustomFields Then Exit Sub
%>
  <tr><td colspan="3"><strong>Custom Import</strong></td></tr>
<%
  For i = 0 To UBound(maryCustomFields)
  %>
  <tr>
    <td nowrap>&nbsp;
    <%
    If Len(maryCustomFields(i)(enDisplayFieldName)) > 0 Then
	%><span title="Target Field <%= maryCustomFields(i)(enTargetFieldName) %>"><%= maryCustomFields(i)(enDisplayFieldName) %></span><%
    Else
	%><span title="Target Field <%= maryCustomFields(i)(enTargetFieldName) %>"><%= maryCustomFields(i)(enTargetFieldName) %></span><%
    End If
    %>
    </td>
    <td>
      <select name="customField<%= i %>" id="Select1">
      <%= GetSelectOptions(marySheetColumns, maryCustomFields(i)(enSourceFieldName)) %>
      </select>
    </td>
    <td><%= writeHTMLFormElement(maryCustomFields(i)(enDisplayType), "40", "customFieldDefault" & i, "customFieldDefault" & i, maryCustomFields(i)(enDefaultValue), "", "") %></td>
    </tr>
  <%
  Next 'i
  
End Sub	'DisplayCustomImportRoutine

'**************************************************************************************************************************************************

Function InitializeCustomFields

	InitializeCustomFields = False
	Exit Function

	If Not isArray(maryCustomFields) Then
		ReDim maryCustomFields(0)
		maryCustomFields(0) = Array("ProductColor","ProductColors","Product Color","",enDatatype_string,enDisplayType_textbox)
	End If
	
	InitializeCustomFields = isArray(maryCustomFields)
	
End Function	'InitializeCustomFields

'**************************************************************************************************************************************************

Sub saveCustom(byRef objrsProducts, byVal lngProductUID)

Dim pstrColor
Dim pstrImprintColor
Dim paryColor
Dim paryImprintColor
Dim i
Dim j
Dim pblnLocalResult

	If Len(lngProductUID) = 0 Then Exit Sub
	If Not InitializeCustomFields Then Exit Sub

	If CBool(mlngImportType = enImportUpdateSelectedFieldsOnly) AND CBool((Len(maryCustomFields(0)(enSourceFieldName)) = 0) OR (Len(maryCustomFields(0)(enSourceFieldName)) = 0))  Then Exit Sub

	pstrImprintColor = getValueFrom(maryCustomFields, objrsProducts, 0, False, "")
	If Len(pstrImprintColor) = 0 Then Exit Sub
	
	paryImprintColor = Split(pstrImprintColor, cstrCustomFieldDelimiter_Multiple)

	pblnLocalResult = True
	For i = 0 To UBound(paryImprintColor)
		pstrColor = paryImprintColor(i)
		paryColor = Split(pstrColor, cstrCustomFieldDelimiter)
		For j = 1 To UBound (paryColor)
			pblnLocalResult = pblnLocalResult And setProductColor(lngProductUID, paryColor(0), paryColor(j))
		Next 'j
	Next 'i

	
	If pblnLocalResult Then
		WriteOutput "&nbsp;&nbsp;&nbsp;Updated Product Colors<br />" & vbcrlf
	Else
		WriteOutput "<font color=red>&nbsp;&nbsp;&nbsp;Error updating Product Colors</font><br />" & vbcrlf
	End If
		
End Sub	'saveCustom

'**************************************************************************************************************************************************

Sub saveCustom_orig(byRef objrsProducts, byVal lngProductUID)

Dim pstrColor
Dim pstrImprintColor
Dim paryColor
Dim paryImprintColor
Dim i
Dim j
Dim k
Dim pblnLocalResult

	If Len(lngProductUID) = 0 Then Exit Sub
	If Not InitializeCustomFields Then Exit Sub

	If CBool(mlngImportType = enImportUpdateSelectedFieldsOnly) AND CBool((Len(maryGiftWrap(0)(enSourceFieldName)) = 0) OR (Len(maryGiftWrap(1)(enSourceFieldName)) = 0))  Then Exit Sub

	pstrColor = getValueFrom(maryCustomFields, objrsProducts, 0, False, "")
	paryColor = Split(pstrColor, cstrCustomFieldDelimiter)
	If Len(pstrColor) = 0 Then paryColor = Array("")
	
	pstrImprintColor = getValueFrom(maryCustomFields, objrsProducts, 1, False, "")
	paryImprintColor = Split(pstrImprintColor, cstrCustomFieldDelimiter)
	If Len(pstrImprintColor) = 0 Then paryImprintColor = Array("")

	k = 1
	pblnLocalResult = True
	For i = 0 To UBound(paryColor)
		For j = 0 To UBound (paryImprintColor)
			pblnLocalResult = pblnLocalResult And setProductColor(lngProductUID, paryColor(i), paryImprintColor(j))
			k = k + 1
		Next 'j
	Next 'i

	If pblnLocalResult Then
		WriteOutput "&nbsp;&nbsp;&nbsp;Updated Product Colors<br />" & vbcrlf
	Else
		WriteOutput "<font color=red>&nbsp;&nbsp;&nbsp;Error updating Product Colors</font><br />" & vbcrlf
	End If
		
End Sub	'saveCustom_orig

'****************************************************************************************************************************************************************

Function setProductColor(byVal lngProductUID, byVal strProductColor, byVal strImprintColor)

Dim pstrSQL
Dim pobjRS
Dim pbytIsActive
Dim pstrPrice

	'Check Data
	If CBool(Len(strProductColor) = 0 And Len(strImprintColor) = 0) Or Len(lngProductUID) = 0 Then
		setProductColor = False
		Exit Function
	End If
	
	pstrSQL = "Select uid From ProductColor" _
			& " Where ProductID=" & wrapSQLValue(lngProductUID, False, enDatatype_number) _
			& "  AND ProductColor = " & wrapSQLValue(strProductColor, False, enDatatype_string) _
			& "  AND ImprintColor = " & wrapSQLValue(strImprintColor, False, enDatatype_string)
	'debugprint "pstrSQL", pstrSQL
	Set pobjRS = GetRS(pstrSQL)
	If pobjRS.EOF Then
		pstrSQL = "Insert Into ProductColor (ProductID, ProductColor, ImprintColor) Values (" _
				& wrapSQLValue(lngProductUID, False, enDatatype_number) & ", " _
				& wrapSQLValue(strProductColor, False, enDatatype_string) & ", " _
				& wrapSQLValue(strImprintColor, False, enDatatype_string) & ")"
	Else
		pstrSQL = "Update ProductColor Set" _
				& " ProductColor = " & wrapSQLValue(strProductColor, False, enDatatype_string) & ", " _
				& " ImprintColor = " & wrapSQLValue(strImprintColor, False, enDatatype_string) & " " _
				& " Where uid=" & wrapSQLValue(pobjRS.Fields("uid").Value, False, enDatatype_number)
				
		'this is blanked out since you wouldn't update this since it is by definition already there
		pstrSQL = ""
	End If
	pobjRS.Close
	Set pobjRS = Nothing
	
	If Len(pstrSQL) > 0 Then
		cnn.Execute pstrSQL,,128
		setProductColor = CBool(Err.number = 0)
	Else
		setProductColor = True
	End If
		
End Function	'setProductColor

'**************************************************************************************************************************************************

%>