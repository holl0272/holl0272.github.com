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

Const enTargetFieldName = 0
Const enSourceFieldName = 1
Const enDisplayFieldName = 2
Const enDefaultValue = 3
Const enFieldDataType = 4
Const enDisplayType = 5

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

Dim marySourceTables
Dim maryTargetTables
Dim mstrAvailableTables
Dim mstrProfileName

'**********************************************************
'*	Functions
'**********************************************************

'Function checkForCorruptData(byVal strData)
'Function checkForSubCategories(byVal strValue, byRef dicCategories, byVal pblnLocalDebug, byRef pstrLocalDebugOut)
'Function checkReplacements(byRef strSource, byRef strProdID)
'Function Counter(ByRef lngCounter)
'Function customBoolean(byVal strValue)
'Function deleteXMLProfile(byVal strProfileID)
'Function dbFieldName(byVal strSpreadSheetFieldName)
'Sub DoImport(objcnnSource, strTableSource)
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
'Sub InitializeDefaults(byRef strProfileID)
'Sub LoadFormValues(byVal blnComplete)
'Sub loadNamedProfiles()
'Function loadSettingsFromProfile(byRef objXMLProfile)
'Function LoadXMLProfiles(byVal strProfileID, byRef XMLDoc)
'Function newXMLProfile(byVal strProfileID, byVal strProfileName)
'Function OpenTableSQL(ByVal strTableSource)
'Function profilePath()
'Sub RecordTime(byVal strMessage, byVal dtStartTime, byRef dtCurrentTime, byRef dtLastTime)
'Sub SaveFailedImport(strProductID, strProductName)
'Function saveXMLProfile(byVal strProfileID, byVal strProfileName)
'Sub SetColumns(byRef objRS)
'Sub setProductDetail(byRef objXMLElement, byRef aryProductDetail)
'Function setInventoryLevels(byRef objrsProducts, byVal aryProduct, byRef blnInventoryUpdateOnly)
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

Function checkReplacements(byRef strSource, byRef strProdID)

Dim pstrTemp

	pstrTemp = Trim(strSource & "")
	pstrTemp = Replace(pstrTemp, "<prodID>", strProdID)
	pstrTemp = Replace(pstrTemp, "<code>", strProdID)
	pstrTemp = Replace(pstrTemp, "<Date()>", Date())
	
	checkReplacements = pstrTemp

End Function	'checkReplacements

'****************************************************************************************************************************************************************

Function Counter(ByRef lngCounter)
	lngCounter = lngCounter + 1
	Counter = lngCounter
End Function	'Counter

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

	'Check to see if this is resuming a previous import
	If Len(Session("ssNextProductIDToImport")) = 0 Then

	Else
		'move to next product to import
		pstrProductID = Trim(Session("ssNextProductIDToImport"))
		
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
	For i = 1 To UBound(maryFields)
		If Len(pstrSQLInsert_Part1) = 0 Then
			pstrSQLInsert_Part1 = "Insert Into " & mstrTargetTable & " (" & WrapFieldName(maryFields(i)(enTargetFieldName))
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
		
		pstrProductID = " "
		pstrProductName = ""
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
			WriteOutput "<font size='-1'>(" & plngCurrentRecord + 1 & " of " & plngRecordsToBeImported & ")</font><strong> Importing " & pstrProductID & " . . .</strong><br />" & vbcrlf
			
			'Clear the variables
			pstrSQLInsert_Values = ""
			pstrSQLUpdate = ""
			pblnAlreadyExists = False
			
			For i = 1 To UBound(maryFields)
				pblnRequiredValueSet = True
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

				End If	'Len(maryFields(i)(enSourceFieldName)) = 0
				
				On error goto 0
				'Check for textboxs and defaults
				If CBool(maryFields(i)(enDisplayType)=enDisplayType_checkbox) And CBool(maryFields(i)(enFieldDataType)=enDatatype_boolean) Then
					pstrFieldValue = customBoolean(pstrFieldValue)
					'Response.Write maryFields(i)(enSourceFieldName) & " (" & maryFields(i)(enTargetFieldName) & "): " & pstrFieldValue & "<br />"
				End If

				'Response.Write maryFields(i)(enSourceFieldName) & " (" & maryFields(i)(enTargetFieldName) & "): " & pstrFieldValue & "<br />"
				'Check for some common import errors
				If Len(pstrSQLInsert_Values) = 0 Then
					pstrSQLInsert_Values = "(" & wrapSQLValue(pstrFieldValue, True, maryFields(i)(enFieldDataType))
				Else
					pstrSQLInsert_Values = pstrSQLInsert_Values & ", " & wrapSQLValue(pstrFieldValue, True, maryFields(i)(enFieldDataType))
				End If
				'Response.Write "pstrSQLInsert_Values: " & pstrSQLInsert_Values & "<br />"
				
				If (mlngImportType = enImportUpdateSelectedFieldsOnly AND Len(maryFields(i)(enSourceFieldName)) > 0) Or (mlngImportType <> enImportUpdateSelectedFieldsOnly) Then
					If Not pblnRequiredValueSet Then
						If Len(pstrSQLUpdate) = 0 Then
							pstrSQLUpdate = makeSQLUpdate(maryFields(i)(enTargetFieldName),pstrFieldValue, True, maryFields(i)(enFieldDataType))
						Else
							pstrSQLUpdate = pstrSQLUpdate & ", " & makeSQLUpdate(maryFields(i)(enTargetFieldName),pstrFieldValue, True, maryFields(i)(enFieldDataType))
						End If
					End If
				End If
				
			Next 'i
			pstrSQLInsert_Values = pstrSQLInsert_Values & ")"
			pstrSQL = pstrSQLInsert_Part1 & pstrSQLInsert_Values
			
			plngProductUID = -1	'getProductUIDByCode(pstrProductID)			
			pblnAlreadyExists = CBool(plngProductUID <> -1)
			'debugprint "mlngImportType",mlngImportType
			If Len(pstrSQLUpdate) > 0 Then pstrSQLUpdate = "Update sfProducts Set " & pstrSQLUpdate & " Where sfProductID=" & wrapSQLValue(plngProductUID, True, enDatatype_number)

			If False Then
				Response.Write "<fieldset><legend></legend>"
				Response.Write "pstrSQL: " & pstrSQL & "<br />"
				Response.Write "pstrSQLInsert_Values: " & pstrSQLInsert_Values & "<br />"
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
				Session("ssNextProductIDToImport") = plngCurrentRecord
			Else
				Session("ssNextProductIDToImport") = ""
			End If
			If CBool(Now() > pdtMustFinishBy) Then
				pblnAborted = True
				Exit Do
			End If
		Else
			'this check added to handle database connection timeouts
			WriteOutput "<font color=red>Unexpected error. Attempting to restart with product <i>" & Trim(pobjRSSource.Fields(maryFields(0)(enSourceFieldName)).Value & "") & "</i></font><br />" & vbcrlf
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
		Session.Contents.Remove("ssFailedImports")
	End If
	
End Sub	'DoImport

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

'**************************************************************************************************************************************************

Function GetSelectOptions(aryData, vntValue)

Dim pstrOut
Dim pblnFound
Dim j

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
		pstrOut = pstrOut & "  <option></option>" & vbcrlf
	Else
		pstrOut = pstrOut & "  <option selected></option>" & vbcrlf
	End If
	
	GetSelectOptions = pstrOut
 
End Function	'GetSelectOptions

'**************************************************************************************************************************************************

Function getAvailableTables

Dim pstrOut

	marySourceTables = getAvailableTables_Array(True)
	maryTargetTables = getAvailableTables_Array(False)
	For i = 0 To UBound(marySourceTables)
		pstrOut = pstrOut & "<option value=""" & marySourceTables(i)(0) & """>" & marySourceTables(i)(1) & "</option>"
	Next 'i

	getAvailableTables = pstrOut

End Function	'getAvailableTables

'**************************************************************************************************************************************************

Function getAvailableTables_Array(byVal blnSource)

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

	If blnSource Then
		Set objSchema = mobjCnn.OpenSchema(20) 
	Else
		Set objSchema = cnn.OpenSchema(20) 
	End If
	'adSchemaTables = 20
	'adSchemaPrimaryKeys = 28

	plngCounter = -1
	Do Until objSchema.EOF
		'debugprint "TABLE_TYPE", objSchema("TABLE_TYPE")
		pstrType = objSchema("TABLE_TYPE")
		Select Case UCase(objSchema("TABLE_TYPE"))
			Case "TABLE"
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
			Case "VIEW"
				If blnSource Then
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
				End If	'blnSource
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
	mstrTargetTable = Trim(LoadRequestValue("TargetTable"))
	
	'Remove any extra carriage returns
	mstrDSN_Source = Replace(mstrDSN_Source, vbcrlf, "")
	mstrDSN_Target = Replace(mstrDSN_Target, vbcrlf, "")
	
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

				'Import Options
				pobjProfileNode.selectSingleNode("ImportOptions/CreateCat").Text = mbytCreateCat
				pobjProfileNode.selectSingleNode("ImportOptions/CreateMfg").Text = mbytCreateMfg
				pobjProfileNode.selectSingleNode("ImportOptions/CreateVend").Text = mbytCreateVend
				pobjProfileNode.selectSingleNode("ImportOptions/DefaultImportType").Text = mlngDefaultImportType
				pobjProfileNode.selectSingleNode("ImportOptions/DefaultCategoryID").Text = mlngDefaultCategoryID
				pobjProfileNode.selectSingleNode("ImportOptions/DefaultManufacturerID").Text = mlngDefaultManufacturerID
				pobjProfileNode.selectSingleNode("ImportOptions/DefaultVendorID").Text = mlngDefaultVendorID
				
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

End Sub	'SetColumns

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
	If Len(mstrDSN_Target) = 0 Then mstrDSN_Target = connectionString
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

Sub LoadFieldData()

Dim i
Dim pobjRS

	Set pobjRS = GetRS(mstrTargetTable)
	If pobjRS.State = 1 Then
		ReDim maryFields(pobjRS.Fields.Count - 1)
		For i = 0 to pobjRS.Fields.Count - 1
		
			maryFields(i) = Array(enTargetFieldName, enSourceFieldName, enDisplayFieldName, enDefaultValue, enFieldDataType, enDisplayType)

			maryFields(i)(enTargetFieldName) = pobjRS.Fields(i).Name
			maryFields(i)(enSourceFieldName) = ""
			maryFields(i)(enDisplayFieldName) = pobjRS.Fields(i).Name
			maryFields(i)(enDefaultValue) = ""
			
			Select Case pobjRS.Fields(i).Type
				Case 11:
					maryFields(i)(enFieldDataType) = enDatatype_boolean
				Case 3:
					maryFields(i)(enFieldDataType) = enDatatype_number
				Case 129, 130, 201, 203:
					maryFields(i)(enFieldDataType) = enDatatype_string
				Case 7, 135:
					maryFields(i)(enFieldDataType) = enDatatype_date
				Case Else
					Response.Write "<strong>The following field type has not been defined in <em>LoadFieldData</em>: " & pobjRS.Fields(i).Name & ": " & pobjRS.Fields(i).Type & "<br />"
			End Select
	
			'maryFields(i)(enFieldDataType) = CLng(pvntTemp)
			maryFields(i)(enDisplayType) = enDisplayType_textbox

		Next 'i
	End If
	Call ReleaseObject(pobjRS)

End Sub	'LoadFieldData

'****************************************************************************************************************************************************************
'****************************************************************************************************************************************************************

Dim mblnDSN_SourceExists
Dim mstrDSN_Source
Dim mblnSourceTableExists
Dim mstrSourceTable
Dim mstrTargetTable

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

	mstrPageTitle  = "Item Import - Beta"
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
	
	if (theForm.Action.value == 'Import'){openOutputWindow('ssOutputWindow.asp');}
	return true;
}

//writeToOutputWindow("<h2>test</h2>");
</script>

<xml id="" />
<form name="frmProfiles" id="frmProfiles" action="ssImportUtility.asp" method="POST">
  <input type="hidden" name="Action" id="Action2" value="Save">
  <input type="hidden" name="ProfileXML" id="ProfileXML" value="">
</form>

<form name="frmData" id="frmData" action="ssImportUtility.asp" method="POST" onsubmit="return validateImportPage(this);">
	  <input type="hidden" name="Action" id="Action" value="Import">
<table class="tbl" cellspacing="0" cellpadding="1" border="0">
  <tr>
    <th colspan="2" align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ProductImport/help_ProductImport.htm')" id="btnHelp" name=btnHelp title="Release Version <%= mstrssAddonVersion %>"></div></th>
    </tr>
  <tr>
    <td colspan="2">
<%

	mstrProfileName = Request.Form("ProfileName")

    mstrAction = LoadRequestValue("Action")
    mstrProfileID = LoadRequestValue("ProfileID")
	mblnProfileOnly = CBool(LoadRequestValue("ProfileOnly") = "1")

	Call InitializeDefaults(mstrProfileID)
	If Len(LoadRequestValue("defaultProfileToUse")) > 0 Then cstrDefaultProfile = mstrProfileName

	'Need to load everything
	Call LoadFormValues(True)
	Call VerifyConnections(mobjCnn, mstrTableSource)
	
	If mblnDSN_SourceExists Then
		mstrAvailableTables = getAvailableTables
		Call LoadFieldData
	ElseIf mblnDSN_TargetExists Then
		maryTargetTables = getAvailableTables_Array(False)
	End If
	
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
			mblnValidImport = mblnDSN_SourceExists And mblnSourceTableExists And mblnDSN_TargetExists
    End Select
	Set mobjXMLProfiles = Nothing
	
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
    <td <% If Not mblnDSN_TargetExists Then Response.Write "style='color:red;'" %>>&nbsp;&nbsp;<label for="DSN_Target" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Target&nbsp;DSN</label></td>
    <td><textarea name="DSN_Target" id="DSN_Target" rows="4" cols="80"><%= mstrDSN_Target %></textarea></td>
  </tr>
  <tr>
    <td <% If Not mblnDSN_TargetExists Then Response.Write "style='color:red;'" %>>&nbsp;&nbsp;<label for="targetTableSelect" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Target&nbsp;Table</label></td>
    <td>
      <% If mblnDSN_TargetExists Then %>
      <br />
      <select name="targetTable" id="targetTable">
      <option>Select a table</option>
      <%= createComboFromArray(maryTargetTables, "", "", mstrTargetTable) %>
      </select>
      <% End If %>
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
      <br /><select name="sourceTableSelect" id="" onchange="this.form.SourceTable.value = getSelectValue(this)"><option>Select a table</option><%= createComboFromArray(marySourceTables, "", "", mstrSourceTable) %></select>
      <% End If %>
    </td>
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
    </table>
    </td></tr>
<% End If 'mblnValidImport %>

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
%>