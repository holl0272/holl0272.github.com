<%
'********************************************************************************
'*   Product Manager Version SF 5.0 					                        *
'*   Release Version:	2.00.002		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		September 11, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Sub LoadUserSettings

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

mblnShowTabs = True					'Set this to true to show tabs
mblnAutoShowTable = True			'Set to true to automatically display database summary
mbytSummaryTableHeight = 500		'Summary Table Height
mblnAutoShowDetailInWindow = 1		'Set this to 1 to show detail in existing window; 0 to open new window
Server.ScriptTimeout = 300			'in seconds. Adjust for large databases or if some products have a lot of attributes. Server Default is usually 90 seconds
mlngShortDescriptionLength = ""		'by default this field is limited to 255 characters. If you desire, you can set it higher; "" is unlimited - this will required a change to the sfProducts table
clngDefaultMaxRecords = 50			'Set your default Maximum Records to show in summary table
mblnExpandAttributesAutomatically = False

cstrDefaultDetailLinkPath = "detail.asp?Product_ID=<prodID>"	'Sets the default text to use for the product detail page link.

'For custom field display
ReDim maryDisplayField(5)
maryDisplayField(0) = Array("prodID", True)	
maryDisplayField(1) = Array("prodName", True)	
maryDisplayField(2) = Array("prodPrice", True)	
maryDisplayField(3) = Array("prodSalePrice", True)	
maryDisplayField(4) = Array("prodDateAdded", True)	
maryDisplayField(5) = Array("prodEnabledIsActive", True)	

'For Use With Pricing Level Manager
mblnAttrPrice = True And cblnAddon_PricingLevelMgr			'Show pricing levels for the attribute price
mblnMTPrice = True And cblnAddon_PricingLevelMgr			'Show pricing levels for the multi-tier price

'For Use With Dynamic Product Display
cblnAddon_DynamicProductDisplay = True				'Show pricing levels for the regular price

'This section is used to enable using multiple images beyond the standard small and large images
ReDim maryImageFields(1)
maryImageFields(0) = Array("Small Image","prodImageSmallPath","","images/<prodID>.jpg","","",enDatatype_string)
maryImageFields(1) = Array("Large Image","prodImageLargePath","","images/<prodID>.jpg","","",enDatatype_string)
'maryImageFields(2) = Array("Huge Image","XLGImage","","detail.asp?Product_ID=<prodID>","","",enDatatype_string)

'This section is only applicable for Attribute Extender enabled sites
maryAttributeTypes = Array( _
							"Select (default)", _
							"Radio", _
							"Text box", _
							"Text box<sup>*</sup>", _
							"Textarea", _
							"Textarea<sup>*</sup>", _
							"Checkbox", _
							"Select (Show price)", _
							"Select (Change Image)", _
							"Select (Update Price)", _
							"Radio (Update Price)", _
							"Qty Box", _
							"Radio Attributes", _
							"Custom", _
							"Custom 1" _
							)

'/
'/////////////////////////////////////////////////

End Sub	'LoadUserSettings

'***********************************************************************************************

Sub InitializeCustomValues(byRef aryCustomProductFields)

	'The line below has been added to disable examples
	Exit Sub
	
	
	'Array Structure		XML Document
	'0) Display Text		displayName
	'1) field name			fieldName
	'2) field value			must be ""
	'3) DisplayType			displayType, must be one of the below values
	'						- enDisplayType_hidden
	'						- enDisplayType_select
	'						- enDisplayType_textarea
	'						- enDisplayType_textbox
	'						- enDisplayType_checkbox
	'						- enDisplayType_listbox
	'						- enDisplayType_textbox_WithDateSelect
	'						- enDisplayType_textbox_WithHTMLSelect
	'4) DisplayLength		displayLength
	'5) sqlSource			sqlSource
	'6) Datatype			dataType, must be one of the below values
	'						- enDatatype_string
	'						- enDatatype_number
	'						- enDatatype_date
	'						- enDatatype_boolean

	ReDim aryCustomProductFields(7)
	aryCustomProductFields(0) = Array("Version","version","",enDisplayType_textbox,"","",enDatatype_string)
	aryCustomProductFields(1) = Array("Release Date","releaseDate","",enDisplayType_textbox,"","", enDatatype_date)
	aryCustomProductFields(2) = Array("Installation Hours","InstallationHours","",enDisplayType_textbox,"","",enDatatype_string)
	aryCustomProductFields(3) = Array("My Product","MyProduct","","checkbox","","",enDatatype_boolean)
	aryCustomProductFields(4) = Array("Installation Required","InstallationRequired","",enDisplayType_checkbox,"","",enDatatype_boolean)
	aryCustomProductFields(5) = Array("Include In Search","IncludeInSearch","",enDisplayType_checkbox,"","",enDatatype_boolean)
	aryCustomProductFields(6) = Array("Include In Random Product","IncludeInRandomProduct","",enDisplayType_checkbox,"","",enDatatype_boolean)
	aryCustomProductFields(7) = Array("Upgrade Notes","UpgradeVersion","",enDisplayType_textbox,"","",enDatatype_boolean)

End Sub	'InitializeCustomValues

'***********************************************************************************************

Function GetDisplayFieldIndex(byVal strFieldName)

Dim i
Dim plngReturn

	For i = 0 To UBound(maryDisplayField)
		If maryDisplayField(i)(0) = strFieldName Then
			plngReturn = i
			Exit For
		End If
	Next 'i
	
	GetDisplayFieldIndex = plngReturn

End Function	'GetDisplayFieldIndex

'***********************************************************************************************

%>
