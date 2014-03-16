<%
'********************************************************************************
'*   Product Import Tool For StoreFront 6.0
'*   Release Version:	1.01.001
'*   Release Date:		August 9, 2003
'*   Revision Date:		November 1, 2003
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.
'********************************************************************************


'****************************************************************************************************************************************************************

'/////////////////////////////////////////////////
'/
'/  This file permits you to set your import settings so you do not need to adjust you import parameters for each import
'/  This file is broken into sections corresponding to the sections you see in your browser:

'Connection information
'Product Table data
'Attributes
'Category Assignments
'Inventory
'Gift Wrap
'Volume Pricing
'Import Options

'Connection information
'This section sets your data source connection; the target connection is set to the database identified in the web.config file
'Note: [MapPath] maps to the current location of this file - ssl\Management\ssAdmin
'      [DBPath] maps to the \db folder
mstrDSN_Source = "[MapPath]\ssSamples\ssProductImportSample.xls"
mstrSourceTable = "Products"

'Product Table data
'This section identifies the fields located in the Product table
	'custom Fields
	'Decoder for products
	'0 - Product field								enTargetFieldName
	'1 - spreadsheet column header					enSourceFieldName
	'2 - display name (optional)					enDisplayFieldName
	'3 - default value if field not in spreadsheet	enDefaultValue
	'4 - field type									enFieldDataType
	'	 - Possible Values
	'	 - enDatatype_string
	'	 - enDatatype_number
	'	 - enDatatype_date
	'	 - enDatatype_boolean
	'5 - display type								enDisplayType
	'	 - Possible Values
	'	 - enDisplayType_hidden = "hidden"
	'	 - enDisplayType_select = "select"
	'	 - enDisplayType_radio = "radio"			'NOT Currently Supported
	'	 - enDisplayType_textarea = "textarea"
	'	 - enDisplayType_textbox = "textbox"
	'	 - enDisplayType_checkbox = "checkbox"
	'	 - enDisplayType_listbox = "listbox"
	
	'If you have additional fields simply add them to the bottom of the list 
	'Make sure you adjust the Dim maryFields() below to reflect the new values
	ReDim maryFields(34)
	i = -1
	maryFields(Counter(i)) = Array("Code","Product Code","Code","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("ManufacturerId","Manufacturer","Manufacturer","1",enDatatype_number,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("VendorId","Vendor","Vendor","1",enDatatype_number,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("IsActive","Activate Product","Activate Product","1",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("Name","Product Name","Name","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("NamePlural","Plural Name","Plural Name","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("ShortDescription","Short Description","Short Description","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Description","Long Description","Long Description","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("UpSellMessage","Confirmation Message","Confirmation Message","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("ImageSmallPath","Small Image","Small Image","images/Products/Smallimages/<prodID>_small.jpg",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("ImageLargePath","Large Image","Large Image","images/Products/Largeimages/<prodID>_large.jpg",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("FileName","Download File Name","Download File Name","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Cost","Cost","Cost","0",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Price","Price","Price","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("IsOnSale","Activate Sale","Activate Sale","",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("SalePrice","Sale Price","Sale Price","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("IsShipable","Ship This Product","Ship This Product","1",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("ShipPrice","Ship Price","Ship Price","0",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Weight","Weight","",0,enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Length","Length","","0",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Width","Width","","0",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Height","Height","","0",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("HasCountryTax","Apply Country Tax","Apply Country Tax","",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("HasStateTax","Apply State Tax","Apply State Tax","1",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("HasLocalTax","Apply Local Tax","Apply Local Tax","",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("DateAdded","DateAdded","DateAdded",Date(),enDatatype_date,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("DateModified","DateModified","DateModified",Date(),enDatatype_date,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Keywords","Keywords","Keywords","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("DetailLink","Detail Link","Detail Link","../detail.aspx?ID=<uid>",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("Inventory_Tracked","Track Inventory","Track Inventory","0",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("DropShip","Ship From Vendor","Ship From Vendor","0",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("DownloadOneTime","Allow Multiple Downloads","Allow Multiple Downloads","0",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("DownloadExpire","Download Expires","Download Expires","",enDatatype_string,enDisplayType_textbox)
	maryFields(Counter(i)) = Array("DealTimeIsActive","Publish To DealTime","Publish To DealTime","0",enDatatype_boolean,enDisplayType_checkbox)
	maryFields(Counter(i)) = Array("MMIsActive","Publish To Marketplace Manager","Publish To Marketplace Manager","0",enDatatype_boolean,enDisplayType_checkbox)


'Attributes
'This section accounts for custom creation/importing of attribute templates
	'Decoder for attributes:
	'0 - Product field								enTargetFieldName
	'1 - spreadsheet column header					enSourceFieldName
	'2 - display name (optional)					enDisplayFieldName
	'3 - default value if field not in spreadsheet	enDefaultValue
	'4 - Attribute Import Style						enFieldDataType
	'	 - Possible Values
	'	 - 0 - cart: Attribute Category||Type||Required|n|Attribute Name||Price||PriceType||Weight||WeightType||AttrOrder||SmallImage||LargeImage||FileLocation|n||n|Start of new attribute . . .
	'			|n||n| - separator between completely new attribute categories
	'			|n| - separator between attributes
	'			|| - separator within attribute fields, either the attribute category or attribute detail
	'	 - Future import styles to be added
	
	i = -1
	ReDim maryAttributes(0)
	maryAttributes(Counter(i)) = Array("N/A","Attributes","Attribute Style",3)

'Category Assignments
'This section sets your default category settings for this profile
	mstrCategoryColumn = "Categories"			'Set to column name if you are importing category information
	cstrMultipleCategoryDelimiter = "|"		'Delimiter to use when importing multiple category assignments
	cstrSubcategoryDelimiter = "~"			'Delimiter to use when importing subcategory structures

'Inventory (AE only)
'The same array decoder for the products section applies here
	i = -1
	ReDim maryInventoryFields(7)
	maryInventoryFields(Counter(i)) = Array("Tracked","Track Inventory","Track Inventory","",enDatatype_boolean,enDisplayType_checkbox)
	maryInventoryFields(Counter(i)) = Array("Status","Show Status","Show Status","",enDatatype_boolean,enDisplayType_checkbox)
	maryInventoryFields(Counter(i)) = Array("Notify","Send Low Stock Notice","Send Low Stock Notice","",enDatatype_boolean,enDisplayType_checkbox)
	maryInventoryFields(Counter(i)) = Array("LowFlag", "Qty to send Low Stock Notice at", "Qty to send Low Stock Notice at", "", enDatatype_number, enDisplayType_textbox)
	maryInventoryFields(Counter(i)) = Array("CanBackOrder","Allow Backorder","Allow Backorder","",enDatatype_boolean,enDisplayType_checkbox)
	maryInventoryFields(Counter(i)) = Array("DefaultQTY","Set default Qty In Stock At","Set default Qty In Stock At","",enDatatype_number,enDisplayType_textbox)
	'maryInventoryFields(Counter(i)) = Array("OnOrder","OnOrder","","",enDatatype_boolean,enDisplayType_checkbox)
	'maryInventoryFields(Counter(i)) = Array("InStock","InStock","","",enDatatype_number,enDisplayType_textbox)

	'this is for the actual inventory
	maryInventoryFields(Counter(i)) = Array("QtyInStock","Qty In Stock","Qty In Stock","",enDatatype_number,enDisplayType_textbox)
	maryInventoryFields(Counter(i)) = Array("QtyLowFlag","Qty to send Low Stock Notice at","Qty Low Flag","",enDatatype_number,enDisplayType_textbox)
	'maryInventoryFields(Counter(i)) = Array("OnOrder","OnOrder","","",enDatatype_boolean,enDisplayType_checkbox)

'Gift Wrap (AE only)
'The same array decoder for the products section applies here
	i = -1
	ReDim maryGiftWrap(1)
	maryGiftWrap(Counter(i)) = Array("IsActive","Gift Wrap This Product","Gift Wrap This Product","",enDatatype_boolean,enDisplayType_checkbox)
	maryGiftWrap(Counter(i)) = Array("Price","Gift Wrap Price","Gift Wrap Price","0",enDatatype_boolean,enDisplayType_textbox)
	'maryGiftWrap(Counter(i)) = Array("Message","Message","Gift Wrap Message","",enDatatype_boolean,enDisplayType_checkbox)

'Volume Pricing (AE only)
'You do not need to identify any columns here as the application will automatically find the colums starting with the below prefix
	cstrMTPImportPrefix = "MTP-"
	cstrMTPImportSeparator = "-"
	mblnDeleteExistingMTPs = True

'Import Options
	mbytCreateCat = cenCreateCat_Create	'cenCreateCat_Default, cenCreateCat_Create, cenCreateCat_CreateAndDelete
	mbytCreateMfg = cenCreateMfg_Create	'cenCreateMfg_Default, cenCreateMfg_Create, cenCreateMfg_CreateAndDelete
	mbytCreateVend = cenCreateVend_Create	'cenCreateVend_Default, cenCreateVend_Create, cenCreateVend_CreateAndDelete
	mlngDefaultImportType = enImportAll		'enImportAll, enImportNewOnly, enImportInvPriceOnly, enImportInformationOnly

	mlngDefaultCategoryID = 1				'Default uid to use from Categories table
	mlngDefaultManufacturerID = 1			'Default uid to use from Manufacturers table
	mlngDefaultVendorID = 1					'Default uid to use from Vendors table

'/
'/////////////////////////////////////////////////

'****************************************************************************************************************************************************************
%>
